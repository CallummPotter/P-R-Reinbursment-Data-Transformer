import io
import re
from typing import Tuple

import pandas as pd
import streamlit as st
import xlsxwriter 
from PIL import Image


st.set_page_config(page_title="PAYG Reimbursement Calculator", layout="wide")
st.image(
    "LOGO.png",
    width=250
)
st.title("PAYG Reimbursement Calculator")

# ========================
# Helpers
# =========================

def normalize_string(value) -> str:
    if pd.isna(value):
        return 
    return str(value).strip()


def sanitize_filename_part(text: str) -> str:
    text = normalize_string(text)
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    return text or "Unknown Organisation"


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [re.sub(r"[^a-z0-9]", "", str(c).strip().lower()) for c in df.columns]
    return df


def parse_datetime_series(series: pd.Series) -> pd.Series:
    s = series.copy()

    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if parsed.notna().all():
        return parsed

    still_bad = parsed.isna() & s.notna() & (s.astype(str).str.strip() != "")
    if still_bad.any():
        parsed2 = pd.to_datetime(s[still_bad], errors="coerce", dayfirst=False)
        parsed.loc[still_bad] = parsed2

    return parsed


def get_quarter(dt: pd.Timestamp) -> int:
    return ((dt.month - 1) // 3) + 1


def determine_period(start_series: pd.Series, end_series: pd.Series) -> Tuple[int, int]:
    valid_start = pd.to_datetime(start_series, errors="coerce").dropna()
    if not valid_start.empty:
        first_date = valid_start.min()
        return int(first_date.year), int(get_quarter(first_date))

    valid_end = pd.to_datetime(end_series, errors="coerce").dropna()
    if not valid_end.empty:
        first_date = valid_end.min()
        return int(first_date.year), int(get_quarter(first_date))

    raise ValueError("Could not determine quarter/year from Start or End times.")


def choose_organisation(devices_df: pd.DataFrame, revenue_df: pd.DataFrame) -> str:
    if "deviceorganisation" in devices_df.columns:
        vals = devices_df["deviceorganisation"].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        if not vals.empty:
            return vals.iloc[0]

    if "organisation" in revenue_df.columns:
        vals = revenue_df["organisation"].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        if not vals.empty:
            return vals.iloc[0]

    return "Unknown Organisation"


def build_group_key(row) -> str:
    station_name = normalize_string(row.get("stationname", ""))
    device_name = normalize_string(row.get("device", ""))
    mpan_revenue = normalize_string(row.get("mpan", ""))
    mpan_device = normalize_string(row.get("mpan_dev", ""))

    if station_name:
        return station_name
    if device_name:
        return device_name
    if mpan_revenue:
        return f"MPAN_{mpan_revenue}"
    if mpan_device:
        return f"MPAN_{mpan_device}"
    return "UNKNOWN"



def deduplicate_sessions_with_audit(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    working = df.copy().reset_index(drop=True)
    working["_original_order"] = range(len(working))
    working["transactionid_num"] = pd.to_numeric(working["transactionid"], errors="coerce")

    can_dedupe = working["start_parsed"].notna() & working["end_parsed"].notna()

    dedupe_candidates = working[can_dedupe].copy()
    non_dedupe_rows = working[~can_dedupe].copy()

    dedupe_candidates = dedupe_candidates.sort_values(
        by=["transactionid_num", "_original_order"],
        ascending=[True, True],
        na_position="last",
    )

    group_cols = ["ResolvedCharger", "connector", "start_parsed", "end_parsed"]

    kept_parts = []
    audit_rows = []

    for _, group in dedupe_candidates.groupby(group_cols, dropna=False, sort=False):
        group = group.sort_values(
            by=["transactionid_num", "_original_order"],
            ascending=[True, True],
            na_position="last",
        )

        kept_row = group.iloc[[0]].copy()
        kept_row["AuditStatus"] = "KEPT"
        kept_parts.append(kept_row)
        audit_rows.append(kept_row)

        if len(group) > 1:
            dropped = group.iloc[1:].copy()
            dropped["AuditStatus"] = "DROPPED_DUPLICATE"
            audit_rows.append(dropped)

    if not non_dedupe_rows.empty:
        non_dedupe_rows = non_dedupe_rows.copy()
        non_dedupe_rows["AuditStatus"] = "KEPT_UNPARSED_TIME"
        kept_parts.append(non_dedupe_rows)
        audit_rows.append(non_dedupe_rows)

    deduped_df = pd.concat(kept_parts, ignore_index=True) if kept_parts else working.iloc[0:0].copy()
    deduped_df = deduped_df.sort_values("_original_order").reset_index(drop=True)

    audit_df = pd.concat(audit_rows, ignore_index=True) if audit_rows else working.iloc[0:0].copy()

    return deduped_df, audit_df

logo_image = Image.open("LOGO.png")

def build_output_excel_bytes(
    agg_df: pd.DataFrame,
    organisation: str,
    year: int,
    quarter: int,
    quote_reference: str = "INSERT QUOTE REFERENCE HERE",
    logo_path: str | None = None,
) -> io.BytesIO:
    output = io.BytesIO()

    period_label = f"{year} Q{quarter}"
    output_sheet_name = "PAYG Revenue"

    total_energy = float(agg_df["TotalEnergy"].sum())
    total_collected_fee = float(agg_df["CollectedFee"].sum())
    total_transaction_fee = float(agg_df["TransactionFee"].sum())
    total_repayment = float(agg_df["Repayment"].sum())

    # VAT-style split based on your formulas
    collected_net = (total_collected_fee / 120) * 100 if total_collected_fee else 0.0
    collected_vat = total_collected_fee - collected_net
    collected_inc_vat = total_collected_fee

    pr_fee_net = (total_transaction_fee / 120) * 100 if total_transaction_fee else 0.0
    pr_fee_vat = total_transaction_fee - pr_fee_net
    pr_fee_inc_vat = total_transaction_fee

    repayment_net = round(collected_net - pr_fee_net, 2)
    repayment_vat = round(collected_vat - pr_fee_vat, 2)
    repayment_inc_vat = round(collected_inc_vat - pr_fee_inc_vat, 2)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet(output_sheet_name)
        writer.sheets[output_sheet_name] = worksheet

        # =========================================================
        # Column widths
        # =========================================================
        worksheet.set_column("A:A", 4)
        worksheet.set_column("B:B", 52)
        worksheet.set_column("C:C", 28)
        worksheet.set_column("D:D", 28)
        worksheet.set_column("E:E", 28)
        worksheet.set_column("F:F", 28)

        # =========================================================
        # Colors
        # =========================================================
        blue = "#0080FF"
        orange = "#F26B21"
        green = "#70AD47"
        red = "#FF0000"
        grey_text = "#666666"
        black = "#000000"

        # =========================================================
        # Border helpers
        # =========================================================
        # xlsxwriter border indexes:
        # 1 = thin, 2 = medium
        thin = 1
        bold = 2

        def fmt(props=None):
            base = {
                "font_name": "Calibri",
                "font_size": 12,
                "valign": "vcenter",
            }
            if props:
                base.update(props)
            return workbook.add_format(base)

        # =========================================================
        # Core formats
        # =========================================================
        fmt_grey_note = fmt({
            "font_color": grey_text,
            "align": "center",
            "text_wrap": True,
        })

        fmt_grey_note_bold = fmt({
            "font_color": grey_text,
            "bold": True,
            "align": "center",
        })

        fmt_black_box = fmt({
            "border": bold,
        })

        fmt_title_top = fmt({
            "bg_color": blue,
            "bold": True,
            "underline": True,
            "align": "center",
            "top": bold,
            "left": bold,
            "right": bold,
            "bottom": 0,
        })

        fmt_title_bottom = fmt({
            "bg_color": blue,
            "bold": True,
            "align": "center",
            "top": 0,
            "left": bold,
            "right": bold,
            "bottom": bold,
        })

        fmt_section_header = fmt({
            "bg_color": blue,
            "bold": True,
            "align": "center",
            "top": bold,
            "left": bold,
            "right": bold,
            "bottom": thin,
        })

        fmt_stats_label = fmt({
            "left": bold,
            "top": thin,
            "bottom": thin,
            "right": thin,
            "align": "left",
        })

        fmt_stats_value = fmt({
            "left": thin,
            "top": thin,
            "bottom": thin,
            "right": 0,
            "align": "right",
            "num_format": "0.00",
        })

        fmt_stats_unit = fmt({
            "left": 0,
            "top": thin,
            "bottom": thin,
            "right": thin,
            "align": "left",
        })

        fmt_stats_blank = fmt({
            "border": thin,
        })

        fmt_fin_hdr_left = fmt({
            "bg_color": blue,
            "bold": True,
            "align": "left",
            "left": bold,
            "top": bold,
            "bottom": 0,
            "right": thin,
        })

        fmt_fin_hdr_mid = fmt({
            "bg_color": blue,
            "bold": True,
            "align": "center",
            "top": bold,
            "left": thin,
            "right": thin,
            "bottom": 0,
        })

        fmt_fin_hdr_right = fmt({
            "bg_color": blue,
            "bold": True,
            "align": "center",
            "top": bold,
            "left": thin,
            "right": bold,
            "bottom": 0,
        })

        fmt_label_left = fmt({
            "left": bold,
            "top": thin,
            "bottom": thin,
            "right": thin,
            "align": "left",
        })

        fmt_label_left_bold = fmt({
            "left": bold,
            "top": thin,
            "bottom": thin,
            "right": thin,
            "align": "left",
            "bold": True,
        })

        fmt_label_left_bottombold = fmt({
            "left": bold,
            "top": thin,
            "bottom": bold,
            "right": thin,
            "align": "left",
            "bold": True,
        })

        fmt_currency_thin = fmt({
            "border": thin,
            "align": "right",
            "num_format": '£#,##0.00',
        })

        fmt_currency_rightbold = fmt({
            "left": thin,
            "top": thin,
            "bottom": thin,
            "right": bold,
            "align": "right",
            "num_format": '£#,##0.00',
        })

        fmt_currency_orange_rightbold = fmt({
            "left": thin,
            "top": thin,
            "bottom": thin,
            "right": bold,
            "align": "right",
            "font_color": orange,
            "num_format": '£#,##0.00',
        })

        fmt_currency_green_bottomthin = fmt({
            "left": thin,
            "top": thin,
            "bottom": 0,
            "right": thin,
            "align": "right",
            "font_color": green,
            "bold": True,
            "num_format": '£#,##0.00',
        })

        fmt_currency_red_bottomthin_rightbold = fmt({
            "left": thin,
            "top": thin,
            "bottom": 0,
            "right": bold,
            "align": "right",
            "font_color": red,
            "bold": True,
            "num_format": '£#,##0.00',
        })

        fmt_currency_bottombold = fmt({
            "left": thin,
            "top": thin,
            "bottom": bold,
            "right": thin,
            "align": "right",
            "num_format": '£#,##0.00',
        })

        fmt_currency_green_bottombold = fmt({
            "left": thin,
            "top": thin,
            "bottom": bold,
            "right": thin,
            "align": "right",
            "font_color": green,
            "bold": True,
            "num_format": '£#,##0.00',
        })

        fmt_currency_red_bottombold_rightbold = fmt({
            "left": thin,
            "top": thin,
            "bottom": bold,
            "right": bold,
            "align": "right",
            "font_color": red,
            "bold": True,
            "num_format": '£#,##0.00',
        })

        fmt_detail_hdr_left = fmt({
            "bg_color": blue,
            "bold": True,
            "left": bold,
            "top": bold,
            "right": 0,
            "bottom": 0,
            "align": "left",
        })

        fmt_detail_hdr_mid = fmt({
            "bg_color": blue,
            "bold": True,
            "top": bold,
            "left": thin,
            "right": thin,
            "bottom": 0,
            "align": "center",
            "text_wrap": True,
        })

        fmt_detail_hdr_right = fmt({
            "bg_color": blue,
            "bold": True,
            "top": bold,
            "left": thin,
            "right": bold,
            "bottom": 0,
            "align": "center",
            "text_wrap": True,
        })

        fmt_detail_text = fmt({
            "border": thin,
            "align": "left",
        })

        fmt_detail_num = fmt({
            "border": thin,
            "align": "right",
            "num_format": "0.00",
        })

        fmt_detail_currency = fmt({
            "border": thin,
            "align": "right",
            "num_format": '£#,##0.00',
        })

        fmt_total_left = fmt({
            "bg_color": blue,
            "bold": True,
            "left": bold,
            "top": thin,
            "right": thin,
            "bottom": bold,
            "align": "left",
        })

        fmt_total_mid_num = fmt({
            "bg_color": blue,
            "bold": True,
            "left": thin,
            "top": thin,
            "right": thin,
            "bottom": bold,
            "align": "right",
            "num_format": "0.00",
        })

        fmt_total_mid_currency = fmt({
            "bg_color": blue,
            "bold": True,
            "left": thin,
            "top": thin,
            "right": thin,
            "bottom": bold,
            "align": "right",
            "num_format": '£#,##0.00',
        })

        fmt_total_right_currency = fmt({
            "bg_color": blue,
            "bold": True,
            "left": thin,
            "top": thin,
            "right": bold,
            "bottom": bold,
            "align": "right",
            "num_format": '£#,##0.00',
        })

        fmt_key_header = fmt({
            "bg_color": blue,
            "bold": True,
            "left": bold,
            "right": bold,
            "top": bold,
            "bottom": thin,
            "align": "left",
        })

        fmt_key_orange = fmt({
            "font_color": orange,
            "left": bold,
            "right": bold,
            "top": thin,
            "bottom": thin,
            "align": "left",
        })

        fmt_key_red = fmt({
            "font_color": red,
            "left": bold,
            "right": bold,
            "top": thin,
            "bottom": thin,
            "align": "left",
        })

        fmt_key_green = fmt({
            "font_color": green,
            "left": bold,
            "right": bold,
            "top": thin,
            "bottom": bold,
            "align": "left",
        })

        # =========================================================
        # Row heights
        # =========================================================
        worksheet.set_row(1, 38)   # row 2
        worksheet.set_row(3, 28)   # row 4
        worksheet.set_row(4, 28)   # row 5
        worksheet.set_row(6, 24)   # row 7
        worksheet.set_row(15, 24)  # row 16

        # Hide gridlines
        worksheet.hide_gridlines(2)

        # =========================================================
        # ROW 2
        # =========================================================
        if logo_path:
            img = Image.open(logo_path)

            img_width, img_height = img.size

            # Target size (adjust these to fit your layout nicely)
            target_width = 180
            target_height = int((target_width / img_width) * img_height)

            worksheet.insert_image(
                "B2",
                "logo.png",
                {
                    "image_data": logo_path,
                    "x_scale": target_width / img_width,
                    "y_scale": target_height / img_height,
                    "x_offset": 5,
                    "y_offset": 5,
                    "object_position": 1,
                },
            )
        worksheet.write("E2", "Please quote this reference\non your invoice:", fmt_grey_note)
        worksheet.write("F2", quote_reference, fmt_grey_note_bold)

        # =========================================================
        # ROW 3
        # =========================================================
        worksheet.merge_range("B3:F3", "", fmt_black_box)

        # =========================================================
        # ROW 4
        # =========================================================
        worksheet.merge_range("B4:F4", "Customer Pay As You Go (PAYG) Finances", fmt_title_top)

        # =========================================================
        # ROW 5
        # =========================================================
        worksheet.merge_range("B5:F5", period_label, fmt_title_bottom)

        # =========================================================
        # ROW 7
        # =========================================================
        worksheet.merge_range("B7:F7", "Statistics:", fmt_section_header)

        # =========================================================
        # ROW 8
        # =========================================================
        worksheet.merge_range(
            "B8:C8",
            f"Total Energy Consumed by EV Charging in {period_label}",
            fmt_stats_label,
        )
        worksheet.write_number("D8", round(total_energy, 2), fmt_stats_value)
        worksheet.write("E8", "kWh", fmt_stats_unit)
        worksheet.write_blank("F8", None, fmt_stats_blank)

        # =========================================================
        # ROW 10
        # =========================================================
        worksheet.merge_range("B10:C10", "Financial summary", fmt_fin_hdr_left)
        worksheet.write("D10", "Totals (Net):", fmt_fin_hdr_mid)
        worksheet.write("E10", "VAT Incurred:", fmt_fin_hdr_mid)
        worksheet.write("F10", "Totals (inc. VAT):", fmt_fin_hdr_right)

        # =========================================================
        # ROW 11
        # =========================================================
        worksheet.merge_range("B11:C11", f"Sum of Collected Fees for {period_label}", fmt_label_left)
        worksheet.write_number("D11", round(collected_net, 2), fmt_currency_thin)
        worksheet.write_number("E11", round(collected_vat, 2), fmt_currency_thin)
        worksheet.write_number("F11", round(collected_inc_vat, 2), fmt_currency_rightbold)

        # =========================================================
        # ROW 12
        # =========================================================
        worksheet.merge_range("B12:C12", f"Sum of P&R transaction Fee for {period_label}", fmt_label_left)
        worksheet.write_number("D12", round(pr_fee_net, 2), fmt_currency_thin)
        worksheet.write_number("E12", round(pr_fee_vat, 2), fmt_currency_thin)
        worksheet.write_number("F12", round(pr_fee_inc_vat, 2), fmt_currency_orange_rightbold)

        # =========================================================
        # ROW 13
        # =========================================================
        worksheet.merge_range("B13:C13", "Sum of Repayment Due to Customer", fmt_label_left_bottombold)
        worksheet.write_number("D13", round(repayment_net, 2), fmt_currency_bottombold)
        worksheet.write_number("E13", round(repayment_vat, 2), fmt_currency_green_bottombold)
        worksheet.write_number("F13", round(repayment_inc_vat, 2), fmt_currency_red_bottombold_rightbold)

        # =========================================================
        # ROW 16
        # =========================================================
        worksheet.write("B16", "Row Labels", fmt_detail_hdr_left)
        worksheet.write("C16", "Sum of Total_energy (kWh)", fmt_detail_hdr_mid)
        worksheet.write("D16", "Sum of Collected_fee", fmt_detail_hdr_mid)
        worksheet.write("E16", "Sum of PR_Transaction_fee", fmt_detail_hdr_mid)
        worksheet.write("F16", "Sum of Repayment", fmt_detail_hdr_right)

        # =========================================================
        # ROW 17 onward - detail rows
        # =========================================================
        start_excel_row = 17
        for idx, row in enumerate(agg_df.itertuples(index=False), start=start_excel_row):
            worksheet.write(f"B{idx}", row.ResolvedCharger, fmt_detail_text)
            worksheet.write_number(f"C{idx}", float(row.TotalEnergy), fmt_detail_num)
            worksheet.write_number(f"D{idx}", float(row.CollectedFee), fmt_detail_currency)
            worksheet.write_number(f"E{idx}", float(row.TransactionFee), fmt_detail_currency)
            worksheet.write_number(f"F{idx}", float(row.Repayment), fmt_detail_currency)

        # =========================================================
        # Grand total row
        # =========================================================
        grand_total_excel_row = start_excel_row + len(agg_df)
        worksheet.write(f"B{grand_total_excel_row}", "Grand Total", fmt_total_left)
        worksheet.write_number(f"C{grand_total_excel_row}", round(total_energy, 2), fmt_total_mid_num)
        worksheet.write_number(f"D{grand_total_excel_row}", round(total_collected_fee, 2), fmt_total_mid_currency)
        worksheet.write_number(f"E{grand_total_excel_row}", round(total_transaction_fee, 2), fmt_total_mid_currency)
        worksheet.write_number(f"F{grand_total_excel_row}", round(total_repayment, 2), fmt_total_right_currency)

        # =========================================================
        # Key box
        # =========================================================
        worksheet.write("B22", "Key", fmt_key_header)
        worksheet.write("B23", "Our Invoice will say amount due £0", fmt_key_orange)
        worksheet.write("B24", "Final figure to invoice us for", fmt_key_red)
        worksheet.write("B25", "Amount in VAT which customer should repay to HMRC", fmt_key_green)

    output.seek(0)
    return output


# =========================
# App
# =========================

revenue_file = st.file_uploader("Upload Revenue File (CSV or Excel)", type=["csv", "xlsx"])
devices_file = st.file_uploader("Upload Devices File (CSV or Excel)", type=["csv", "xlsx"])

if revenue_file and devices_file:
    try:
        # Determine file type and read accordingly
        if revenue_file.name.endswith(".csv"):
            revenue_raw = pd.read_csv(revenue_file)
        elif revenue_file.name.endswith(".xlsx"):
            revenue_raw = pd.read_excel(revenue_file)

        if devices_file.name.endswith(".csv"):
            devices_raw = pd.read_csv(devices_file)
        elif devices_file.name.endswith(".xlsx"):
            devices_raw = pd.read_excel(devices_file)

        revenue_df = standardize_columns(revenue_raw)
        devices_df = standardize_columns(devices_raw)

        required_revenue_cols = ["transactionid", "connector", "repayment", "start", "end"]
        missing = [c for c in required_revenue_cols if c not in revenue_df.columns]
        if missing:
            st.error("Revenue CSV is missing required columns: " + ", ".join(missing))
            st.stop()

        optional_numeric_cols = {
            "totalenergykwh": "TotalEnergy",
            "collectedfee": "CollectedFee",
            "prtransactionfee": "TransactionFee",
        }

        for source_col in optional_numeric_cols:
            if source_col not in revenue_df.columns:
                revenue_df[source_col] = 0

        if "device" in revenue_df.columns and "device" in devices_df.columns:
            merged = revenue_df.merge(
                devices_df,
                on="device",
                how="left",
                suffixes=("", "_dev")
            )
        else:
            merged = revenue_df.copy()
            if "mpan_dev" not in merged.columns:
                merged["mpan_dev"] = None

        for col in ["stationname", "device", "mpan", "mpan_dev", "organisation"]:
            if col not in merged.columns:
                merged[col] = None

        merged["start_raw"] = merged["start"]
        merged["end_raw"] = merged["end"]

        merged["start_parsed"] = parse_datetime_series(merged["start"])
        merged["end_parsed"] = parse_datetime_series(merged["end"])

        merged["repayment"] = pd.to_numeric(merged["repayment"], errors="coerce").fillna(0)
        merged["totalenergykwh"] = pd.to_numeric(merged["totalenergykwh"], errors="coerce").fillna(0)
        merged["collectedfee"] = pd.to_numeric(merged["collectedfee"], errors="coerce").fillna(0)
        merged["prtransactionfee"] = pd.to_numeric(merged["prtransactionfee"], errors="coerce").fillna(0)

        merged["ResolvedCharger"] = merged.apply(build_group_key, axis=1)

        deduped, audit_df = deduplicate_sessions_with_audit(merged)

        agg_df = (
            deduped.groupby("ResolvedCharger", dropna=False)
            .agg(
                SessionCount=("repayment", "count"),
                Repayment=("repayment", "sum"),
                TotalEnergy=("totalenergykwh", "sum"),
                CollectedFee=("collectedfee", "sum"),
                TransactionFee=("prtransactionfee", "sum"),
            )
            .reset_index()
            .sort_values("ResolvedCharger")
        )

        organisation = choose_organisation(devices_df, revenue_df)
        year, quarter = determine_period(deduped["start_parsed"], deduped["end_parsed"])
        output_filename = f"{year} Q{quarter} PAYG Revenue- {sanitize_filename_part(organisation)}.xlsx"

        # Convert to bytes (needed for Excel)
        img_bytes = io.BytesIO()
        logo_image.save(img_bytes, format="PNG")
        img_bytes.seek(0)

        # Build Excel with logo
        output = build_output_excel_bytes(
            agg_df=agg_df,
            organisation=organisation,
            year=year,
            quarter=quarter,
            quote_reference="INSERT QUOTE REFERENCE HERE",
            logo_path=img_bytes,
        )

        unparsed_start = int(((merged["start_raw"].notna()) & (merged["start_parsed"].isna())).sum())
        unparsed_end = int(((merged["end_raw"].notna()) & (merged["end_parsed"].isna())).sum())
        dropped_count = int((audit_df["AuditStatus"] == "DROPPED_DUPLICATE").sum())

        total_sessions = int(agg_df["SessionCount"].sum())
        total_repayment = float(agg_df["Repayment"].sum())
        total_energy = float(agg_df["TotalEnergy"].sum())
        total_collected_fee = float(agg_df["CollectedFee"].sum())
        total_transaction_fee = float(agg_df["TransactionFee"].sum())

        st.success("Processing complete.")

        if unparsed_start or unparsed_end:
            st.warning(
                f"Some times could not be parsed. Start parse failures: {unparsed_start}, "
                f"End parse failures: {unparsed_end}. These rows were kept, but not deduped by time."
            )

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Chargers", len(agg_df))
        col2.metric("Final Sessions", total_sessions)
        col3.metric("Dropped Duplicates", dropped_count)
        col4.metric("Total Repayment", f"£{total_repayment:,.2f}")

        col5, col6, col7 = st.columns(3)
        col5.metric("Total Energy Consumed", f"{total_energy:,.2f} kWh")
        col6.metric("Total Collected Fee", f"£{total_collected_fee:,.2f}")
        col7.metric("Total Transaction Fee", f"£{total_transaction_fee:,.2f}")

        st.subheader("Summary")
        st.dataframe(agg_df, use_container_width=True)

        st.subheader("Dedupe Audit Preview")
        audit_display_cols = [
            c for c in [
                "AuditStatus",
                "ResolvedCharger",
                "transactionid",
                "device",
                "stationname",
                "start_raw",
                "end_raw",
                "connector",
                "repayment",
                "totalenergykwh",
                "collectedfee",
                "prtransactionfee",
                "mpan",
                "mpan_dev",
                "organisation",
            ]
            if c in audit_df.columns
        ]

        def audit_row_style(row):
            status = row.get("AuditStatus", "")
            if status == "DROPPED_DUPLICATE":
                return ["background-color: #FCE4D6; color: #C00000;"] * len(row)
            if status == "KEPT":
                return ["background-color: #E2F0D9;"] * len(row)
            if status == "KEPT_UNPARSED_TIME":
                return ["background-color: #FFF2CC;"] * len(row)
            return [""] * len(row)

        st.dataframe(
            audit_df[audit_display_cols].style.apply(audit_row_style, axis=1),
            use_container_width=True,
        )

        st.download_button(
            label="Download Excel Output",
            data=output.getvalue(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Processing failed: {e}")

else:
    st.info("Please upload both the Revenue file and the Devices file.")