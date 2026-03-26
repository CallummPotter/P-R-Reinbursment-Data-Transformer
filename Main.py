import io
import re
import zipfile
from collections import defaultdict
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
        return ""
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
    if "organisation" in revenue_df.columns:
        vals = revenue_df["organisation"].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        if not vals.empty:
            return vals.value_counts().idxmax()

    if "deviceorganisation" in devices_df.columns:
        vals = devices_df["deviceorganisation"].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        if not vals.empty:
            return vals.value_counts().idxmax()

    return "Unknown Organisation"


def build_quote_reference(organisation: str, quarter: int, year: int) -> str:
    org = normalize_string(organisation)
    org = re.sub(r'[\\/*?:"<>|]', "", org)
    org = re.sub(r"\s+", "_", org)
    return f"PAG_{org}_Q{quarter}_{year}"


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


def read_table_file(file_name: str, file_bytes: bytes) -> pd.DataFrame:
    lower_name = file_name.lower()

    if lower_name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes))
    if lower_name.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(file_bytes))

    raise ValueError("Unsupported file type. Only CSV and XLSX are allowed.")


def detect_file_kind(df: pd.DataFrame) -> str:
    cols = set(df.columns)

    revenue_required = {"transactionid", "connector", "repayment", "start", "end"}
    device_signals = {"deviceorganisation", "mpan", "stationname", "device"}

    if revenue_required.issubset(cols):
        return "revenue"

    if "deviceorganisation" in cols:
        return "device"

    if len(device_signals.intersection(cols)) >= 2 and "transactionid" not in cols:
        return "device"

    return "unknown"


def get_most_common_org_from_df(df: pd.DataFrame, file_kind: str) -> str:
    if file_kind == "revenue" and "organisation" in df.columns:
        vals = df["organisation"].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        if not vals.empty:
            return vals.value_counts().idxmax()

    if file_kind == "device" and "deviceorganisation" in df.columns:
        vals = df["deviceorganisation"].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        if not vals.empty:
            return vals.value_counts().idxmax()

    return "Unknown Organisation"


logo_image = Image.open("LOGO.png")


def build_output_excel_bytes(
    agg_df: pd.DataFrame,
    organisation: str,
    year: int,
    quarter: int,
    quote_reference: str = "INSERT QUOTE REFERENCE HERE",
    logo_path: io.BytesIO | None = None,
) -> io.BytesIO:
    output = io.BytesIO()

    period_label = f"{year} Q{quarter}"
    output_sheet_name = "PAYG Revenue"

    total_energy = float(agg_df["TotalEnergy"].sum())
    total_collected_fee = float(agg_df["CollectedFee"].sum())
    total_transaction_fee = float(agg_df["TransactionFee"].sum())
    total_repayment = float(agg_df["Repayment"].sum())

    collected_net = (total_collected_fee / 120) * 100 if total_collected_fee else 0.0
    collected_vat = total_collected_fee - collected_net
    collected_inc_vat = total_collected_fee

    pr_fee_net = (total_transaction_fee / 120) * 100 if total_transaction_fee else 0.0
    pr_fee_vat = total_transaction_fee - pr_fee_net
    pr_fee_inc_vat = total_transaction_fee

    repayment_net = round(collected_net - pr_fee_net, 2)
    repayment_vat = round(collected_vat - pr_fee_vat, 2)
    repayment_inc_vat = round(total_repayment, 2)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet(output_sheet_name)
        writer.sheets[output_sheet_name] = worksheet

        worksheet.set_column("A:A", 4)
        worksheet.set_column("B:B", 52)
        worksheet.set_column("C:C", 28)
        worksheet.set_column("D:D", 28)
        worksheet.set_column("E:E", 28)
        worksheet.set_column("F:F", 28)
        worksheet.set_column("I:L", 18)

        blue = "#0080FF"
        orange = "#F26B21"
        green = "#70AD47"
        red = "#FF0000"
        grey_text = "#666666"

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
            "align": "center",
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

        worksheet.set_row(1, 38)
        worksheet.set_row(3, 28)
        worksheet.set_row(4, 28)
        worksheet.set_row(6, 24)
        worksheet.set_row(15, 24)

        worksheet.hide_gridlines(2)

        if logo_path:
            logo_path.seek(0)
            img = Image.open(logo_path)
            img_width, img_height = img.size
            target_width = 180
            target_height = int((target_width / img_width) * img_height)

            logo_path.seek(0)
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

        worksheet.merge_range("B3:F3", "", fmt_black_box)
        worksheet.merge_range("B4:F4", "Customer Pay As You Go (PAYG) Finances", fmt_title_top)
        worksheet.merge_range("B5:F5", period_label, fmt_title_bottom)

        worksheet.merge_range("B7:F7", "Statistics:", fmt_section_header)

        worksheet.merge_range(
            "B8:C8",
            f"Total Energy Consumed by EV Charging in {period_label}",
            fmt_stats_label,
        )
        worksheet.write_number("D8", round(total_energy, 2), fmt_stats_value)
        worksheet.write("E8", "kWh", fmt_stats_unit)
        worksheet.write_blank("F8", None, fmt_stats_blank)

        worksheet.merge_range("B10:C10", "Financial summary", fmt_fin_hdr_left)
        worksheet.write("D10", "Totals (Net):", fmt_fin_hdr_mid)
        worksheet.write("E10", "VAT Incurred:", fmt_fin_hdr_mid)
        worksheet.write("F10", "Totals (inc. VAT):", fmt_fin_hdr_right)

        worksheet.merge_range("B11:C11", f"Sum of Collected Fees for {period_label}", fmt_label_left)
        worksheet.write_number("D11", round(collected_net, 2), fmt_currency_thin)
        worksheet.write_number("E11", round(collected_vat, 2), fmt_currency_thin)
        worksheet.write_number("F11", round(collected_inc_vat, 2), fmt_currency_rightbold)

        worksheet.merge_range("B12:C12", f"Sum of P&R transaction Fee for {period_label}", fmt_label_left)
        worksheet.write_number("D12", round(pr_fee_net, 2), fmt_currency_thin)
        worksheet.write_number("E12", round(pr_fee_vat, 2), fmt_currency_thin)
        worksheet.write_number("F12", round(pr_fee_inc_vat, 2), fmt_currency_orange_rightbold)

        worksheet.merge_range("B13:C13", "Sum of Repayment Due to Customer", fmt_label_left_bottombold)
        worksheet.write_number("D13", round(repayment_net, 2), fmt_currency_bottombold)
        worksheet.write_number("E13", round(repayment_vat, 2), fmt_currency_green_bottombold)
        worksheet.write_number("F13", round(repayment_inc_vat, 2), fmt_currency_red_bottombold_rightbold)

        worksheet.write("B16", "Row Labels", fmt_detail_hdr_left)
        worksheet.write("C16", "Sum of Total_energy (kWh)", fmt_detail_hdr_mid)
        worksheet.write("D16", "Sum of Collected_fee", fmt_detail_hdr_mid)
        worksheet.write("E16", "Sum of PR_Transaction_fee", fmt_detail_hdr_mid)
        worksheet.write("F16", "Sum of Repayment", fmt_detail_hdr_right)

        start_excel_row = 17
        for idx, row in enumerate(agg_df.itertuples(index=False), start=start_excel_row):
            worksheet.write(f"B{idx}", row.ResolvedCharger, fmt_detail_text)
            worksheet.write_number(f"C{idx}", float(row.TotalEnergy), fmt_detail_num)
            worksheet.write_number(f"D{idx}", float(row.CollectedFee), fmt_detail_currency)
            worksheet.write_number(f"E{idx}", float(row.TransactionFee), fmt_detail_currency)
            worksheet.write_number(f"F{idx}", float(row.Repayment), fmt_detail_currency)

        grand_total_excel_row = start_excel_row + len(agg_df)
        worksheet.write(f"B{grand_total_excel_row}", "Grand Total", fmt_total_left)
        worksheet.write_number(f"C{grand_total_excel_row}", round(total_energy, 2), fmt_total_mid_num)
        worksheet.write_number(f"D{grand_total_excel_row}", round(total_collected_fee, 2), fmt_total_mid_currency)
        worksheet.write_number(f"E{grand_total_excel_row}", round(total_transaction_fee, 2), fmt_total_mid_currency)
        worksheet.write_number(f"F{grand_total_excel_row}", round(total_repayment, 2), fmt_total_right_currency)

        worksheet.merge_range("I10:L10", "Key", fmt_key_header)
        worksheet.merge_range("I11:L11", "Our Invoice will say amount due £0", fmt_key_orange)
        worksheet.merge_range("I12:L12", "Final figure to invoice us for", fmt_key_red)
        worksheet.merge_range("I13:L13", "Amount in VAT which customer should repay to HMRC", fmt_key_green)

    output.seek(0)
    return output


def process_revenue_and_devices(
    revenue_raw: pd.DataFrame,
    devices_raw: pd.DataFrame,
    logo_image: Image.Image,
) -> dict:
    revenue_df = standardize_columns(revenue_raw)
    devices_df = standardize_columns(devices_raw)

    required_revenue_cols = ["transactionid", "connector", "repayment", "start", "end"]
    missing = [c for c in required_revenue_cols if c not in revenue_df.columns]
    if missing:
        raise ValueError("Revenue file is missing required columns: " + ", ".join(missing))

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
    quote_reference = build_quote_reference(organisation, quarter, year)
    output_filename = f"{year} Q{quarter} PAYG Revenue- {sanitize_filename_part(organisation)}.xlsx"

    img_bytes = io.BytesIO()
    logo_image.save(img_bytes, format="PNG")
    img_bytes.seek(0)

    output = build_output_excel_bytes(
        agg_df=agg_df,
        organisation=organisation,
        year=year,
        quarter=quarter,
        quote_reference=quote_reference,
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

    return {
        "organisation": organisation,
        "year": year,
        "quarter": quarter,
        "quote_reference": quote_reference,
        "output_filename": output_filename,
        "output_bytes": output.getvalue(),
        "agg_df": agg_df,
        "audit_df": audit_df,
        "unparsed_start": unparsed_start,
        "unparsed_end": unparsed_end,
        "dropped_count": dropped_count,
        "total_sessions": total_sessions,
        "total_repayment": total_repayment,
        "total_energy": total_energy,
        "total_collected_fee": total_collected_fee,
        "total_transaction_fee": total_transaction_fee,
    }


def process_bulk_zip(zip_file, logo_image: Image.Image):
    successes = []
    skips = []

    revenue_files_by_org = defaultdict(list)
    device_files_by_org = defaultdict(list)

    with zipfile.ZipFile(zip_file) as zf:
        members = [m for m in zf.infolist() if not m.is_dir()]

        if not members:
            raise ValueError("The ZIP file is empty.")

        supported_members = [
            m for m in members
            if m.filename.lower().endswith((".csv", ".xlsx"))
        ]

        if not supported_members:
            raise ValueError("No supported CSV or XLSX files were found in the ZIP.")

        for member in supported_members:
            try:
                file_bytes = zf.read(member.filename)
                raw_df = read_table_file(member.filename, file_bytes)
                standardized_df = standardize_columns(raw_df)
                kind = detect_file_kind(standardized_df)
                org = get_most_common_org_from_df(standardized_df, kind)

                if kind == "revenue":
                    revenue_files_by_org[org].append({
                        "name": member.filename,
                        "raw_df": raw_df,
                    })
                elif kind == "device":
                    device_files_by_org[org].append({
                        "name": member.filename,
                        "raw_df": raw_df,
                    })
                else:
                    skips.append({
                        "file": member.filename,
                        "organisation": org,
                        "reason": "Could not determine whether file is Revenue or Devices.",
                    })

            except Exception as e:
                skips.append({
                    "file": member.filename,
                    "organisation": "Unknown Organisation",
                    "reason": f"Failed to read file: {e}",
                })

    all_orgs = sorted(set(revenue_files_by_org.keys()) | set(device_files_by_org.keys()))

    zip_output = io.BytesIO()

    with zipfile.ZipFile(zip_output, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for org in all_orgs:
            revenue_candidates = revenue_files_by_org.get(org, [])
            device_candidates = device_files_by_org.get(org, [])

            if not revenue_candidates:
                for device_file in device_candidates:
                    skips.append({
                        "file": device_file["name"],
                        "organisation": org,
                        "reason": "No matching Revenue file found for this organisation.",
                    })
                continue

            if not device_candidates:
                for revenue_file in revenue_candidates:
                    skips.append({
                        "file": revenue_file["name"],
                        "organisation": org,
                        "reason": "No matching Devices file found for this organisation.",
                    })
                continue

            if len(revenue_candidates) > 1:
                file_list = ", ".join(f["name"] for f in revenue_candidates)
                for revenue_file in revenue_candidates:
                    skips.append({
                        "file": revenue_file["name"],
                        "organisation": org,
                        "reason": f"Multiple Revenue files found for this organisation: {file_list}",
                    })
                continue

            if len(device_candidates) > 1:
                file_list = ", ".join(f["name"] for f in device_candidates)
                for device_file in device_candidates:
                    skips.append({
                        "file": device_file["name"],
                        "organisation": org,
                        "reason": f"Multiple Devices files found for this organisation: {file_list}",
                    })
                continue

            revenue_file = revenue_candidates[0]
            device_file = device_candidates[0]

            try:
                result = process_revenue_and_devices(
                    revenue_raw=revenue_file["raw_df"],
                    devices_raw=device_file["raw_df"],
                    logo_image=logo_image,
                )

                out_zip.writestr(result["output_filename"], result["output_bytes"])

                successes.append({
                    "organisation": result["organisation"],
                    "output_filename": result["output_filename"],
                    "revenue_file": revenue_file["name"],
                    "device_file": device_file["name"],
                    "year": result["year"],
                    "quarter": result["quarter"],
                    "total_sessions": result["total_sessions"],
                    "total_repayment": result["total_repayment"],
                })

            except Exception as e:
                skips.append({
                    "file": f"{revenue_file['name']} + {device_file['name']}",
                    "organisation": org,
                    "reason": f"Matched pair failed to process: {e}",
                })

    zip_output.seek(0)
    return zip_output, successes, skips


# =========================
# App
# =========================

bulk_mode = st.toggle("Bulk upload mode (ZIP input)", value=False)

if not bulk_mode:
    revenue_file = st.file_uploader("Upload Revenue File (CSV or Excel)", type=["csv", "xlsx"])
    devices_file = st.file_uploader("Upload Devices File (CSV or Excel)", type=["csv", "xlsx"])

    if revenue_file and devices_file:
        try:
            revenue_raw = read_table_file(revenue_file.name, revenue_file.getvalue())
            devices_raw = read_table_file(devices_file.name, devices_file.getvalue())

            result = process_revenue_and_devices(
                revenue_raw=revenue_raw,
                devices_raw=devices_raw,
                logo_image=logo_image,
            )

            st.success("Processing complete.")

            if result["unparsed_start"] or result["unparsed_end"]:
                st.warning(
                    f"Some times could not be parsed. Start parse failures: {result['unparsed_start']}, "
                    f"End parse failures: {result['unparsed_end']}. These rows were kept, but not deduped by time."
                )

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Chargers", len(result["agg_df"]))
            col2.metric("Final Sessions", result["total_sessions"])
            col3.metric("Dropped Duplicates", result["dropped_count"])
            col4.metric("Total Repayment", f"£{result['total_repayment']:,.2f}")

            col5, col6, col7 = st.columns(3)
            col5.metric("Total Energy Consumed", f"{result['total_energy']:,.2f} kWh")
            col6.metric("Total Collected Fee", f"£{result['total_collected_fee']:,.2f}")
            col7.metric("Total Transaction Fee", f"£{result['total_transaction_fee']:,.2f}")

            st.subheader("Summary")
            st.dataframe(result["agg_df"], use_container_width=True)

            st.subheader("Dedupe Audit Preview")
            audit_df = result["audit_df"]
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
                data=result["output_bytes"],
                file_name=result["output_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Processing failed: {e}")

    else:
        st.info("Please upload both the Revenue file and the Devices file.")

else:
    bulk_zip = st.file_uploader(
        "Upload ZIP containing Revenue and Devices CSV/XLSX files",
        type=["zip"]
    )

    if bulk_zip:
        try:
            zip_output, successes, skips = process_bulk_zip(bulk_zip, logo_image)

            if successes:
                st.success(f"Bulk processing complete. Generated {len(successes)} report(s).")
            else:
                st.warning("Bulk processing completed, but no reports could be generated.")

            if successes:
                st.subheader("Generated Reports")
                success_df = pd.DataFrame(successes)
                st.dataframe(success_df, use_container_width=True)

                st.download_button(
                    label="Download All Reports (ZIP)",
                    data=zip_output.getvalue(),
                    file_name="PAYG_Reports_Bulk_Output.zip",
                    mime="application/zip",
                )

            if skips:
                st.subheader("Skipped Files / Organisations")
                skip_df = pd.DataFrame(skips)
                st.dataframe(skip_df, use_container_width=True)

        except Exception as e:
            st.error(f"Bulk processing failed: {e}")

    else:
        st.info("Please upload a ZIP file containing Revenue and Devices CSV/XLSX files.")