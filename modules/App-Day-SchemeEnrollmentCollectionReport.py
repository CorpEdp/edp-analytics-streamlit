import streamlit as st
import pandas as pd
import numpy as np
import calendar
import re
import io
from io import BytesIO
from datetime import datetime
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings

warnings.filterwarnings("ignore")

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Scheme Enrollment & Collection Report",
    page_icon="💰",
    layout="wide"
)

# ============================================================
# CONSTANTS & MAPPINGS
# ============================================================

REQUIRED_COLUMNS = [
    "Id", "Scheme Participation Id", "Date", "Status", "Saved Amount",
    "Reward Amount", "Transaction Reference", "Installment number",
    "Metal Type", "Metal Rate", "Saved Metal Weight", "Rewards Metal Weight",
    "Benefit Metal Amount", "Benefit Metal Weight", "Benefit Metal Percentage",
    "Receipt ID", "Customer Name", "Customer Phone Number",
    "Passbook number", "Scheme Name"
]

DAILY_SCHEME_KEYS = ["e-gold", "egold", "e gold", "e-silver", "esilver", "e silver"]

SCHEME_NORMALISE_MAP = {
    "valatine's day": "Valentine's Day", "valentine's day": "Valentine's Day",
    "valentines day": "Valentine's Day", "valantine's day": "Valentine's Day",
    "akshaya thrithiyai": "Akshaya Tritiya", "akshaya trithiya": "Akshaya Tritiya",
    "akshaya thritiya": "Akshaya Tritiya", "akshaya tritiya": "Akshaya Tritiya",
    "e gold": "e-Gold", "egold": "e-Gold", "e-gold": "e-Gold",
    "e silver": "e-Silver", "esilver": "e-Silver", "e-silver": "e-Silver",
}

SESIONAL_ORDER = {
    "pongal": 10, "akshaya tritiya": 20, "diwali": 30, "christmas": 40, "valentine's day": 50,
}

BRANCH_MAPPING = {
    "bhima jewellery - madurai": "MDU", "head office": "MDU", "madurai branch": "MDU",
    "bhima jewellery - telecaller": "MDU-Telecalling", "bhima jewellery - marthandam": "MDM",
    "bhima jewellery - salem": "SLM", "app showroom location": "SLM", "in-transit- locations": "SLM",
    "in-transit locations": "SLM", "n/a": "Unassigned", "bhima jewellery - tirunelveli": "TVL",
    "bhima jewellery - tiruchirappalli": "TCY", "trichy branch": "TCY",
    "bhima jewellery -  rajapalayam": "RJPM", "bhima jewellery - rajapalayam": "RJPM",
    "rajapalayam branch": "RJPM", "bhima jewellery - dindigul": "DGL", "dindigul branch": "DGL",
    "bhima jewellery - noida": "ND", "bhima jewellery - virudhunagar": "VNR",
    "bhima jewellery -anna nagar": "AN", "bhima jewellery - anna nagar": "AN",
    "bhima jewellery - thanjavur": "TJR", "bhima jewellery -thanjavur": "TJR",
}

BRANCH_ORDER = [
    "MDU", "MDM", "SLM", "TVL", "TCY", "RJPM", "DGL", "ND", "VNR", "TJR", "AN", "MDU-Telecalling", "Unassigned",
]

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def classify_scheme_type(scheme_name):
    if pd.isna(scheme_name): return "sessional"
    text = str(scheme_name).lower().strip()
    return "daily" if any(key in text for key in DAILY_SCHEME_KEYS) else "sessional"

def normalise_scheme(name):
    if pd.isna(name): return name
    key = str(name).strip().lower()
    if key == "" or key == "nan": return str(name).strip()
    return SCHEME_NORMALISE_MAP.get(key, str(name).strip())

def sessional_sort_key(name):
    return SESIONAL_ORDER.get(str(name).strip().lower(), 999)

def normalize_branch(branch_name):
    if pd.isna(branch_name): return "Skip"
    key = " ".join(str(branch_name).strip().lower().split())
    if not key: return "Skip"
    if key in BRANCH_MAPPING: return BRANCH_MAPPING[key]
    for k, v in BRANCH_MAPPING.items():
        if " ".join(k.split()) in key: return v
    return "Skip"

def clean_phone_scalar(value):
    if pd.isna(value): return ""
    s = str(value).strip()
    if s.endswith(".0"): s = s[:-2]
    digits = re.sub(r'\D', '', s)
    if len(digits) == 12 and digits.startswith("91"): digits = digits[2:]
    elif len(digits) == 11 and digits.startswith("0"): digits = digits[1:]
    return digits[-10:] if len(digits) >= 10 else digits

def _parse_date_flexible(date_val):
    if pd.isna(date_val) or date_val == '': return pd.NaT
    try:
        if isinstance(date_val, (datetime, pd.Timestamp)): return pd.Timestamp(date_val).normalize()
        date_str = str(date_val).strip()
        formats = ['%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%y', '%d/%m/%y', '%b %d, %Y', '%d %b %Y']
        for fmt in formats:
            try: return pd.Timestamp(datetime.strptime(date_str, fmt)).normalize()
            except: continue
        return pd.to_datetime(date_str).normalize()
    except: return pd.NaT

def clean_amount(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", "").str.replace("₹", "").str.replace("Rs.", "").str.replace("Rs", "").str.strip(),
        errors="coerce"
    ).fillna(0)

def clean_columns(df):
    df.columns = df.columns.astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    return df

def round_df(df, exclude_columns=None):
    res = df.copy()
    ex = exclude_columns or []
    for col in res.columns:
        if col not in ex and pd.api.types.is_numeric_dtype(res[col]):
            res[col] = res[col].apply(lambda x: round(x) if pd.notna(x) else 0)
    return res

def get_scheme_list(df):
    return sorted(df["Scheme"].dropna().unique().tolist())

def get_all_transaction_schemes(df):
    return sorted(df["Scheme"].dropna().apply(normalise_scheme).unique().tolist())

def _scheme_priority(scheme):
    if not scheme: return 100
    s = str(scheme).lower().strip()
    if "e-gold" in s or "egold" in s: return 0
    if "e-silver" in s or "esilver" in s: return 1
    return 50

# ============================================================
# EXCEL FORMATTING ENGINE
# ============================================================

def apply_cell_style(cell, fill=None, font=None, alignment=None, number_format=None, border=None):
    if fill: cell.fill = fill
    if font: cell.font = font
    if alignment: cell.alignment = alignment
    if number_format: cell.number_format = number_format
    if border: cell.border = border

def write_section_header(worksheet, row, title, max_column, color="4472C4"):
    if max_column > 1:
        worksheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_column)
    cell = worksheet.cell(row=row, column=1)
    cell.value = title
    cell.font = Font(size=14, bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    for c in range(1, max_column + 1):
        cc = worksheet.cell(row=row, column=c)
        cc.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cc.border = Border(left=Side(style="thin", color=color), right=Side(style="thin", color=color),
                            top=Side(style="thin", color=color), bottom=Side(style="thin", color=color))
    worksheet.row_dimensions[row].height = 25

def format_dataframe_section(worksheet, dataframe, start_row, total_identifier=None):
    if dataframe.empty: return
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                         top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
    total_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    for col_idx, column_name in enumerate(dataframe.columns, start=1):
        cell = worksheet.cell(row=start_row, column=col_idx)
        cell.value = column_name
        apply_cell_style(cell, fill=header_fill, font=Font(bold=True), 
                         alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=thin_border)
    
    for row_idx in range(start_row + 1, start_row + 1 + len(dataframe)):
        is_total = False
        if total_identifier and dataframe.iloc[row_idx - start_row - 1, 0]:
            if total_identifier in str(dataframe.iloc[row_idx - start_row - 1, 0]): is_total = True
        for col_idx in range(1, len(dataframe.columns) + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            apply_cell_style(cell, fill=total_fill if is_total else None, 
                             font=Font(bold=True) if is_total else None,
                             alignment=Alignment(horizontal="center", vertical="center"),
                             number_format='#,##0' if col_idx > 1 else None, border=thin_border)

def write_projection_section(worksheet, start_row, enrollment_df, collection_df, month_label, thin_border):
    if enrollment_df.empty and collection_df.empty: return
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    sub_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    title = f"📈 MONTHLY PROJECTION - {month_label}" if month_label else "📈 MONTHLY PROJECTION"
    write_section_header(worksheet, start_row, title, 7, "ED7D31")
    
    sub_row, header_row, data_start_row = start_row + 1, start_row + 2, start_row + 3
    worksheet.merge_cells(start_row=sub_row, start_column=1, end_row=sub_row, end_column=3)
    c1 = worksheet.cell(row=sub_row, column=1)
    c1.value = "Projection of this month's first enrollments"
    c1.font, c1.fill, c1.alignment = Font(bold=True, italic=True), sub_fill, Alignment(horizontal="center")
    
    worksheet.merge_cells(start_row=sub_row, start_column=5, end_row=sub_row, end_column=7)
    c2 = worksheet.cell(row=sub_row, column=5)
    c2.value = "Projection of this month's overall collection"
    c2.font, c2.fill, c2.alignment = Font(bold=True, italic=True), sub_fill, Alignment(horizontal="center")

    for offset, label in enumerate(["Scheme", "Projected Count", "Projected Amount"]):
        for base_col in (1, 5):
            cell = worksheet.cell(row=header_row, column=base_col + offset)
            cell.value, cell.font, cell.fill, cell.border = label, Font(bold=True), header_fill, thin_border
            cell.alignment = Alignment(horizontal="center")

    row_count = max(len(enrollment_df), len(collection_df))
    for r in range(row_count):
        row_idx = data_start_row + r
        if r < len(enrollment_df):
            er = enrollment_df.iloc[r]
            worksheet.cell(row=row_idx, column=1).value = er["Scheme"]
            worksheet.cell(row=row_idx, column=2).value = er["Projected Count"]
            worksheet.cell(row=row_idx, column=3).value = er["Projected Amount"]
        if r < len(collection_df):
            cr = collection_df.iloc[r]
            worksheet.cell(row=row_idx, column=5).value = cr["Scheme"]
            worksheet.cell(row=row_idx, column=6).value = cr["Projected Count"]
            worksheet.cell(row=row_idx, column=7).value = cr["Projected Amount"]
        for col_idx in (1, 2, 3, 5, 6, 7):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.alignment, cell.border = Alignment(horizontal="center"), thin_border
            if col_idx in (2, 3, 6, 7): cell.number_format = '#,##0'

def write_avg_ticket_section(worksheet, start_row, avg_ticket_data, thin_border):
    if avg_ticket_data.empty: return
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    neg_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    neg_font, pos_font = Font(color="C00000"), Font(color="006100")
    total_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    write_section_header(worksheet, start_row, "📊 AVERAGE TICKET SIZE COMPARISON", len(avg_ticket_data.columns), "70AD47")
    header_row, data_start_row = start_row + 1, start_row + 2
    for col_idx, col_name in enumerate(avg_ticket_data.columns, start=1):
        cell = worksheet.cell(row=header_row, column=col_idx)
        cell.value, cell.font, cell.fill, cell.border = col_name, Font(bold=True), header_fill, thin_border
        cell.alignment = Alignment(horizontal="center")

    for r, row in enumerate(avg_ticket_data.values):
        row_idx = data_start_row + r
        is_total = "Total" in str(row[0])
        for col_idx, value in enumerate(row, start=1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.value, cell.alignment, cell.border = value, Alignment(horizontal="center"), thin_border
            col_name = avg_ticket_data.columns[col_idx - 1]
            if col_name in ["First Enrollment Amount", "Collection Amount", "First Enrollment Avg Ticket", "Collection Avg Ticket", "Difference"]:
                cell.number_format = '#,##0'
            elif col_name == "% Change":
                cell.number_format = '0.00%'
            
            if col_name in ["Difference", "% Change"] and isinstance(value, (int, float)):
                if value < 0: cell.fill, cell.font = neg_fill, neg_font
                elif value > 0: cell.font = pos_font
            if is_total:
                if not (col_name in ["Difference", "% Change"] and value < 0): cell.fill = total_fill
                cell.font = Font(bold=True)

def write_sheet(workbook, sheet_data):
    ws = workbook.create_sheet(sheet_data["sheet_name"])
    summary_df = sheet_data["summary"]
    enrollment_df = sheet_data["enrollment"]
    collection_df = sheet_data["collection"]
    unique_df = sheet_data["unique"]
    date_range, report_title = sheet_data["date_range"], sheet_data["report_title"]
    avg_ticket_data = sheet_data.get("avg_ticket_data", pd.DataFrame())

    title_color, section_color, scheme_color = "203764", "4472C4", "5B9BD5"
    total_color, light_blue, gold_color = "C00000", "D9E1F2", "FFC000"
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                         top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    max_cols = max(len(summary_df.columns) if not summary_df.empty else 1, 7)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
    tc = ws.cell(row=1, column=1)
    tc.value, tc.font, tc.fill, tc.alignment = report_title, Font(size=16, bold=True, color="FFFFFF"), PatternFill(start_color=title_color, end_color=title_color, fill_type="solid"), Alignment(horizontal="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_cols)
    dc = ws.cell(row=2, column=1)
    dc.value, dc.font, dc.alignment = f"Report Period: {date_range[0]} to {date_range[1]}", Font(size=12, bold=True), Alignment(horizontal="center")

    # SUMMARY
    summary_section_row = 4
    write_section_header(ws, summary_section_row, "📊 SCHEME SUMMARY", len(summary_df.columns) if not summary_df.empty else max_cols, section_color)
    format_dataframe_section(ws, summary_df, summary_section_row + 1, total_identifier="Grand Total")

    # AVG TICKET
    avg_ticket_row = summary_section_row + len(summary_df) + 3
    if not avg_ticket_data.empty:
        write_avg_ticket_section(ws, avg_ticket_row, avg_ticket_data, thin_border)
        last_row = avg_ticket_row + len(avg_ticket_data) + 2
    else:
        last_row = summary_section_row + len(summary_df)

    # PROJECTION
    proj_enroll = sheet_data.get("enrollment_projection", pd.DataFrame())
    proj_coll = sheet_data.get("collection_projection", pd.DataFrame())
    proj_label = sheet_data.get("projection_month_label", "")
    if not proj_enroll.empty or not proj_coll.empty:
        write_projection_section(ws, last_row + 3, proj_enroll, proj_coll, proj_label, thin_border)
        last_row = last_row + 3 + max(len(proj_enroll), len(proj_coll)) + 3

    # DAY-WISE ENROLLMENT
    enroll_row = last_row + 3
    write_section_header(ws, enroll_row, "📅 DAY-WISE NEW ENROLLMENT", len(enrollment_df.columns) if not enrollment_df.empty else max_cols, section_color)
    format_dataframe_section(ws, enrollment_df, enroll_row + 1, total_identifier="Days")
    # Specialized formatting for Enrollment (Metal Rates)
    for c_idx, col in enumerate(enrollment_df.columns, start=1):
        cell = ws.cell(row=enroll_row + 1, column=c_idx)
        if "Total" in col: fill = total_color
        elif "Rate" in col: fill = scheme_color
        else: fill = light_blue
        cell.fill = PatternFill(start_color=fill, end_color=fill, fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFF" if fill in [total_color, scheme_color] else "000000")

    # DAY-WISE COLLECTION
    coll_row = enroll_row + len(enrollment_df) + 3
    write_section_header(ws, coll_row, "💰 DAY-WISE COLLECTION", len(collection_df.columns) if not collection_df.empty else max_cols, section_color)
    format_dataframe_section(ws, collection_df, coll_row + 1, total_identifier="Days")

    # UNIQUE
    uniq_row = coll_row + len(collection_df) + 3
    write_section_header(ws, uniq_row, "👥 UNIQUE ENROLLMENT", len(unique_df.columns) if not unique_df.empty else max_cols, section_color)
    format_dataframe_section(ws, unique_df, uniq_row + 1, total_identifier="Days")

    for col in ws.columns:
        max_len = max([len(str(cell.value)) for cell in col if cell.value], default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, 12), 25)

    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

# ============================================================
# REFERRAL REPORTING LOGIC
# ============================================================

def _fmt_cell(val):
    if pd.isna(val) or val == 0 or val == "": return "-"
    return val

def read_uploaded_file(uploaded):
    if uploaded is None: return None
    try:
        if uploaded.name.lower().endswith(".csv"): return pd.read_csv(uploaded, keep_default_na=False, na_values=[])
        return pd.read_excel(uploaded, keep_default_na=False, na_values=[])
    except Exception as e:
        st.error(f"Error reading {uploaded.name}: {e}")
        return None

def _transform_branch_pretty(branch):
    if pd.isna(branch) or str(branch).strip() == "": return "Bhima Jewellery - Customer"
    b = str(branch).upper().strip()
    if b == "HEAD OFFICE": return "Bhima Jewellery - Madurai"
    if b == "IN-TRANSIT- LOCATIONS": return "Bhima Jewellery - Salem"
    if b == "APP SHOWROOM LOCATION": return "Bhima Jewellery - Tirunelveli"
    if "BHIMA JEWELLERY -" in b:
        return f"Bhima Jewellery - {str(branch).split('Bhima Jewellery -', 1)[-1].strip().title()}"
    return f"Bhima Jewellery - {str(branch).title()}"

def build_referral_report(transactions_df, employees_df, referrals_df, start_date=None, end_date=None, master_scheme_list=None):
    empty = pd.DataFrame()
    diagnostics, debug_info = {}, {}
    if employees_df is None or referrals_df is None: return empty, [], empty, diagnostics, empty, debug_info, []

    emp = clean_columns(employees_df.copy())
    ref = clean_columns(referrals_df.copy())
    txn = transactions_df.copy()

    # Find columns
    emp_ref_col = find_column(emp, ["Referral Code", "Referal Code", "Ref Code"])
    emp_br_col = find_column(emp, ["Branch", "Branch Name", "Showroom"])
    emp_type_col = find_column(emp, ["Employee Type", "Emp Type", "Category"])
    emp_code_col = find_column(emp, ["Employee Code", "Emp Code"])

    ref_code_col = find_column(ref, ["Referral Code", "Referal Code", "Ref Code"])
    ref_phone_col = find_column(ref, ["Referee Phone", "Referee Mobile", "Referee Contact", "Mobile", "Phone"])
    ref_name_col = find_column(ref, ["Referee Name", "Referee", "Customer Name"])
    ref_amt_col = find_column(ref, ["Enrollment Amount", "Enroll Amount", "Amount"])
    ref_stat_col = find_column(ref, ["Status"])
    ref_r_name_col = find_column(ref, ["Referrer Name", "Referer Name"])
    ref_r_phone_col = find_column(ref, ["Referrer Phone", "Referer Phone"])
    ref_reg_col = find_column(ref, ["Registered Date", "Registration Date", "Register Date"])
    ref_join_col = find_column(ref, ["Joined Date", "Join Date", "Joining Date"])
    ref_sch_col = find_column(ref, ["Scheme", "Scheme Name"])

    if any(c is None for c in [emp_ref_col, emp_br_col, ref_code_col, ref_phone_col, ref_join_col]):
        st.error("❌ Missing critical columns in Employee or Referral files.")
        return empty, [], empty, diagnostics, empty, debug_info, []

    # Mappings
    ref_code_to_branch = dict(zip(emp[emp_ref_col].astype(str).str.strip(), emp[emp_br_col]))
    ref_code_to_empcode = dict(zip(emp[emp_ref_col].astype(str).str.strip(), emp[emp_code_col].astype(str))) if emp_code_col else {}
    ref_code_to_type = dict(zip(emp[emp_ref_col].astype(str).str.strip(), emp[emp_type_col].astype(str))) if emp_type_col else {}

    # Construct Dataframe
    fr = pd.DataFrame()
    fr['Customer Name'] = ref[ref_name_col] if ref_name_col else ""
    fr['Customer Phone'] = ref[ref_phone_col].apply(clean_phone_scalar)
    fr['Customer Enrollment Amount'] = ref[ref_amt_col].apply(lambda x: float(re.sub(r'[^\d.]', '', str(x))) if pd.notna(x) else np.nan) if ref_amt_col else np.nan
    fr['Status'] = ref[ref_stat_col].astype(str) if ref_stat_col else ""
    fr['Employee Name'] = ref[ref_r_name_col] if ref_r_name_col else ""
    fr['Referral Code'] = ref[ref_code_col].astype(str).str.strip()
    fr['Employee Code'] = fr['Referral Code'].map(ref_code_to_empcode).fillna('')
    fr['Branch'] = fr['Referral Code'].map(ref_code_to_branch).apply(_transform_branch_pretty)
    fr['Category'] = fr['Referral Code'].map(ref_code_to_type).fillna('Customer')
    fr['Joined Date'] = ref[ref_join_col].apply(_parse_date_flexible)
    fr['Registered Date'] = ref[ref_reg_col].apply(_parse_date_flexible) if ref_reg_col else pd.NaT
    fr['Referral Scheme'] = ref[ref_sch_col].astype(str) if ref_sch_col else ""

    # Scope filtering
    if start_date and end_date:
        start_ts, end_ts = pd.Timestamp(start_date), pd.Timestamp(end_date)
        enrolled_mask = fr['Joined Date'].between(start_ts, end_ts)
        not_enrolled_mask = (fr['Joined Date'].isna()) & (fr['Registered Date'].between(start_ts, end_ts))
        fr_filtered = fr[enrolled_mask | not_enrolled_mask].copy()
    else:
        fr_filtered = fr.copy()

    # Transaction Matching
    txn_f1 = txn[txn['Installment number'] == 1].copy() if 'Installment number' in txn.columns else txn.copy()
    txn_f1['_phone'] = txn_f1['Customer Phone Number'].apply(clean_phone_scalar)
    txn_f1['_date'] = txn_f1['Date'].apply(_parse_date_flexible)
    txn_f1['_amount'] = txn_f1['Saved Amount'].apply(lambda x: float(re.sub(r'[^\d.]', '', str(x))) if pd.notna(x) else np.nan)
    
    txn_lookup = defaultdict(list)
    for idx, row in txn_f1.iterrows():
        if row['_phone'] and pd.notna(row['_date']):
            txn_lookup[(row['_phone'], row['_date'].date())].append(row)

    matched_rows = []
    for idx, row in fr_filtered.iterrows():
        phone, joined_date, amt = row['Customer Phone'], row['Joined Date'], row['Customer Enrollment Amount']
        
        if pd.isna(joined_date):
            matched_rows.append({**row, "Match Reason": "Not Enrolled", "Final Category": "Not Enrolled", "Passbook": "", "Paid Amount": None})
            continue
            
        match = None
        if phone:
            candidates = txn_lookup.get((phone, joined_date.date()), [])
            exact = [c for c in candidates if abs(c['_amount'] - amt) <= 0.01]
            if exact: match = exact[0]

        if match:
            scheme = normalise_scheme(match.get('Scheme Name', 'Unknown'))
            cat = "e-Silver" if "silver" in str(scheme).lower() else "e-Gold" if "gold" in str(scheme).lower() else scheme
            matched_rows.append({**row, "Match Reason": "Enrolled (Strict)", "Final Category": cat, "Passbook": match.get('Passbook number', ''), "Paid Amount": match['_amount']})
        else:
            scheme = normalise_scheme(row['Referral Scheme'])
            cat = "e-Silver" if "silver" in str(scheme).lower() else "e-Gold" if "gold" in str(scheme).lower() else scheme
            matched_rows.append({**row, "Match Reason": "Enrolled (No Txn)", "Final Category": cat, "Passbook": "", "Paid Amount": None})

    detail_df = pd.DataFrame(matched_rows)
    
    # Duplicates and Collisions
    detail_df['Is Duplicate'] = False
    if not detail_df.empty and 'Passbook' in detail_df.columns:
        pb_counts = detail_df[detail_df['Passbook'] != ""].groupby('Passbook').size()
        for pb in pb_counts[pb_counts > 1].index:
            mask = detail_df['Passbook'] == pb
            detail_df.loc[mask, 'Is Duplicate'] = False
            detail_df.loc[mask.index[1:], 'Is Duplicate'] = True

    # Final Pivoting
    ordered_branches = build_branch_order(detail_df['Branch'].unique().tolist()) if not detail_df.empty else []
    
    pivot = detail_df.pivot_table(index="Branch", columns="Final Category", values="Customer Phone", aggfunc="count", fill_value=0)
    pivot = pivot.reindex(ordered_branches).fillna(0)
    pivot["Grand Total"] = pivot.sum(axis=1)
    pivot.loc["Total"] = pivot.sum()
    pivot = pivot.reset_index().rename(columns={"index": "Showroom"})

    # Daily Breakdowns
    daily_list = []
    all_days = sorted(detail_df['Joined Date'].dropna().unique())
    for d in all_days:
        day_df = detail_df[detail_df['Joined Date'] == d]
        p = day_df.pivot_table(index="Branch", columns="Final Category", values="Customer Phone", aggfunc="count", fill_value=0)
        p = p.reindex(ordered_branches).fillna(0)
        p["Grand Total"] = p.sum(axis=1)
        p.loc["Total"] = p.sum()
        daily_list.append((d.strftime('%d-%m-%Y'), p.reset_index().rename(columns={"index": "Showroom"})))

    diagnostics = {
        "total_referrals": len(detail_df),
        "matched_referrals": len(detail_df[detail_df['Final Category'] != 'Not Enrolled']),
        "not_enrolled": len(detail_df[detail_df['Final Category'] == 'Not Enrolled']),
        "duplicates": detail_df['Is Duplicate'].sum() if 'Is Duplicate' in detail_df.columns else 0,
        "date_filter": f"{start_date} to {end_date}"
    }

    return pivot, daily_list, detail_df, diagnostics, detail_df[detail_df['Is Duplicate'] == True], {}, []

def find_column(df, candidates):
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lookup: return lookup[cand.lower()]
    return None

def build_branch_order(all_branches):
    if not all_branches: return []
    fixed_groups = [["madurai"], ["marthandam"], ["salem"], ["tirunelveli"], ["trichy", "tiruchirappalli"], ["rajapalayam"], ["dindigul"], ["noida"], ["virudhunagar"], ["thanjavur"]]
    
    def get_rank(b):
        b_l = str(b).lower()
        for i, group in enumerate(fixed_groups):
            if any(g in b_l for g in group): return i
        return 99
    
    return sorted(all_branches, key=get_rank)

# ============================================================
# CORE REPORTING ENGINE (VECTORIZED)
# ============================================================

@st.cache_data(show_spinner=False)
def generate_report_data(df, schemes, start_date=None, end_date=None):
    if not schemes or df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # First Installment Logic
    first_df = df[df["Installment number"] == 1].sort_values(["Passbook number", "Date", "Id"]).drop_duplicates(subset=["Passbook number"], keep="first").copy()
    coll_df = df[df["Installment number"] != 1].copy()

    # Denominator for Averages
    if start_date and end_date:
        report_day_count = (pd.Timestamp(end_date) - pd.Timestamp(start_date)).days + 1
    else:
        report_day_count = df["Date"].nunique()
    report_day_count = max(report_day_count, 1)

    # 1. SUMMARY TABLE (Vectorized)
    s_first_counts = first_df[first_df["Scheme"].isin(schemes)].groupby("Scheme").size()
    s_first_amts = first_df[first_df["Scheme"].isin(schemes)].groupby("Scheme")["Saved Amount"].sum()
    s_coll_counts = coll_df[coll_df["Scheme"].isin(schemes)].groupby("Scheme").size()
    s_coll_amts = coll_df[coll_df["Scheme"].isin(schemes)].groupby("Scheme")["Saved Amount"].sum()

    summary_rows = []
    for s in schemes:
        fc = s_first_counts.get(s, 0)
        fa = s_first_amts.get(s, 0)
        cc = s_coll_counts.get(s, 0)
        ca = s_coll_amts.get(s, 0)
        summary_rows.append({
            "Scheme": s, "First Enrollment Count": fc, "First Enrollment Amount": fa,
            "Avg Count/Day": fc / report_day_count, "Avg Amount/Day": fa / report_day_count,
            "Collection Count": cc, "Collection Amount": ca,
            "Avg Collection Count/Day": cc / report_day_count, "Avg Collection Amount/Day": ca / report_day_count,
        })
    
    summary_df = pd.DataFrame(summary_rows)
    if not summary_df.empty:
        total_row = summary_df.sum(numeric_only=True)
        total_row["Scheme"] = "Grand Total"
        summary_df = pd.concat([summary_df, pd.DataFrame([total_row])], ignore_index=True)

    # 2. DAY-WISE ENROLLMENT (Vectorized)
    # Counts and Sums
    e_counts = first_df[first_df["Scheme"].isin(schemes)].pivot_table(index="Date", columns="Scheme", values="Id", aggfunc="count", fill_value=0)
    e_amts = first_df[first_df["Scheme"].isin(schemes)].pivot_table(index="Date", columns="Scheme", values="Saved Amount", aggfunc="sum", fill_value=0)
    
    # Metal Rates
    rates_df = df.pivot_table(index="Date", columns="Scheme", values="Metal Rate", aggfunc="first")
    
    enroll_list = []
    dates = sorted(df["Date"].unique())
    for d in dates:
        row = {"Date": d.strftime("%d-%m-%Y")}
        t_cnt, t_amt = 0, 0
        for s in schemes:
            rate_col = get_rate_column_name(s)
            if rate_col: row[rate_col] = rates_df.loc[d, s] if s in rates_df.columns else 0
            cnt = e_counts.loc[d, s] if s in e_counts.columns else 0
            amt = e_amts.loc[d, s] if s in e_amts.columns else 0
            row[f"{s} Count"], row[f"{s} Amount"] = cnt, amt
            row[f"{s} Avg Ticket"] = amt / cnt if cnt > 0 else 0
            t_cnt += cnt; t_amt += amt
        row["Total Count"], row["Total Amount"] = t_cnt, t_amt
        row["Total Avg Ticket"] = t_amt / t_cnt if t_cnt > 0 else 0
        enroll_list.append(row)
    
    enrollment_df = pd.DataFrame(enroll_list)
    if not enrollment_df.empty:
        total_row = {"Date": f"{len(enrollment_df)} Days"}
        for s in schemes:
            total_row[f"{s} Count"] = enrollment_df[f"{s} Count"].sum()
            total_row[f"{s} Amount"] = enrollment_df[f"{s} Amount"].sum()
            total_row[f"{s} Avg Ticket"] = total_row[f"{s} Amount"] / total_row[f"{s} Count"] if total_row[f"{s} Count"] > 0 else 0
        total_row["Total Count"] = enrollment_df["Total Count"].sum()
        total_row["Total Amount"] = enrollment_df["Total Amount"].sum()
        total_row["Total Avg Ticket"] = total_row["Total Amount"] / total_row["Total Count"] if total_row["Total Count"] > 0 else 0
        enrollment_df = pd.concat([enrollment_df, pd.DataFrame([total_row])], ignore_index=True)

    # 3. DAY-WISE COLLECTION (Vectorized)
    c_counts = coll_df[coll_df["Scheme"].isin(schemes)].pivot_table(index="Date", columns="Scheme", values="Id", aggfunc="count", fill_value=0)
    c_amts = coll_df[coll_df["Scheme"].isin(schemes)].pivot_table(index="Date", columns="Scheme", values="Saved Amount", aggfunc="sum", fill_value=0)
    
    coll_list = []
    for d in dates:
        row = {"Date": d.strftime("%d-%m-%Y")}
        t_cnt, t_amt = 0, 0
        for s in schemes:
            cnt = c_counts.loc[d, s] if s in c_counts.columns else 0
            amt = c_amts.loc[d, s] if s in c_amts.columns else 0
            row[f"{s} Count"], row[f"{s} Amount"] = cnt, amt
            row[f"{s} Avg Ticket"] = amt / cnt if cnt > 0 else 0
            t_cnt += cnt; t_amt += amt
        row["Total Count"], row["Total Amount"] = t_cnt, t_amt
        row["Total Avg Ticket"] = t_amt / t_cnt if t_cnt > 0 else 0
        coll_list.append(row)
    
    collection_df = pd.DataFrame(coll_list)
    if not collection_df.empty:
        t_row = {"Date": f"{len(collection_df)} Days"}
        for s in schemes:
            t_row[f"{s} Count"] = collection_df[f"{s} Count"].sum()
            t_row[f"{s} Amount"] = collection_df[f"{s} Amount"].sum()
            t_row[f"{s} Avg Ticket"] = t_row[f"{s} Amount"] / t_row[f"{s} Count"] if t_row[f"{s} Count"] > 0 else 0
        t_row["Total Count"] = collection_df["Total Count"].sum()
        t_row["Total Amount"] = collection_df["Total Amount"].sum()
        t_row["Total Avg Ticket"] = t_row["Total Amount"] / t_row["Total Count"] if t_row["Total Count"] > 0 else 0
        collection_df = pd.concat([collection_df, pd.DataFrame([t_row])], ignore_index=True)

    # 4. UNIQUE (Vectorized)
    unique_list = []
    for d in dates:
        row = {"Date": d.strftime("%d-%m-%Y")}
        gt = 0
        for s in schemes:
            uc = first_df[(first_df["Date"] == d) & (first_df["Scheme"] == s) & (first_df["Customer Phone Number"].notna())]["Customer Phone Number"].nunique()
            row[s] = uc
            gt += uc
        row["Grand Total"] = gt
        unique_list.append(row)
    unique_df = pd.DataFrame(unique_list)
    if not unique_df.empty:
        t_row = {"Date": f"{len(unique_df)} Days"}
        for s in schemes: t_row[s] = unique_df[s].sum()
        t_row["Grand Total"] = unique_df["Grand Total"].sum()
        unique_df = pd.concat([unique_df, pd.DataFrame([t_row])], ignore_index=True)

    return round_df(summary_df), round_df(enrollment_df), round_df(collection_df), round_df(unique_df), add_avg_ticket_size_comparison(summary_df)

def get_rate_column_name(scheme):
    sl = str(scheme).lower()
    if any(k in sl for k in DAILY_SCHEME_KEYS):
        return "Silver Rate" if "silver" in sl else "Gold Rate"
    return f"{scheme} Rate"

def add_avg_ticket_size_comparison(summary_df):
    if summary_df.empty: return pd.DataFrame()
    res = []
    for _, row in summary_df[summary_df["Scheme"] != "Grand Total"].iterrows():
        f_avg = row["First Enrollment Amount"] / row["First Enrollment Count"] if row["First Enrollment Count"] > 0 else 0
        c_avg = row["Collection Amount"] / row["Collection Count"] if row["Collection Count"] > 0 else 0
        diff = c_avg - f_avg
        pct = diff / f_avg if f_avg > 0 else 0
        res.append({"Scheme": row["Scheme"], "First Enrollment Avg Ticket": f_avg, "Collection Avg Ticket": c_avg, "Difference": diff, "% Change": pct})
    
    t_f_avg = summary_df.loc[summary_df["Scheme"] == "Grand Total", "First Enrollment Amount"].values[0] / summary_df.loc[summary_df["Scheme"] == "Grand Total", "First Enrollment Count"].values[0] if not summary_df.empty else 0
    t_c_avg = summary_df.loc[summary_df["Scheme"] == "Grand Total", "Collection Amount"].values[0] / summary_df.loc[summary_df["Scheme"] == "Grand Total", "Collection Count"].values[0] if not summary_df.empty else 0
    res.append({"Scheme": "Grand Total", "First Enrollment Avg Ticket": t_f_avg, "Collection Avg Ticket": t_c_avg, "Difference": t_c_avg - t_f_avg, "% Change": (t_c_avg - t_f_avg)/t_f_avg if t_f_avg > 0 else 0})
    return pd.DataFrame(res)

def generate_monthly_projection(full_df, schemes, reference_date=None):
    if full_df.empty or not schemes: return pd.DataFrame(), pd.DataFrame(), ""
    latest_date = pd.Timestamp(reference_date) if reference_date else full_df["Date"].max()
    month_start = latest_date.replace(day=1)
    days_in_month = calendar.monthrange(latest_date.year, latest_date.month)[1]
    days_elapsed = max(latest_date.day, 1)
    
    month_label = f"{latest_date.strftime('%B %Y')} ({days_elapsed}/{days_in_month} days elapsed)"
    
    # Data for current month
    curr_month_df = full_df[(full_df["Date"] >= month_start) & (full_df["Date"] <= latest_date)]
    first_month_df = full_df[(full_df["Installment number"] == 1) & (full_df["Date"] >= month_start) & (full_df["Date"] <= latest_date)]
    
    e_proj, c_proj = [], []
    for s in schemes:
        e_data = first_month_df[first_month_df["Scheme"] == s]
        c_data = curr_month_df[(curr_month_df["Scheme"] == s) & (curr_month_df["Installment number"] != 1)]
        
        e_proj.append({"Scheme": s, "Projected Count": round((len(e_data)/days_elapsed)*days_in_month), "Projected Amount": round((e_data["Saved Amount"].sum()/days_elapsed)*days_in_month)})
        c_proj.append({"Scheme": s, "Projected Count": round((len(c_data)/days_elapsed)*days_in_month), "Projected Amount": round((c_data["Saved Amount"].sum()/days_elapsed)*days_in_month)})
        
    return pd.DataFrame(e_proj), pd.DataFrame(c_proj), month_label

def create_formatted_excel(daily_data, sessional_data, date_range, report_title, **kwargs):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        workbook = writer.book
        if "Sheet" in workbook.sheetnames: del workbook["Sheet"]
        
        if daily_data["schemes"]: write_sheet(workbook, daily_data)
        if sessional_data["schemes"]: write_sheet(workbook, sessional_data)
        
        if kwargs.get("referral_summary") is not None and not kwargs["referral_summary"].empty:
            # Write referral report using the custom function
            write_referral_sheet(workbook, kwargs["referral_summary"], kwargs.get("referral_daily", []), 
                                kwargs.get("referral_detail", pd.DataFrame()), kwargs.get("referral_duplicates", pd.DataFrame()), date_range)
            
    output.seek(0)
    return output.getvalue()

# ============================================================
# REFERRAL EXCEL WRITER
# ============================================================

def write_referral_sheet(workbook, summary_df, daily_list, detail_df, duplicates_df, date_range, sheet_name="Referral Report"):
    ws = workbook.create_sheet(sheet_name)
    thin = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    
    write_section_header(ws, 4, f"🏢 BRANCH-WISE REFERRAL SUMMARY ({date_range[0]} to {date_range[1]})", len(summary_df.columns), "7030A0")
    format_dataframe_section(ws, summary_df, 5, total_identifier="Total")
    
    # This is a simplified version of your original writer to ensure stability
    last_row = 5 + len(summary_df)
    if daily_list:
        for date_label, day_df in daily_list:
            last_row += 3
            write_section_header(ws, last_row, f"📅 REFERRAL — {date_label}", len(day_df.columns), "7030A0")
            format_dataframe_section(ws, day_df, last_row + 1, total_identifier="Total")
            last_row = last_row + 1 + len(day_df)

    if not detail_df.empty:
        last_row += 3
        write_section_header(ws, last_row, "📋 REFERRAL MATCH DETAILS", len(detail_df.columns), "7030A0")
        format_dataframe_section(ws, detail_df, last_row + 1)

# ============================================================
# MAIN STREAMLIT APP
# ============================================================

st.title("💰 Scheme Enrollment & Collection Report")
st.caption("High-performance BI tool for Scheme and Referral Reconciliation")

# SIDEBAR
st.sidebar.header("📂 File Uploads")
u_main = st.sidebar.file_uploader("1️⃣ Main Transaction File", type=["xlsx", "csv"])
u_emp = st.sidebar.file_uploader("2️⃣ Employee File", type=["xlsx", "csv"])
u_ref = st.sidebar.file_uploader("3️⃣ Referral File", type=["xlsx", "csv"])

if u_main is None:
    st.info("Please upload the Main Transaction File to begin.")
    st.stop()

# LOAD DATA
try:
    df = pd.read_csv(u_main) if u_main.name.endswith(".csv") else pd.read_excel(u_main)
except Exception as e:
    st.error(f"Error: {e}"); st.stop()

df = clean_columns(df)
if not all(col in df.columns for col in REQUIRED_COLUMNS):
    st.error(f"Missing columns: {[c for c in REQUIRED_COLUMNS if c not in df.columns]}"); st.stop()

df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True).dt.normalize()
df["Saved Amount"] = clean_amount(df["Saved Amount"])
df["Scheme"] = df["Scheme Name"].apply(normalise_scheme)
df = df[df["Date"].notna() & df["Scheme"].notna()].copy()

# FILTERS
st.sidebar.header("🔎 Filters")
min_d, max_d = df["Date"].min().date(), df["Date"].max().date()
date_range = st.sidebar.date_input("Date Range", [min_d, max_d], min_value=min_d, max_value=max_d)

if len(date_range) == 2:
    start_d, end_d = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
else:
    start_d = end_d = pd.Timestamp(date_range[0])

all_schemes = get_scheme_list(df)
sel_schemes = st.sidebar.multiselect("Select Schemes", all_schemes, default=all_schemes)

filtered_df = df[(df["Date"] >= start_d) & (df["Date"] <= end_d) & (df["Scheme"].isin(sel_schemes))].copy()

if filtered_df.empty:
    st.warning("No data found for selected filters."); st.stop()

# REPORT GENERATION
daily_schemes = [s for s in sel_schemes if classify_scheme_type(s) == "daily"]
sess_schemes = [s for s in sel_schemes if classify_scheme_type(s) == "sessional"]

with st.spinner("Generating reports..."):
    d_sum, d_enr, d_col, d_uni, d_avg = generate_report_data(filtered_df[filtered_df["Scheme"].isin(daily_schemes)], daily_schemes, start_d, end_d)
    s_sum, s_enr, s_col, s_uni, s_avg = generate_report_data(filtered_df[filtered_df["Scheme"].isin(sess_schemes)], sess_schemes, start_d, end_d)
    
    d_e_proj, d_c_proj, d_proj_lbl = generate_monthly_projection(df, daily_schemes, end_d)
    s_e_proj, s_c_proj, s_proj_lbl = generate_monthly_projection(df, sess_schemes, end_d)

# REFERRALS
ref_sum, ref_daily, ref_det, ref_diag, ref_dup, _, _ = build_referral_report(filtered_df, read_uploaded_file(u_emp), read_uploaded_file(u_ref), start_d, end_d)

# DASHBOARD PREVIEW
st.header("📋 Report Preview")
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Enrollments", int(d_sum["First Enrollment Count"].sum() + s_sum["First Enrollment Count"].sum()) if not d_sum.empty else 0)
col2.metric("Total Collection", f"₹{int(d_sum["Collection Amount"].sum() + s_sum["Collection Amount"].sum()):,}" if not d_sum.empty else 0)
col3.metric("Referral Conv %", f"{(ref_diag.get('matched_referrals',0)/ref_diag.get('total_referrals',1)*100):.1f}%" if ref_diag else "N/A")
col4.metric("Unique Customers", int(d_uni["Grand Total"].sum() + s_uni["Grand Total"].sum()) if not d_uni.empty else 0)

# SCHEME PREVIEWS
if not d_sum.empty:
    with st.expander("📅 Daily Schemes Preview", expanded=True):
        st.dataframe(d_sum, use_container_width=True, hide_index=True)
        st.dataframe(d_avg, use_container_width=True, hide_index=True)
        st.write(f"**Projection: {d_proj_lbl}**")
        c1, c2 = st.columns(2)
        c1.dataframe(d_e_proj); c2.dataframe(d_c_proj)

if not s_sum.empty:
    with st.expander("🎯 Sessional Schemes Preview"):
        st.dataframe(s_sum, use_container_width=True, hide_index=True)
        st.dataframe(s_avg, use_container_width=True, hide_index=True)

if not ref_sum.empty:
    with st.expander("🎁 Referral Preview"):
        st.dataframe(ref_sum, use_container_width=True, hide_index=True)
        st.dataframe(ref_det, use_container_width=True, hide_index=True)

# DOWNLOAD
st.header("📥 Export")
daily_data = {"summary": d_sum, "enrollment": d_enr, "collection": d_col, "unique": d_uni, "schemes": daily_schemes, "date_range": (start_d.strftime("%d-%m-%Y"), end_d.strftime("%d-%m-%Y")), "report_title": "Daily Schemes Report", "sheet_name": "Daily", "enrollment_projection": d_e_proj, "collection_projection": d_c_proj, "projection_month_label": d_proj_lbl, "avg_ticket_data": d_avg}
sess_data = {"summary": s_sum, "enrollment": s_enr, "collection": s_col, "unique": s_uni, "schemes": sess_schemes, "date_range": (start_d.strftime("%d-%m-%Y"), end_d.strftime("%d-%m-%Y")), "report_title": "Sessional Schemes Report", "sheet_name": "Sessional", "enrollment_projection": s_e_proj, "collection_projection": s_c_proj, "projection_month_label": s_proj_lbl, "avg_ticket_data": s_avg}

excel_file = create_formatted_excel(daily_data, sess_data, (start_d.strftime("%d-%m-%Y"), end_d.strftime("%d-%m-%Y")), u_main.name, referral_summary=ref_sum, referral_daily=ref_daily, referral_detail=ref_det, referral_duplicates=ref_dup)

st.download_button("⬇️ Download Full Excel Report", data=excel_file, file_name="Scheme_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)