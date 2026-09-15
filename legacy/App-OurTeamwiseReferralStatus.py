import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime, date, timedelta
import calendar
import json
import warnings
import math
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, Rule
from openpyxl.utils import get_column_letter
warnings.filterwarnings('ignore')

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Employee Referral Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown("""
<style>
    .main-header {
        font-size: 2.4rem;
        font-weight: 700;
        color: #1E88E5;
        margin-bottom: 0;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #808495;
        margin-top: -6px;
        margin-bottom: 18px;
    }
    .metric-card {
        background: var(--background-color, #f8f9fa);
        border: 1px solid rgba(128,128,128,0.15);
        border-radius: 12px;
        padding: 16px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
        text-align: center;
        transition: transform 0.15s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
    }
    .stButton > button {
        width: 100%;
        background-color: #1E88E5;
        color: white;
        font-weight: 600;
        border-radius: 8px;
    }
    .stButton > button:hover {
        background-color: #1565C0;
        color: white;
    }
    .warning-box {
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 10px 14px;
        margin: 10px 0;
        border-radius: 6px;
        color: #664d03;
    }
    .success-box {
        background-color: #d1e7dd;
        border-left: 4px solid #43A047;
        padding: 10px 14px;
        margin: 10px 0;
        border-radius: 6px;
        color: #0f5132;
    }
    .info-box {
        background-color: #cfe2ff;
        border-left: 4px solid #0d6efd;
        padding: 10px 14px;
        margin: 10px 0;
        border-radius: 6px;
        color: #084298;
    }
    .registered-header {
        background-color: #E3F2FD;
        padding: 12px 16px;
        border-radius: 8px;
        border-left: 4px solid #1E88E5;
    }
    .joined-header {
        background-color: #E8F5E9;
        padding: 12px 16px;
        border-radius: 8px;
        border-left: 4px solid #43A047;
    }
    .badge {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-left: 6px;
    }
    .badge-up { background:#d1e7dd; color:#0f5132; }
    .badge-down { background:#f8d7da; color:#842029; }
    .badge-flat { background:#e2e3e5; color:#41464b; }
    .badge-warning { background:#fff3cd; color:#664d03; }
    .badge-inbound { background:#d4edda; color:#155724; }
    .badge-outbound { background:#f8d7da; color:#721c24; }
    .filter-section {
        background: var(--background-color, #f8f9fa);
        padding: 12px 16px;
        border-radius: 8px;
        border: 1px solid rgba(128,128,128,0.1);
        margin-bottom: 16px;
    }
    .zero-value {
        color: #999;
        font-style: italic;
    }
    .insight-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 16px 20px;
        border-radius: 12px;
        margin: 8px 0;
    }
    .dark-mode-aware {
        color: var(--text-color);
    }
    .lookup-card {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white;
        padding: 18px 22px;
        border-radius: 12px;
        margin: 10px 0 18px 0;
    }
    .decimal-info {
        font-size: 0.8rem;
        color: #666;
        margin-top: 2px;
    }
    .report-header {
        background: linear-gradient(135deg, #1E88E5 0%, #1565C0 100%);
        color: white;
        padding: 12px 20px;
        border-radius: 8px;
        margin: 10px 0 15px 0;
        font-weight: 600;
    }
    .performance-high { background-color: #d4edda; color: #155724; }
    .performance-medium { background-color: #fff3cd; color: #856404; }
    .performance-low { background-color: #f8d7da; color: #721c24; }
    .reporting-section {
        background: var(--background-color, #f8f9fa);
        padding: 20px;
        border-radius: 12px;
        border: 2px solid #1E88E5;
        margin-bottom: 20px;
    }
    .reporting-section-header {
        background: #1E88E5;
        color: white;
        padding: 10px 16px;
        border-radius: 8px;
        margin-bottom: 15px;
        font-weight: 600;
    }
    .reporting-sub-header {
        background: #E3F2FD;
        padding: 10px 16px;
        border-radius: 8px;
        margin-bottom: 15px;
        font-weight: 600;
        border-left: 4px solid #1E88E5;
    }
    .reporting-sub-header-current {
        background: #E8F5E9;
        padding: 10px 16px;
        border-radius: 8px;
        margin-bottom: 15px;
        font-weight: 600;
        border-left: 4px solid #43A047;
    }
    .reporting-sub-header-today {
        background: #FFF3E0;
        padding: 10px 16px;
        border-radius: 8px;
        margin-bottom: 15px;
        font-weight: 600;
        border-left: 4px solid #FB8C00;
    }
    .note-box {
        background-color: #e3f2fd;
        border-left: 4px solid #1976d2;
        padding: 8px 12px;
        margin: 8px 0 16px 0;
        border-radius: 4px;
        font-size: 0.9rem;
        color: #1565c0;
    }
    .conversion-high { color: #2e7d32; font-weight: bold; }
    .conversion-medium { color: #f57f17; font-weight: bold; }
    .conversion-low { color: #c62828; font-weight: bold; }
    .employee-category-legend {
        background: #f0f2f6;
        padding: 8px 12px;
        border-radius: 6px;
        margin-bottom: 12px;
        font-size: 0.85rem;
    }
    .employee-category-legend .inbound { color: #155724; font-weight: bold; }
    .employee-category-legend .outbound { color: #721c24; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# INITIAL EMPLOYEE LIST WITH CATEGORIES
# ============================================================

ENROLMENT_AMOUNT_COLUMN_CANDIDATES = [
    "Enrolment Amount", "Enrollment Amount", "EnrolmentAmount", "EnrollmentAmount",
    "Enrolment Amt", "Enrollment Amt", "Enrolment Fee", "Enrollment Fee",
    "Amount", "Enrolment Value", "Enrollment Value", "Enrolment", "Enrollment",
    "Join Amount", "Joined Amount", "Join Value", "Joined Value", "Incentive Amount"
]

# Employee list with status categories (for display)
DEFAULT_EMPLOYEES = {
    "Empd69580": "Ramvinoth (IN)",
    "Empbc5e29": "Lavanya M (IN)",
    "Emp948406": "Srividya (OUT)",
    "Emp8a3956": "Shreen (OUT-Left)"
}

if "employees" not in st.session_state:
    st.session_state.employees = DEFAULT_EMPLOYEES.copy()

if "date_preset" not in st.session_state:
    st.session_state.date_preset = "All time"

if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False

if "decimal_places" not in st.session_state:
    st.session_state.decimal_places = 3

if "rounding_mode" not in st.session_state:
    st.session_state.rounding_mode = "Standard Rounding"

# ============================================================
# ROUNDING UTILITY FUNCTIONS
# ============================================================

def round_decimal(value, decimals=3):
    if value is None or pd.isna(value):
        return value
    try:
        return round(float(value), decimals)
    except (ValueError, TypeError):
        return value

def apply_rounding(value, decimals=3, mode="Standard Rounding"):
    if value is None or pd.isna(value):
        return value
    
    try:
        val = float(value)
        multiplier = 10 ** decimals
        
        if mode == "Standard Rounding":
            return round(val, decimals)
        elif mode == "Round Up (Ceiling)":
            return math.ceil(val * multiplier) / multiplier
        elif mode == "Round Down (Floor)":
            return math.floor(val * multiplier) / multiplier
        else:
            return round(val, decimals)
    except (ValueError, TypeError):
        return value

def format_currency(value, decimals=3, currency_symbol="₹"):
    if value is None or pd.isna(value):
        return "N/A"
    try:
        rounded_value = apply_rounding(value, decimals, st.session_state.get('rounding_mode', 'Standard Rounding'))
        if decimals == 0:
            return f"{currency_symbol}{rounded_value:,.0f}"
        else:
            return f"{currency_symbol}{rounded_value:,.{decimals}f}"
    except (ValueError, TypeError):
        return "N/A"

def format_number(value, decimals=3):
    if value is None or pd.isna(value):
        return "N/A"
    try:
        val = float(value)
        rounded_val = apply_rounding(val, decimals, st.session_state.get('rounding_mode', 'Standard Rounding'))
        if decimals == 0:
            return f"{rounded_val:,.0f}"
        else:
            return f"{rounded_val:,.{decimals}f}"
    except (ValueError, TypeError):
        return str(value)

def apply_rounding_to_df(df, columns, decimals=None, mode=None):
    if decimals is None:
        decimals = st.session_state.get('decimal_places', 3)
    if mode is None:
        mode = st.session_state.get('rounding_mode', 'Standard Rounding')
    
    df = df.copy()
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            df[col] = df[col].apply(lambda x: apply_rounding(x, decimals, mode))
    return df

def get_decimal_settings():
    return {
        'decimals': st.session_state.get('decimal_places', 3),
        'mode': st.session_state.get('rounding_mode', 'Standard Rounding')
    }

# ============================================================
# EXCEL EXPORT FUNCTIONS
# ============================================================

def create_reporting_excel(overall_df, month_df, today_df, company_name="BHIMA Jewellery – e-Gold App Telecaller Referral Report"):
    """Create a single-sheet Excel report with all three reporting sections."""
    output = BytesIO()
    
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.utils import get_column_letter
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Referral Report"
    
    # Define styles
    header_fill = PatternFill(start_color="1E88E5", end_color="1E88E5", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    sub_header_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
    sub_header_font = Font(bold=True, size=11)
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    total_font = Font(bold=True, size=10)
    note_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    current_row = 1
    
    # Title
    ws.merge_cells(f'A{current_row}:J{current_row}')
    title_cell = ws.cell(row=current_row, column=1, value=company_name)
    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = Alignment(horizontal='center')
    current_row += 1
    
    # Date and time info
    latest_date = datetime.now().strftime("%I:%M %p")
    ws.merge_cells(f'A{current_row}:J{current_row}')
    info_cell = ws.cell(row=current_row, column=1, value=f"Data Till {latest_date}")
    info_cell.font = Font(size=11, italic=True)
    info_cell.alignment = Alignment(horizontal='center')
    current_row += 2
    
    # Helper function to format data for Excel
    def format_value_for_excel(value):
        if pd.isna(value):
            return ""
        elif isinstance(value, (int, float)):
            if value == int(value):
                return int(value)
            else:
                return float(value)
        else:
            return str(value)
    
    # Helper function to set cell number format for percentage columns
    def apply_percentage_format(ws, row, col_idx, headers, value):
        if col_idx <= len(headers):
            header_name = str(headers[col_idx - 1])
            percentage_columns = {"Reg %", "Join %", "Conversion %", "Join/Reg Ratio %", "Enrolment %"}
            if header_name in percentage_columns and isinstance(value, (int, float)):
                cell = ws.cell(row=row, column=col_idx)
                cell.number_format = '0.00"%"'  # Shows as 14.37%
                return True
        return False
    
    # ===== SECTION 1: OVERALL =====
    if not overall_df.empty:
        # Header
        ws.merge_cells(f'A{current_row}:J{current_row}')
        header_cell = ws.cell(row=current_row, column=1, value="📊 Reporting Period: Overall Details")
        header_cell.font = Font(bold=True, size=12)
        header_cell.fill = sub_header_fill
        header_cell.alignment = Alignment(horizontal='left')
        current_row += 1
        
        # Table headers
        headers = list(overall_df.columns)
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        current_row += 1
        
        # Data rows
        for _, row in overall_df.iterrows():
            for col_idx, value in enumerate(row, 1):
                cell = ws.cell(row=current_row, column=col_idx)
                formatted_val = format_value_for_excel(value)
                cell.value = formatted_val
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
                
                # Apply percentage format
                if not apply_percentage_format(ws, current_row, col_idx, headers, value):
                    # Apply number format for other numeric values
                    if isinstance(value, (int, float)):
                        if value == int(value):
                            cell.number_format = '#,##0'
                        else:
                            cell.number_format = '#,##0.00'
                
                # Highlight total row
                if col_idx == 1 and str(value) == "OVERALL TOTAL":
                    for c in range(1, len(headers) + 1):
                        ws.cell(row=current_row, column=c).fill = total_fill
                        ws.cell(row=current_row, column=c).font = total_font
            current_row += 1
        
        # Add note about conversion calculation
        current_row += 1
        ws.merge_cells(f'A{current_row}:J{current_row}')
        note_cell = ws.cell(row=current_row, column=1, value="💡 Note: Conversion % = (Lifetime Joined / Lifetime Registered) × 100 (capped at 100%)")
        note_cell.font = Font(size=9, italic=True, color="1565C0")
        note_cell.fill = note_fill
        note_cell.alignment = Alignment(horizontal='left')
        current_row += 1
    
    # ===== SECTION 2: CURRENT MONTH =====
    if not month_df.empty:
        current_row += 1
        ws.merge_cells(f'A{current_row}:J{current_row}')
        header_cell = ws.cell(row=current_row, column=1, value=f"📅 Reporting Period: {date.today().replace(day=1).strftime('%B %d')}–{date.today().strftime('%B %d, %Y')}")
        header_cell.font = Font(bold=True, size=12)
        header_cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
        header_cell.alignment = Alignment(horizontal='left')
        current_row += 1
        
        # Table headers
        headers = list(month_df.columns)
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = PatternFill(start_color="43A047", end_color="43A047", fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        current_row += 1
        
        # Data rows
        for _, row in month_df.iterrows():
            for col_idx, value in enumerate(row, 1):
                cell = ws.cell(row=current_row, column=col_idx)
                formatted_val = format_value_for_excel(value)
                cell.value = formatted_val
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
                
                # Apply percentage format
                if not apply_percentage_format(ws, current_row, col_idx, headers, value):
                    if isinstance(value, (int, float)):
                        if value == int(value):
                            cell.number_format = '#,##0'
                        else:
                            cell.number_format = '#,##0.00'
                
                if col_idx == 1 and str(value) == "OVERALL TOTAL":
                    for c in range(1, len(headers) + 1):
                        ws.cell(row=current_row, column=c).fill = total_fill
                        ws.cell(row=current_row, column=c).font = total_font
            current_row += 1
        
        # Add note about period conversion
        current_row += 1
        ws.merge_cells(f'A{current_row}:J{current_row}')
        note_cell = ws.cell(row=current_row, column=1, value="💡 Note: Join/Registration Ratio % = (Joined this period / Registered this period) × 100. Values > 100% indicate joins from registrations made in previous periods.")
        note_cell.font = Font(size=9, italic=True, color="1565C0")
        note_cell.fill = note_fill
        note_cell.alignment = Alignment(horizontal='left')
        current_row += 1
    
    # ===== SECTION 3: TODAY =====
    if not today_df.empty:
        current_row += 1
        ws.merge_cells(f'A{current_row}:J{current_row}')
        header_cell = ws.cell(row=current_row, column=1, value=f"📆 Date: {date.today().strftime('%d-%m-%Y')}")
        header_cell.font = Font(bold=True, size=12)
        header_cell.fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
        header_cell.alignment = Alignment(horizontal='left')
        current_row += 1
        
        # Table headers
        headers = list(today_df.columns)
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = PatternFill(start_color="FB8C00", end_color="FB8C00", fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        current_row += 1
        
        # Data rows
        for _, row in today_df.iterrows():
            for col_idx, value in enumerate(row, 1):
                cell = ws.cell(row=current_row, column=col_idx)
                formatted_val = format_value_for_excel(value)
                cell.value = formatted_val
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
                
                # Apply percentage format
                if not apply_percentage_format(ws, current_row, col_idx, headers, value):
                    if isinstance(value, (int, float)):
                        if value == int(value):
                            cell.number_format = '#,##0'
                        else:
                            cell.number_format = '#,##0.00'
                
                if col_idx == 1 and str(value) == "OVERALL TOTAL":
                    for c in range(1, len(headers) + 1):
                        ws.cell(row=current_row, column=c).fill = total_fill
                        ws.cell(row=current_row, column=c).font = total_font
            current_row += 1
        
        # Add note about today's conversion
        current_row += 1
        ws.merge_cells(f'A{current_row}:J{current_row}')
        note_cell = ws.cell(row=current_row, column=1, value="💡 Note: Join/Registration Ratio % = (Joined today / Registered today) × 100. Values > 100% indicate joins from registrations made on previous days.")
        note_cell.font = Font(size=9, italic=True, color="1565C0")
        note_cell.fill = note_fill
        note_cell.alignment = Alignment(horizontal='left')
        current_row += 1
    
    # Auto-adjust column widths
    if len(headers) > 0:
        for col_idx in range(1, len(headers) + 1):
            max_length = 0
            col_letter = get_column_letter(col_idx)
            for row_idx in range(1, current_row):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.coordinate in ws.merged_cells:
                    continue
                try:
                    if cell.value:
                        length = len(str(cell.value))
                        if length > max_length:
                            max_length = length
                except:
                    pass
            adjusted_width = min(max_length + 2, 30)
            ws.column_dimensions[col_letter].width = adjusted_width
    
    wb.save(output)
    output.seek(0)
    return output

def create_styled_excel_report(sheets: dict, date_range_text="All Time"):
    """Create Excel report with multiple sheets, formatting, and colors."""
    output = BytesIO()
    
    from openpyxl import Workbook
    wb = Workbook()
    
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    total_font = Font(bold=True, size=10)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    positive_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    negative_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    neutral_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    
    blue_gradient = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    green_gradient = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    orange_gradient = PatternFill(start_color="FDEBD0", end_color="FDEBD0", fill_type="solid")
    purple_gradient = PatternFill(start_color="E4DFEC", end_color="E4DFEC", fill_type="solid")
    
    wb.remove(wb.active)
    
    ws_cover = wb.create_sheet(title="Report Info", index=0)
    cover_fill = PatternFill(start_color="1E88E5", end_color="1E88E5", fill_type="solid")
    cover_font = Font(color="FFFFFF", bold=True, size=14)
    
    cover_data = [
        ["EMPLOYEE REFERRAL DASHBOARD REPORT"],
        [""],
        [f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"],
        [f"Date Range: {date_range_text}"],
        [""],
        ["Report Summary"],
        [""],
        ["Metric", "Value"],
        ["Total Registered", sheets.get("Employee Summary", pd.DataFrame())["Registered"].sum() if not sheets.get("Employee Summary", pd.DataFrame()).empty else 0],
        ["Total Joined", sheets.get("Employee Summary", pd.DataFrame())["Joined"].sum() if not sheets.get("Employee Summary", pd.DataFrame()).empty else 0],
        ["Total Enrolment Amount", sheets.get("Employee Summary", pd.DataFrame())["Enrolment Amount"].sum() if not sheets.get("Employee Summary", pd.DataFrame()).empty else 0],
        ["Total Balance", sheets.get("Employee Summary", pd.DataFrame())["Balance"].sum() if not sheets.get("Employee Summary", pd.DataFrame()).empty else 0],
        [""],
        ["Decimal Precision", f"{st.session_state.get('decimal_places', 3)} places"],
        ["Rounding Mode", st.session_state.get('rounding_mode', 'Standard Rounding')],
        [""],
        ["Report Sheets Included:"],
        [", ".join([name for name in sheets.keys() if not sheets[name].empty])],
        [""],
        ["📌 Conversion Rate Calculation Notes:"],
        ["• Overall Report: Conversion % = (Total Joined / Total Registered) × 100 (capped at 100%)"],
        ["• Period Reports: Join/Registration Ratio % = (Joined in period / Registered in period) × 100"],
        ["• Join/Registration Ratio may exceed 100% as it includes joins from previous registrations"],
        ["• Individual employee conversion rates are capped at 100% for the Overall report only"]
    ]
    
    for row_idx, row_data in enumerate(cover_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws_cover.cell(row=row_idx, column=col_idx, value=value)
            if row_idx == 1:
                cell.font = Font(color="FFFFFF", bold=True, size=16)
                cell.fill = cover_fill
            elif row_idx in [8, 16]:
                cell.font = Font(bold=True)
            elif row_idx >= 20:
                cell.font = Font(size=9, italic=True, color="1565C0")
    
    for column in ws_cover.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
            except:
                pass
        ws_cover.column_dimensions[column_letter].width = min(max_length + 2, 50)
    
    for sheet_name, df in sheets.items():
        if df.empty:
            continue
            
        ws = wb.create_sheet(title=sheet_name[:31])
        
        for c_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=c_idx, value=str(col_name))
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
        
        for r_idx, row in enumerate(df.values, 2):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                
                if pd.isna(value):
                    cell.value = ""
                elif isinstance(value, (int, float)):
                    if value == int(value):
                        cell.value = int(value)
                    else:
                        cell.value = float(value)
                else:
                    cell.value = str(value)
                
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
                
                col_name = str(df.columns[c_idx - 1]).strip()
                percentage_columns = {"Reg %", "Join %", "Conversion %", "Join/Reg Ratio %", "Enrolment %"}
                
                # Apply percentage format with 2 decimal places
                if col_name in percentage_columns and isinstance(value, (int, float)):
                    cell.number_format = '0.00"%"'  # Shows as 14.37%
                elif isinstance(value, (int, float)):
                    cell.number_format = '#,##0.00' if not float(value).is_integer() else '#,##0'

                if col_name.lower() in ['balance', 'day balance', 'total balance']:
                    try:
                        if isinstance(value, (int, float)):
                            if value > 0:
                                cell.fill = positive_fill
                            elif value < 0:
                                cell.fill = negative_fill
                    except:
                        pass
                
                if any(keyword in col_name.lower() for keyword in ['registered', 'registration']):
                    if isinstance(value, (int, float)) and value > 0:
                        cell.fill = blue_gradient
                elif any(keyword in col_name.lower() for keyword in ['joined', 'join', 'enrolment']):
                    if isinstance(value, (int, float)) and value > 0:
                        cell.fill = green_gradient
                elif any(keyword in col_name.lower() for keyword in ['total', 'overall']):
                    if isinstance(value, (int, float)) and value > 0:
                        cell.fill = orange_gradient
                
                if isinstance(value, str) and ('total' in value.lower() or 'overall' in value.lower()):
                    for col_idx in range(1, len(df.columns) + 1):
                        total_cell = ws.cell(row=r_idx, column=col_idx)
                        total_cell.font = total_font
                        total_cell.fill = total_fill
                        total_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        length = len(str(cell.value))
                        if length > max_length:
                            max_length = length
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        ws.freeze_panes = 'A2'
        
        balance_cols = []
        for idx, col in enumerate(df.columns, 1):
            if col.lower() in ['balance', 'day balance', 'total balance']:
                balance_cols.append(idx)
        
        for col_idx in balance_cols:
            col_letter = ws.cell(row=1, column=col_idx).column_letter
            from openpyxl.formatting.rule import ColorScaleRule
            color_scale = ColorScaleRule(
                start_type='num',
                start_value=-10,
                start_color='FF0000',
                mid_type='num',
                mid_value=0,
                mid_color='FFFF00',
                end_type='num',
                end_value=10,
                end_color='00FF00'
            )
            ws.conditional_formatting.add(f'{col_letter}2:{col_letter}{len(df)+1}', color_scale)
    
    wb.save(output)
    output.seek(0)
    return output

def create_enhanced_excel_report(sheets: dict, date_range_text="All Time"):
    """Create enhanced Excel report with advanced formatting."""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            if df.empty:
                continue
                
            safe_name = name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
            
            workbook = writer.book
            worksheet = workbook[safe_name]
            
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
            
            header_fill = PatternFill(start_color="1E88E5", end_color="1E88E5", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True, size=11)
            total_fill = PatternFill(start_color="FFC107", end_color="FFC107", fill_type="solid")
            total_font = Font(bold=True, size=10, color="000000")
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            positive_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            negative_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            neutral_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            
            for cell in worksheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = border
            
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = border
                    
                    header_name = str(worksheet.cell(row=1, column=cell.column).value or '').strip()
                    percentage_columns = {'Reg %', 'Join %', 'Conversion %', 'Join/Reg Ratio %', 'Enrolment %'}

                    # Apply percentage format with 2 decimal places
                    if isinstance(cell.value, (int, float)) and header_name in percentage_columns:
                        cell.number_format = '0.00"%"'  # Shows as 14.37%
                    elif isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.000'
                        if cell.value == int(cell.value):
                            cell.number_format = '#,##0'
            
            balance_cols = [idx for idx, col in enumerate(df.columns, 1) if col.lower() in ['balance', 'day balance', 'total balance']]
            
            for col_idx in balance_cols:
                col_letter = worksheet.cell(row=1, column=col_idx).column_letter
                from openpyxl.formatting.rule import ColorScaleRule
                color_scale = ColorScaleRule(
                    start_type='num',
                    start_value=-10,
                    start_color='FF0000',
                    mid_type='num',
                    mid_value=0,
                    mid_color='FFFF00',
                    end_type='num',
                    end_value=10,
                    end_color='00FF00'
                )
                range_str = f'{col_letter}2:{col_letter}{worksheet.max_row}'
                worksheet.conditional_formatting.add(range_str, color_scale)
            
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
                first_cell = row[0]
                if first_cell.value and 'total' in str(first_cell.value).lower():
                    for cell in row:
                        cell.fill = total_fill
                        cell.font = total_font
            
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            length = len(str(cell.value))
                            if length > max_length:
                                max_length = length
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            worksheet.freeze_panes = 'A2'
    
    output.seek(0)
    return output

# ============================================================
# DATE PROCESSING FUNCTIONS
# ============================================================

@st.cache_data(show_spinner=False)
def load_data(file_bytes, file_name):
    try:
        buf = BytesIO(file_bytes)
        if file_name.lower().endswith(".csv"):
            try:
                df = pd.read_csv(buf)
            except UnicodeDecodeError:
                buf.seek(0)
                df = pd.read_csv(buf, encoding='latin1')
            except:
                buf.seek(0)
                df = pd.read_csv(buf, encoding='cp1252')
        else:
            df = pd.read_excel(buf)
        return df
    except Exception as e:
        st.error(f"Unable to read file: {e}")
        st.info("Please ensure the file is a valid Excel (.xlsx, .xls) or CSV file.")
        return None

def parse_dates_robust(series):
    if series is None:
        return pd.Series(dtype='datetime64[ns]')

    s = series.copy()
    result = pd.Series(pd.NaT, index=s.index, dtype='datetime64[ns]')

    native_mask = s.map(lambda x: isinstance(x, (pd.Timestamp, datetime, date)) and not pd.isna(x))
    if native_mask.any():
        result.loc[native_mask] = pd.to_datetime(s.loc[native_mask], errors='coerce')

    remaining = result.isna() & s.notna()
    if remaining.any():
        numeric = pd.to_numeric(s.loc[remaining], errors='coerce')
        excel_mask = numeric.between(1, 60000, inclusive='both')
        if excel_mask.any():
            result.loc[numeric.index[excel_mask]] = pd.to_datetime(
                numeric.loc[excel_mask], unit='D', origin='1899-12-30', errors='coerce'
            )

    remaining = result.isna() & s.notna()
    if remaining.any():
        text = s.loc[remaining].astype(str).str.strip()
        parsed = pd.to_datetime(text, errors='coerce', dayfirst=True)
        result.loc[text.index] = parsed

    remaining = result.isna() & s.notna()
    if remaining.any():
        text = s.loc[remaining].astype(str).str.strip()
        parsed = pd.to_datetime(text, errors='coerce', format='mixed')
        result.loc[text.index] = parsed

    return result

def fix_two_digit_years(df, date_columns):
    current_year = datetime.now().year
    
    for col in date_columns:
        if col in df.columns:
            mask_1900s = (df[col].dt.year >= 1900) & (df[col].dt.year < 1970) & df[col].notna()
            if mask_1900s.any():
                df.loc[mask_1900s, col] = df.loc[mask_1900s, col].apply(
                    lambda x: x.replace(year=x.year + 100) if pd.notna(x) else x
                )
                st.info(f"✅ Fixed {mask_1900s.sum()} dates in {col} (2-digit year issue)")
    
    return df

def detect_date_format(df):
    if df.empty:
        return "Unknown"
    
    formats_detected = []
    
    for col in ["Registered Date", "Joined Date"]:
        if col in df.columns and not df[col].isna().all():
            sample_dates = df[col].dropna().head(10)
            for date_val in sample_dates:
                if isinstance(date_val, pd.Timestamp):
                    formats_detected.append("ISO format (YYYY-MM-DD)")
                elif isinstance(date_val, str):
                    if '-' in date_val:
                        parts = date_val.split('-')
                        if len(parts) == 3:
                            if len(parts[0]) == 4:
                                formats_detected.append("YYYY-MM-DD")
                            elif len(parts[2]) == 4:
                                if len(parts[0]) == 2:
                                    formats_detected.append("DD-MM-YYYY")
                                else:
                                    formats_detected.append("MM-DD-YYYY")
                            elif len(parts[2]) == 2:
                                if len(parts[0]) == 2:
                                    formats_detected.append("DD-MM-YY")
                                else:
                                    formats_detected.append("MM-DD-YY")
                    elif '/' in date_val:
                        parts = date_val.split('/')
                        if len(parts) == 3:
                            if len(parts[0]) == 4:
                                formats_detected.append("YYYY/MM/DD")
                            elif len(parts[2]) == 4:
                                if len(parts[0]) == 2:
                                    formats_detected.append("DD/MM/YYYY")
                                else:
                                    formats_detected.append("MM/DD/YYYY")
                            elif len(parts[2]) == 2:
                                if len(parts[0]) == 2:
                                    formats_detected.append("DD/MM/YY")
                                else:
                                    formats_detected.append("MM/DD/YY")
    
    if formats_detected:
        from collections import Counter
        most_common = Counter(formats_detected).most_common(1)
        return most_common[0][0]
    return "Unknown"

@st.cache_data(show_spinner=False)
def process_data(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    required_columns = ["Registered Date", "Joined Date"]
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        return None, missing_columns

    if "Registered Date" in df.columns:
        df["Original Registered Date"] = df["Registered Date"].astype(str)
    if "Joined Date" in df.columns:
        df["Original Joined Date"] = df["Joined Date"].astype(str)

    df["Registered Date"] = parse_dates_robust(df["Registered Date"])
    df["Joined Date"] = parse_dates_robust(df["Joined Date"])

    df = fix_two_digit_years(df, ["Registered Date", "Joined Date"])

    current_date = pd.Timestamp.now().normalize()
    
    future_reg = df[df["Registered Date"] > current_date]
    future_join = df[df["Joined Date"] > current_date]
    
    if not future_reg.empty or not future_join.empty:
        st.markdown("""
        <div class="info-box">
            <strong>📅 Date Parsing Information</strong><br>
            Please review the date parsing results below. Future dates may indicate incorrect parsing.
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander("🔍 Date Parsing Diagnostic", expanded=True):
            st.subheader("Date Parsing Results")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Registered Date Samples:**")
                if not df["Registered Date"].isna().all():
                    sample_reg = df[["Original Registered Date", "Registered Date"]].head(10)
                    st.dataframe(sample_reg, use_container_width=True)
                    
                    if df["Registered Date"].notna().any():
                        reg_years = df[df["Registered Date"].notna()]["Registered Date"].dt.year.value_counts().sort_index()
                        st.write("**Year Distribution:**")
                        st.bar_chart(reg_years)
            
            with col2:
                st.write("**Joined Date Samples:**")
                if not df["Joined Date"].isna().all():
                    sample_join = df[["Original Joined Date", "Joined Date"]].head(10)
                    st.dataframe(sample_join, use_container_width=True)
                    
                    if df["Joined Date"].notna().any():
                        join_years = df[df["Joined Date"].notna()]["Joined Date"].dt.year.value_counts().sort_index()
                        st.write("**Year Distribution:**")
                        st.bar_chart(join_years)
            
            if not future_join.empty:
                st.warning(f"⚠️ {len(future_join)} future Joined Dates found")
                st.dataframe(
                    future_join[["Original Joined Date", "Joined Date", "Employee Code", "Employee Name"]].head(20),
                    use_container_width=True
                )
                
                future_years = future_join["Joined Date"].dt.year.value_counts().sort_index()
                st.write("**Future Years Distribution:**")
                st.dataframe(future_years, use_container_width=True)
                
                if st.button("🔧 Auto-fix Future Dates (subtract 100 years from 1900s dates)"):
                    st.info("Will fix in next step")

    df["Is Registered Future"] = df["Registered Date"] > current_date
    df["Is Joined Future"] = df["Joined Date"] > current_date

    if "Referrer Name" not in df.columns:
        if "Employee Code" in df.columns:
            df["Referrer Name"] = df["Employee Code"]
        else:
            df["Referrer Name"] = "Unknown"

    df["Referrer Name"] = df["Referrer Name"].fillna("").astype(str).str.strip()
    
    validation_issues = validate_data(df)
    if validation_issues:
        st.warning("Data validation issues found:")
        for issue in validation_issues:
            st.write(f"• {issue}")

    return df, None

def validate_data(df):
    issues = []
    
    reg_empty = df["Registered Date"].isna().all()
    join_empty = df["Joined Date"].isna().all()
    if reg_empty and join_empty:
        issues.append("⚠️ Both Registered Date and Joined Date columns are completely empty")
    elif reg_empty:
        issues.append("⚠️ Registered Date column is completely empty")
    elif join_empty:
        issues.append("⚠️ Joined Date column is completely empty")
    
    if not reg_empty:
        min_date = df["Registered Date"].min()
        max_date = df["Registered Date"].max()
        if pd.notna(min_date) and pd.notna(max_date):
            days_diff = (max_date - min_date).days
            if days_diff > 730:
                issues.append(f"⚠️ Date range spans {days_diff} days. This might be unusually large")
    
    current_date = pd.Timestamp.now().normalize()
    future_reg = (df["Registered Date"] > current_date).sum()
    future_join = (df["Joined Date"] > current_date).sum()
    if future_reg > 0:
        issues.append(f"⚠️ {future_reg} records have future Registered Dates")
        sample_dates = df[df["Registered Date"] > current_date]["Registered Date"].head(5)
        issues.append(f"   Sample future dates: {sample_dates.tolist()}")
    if future_join > 0:
        issues.append(f"⚠️ {future_join} records have future Joined Dates")
        sample_dates = df[df["Joined Date"] > current_date]["Joined Date"].head(5)
        issues.append(f"   Sample future dates: {sample_dates.tolist()}")
    
    for col in ["Registered Date", "Joined Date"]:
        if col in df.columns and not df[col].isna().all():
            year_counts = df[df[col].notna()][col].dt.year.value_counts()
            for year, count in year_counts.items():
                if year < 1970 and count > 0:
                    issues.append(f"⚠️ {count} records in {col} have year {year} (likely 2-digit year issue)")
    
    return issues

@st.cache_data(show_spinner=False)
def identify_employees(df, employee_items):
    employee_dict = dict(employee_items)
    
    # Create a clean mapping without suffixes for matching
    clean_name_map = {}
    for code, name in employee_dict.items():
        clean_name = name
        # Remove common suffixes for matching
        suffixes_to_remove = [" (IN)", " (OUT)", " (OUT-Left)", "(IN)", "(OUT)", "(OUT-Left)"]
        for suffix in suffixes_to_remove:
            if clean_name.endswith(suffix):
                clean_name = clean_name[:-len(suffix)]
        clean_name_map[code] = clean_name.strip()

    def identify_employee(referrer):
        referrer = str(referrer).strip().lower()
        # Try matching with clean names first
        for code, name in clean_name_map.items():
            if code.lower() in referrer or name.lower() in referrer:
                return code
        # Try matching with original names (with suffixes)
        for code, name in employee_dict.items():
            if code.lower() in referrer or name.lower() in referrer:
                return code
        return None

    df = df.copy()
    df["Employee Code"] = df["Referrer Name"].apply(identify_employee)

    if df["Employee Code"].isna().all():
        if "Employee Code" in df.columns:
            df["Employee Code"] = df["Employee Code"].apply(identify_employee)
        if "Employee Name" in df.columns:
            df["Employee Name"] = df["Employee Name"].apply(identify_employee)

    unmatched_count = df["Employee Code"].isna().sum()
    df_matched = df[df["Employee Code"].notna()].copy()
    
    # Map to the full name with category for display
    df_matched["Employee Name"] = df_matched["Employee Code"].map(employee_dict)

    return df_matched, unmatched_count

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def color_balance(val):
    """Style function for balance column."""
    if val > 0:
        return 'color: green; font-weight: bold'
    elif val < 0:
        return 'color: red; font-weight: bold'
    return 'color: gray'

def color_conversion(val):
    """Style function for conversion rate column."""
    if val >= 80:
        return 'color: #2e7d32; font-weight: bold'
    elif val >= 50:
        return 'color: #f57f17; font-weight: bold'
    elif val >= 25:
        return 'color: #e65100; font-weight: bold'
    else:
        return 'color: #c62828; font-weight: bold'

def trend_badge(current, previous):
    if previous == 0 and current == 0:
        return '<span class="badge badge-flat">—</span>'
    if previous == 0:
        return '<span class="badge badge-up">new</span>'
    pct = (current - previous) / previous * 100
    if pct > 0.5:
        return f'<span class="badge badge-up">▲ {pct:.0f}%</span>'
    elif pct < -0.5:
        return f'<span class="badge badge-down">▼ {abs(pct):.0f}%</span>'
    return '<span class="badge badge-flat">flat</span>'

def render_amount_cards(items):
    if not items:
        return
    
    decimals = st.session_state.get('decimal_places', 3)
    cols = st.columns(len(items))
    
    for col, (label, value, color) in zip(cols, items):
        with col:
            if "Amount" in label or "₹" in str(value) or "Enrolment" in label:
                try:
                    num_val = float(str(value).replace('₹', '').replace(',', '').strip())
                    formatted_value = format_currency(num_val, decimals)
                except:
                    formatted_value = str(value)
            else:
                try:
                    num_val = float(str(value).replace(',', '').strip())
                    if num_val == int(num_val):
                        formatted_value = f"{int(num_val):,}"
                    else:
                        formatted_value = format_number(num_val, decimals)
                except:
                    formatted_value = str(value)
            
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color:{color};margin:0;">{label}</h3>
                <h2 style="margin:5px 0;">{formatted_value}</h2>
                <div class="decimal-info">Rounded to {decimals} decimal place{'s' if decimals != 1 else ''}</div>
            </div>
            """, unsafe_allow_html=True)

def detect_enrolment_amount_column(df):
    if df is None or df.empty:
        return None
    normalized = {str(c).strip().lower(): c for c in df.columns}
    for candidate in ENROLMENT_AMOUNT_COLUMN_CANDIDATES:
        if candidate.lower() in normalized:
            return normalized[candidate.lower()]
    for c in df.columns:
        name = str(c).strip().lower()
        if ("enrol" in name or "enroll" in name) and any(x in name for x in ("amount", "amt", "fee", "value")):
            return c
    return None

def prepare_enrolment_amount(df, amount_column=None):
    out = df.copy()
    if out.empty:
        out["Enrolment Amount"] = pd.Series(dtype="float64")
        return out, amount_column

    source_col = amount_column or detect_enrolment_amount_column(out)
    if source_col and source_col in out.columns:
        raw = out[source_col].astype(str).str.replace(r"[^0-9.\-]", "", regex=True)
        out["Enrolment Amount"] = pd.to_numeric(raw, errors="coerce")
        decimals = st.session_state.get('decimal_places', 3)
        mode = st.session_state.get('rounding_mode', 'Standard Rounding')
        out["Enrolment Amount"] = out["Enrolment Amount"].apply(
            lambda x: apply_rounding(x, decimals, mode)
        )
        return out, source_col

    out["Enrolment Amount"] = pd.Series(float("nan"), index=out.index, dtype="float64")
    return out, None

def enrolment_amount_sum(df):
    if df is None or df.empty or "Enrolment Amount" not in df.columns:
        return None
    values = pd.to_numeric(df["Enrolment Amount"], errors="coerce").dropna()
    if values.empty:
        return None
    decimals = st.session_state.get('decimal_places', 3)
    mode = st.session_state.get('rounding_mode', 'Standard Rounding')
    return apply_rounding(float(values.sum()), decimals, mode)

def format_enrolment_amount(value, decimals=None):
    if value is None or pd.isna(value):
        return "N/A"
    if decimals is None:
        decimals = st.session_state.get('decimal_places', 3)
    try:
        return format_currency(value, decimals)
    except (ValueError, TypeError):
        return "N/A"

def filter_by_employees(df, employee_names):
    if employee_names and not df.empty:
        return df[df["Employee Name"].isin(employee_names)]
    return df

def paginate_dataframe(df, page_size=50, key_prefix="page"):
    if df.empty or len(df) <= page_size:
        return df
    
    total_pages = (len(df) + page_size - 1) // page_size
    page = st.selectbox(
        f"Page (Total: {len(df)} rows)", 
        range(1, total_pages + 1),
        key=f"{key_prefix}_paginator"
    )
    start_idx = (page - 1) * page_size
    end_idx = min(start_idx + page_size, len(df))
    return df.iloc[start_idx:end_idx]

def create_day_wise_summary(df, date_column, metric_name):
    if df.empty:
        return pd.DataFrame(columns=["Date", metric_name])

    df_copy = df[df[date_column].notna()].copy()
    if df_copy.empty:
        return pd.DataFrame(columns=["Date", metric_name])

    df_copy["Day"] = df_copy[date_column].dt.date
    day_rows = []

    for d in sorted(df_copy["Day"].unique()):
        day_data = df_copy[df_copy["Day"] == d]
        day_rows.append({"Date": d, metric_name: len(day_data)})

    summary = pd.DataFrame(day_rows)
    if len(summary) > 0:
        total = pd.DataFrame([{"Date": "OVERALL TOTAL", metric_name: summary[metric_name].sum()}])
        return pd.concat([summary, total], ignore_index=True)
    return pd.DataFrame(columns=["Date", metric_name])

def create_month_wise_summary(registered_df, joined_df):
    rows = {}
    if not registered_df.empty:
        reg = registered_df.copy()
        reg["Month"] = reg["Registered Date"].dt.to_period("M").astype(str)
        for month, cnt in reg.groupby("Month").size().items():
            rows.setdefault(month, {"Month": month, "Registered": 0, "Joined": 0, "Enrolment Amount": 0.0})
            rows[month]["Registered"] = cnt
    if not joined_df.empty:
        joi = joined_df.copy()
        joi["Month"] = joi["Joined Date"].dt.to_period("M").astype(str)
        for month, group in joi.groupby("Month"):
            rows.setdefault(month, {"Month": month, "Registered": 0, "Joined": 0, "Enrolment Amount": 0.0})
            rows[month]["Joined"] = len(group)
            rows[month]["Enrolment Amount"] = float(group["Enrolment Amount"].sum())
    if not rows:
        return pd.DataFrame(columns=["Month", "Registered", "Joined", "Enrolment Amount", "Balance"])
    out = pd.DataFrame(list(rows.values())).sort_values("Month")
    out["Balance"] = out["Registered"] - out["Joined"]
    decimals = st.session_state.get('decimal_places', 3)
    out["Enrolment Amount"] = out["Enrolment Amount"].apply(lambda x: apply_rounding(x, decimals))
    return out.reset_index(drop=True)

def create_employee_monthly_summary_with_zeros(registered_df, joined_df, all_employees=None, all_months=None):
    rows = []
    
    if all_employees is None:
        all_employees = set()
        if not registered_df.empty:
            all_employees.update(registered_df["Employee Name"].unique())
        if not joined_df.empty:
            all_employees.update(joined_df["Employee Name"].unique())
    
    if not all_employees:
        return pd.DataFrame(columns=["Month", "Employee Name", "Registered", "Joined", "Enrolment Amount", "Balance"])
    
    if all_months is None:
        all_months = set()
        if not registered_df.empty:
            reg_temp = registered_df.copy()
            reg_temp["Month"] = reg_temp["Registered Date"].dt.to_period("M").astype(str)
            all_months.update(reg_temp["Month"].unique())
        if not joined_df.empty:
            join_temp = joined_df.copy()
            join_temp["Month"] = join_temp["Joined Date"].dt.to_period("M").astype(str)
            all_months.update(join_temp["Month"].unique())
    
    if not all_months:
        return pd.DataFrame(columns=["Month", "Employee Name", "Registered", "Joined", "Enrolment Amount", "Balance"])
    
    reg_dict = {}
    if not registered_df.empty:
        reg_temp = registered_df.copy()
        reg_temp["Month"] = reg_temp["Registered Date"].dt.to_period("M").astype(str)
        for (month, employee), group in reg_temp.groupby(["Month", "Employee Name"]):
            reg_dict[(month, employee)] = len(group)
    
    join_dict = {}
    join_amount_dict = {}
    if not joined_df.empty:
        join_temp = joined_df.copy()
        join_temp["Month"] = join_temp["Joined Date"].dt.to_period("M").astype(str)
        for (month, employee), group in join_temp.groupby(["Month", "Employee Name"]):
            join_dict[(month, employee)] = len(group)
            join_amount_dict[(month, employee)] = float(group["Enrolment Amount"].sum())
    
    decimals = st.session_state.get('decimal_places', 3)
    for employee in sorted(all_employees):
        for month in sorted(all_months):
            reg_count = reg_dict.get((month, employee), 0)
            join_count = join_dict.get((month, employee), 0)
            amount = join_amount_dict.get((month, employee), 0.0)
            rows.append({
                "Month": month,
                "Employee Name": employee,
                "Registered": reg_count,
                "Joined": join_count,
                "Enrolment Amount": apply_rounding(amount, decimals),
                "Balance": reg_count - join_count
            })
    
    return pd.DataFrame(rows).sort_values(["Month", "Employee Name"]).reset_index(drop=True)

def create_employee_day_summary_with_zeros(registered_df, joined_df, all_employees=None, all_dates=None):
    rows = []
    
    if all_employees is None:
        all_employees = set()
        if not registered_df.empty:
            all_employees.update(registered_df["Employee Name"].unique())
        if not joined_df.empty:
            all_employees.update(joined_df["Employee Name"].unique())
    
    if not all_employees:
        return pd.DataFrame(columns=["Date", "Employee Name", "Registered", "Joined", "Enrolment Amount", "Balance"])
    
    if all_dates is None:
        all_dates = set()
        if not registered_df.empty:
            reg_temp = registered_df.copy()
            reg_temp["Date"] = reg_temp["Registered Date"].dt.date
            all_dates.update(reg_temp["Date"].unique())
        if not joined_df.empty:
            join_temp = joined_df.copy()
            join_temp["Date"] = join_temp["Joined Date"].dt.date
            all_dates.update(join_temp["Date"].unique())
    
    if not all_dates:
        return pd.DataFrame(columns=["Date", "Employee Name", "Registered", "Joined", "Enrolment Amount", "Balance"])
    
    reg_dict = {}
    if not registered_df.empty:
        reg_temp = registered_df.copy()
        reg_temp["Date"] = reg_temp["Registered Date"].dt.date
        for (date_d, employee), group in reg_temp.groupby(["Date", "Employee Name"]):
            reg_dict[(date_d, employee)] = len(group)
    
    join_dict = {}
    join_amount_dict = {}
    if not joined_df.empty:
        join_temp = joined_df.copy()
        join_temp["Date"] = join_temp["Joined Date"].dt.date
        for (date_d, employee), group in join_temp.groupby(["Date", "Employee Name"]):
            join_dict[(date_d, employee)] = len(group)
            join_amount_dict[(date_d, employee)] = float(group["Enrolment Amount"].sum())
    
    decimals = st.session_state.get('decimal_places', 3)
    for employee in sorted(all_employees):
        for date_d in sorted(all_dates):
            reg_count = reg_dict.get((date_d, employee), 0)
            join_count = join_dict.get((date_d, employee), 0)
            amount = join_amount_dict.get((date_d, employee), 0.0)
            rows.append({
                "Date": date_d,
                "Employee Name": employee,
                "Registered": reg_count,
                "Joined": join_count,
                "Enrolment Amount": apply_rounding(amount, decimals),
                "Balance": reg_count - join_count
            })
    
    return pd.DataFrame(rows).sort_values(["Date", "Employee Name"]).reset_index(drop=True)

def create_trend_chart(registered_day, joined_day):
    trend_data = pd.merge(
        registered_day[registered_day["Date"] != "OVERALL TOTAL"],
        joined_day[joined_day["Date"] != "OVERALL TOTAL"],
        on="Date", how="outer"
    ).fillna(0)

    if trend_data.empty:
        return None

    trend_data = trend_data.sort_values("Date")

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=trend_data["Date"], y=trend_data["Registered"],
        mode="lines+markers", name="Registered",
        line=dict(color="#1E88E5", width=3), marker=dict(size=8)
    ))
    fig.add_trace(go.Scatter(
        x=trend_data["Date"], y=trend_data["Joined"],
        mode="lines+markers", name="Joined",
        line=dict(color="#43A047", width=3), marker=dict(size=8)
    ))
    fig.update_layout(
        title="Trend: Registered vs Joined Over Time",
        xaxis_title="Date", yaxis_title="Count",
        hovermode="x unified",
        legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01),
        height=400
    )
    return fig

def create_employee_performance_chart(employee_summary):
    if employee_summary.empty:
        return None
    fig = go.Figure(data=[
        go.Bar(name="Registered", x=employee_summary["Employee Name"],
               y=employee_summary["Registered"], marker_color="#1E88E5"),
        go.Bar(name="Joined", x=employee_summary["Employee Name"],
               y=employee_summary["Joined"], marker_color="#43A047")
    ])
    fig.update_layout(
        title="Employee Performance", xaxis_title="Employee", yaxis_title="Count",
        barmode="group", height=400
    )
    return fig

def create_conversion_chart(registered_count, joined_count):
    fig = go.Figure(data=[
        go.Pie(labels=["Registered", "Joined"], values=[registered_count, joined_count],
               marker=dict(colors=["#1E88E5", "#43A047"]), hole=0.4)
    ])
    fig.update_layout(title="Registration to Join Conversion", height=300)
    return fig

def create_balance_chart(registered_day, joined_day):
    if registered_day.empty and joined_day.empty:
        return None, None

    balance_data = pd.merge(
        registered_day[registered_day["Date"] != "OVERALL TOTAL"] if not registered_day.empty else pd.DataFrame(columns=["Date", "Registered"]),
        joined_day[joined_day["Date"] != "OVERALL TOTAL"] if not joined_day.empty else pd.DataFrame(columns=["Date", "Joined"]),
        on="Date", how="outer"
    ).fillna(0)

    if balance_data.empty:
        return None, None

    balance_data["Balance"] = balance_data["Registered"] - balance_data["Joined"]
    balance_data = balance_data.sort_values("Date")

    fig = go.Figure()
    fig.add_trace(go.Bar(name="Registered", x=balance_data["Date"], y=balance_data["Registered"],
                          marker_color="#1E88E5", opacity=0.7))
    fig.add_trace(go.Bar(name="Joined", x=balance_data["Date"], y=balance_data["Joined"],
                          marker_color="#43A047", opacity=0.7))
    fig.add_trace(go.Scatter(
        name="Balance (R - J)", x=balance_data["Date"], y=balance_data["Balance"],
        mode="lines+markers", line=dict(color="#FF6F00", width=3, dash="dash"),
        marker=dict(size=10, color=balance_data["Balance"].apply(lambda x: "green" if x > 0 else "red" if x < 0 else "gray")),
        yaxis="y2"
    ))
    fig.add_shape(
        type="line", x0=balance_data["Date"].min(), y0=0,
        x1=balance_data["Date"].max(), y1=0,
        line=dict(color="gray", width=1, dash="dot"), yref="y2"
    )
    fig.update_layout(
        title="Daily Registered vs Joined with Balance",
        xaxis_title="Date",
        yaxis=dict(title="Count", side="left"),
        yaxis2=dict(title="Balance", overlaying="y", side="right", showgrid=False),
        hovermode="x unified",
        legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01),
        height=450, barmode="group"
    )
    return fig, balance_data

def apply_date_preset(preset, min_date, max_date):
    today = date.today()
    if preset == "Last 7 days":
        return max(min_date, today - timedelta(days=7)), min(max_date, today)
    if preset == "Last 30 days":
        return max(min_date, today - timedelta(days=30)), min(max_date, today)
    if preset == "Last 90 days":
        return max(min_date, today - timedelta(days=90)), min(max_date, today)
    if preset == "This month":
        start = today.replace(day=1)
        return max(min_date, start), min(max_date, today)
    return min_date, max_date

def display_insights(registered_df, joined_df, employee_summary):
    if registered_df.empty and joined_df.empty:
        return
    
    with st.expander("💡 Key Insights", expanded=False):
        insights = []
        
        if not employee_summary.empty and "Registered" in employee_summary.columns:
            try:
                top_employee = employee_summary.loc[employee_summary["Registered"].idxmax()]
                if top_employee["Registered"] > 0:
                    insights.append(f"🏆 Top Performer: **{top_employee['Employee Name']}** with **{top_employee['Registered']}** referrals")
            except:
                pass
        
        if not registered_df.empty:
            reg_dates = registered_df["Registered Date"].dt.date
            if len(reg_dates.unique()) > 1:
                days_range = (reg_dates.max() - reg_dates.min()).days
                if days_range > 0:
                    avg_per_day = len(registered_df) / days_range
                    insights.append(f"📈 Average registrations per day: **{avg_per_day:.1f}**")
        
        if not registered_df.empty and not joined_df.empty:
            conv_rate = (len(joined_df) / len(registered_df)) * 100
            insights.append(f"🔄 Overall conversion rate: **{conv_rate:.1f}%**")
        
        if not registered_df.empty:
            most_active_day = registered_df["Registered Date"].dt.date.mode()
            if not most_active_day.empty:
                insights.append(f"📅 Most active registration day: **{most_active_day.iloc[0]}**")
        
        for insight in insights:
            st.markdown(f"• {insight}")

def handle_future_dates(df):
    return df.copy(), False

# ============================================================
# REPORTING TABLE FUNCTIONS - UPDATED WITH FIXED PERCENTAGES
# ============================================================

def format_indian_amount(value):
    """Format amount in Indian numbering system (lakhs, crores)."""
    try:
        value = float(value)
        sign = "-" if value < 0 else ""
        value = abs(value)
        
        # Format with 2 decimal places
        formatted = f"{sign}₹{value:,.2f}"
        return formatted
    except (ValueError, TypeError):
        return "₹0"

def get_employee_category(name):
    """Get employee category based on name suffix."""
    if "(IN)" in name:
        return "Inbound"
    elif "(OUT-Left)" in name:
        return "Outbound (Left)"
    elif "(OUT)" in name:
        return "Outbound"
    return "Unknown"

def split_codes_by_category(employee_dict, codes):
    """Split a list of employee codes into (inbound_codes, outbound_codes) based on category."""
    inbound_codes = [c for c in codes if get_employee_category(employee_dict.get(c, c)) == "Inbound"]
    outbound_codes = [c for c in codes if get_employee_category(employee_dict.get(c, c)) in ("Outbound", "Outbound (Left)")]
    return inbound_codes, outbound_codes

def create_reporting_table(registered_df, joined_df, employee_dict, selected_codes, title, date_range_text="", show_note=True, total_label="OVERALL TOTAL"):
    """
    Create a formatted reporting table.

    IMPORTANT CATEGORY-TOTAL RULE:
    The Registered / Joined / Enrolment totals are always calculated from
    rows belonging to ``selected_codes``. This is critical for the Inbound
    and Outbound breakdown tables: each category must show its own subtotal,
    not the grand total of all employees.

    - Overall (Lifetime): Conversion % = (Lifetime Joined / Lifetime Registered) × 100, capped at 100%.
    - Period reports (Current Month, Today): Join/Registration Ratio % =
      (Joined / Registered) × 100 and may exceed 100%.
    """
    rows = []

    # If no employees selected, return empty.
    if not selected_codes:
        return pd.DataFrame().style, pd.DataFrame()

    # ------------------------------------------------------------------
    # CRITICAL FIX:
    # Filter BOTH dataframes to the selected employee codes BEFORE
    # calculating totals. Previously the category tables passed the full
    # dataframes plus inbound/outbound codes, causing each category's
    # OVERALL TOTAL row to incorrectly display the grand total (294/174).
    # ------------------------------------------------------------------
    selected_codes_set = set(selected_codes)
    registered_df = registered_df[
        registered_df["Employee Code"].isin(selected_codes_set)
    ].copy()
    joined_df = joined_df[
        joined_df["Employee Code"].isin(selected_codes_set)
    ].copy()

    # Totals are now category-specific when selected_codes is a category.
    total_registered = len(registered_df)
    total_joined = len(joined_df)
    total_enrolment = enrolment_amount_sum(joined_df) or 0

    # Determine report type
    is_period_report = title in ["Current Month", "Today"]
    is_overall_report = title == "Overall Details"

    # Set the ratio column name
    ratio_column_name = "Conversion %" if is_overall_report else "Join/Reg Ratio %"

    for code in selected_codes:
        name = employee_dict.get(code, code)
        emp_reg = registered_df[registered_df["Employee Code"] == code]
        emp_joi = joined_df[joined_df["Employee Code"] == code]
        reg_count = len(emp_reg)
        join_count = len(emp_joi)
        enrol_amount = enrolment_amount_sum(emp_joi) or 0

        # Percentages are relative to this table's total.
        reg_pct = (reg_count / total_registered * 100) if total_registered > 0 else 0
        join_pct = (join_count / total_joined * 100) if total_joined > 0 else 0

        # Calculate conversion / join-registration ratio.
        if reg_count > 0:
            raw_ratio = (join_count / reg_count * 100)
            if is_overall_report:
                ratio_value = min(raw_ratio, 100.0)
            else:
                ratio_value = raw_ratio
        else:
            ratio_value = 0

        enrol_pct = (enrol_amount / total_enrolment * 100) if total_enrolment > 0 else 0

        rows.append({
            "Employee Code": code,
            "Employee Name": name,
            "Registered": reg_count,
            "Joined": join_count,
            "Reg %": reg_pct,
            "Join %": join_pct,
            ratio_column_name: ratio_value,
            "Enrolment Amount": enrol_amount,
            "Enrolment %": enrol_pct,
            "Balance": reg_count - join_count
        })

    # Category subtotal (or grand total for the main Overall table).
    total_row = {
        "Employee Code": total_label,
        "Employee Name": "",
        "Registered": total_registered,
        "Joined": total_joined,
        "Reg %": 100.0 if total_registered > 0 else 0,
        "Join %": 100.0 if total_joined > 0 else 0,
        "Enrolment Amount": total_enrolment,
        "Enrolment %": 100.0 if total_enrolment > 0 else 0,
        "Balance": total_registered - total_joined
    }

    # Calculate total ratio from this table's subtotal.
    if total_registered > 0:
        total_ratio = (total_joined / total_registered * 100)
        total_row[ratio_column_name] = (
            min(total_ratio, 100.0) if is_overall_report else total_ratio
        )
    else:
        total_row[ratio_column_name] = 0

    rows.append(total_row)

    df = pd.DataFrame(rows)
    
    # Build format dict dynamically
    format_dict = {
        "Registered": "{:,.0f}",
        "Joined": "{:,.0f}",
        "Reg %": "{:.1f}%",
        "Join %": "{:.1f}%",
        "Enrolment Amount": format_indian_amount,
        "Enrolment %": "{:.1f}%",
        "Balance": "{:,.0f}"
    }
    
    # Add the ratio column with appropriate formatting
    format_dict[ratio_column_name] = "{:.1f}%"
    
    # Apply styling
    styled_df = df.style.map(color_balance, subset=["Balance"])
    styled_df = styled_df.map(color_conversion, subset=[ratio_column_name])
    styled_df = styled_df.format(format_dict)
    
    # Add explanatory note
    if show_note:
        if is_period_report:
            st.markdown(f"""
            <div class="note-box">
                💡 <strong>Join/Registration Ratio %</strong> = (Joined during period ÷ Registered during period) × 100.<br>
                Values above 100% are possible because customers joining during the period may have registered on previous dates.
            </div>
            """, unsafe_allow_html=True)
        elif is_overall_report:
            st.markdown("""
            <div class="note-box">
                💡 <strong>Conversion %</strong> = (Lifetime Joined ÷ Lifetime Registered) × 100 (capped at 100%).
            </div>
            """, unsafe_allow_html=True)
    
    return styled_df, df

def render_category_comparison(registered_df, joined_df, employee_dict, selected_codes):
    """Show a quick side-by-side Inbound vs Outbound totals comparison."""
    inbound_codes, outbound_codes = split_codes_by_category(employee_dict, selected_codes)

    def totals_for(codes):
        reg = registered_df[registered_df["Employee Code"].isin(codes)] if codes else registered_df.iloc[0:0]
        joi = joined_df[joined_df["Employee Code"].isin(codes)] if codes else joined_df.iloc[0:0]
        r, j = len(reg), len(joi)
        conv = min((j / r * 100), 100.0) if r > 0 else 0.0
        amt = enrolment_amount_sum(joi) or 0
        return r, j, conv, amt

    in_r, in_j, in_c, in_a = totals_for(inbound_codes)
    out_r, out_j, out_c, out_a = totals_for(outbound_codes)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
        <div class="metric-card" style="text-align:left;">
            <h3 style="color:#155724;margin:0;">🟢 Inbound Total</h3>
            <h2 style="margin:5px 0;">{in_r} Reg &nbsp;|&nbsp; {in_j} Joined</h2>
            <div class="decimal-info">Conversion: {in_c:.1f}% &nbsp;|&nbsp; Enrolment: {format_enrolment_amount(in_a)} &nbsp;|&nbsp; {len(inbound_codes)} employee(s)</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="metric-card" style="text-align:left;">
            <h3 style="color:#721c24;margin:0;">🔴 Outbound Total</h3>
            <h2 style="margin:5px 0;">{out_r} Reg &nbsp;|&nbsp; {out_j} Joined</h2>
            <div class="decimal-info">Conversion: {out_c:.1f}% &nbsp;|&nbsp; Enrolment: {format_enrolment_amount(out_a)} &nbsp;|&nbsp; {len(outbound_codes)} employee(s)</div>
        </div>
        """, unsafe_allow_html=True)

def render_category_split_tables(registered_df, joined_df, employee_dict, selected_codes, title):
    """Render separate Inbound / Outbound reporting tables side by side, plus a totals comparison."""
    inbound_codes, outbound_codes = split_codes_by_category(employee_dict, selected_codes)

    st.markdown("##### 🟢🔴 Inbound vs Outbound Breakdown")
    col_in, col_out = st.columns(2)
    with col_in:
        st.markdown('<div class="employee-category-legend"><span class="inbound">🟢 Inbound Employees</span></div>', unsafe_allow_html=True)
        if inbound_codes:
            in_styled, _ = create_reporting_table(
                registered_df,
                joined_df,
                employee_dict,
                inbound_codes,
                title,
                show_note=False,
                total_label="INBOUND TOTAL"
            )
            st.dataframe(in_styled, use_container_width=True, hide_index=True)
        else:
            st.info("No Inbound employees in the current selection.")
    with col_out:
        st.markdown('<div class="employee-category-legend"><span class="outbound">🔴 Outbound Employees</span></div>', unsafe_allow_html=True)
        if outbound_codes:
            out_styled, _ = create_reporting_table(
                registered_df,
                joined_df,
                employee_dict,
                outbound_codes,
                title,
                show_note=False,
                total_label="OUTBOUND TOTAL"
            )
            st.dataframe(out_styled, use_container_width=True, hide_index=True)
        else:
            st.info("No Outbound employees in the current selection.")

    render_category_comparison(registered_df, joined_df, employee_dict, selected_codes)

def get_latest_data_time(df):
    """Get the latest date/time from the data."""
    latest_time = "N/A"
    all_dates = []
    
    if "Registered Date" in df.columns:
        reg_dates = df["Registered Date"].dropna()
        if not reg_dates.empty:
            all_dates.extend(reg_dates.tolist())
    
    if "Joined Date" in df.columns:
        join_dates = df["Joined Date"].dropna()
        if not join_dates.empty:
            all_dates.extend(join_dates.tolist())
    
    if all_dates:
        latest_date = max(all_dates)
        if isinstance(latest_date, (pd.Timestamp, datetime)):
            latest_time = latest_date.strftime("%I:%M %p")
        else:
            latest_time = datetime.now().strftime("%I:%M %p")
    
    return latest_time

# ============================================================
# MAIN APP
# ============================================================

def main():
    st.markdown('<p class="main-header">📊 Employee Referral Status Dashboard</p>',
                unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Registered & Joined — Employee-wise, Day-wise and Consolidated with Balance Analysis</p>',
                unsafe_allow_html=True)

    # ============================================================
    # QUICK ROUNDING TOGGLE
    # ============================================================
    with st.container():
        col1, col2, col3, col4, col5, col6 = st.columns([2, 1, 1, 1, 1, 1])
        with col2:
            if st.button("🔢 0", help="Show whole numbers (no decimals)"):
                st.session_state.decimal_places = 0
                st.rerun()
        with col3:
            if st.button("🔢 2", help="Show 2 decimal places"):
                st.session_state.decimal_places = 2
                st.rerun()
        with col4:
            if st.button("🔢 3", help="Show 3 decimal places (default)"):
                st.session_state.decimal_places = 3
                st.rerun()
        with col5:
            if st.button("🔢 4", help="Show 4 decimal places"):
                st.session_state.decimal_places = 4
                st.rerun()
        with col6:
            if st.button("🔢 6", help="Show 6 decimal places"):
                st.session_state.decimal_places = 6
                st.rerun()
        st.caption(f"💡 Current decimal precision: {st.session_state.decimal_places} place{'s' if st.session_state.decimal_places != 1 else ''} | Rounding mode: {st.session_state.rounding_mode}")

    # ============================================================
    # SIDEBAR - EMPLOYEE MANAGEMENT
    # ============================================================
    with st.sidebar:
        st.header("⚙️ Configuration")

        # Display Employee Category Legend
        st.markdown("""
        <div class="employee-category-legend">
            📌 <b>Employee Categories:</b><br>
            <span class="inbound">🟢 IN</span> = Inbound &nbsp;|&nbsp; 
            <span class="outbound">🔴 OUT</span> = Outbound &nbsp;|&nbsp;
            <span class="outbound">🔴 OUT-Left</span> = Outbound employee who left the company
        </div>
        """, unsafe_allow_html=True)

        with st.expander("👥 Manage Employees", expanded=True):
            tab_list, tab_add, tab_bulk = st.tabs(["List", "Add / Edit", "Bulk Import"])

            with tab_list:
                if st.session_state.employees:
                    # Separate inbound and outbound employees
                    inbound = {}
                    outbound = {}
                    for code, name in st.session_state.employees.items():
                        if "(IN)" in name:
                            inbound[code] = name
                        else:
                            outbound[code] = name
                    
                    # Show Inbound Employees with green indicator
                    if inbound:
                        st.markdown("**🟢 Inbound Employees**")
                        for code, name in list(inbound.items()):
                            c1, c2, c3 = st.columns([3, 3, 1])
                            c1.write(f"**{code}**")
                            c2.write(name.replace("(IN)", "").strip())
                            if c3.button("🗑️", key=f"del_{code}", help=f"Delete {name}"):
                                del st.session_state.employees[code]
                                st.rerun()
                    
                    # Show Outbound Employees with red indicator
                    if outbound:
                        st.markdown("**🔴 Outbound Employees**")
                        for code, name in list(outbound.items()):
                            c1, c2, c3 = st.columns([3, 3, 1])
                            c1.write(f"**{code}**")
                            clean_name = name.replace("(OUT-Left)", "").replace("(OUT)", "").strip()
                            c2.write(clean_name)
                            status = "🔴 Outbound employee who left the company" if "Left" in name else "🔴 Outbound"
                            c2.caption(status)
                            if c3.button("🗑️", key=f"del_{code}", help=f"Delete {name}"):
                                del st.session_state.employees[code]
                                st.rerun()
                else:
                    st.warning("No employees configured. Add some in the next tab.")

                st.download_button(
                    "📥 Export employee list (JSON)",
                    data=json.dumps(st.session_state.employees, indent=2),
                    file_name="employees.json",
                    mime="application/json",
                    use_container_width=True
                )

                if st.button("🔄 Reset to Default Employees", use_container_width=True):
                    st.session_state.employees = DEFAULT_EMPLOYEES.copy()
                    st.rerun()

            with tab_add:
                st.markdown("""
                <div style="background: #e8f5e9; padding: 6px 12px; border-radius: 4px; margin-bottom: 8px;">
                    <small>💡 Use <b>(IN)</b> for Inbound, <b>(OUT)</b> for Outbound, or <b>(OUT-Left)</b> for an outbound employee who left the company</small>
                </div>
                """, unsafe_allow_html=True)
                
                edit_code = st.selectbox(
                    "Edit existing (optional)",
                    ["— New employee —"] + list(st.session_state.employees.keys()),
                    format_func=lambda x: x if x == "— New employee —" else f"{x} - {st.session_state.employees[x]}"
                )
                prefill_code = "" if edit_code == "— New employee —" else edit_code
                prefill_name = "" if edit_code == "— New employee —" else st.session_state.employees[edit_code]

                new_code = st.text_input("Employee Code", value=prefill_code, placeholder="e.g., Emp123456")
                
                col1, col2 = st.columns(2)
                with col1:
                    new_name = st.text_input("Employee Name", value=prefill_name, placeholder="e.g., John Doe")
                with col2:
                    employee_type = st.selectbox(
                        "Employee Type",
                        ["Inbound (IN)", "Outbound (OUT)", "Outbound employee who left the company (OUT-Left)"],
                        index=0
                    )
                    # Append status to name if not already present
                    type_suffix = {
                        "Inbound (IN)": "(IN)",
                        "Outbound (OUT)": "(OUT)",
                        "Outbound employee who left the company (OUT-Left)": "(OUT-Left)"
                    }[employee_type]
                    
                    # Remove any existing status suffix
                    clean_name = new_name
                    for suffix in ["(IN)", "(OUT)", "(OUT-Left)"]:
                        if clean_name.endswith(suffix):
                            clean_name = clean_name[:-len(suffix)].strip()
                    full_name = f"{clean_name} {type_suffix}".strip()

                if st.button("💾 Save Employee", use_container_width=True):
                    if new_code and new_name:
                        if edit_code != "— New employee —" and edit_code != new_code:
                            del st.session_state.employees[edit_code]
                        if new_code in st.session_state.employees and new_code != edit_code:
                            st.error(f"Employee code {new_code} already exists!")
                        else:
                            st.session_state.employees[new_code] = full_name
                            st.success(f"✅ Saved {full_name} ({new_code})")
                            st.rerun()
                    else:
                        st.error("Please enter both Employee Code and Name")

            with tab_bulk:
                st.caption("Paste rows as `code,name` (one per line) to add many employees at once.")
                st.markdown("""
                <div style="background: #fff3cd; padding: 6px 12px; border-radius: 4px; margin-bottom: 8px;">
                    <small>💡 Include <b>(IN)</b> or <b>(OUT)</b> in the name field</small>
                </div>
                """, unsafe_allow_html=True)
                bulk_text = st.text_area(
                    "Bulk paste", 
                    placeholder="Emp111111,Alex Kumar (IN)\nEmp222222,Priya Rao (OUT)\nEmp333333,Ravi Singh (OUT-Left)", 
                    height=120
                )
                if st.button("➕ Add All", use_container_width=True):
                    added, skipped = 0, 0
                    for line in bulk_text.splitlines():
                        line = line.strip()
                        if not line or "," not in line:
                            continue
                        code, name = [p.strip() for p in line.split(",", 1)]
                        if code and name and code not in st.session_state.employees:
                            st.session_state.employees[code] = name
                            added += 1
                        else:
                            skipped += 1
                    st.success(f"Added {added} employee(s), skipped {skipped}.")
                    if added:
                        st.rerun()

        st.info("💰 Enrolment Amount is calculated by summing the actual amount column from the joined records. Joined Count is never multiplied by a fixed rate.")

        st.header("📂 Upload Data")
        uploaded_file = st.file_uploader(
            "Upload Referral Excel / CSV File",
            type=["xlsx", "xls", "csv"],
            help="File must contain columns: 'Registered Date' and 'Joined Date'"
        )

        if uploaded_file is None:
            st.info("📌 Please upload your referral file to get started.")
            st.stop()

        # Decimal Settings in Sidebar
        st.header("⚙️ Decimal Settings")
        
        decimal_places = st.selectbox(
            "Decimal Places for Currency",
            options=[0, 1, 2, 3, 4, 5, 6],
            index=3,
            format_func=lambda x: f"{x} decimal place{'s' if x != 1 else ''}"
        )
        st.session_state.decimal_places = decimal_places
        
        rounding_mode = st.selectbox(
            "Rounding Mode",
            options=["Standard Rounding", "Round Up (Ceiling)", "Round Down (Floor)"],
            index=0
        )
        st.session_state.rounding_mode = rounding_mode
        
        st.caption(f"Current: {decimal_places} decimals, {rounding_mode}")

    employee_dict = st.session_state.employees
    employee_codes = list(employee_dict.keys())

    if not employee_codes:
        st.error("⚠️ No employees configured. Please add employees in the sidebar.")
        st.stop()

    # ============================================================
    # LOAD + PROCESS
    # ============================================================
    with st.spinner("📥 Loading data..."):
        file_bytes = uploaded_file.getvalue()
        df = load_data(file_bytes, uploaded_file.name)
        if df is None:
            st.stop()

    with st.spinner("🔄 Processing data..."):
        date_format = detect_date_format(df)
        st.info(f"📅 Detected date format: {date_format}")
        
        df, missing_columns = process_data(df)

        if missing_columns:
            st.error("🚨 Missing required columns in the uploaded file:")
            for col in missing_columns:
                st.write(f"- {col}")
            st.info("Please ensure your file contains these columns exactly as specified:")
            st.code("Registered Date, Joined Date")
            st.stop()

        with st.expander("📋 Detected Columns in File"):
            st.write(df.columns.tolist())

        df, unmatched_count = identify_employees(df, tuple(employee_dict.items()))

        if df.empty:
            st.warning("⚠️ No data found for the configured employees.")
            st.info("Configured employees:")
            for code, name in employee_dict.items():
                st.write(f"• {code} - {name}")
            st.stop()

        st.success(f"✅ Found {len(df)} matching records for {len(employee_codes)} employees")
        if unmatched_count:
            st.caption(f"ℹ️ {unmatched_count} row(s) in the file didn't match any configured employee and were excluded.")

    # ============================================================
    # DATE FILTERING
    # ============================================================
    df_filtered = df.copy()

    current_date = pd.Timestamp.now().normalize()
    df_filtered["Is Future Date"] = (
        (df_filtered["Registered Date"] > current_date) |
        (df_filtered["Joined Date"] > current_date)
    )

    referee_phone_col = next(
        (c for c in df.columns if str(c).strip().casefold() == "referee phone"),
        None
    )

    if referee_phone_col is not None:
        phone_series = df[referee_phone_col].astype("string").str.strip()
        valid_phone = phone_series.notna() & phone_series.ne("") & phone_series.ne("nan")
        duplicate_phone_mask = valid_phone & phone_series.duplicated(keep=False)
        duplicates = df.loc[duplicate_phone_mask].copy()
    else:
        duplicates = pd.DataFrame()

    if not duplicates.empty:
        with st.expander(
            f"🔁 {len(duplicates)} possible duplicate row(s) detected "
            f"(same Referee Phone)"
        ):
            st.dataframe(duplicates, use_container_width=True, hide_index=True)
    elif referee_phone_col is not None:
        st.success("✅ No duplicate Referee Phone numbers found.")

    if df_filtered.empty:
        st.warning("⚠️ No valid data found after processing.")
        st.stop()

    # ============================================================
    # SIDEBAR FILTERS
    # ============================================================
    with st.sidebar:
        st.header("🔎 Report Filters")

        category_filter = st.radio(
            "👤 Employee Category",
            ["All", "🟢 Inbound only", "🔴 Outbound only"],
            index=0,
            horizontal=False,
            help="Filter the whole dashboard down to Inbound or Outbound employees only."
        )

        search_term = st.text_input("Search employee by name/code", "")
        filterable_codes = [
            c for c in employee_codes
            if search_term.lower() in c.lower() or search_term.lower() in employee_dict.get(c, "").lower()
        ] if search_term else employee_codes

        if category_filter == "🟢 Inbound only":
            filterable_codes = [c for c in filterable_codes if get_employee_category(employee_dict.get(c, c)) == "Inbound"]
        elif category_filter == "🔴 Outbound only":
            filterable_codes = [c for c in filterable_codes if get_employee_category(employee_dict.get(c, c)) in ("Outbound", "Outbound (Left)")]

        selected_codes = st.multiselect(
            "Select Employees",
            filterable_codes,
            default=filterable_codes,
            format_func=lambda x: f"{x} - {employee_dict.get(x, x)}"
        )

        if not selected_codes:
            st.warning("⚠️ Please select at least one employee.")
            st.stop()

        all_dates = []
        if not df_filtered.empty:
            all_dates.extend(df_filtered["Registered Date"].dropna().dt.date.tolist())
            all_dates.extend(df_filtered["Joined Date"].dropna().dt.date.tolist())

        if all_dates:
            min_date = min(all_dates)
            max_date = max(all_dates)

            preset = st.selectbox(
                "📅 Quick range",
                ["All time", "Last 7 days", "Last 30 days", "Last 90 days", "This month", "Custom"],
                index=0
            )

            if preset == "Custom":
                unrestricted_end_date = max(max_date, date.today())
                date_range = st.date_input(
                    "Custom date range",
                    value=(min_date, unrestricted_end_date)
                )
                if isinstance(date_range, tuple) and len(date_range) == 2:
                    start_date, end_date = date_range
                else:
                    start_date = end_date = date_range[0] if isinstance(date_range, tuple) else date_range
            elif preset == "All time":
                start_date, end_date = min_date, max(max_date, date.today())
            else:
                start_date, end_date = apply_date_preset(preset, min_date, max_date)
                st.caption(f"{start_date} → {end_date}")

            if start_date > end_date:
                st.error("⚠️ Start date cannot be after end date")
                st.stop()
        else:
            start_date = end_date = None

        st.divider()
        if st.button("♻️ Reset All Filters", use_container_width=True):
            st.rerun()

    # Store date range text for report headers
    if start_date and end_date:
        if start_date == end_date:
            date_range_text = f"{start_date.strftime('%Y-%m-%d')}"
        else:
            date_range_text = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
    else:
        date_range_text = "All Time"

    # ============================================================
    # FILTER DATA
    # ============================================================
    employee_filtered_df = df_filtered[df_filtered["Employee Code"].isin(selected_codes)].copy()

    registered_df = employee_filtered_df[employee_filtered_df["Registered Date"].notna()].copy()
    if start_date is not None:
        registered_df = registered_df[
            (registered_df["Registered Date"].dt.date >= start_date) &
            (registered_df["Registered Date"].dt.date <= end_date)
        ].copy()

    joined_df = employee_filtered_df[employee_filtered_df["Joined Date"].notna()].copy()
    if start_date is not None:
        joined_df = joined_df[
            (joined_df["Joined Date"].dt.date >= start_date) &
            (joined_df["Joined Date"].dt.date <= end_date)
        ].copy()

    detected_amount_column = detect_enrolment_amount_column(employee_filtered_df)
    joined_df, _amount_source = prepare_enrolment_amount(joined_df, detected_amount_column)
    if _amount_source:
        st.sidebar.success(f"💰 Enrolment amount source: {_amount_source}")
    else:
        st.sidebar.error("❌ No enrolment amount column found in the uploaded data. Joined Count will NOT be converted into money.")

    prev_registered_count = prev_joined_count = 0
    if start_date is not None and end_date is not None:
        period_len = (end_date - start_date).days + 1
        prev_start = start_date - timedelta(days=period_len)
        prev_end = start_date - timedelta(days=1)
        prev_reg = employee_filtered_df[
            employee_filtered_df["Registered Date"].notna() &
            (employee_filtered_df["Registered Date"].dt.date >= prev_start) &
            (employee_filtered_df["Registered Date"].dt.date <= prev_end)
        ]
        prev_joi = employee_filtered_df[
            employee_filtered_df["Joined Date"].notna() &
            (employee_filtered_df["Joined Date"].dt.date >= prev_start) &
            (employee_filtered_df["Joined Date"].dt.date <= prev_end)
        ]
        prev_registered_count = len(prev_reg)
        prev_joined_count = len(prev_joi)

    # ============================================================
    # DISPLAY DATE RANGE IN REPORT HEADER
    # ============================================================
    st.markdown(f"""
    <div class="report-header">
        📊 Report Period: {date_range_text} | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    </div>
    """, unsafe_allow_html=True)

    # ============================================================
    # KPI METRICS
    # ============================================================
    registered_count = len(registered_df)
    joined_count = len(joined_df)
    
    total_count = registered_count
    
    conversion_rate = (joined_count / registered_count * 100) if registered_count > 0 else 0
    balance = registered_count - joined_count

    st.markdown("### 📊 Key Metrics")
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color:#1E88E5;margin:0;">📝 Registered</h3>
            <h2 style="margin:5px 0;">{registered_count} {trend_badge(registered_count, prev_registered_count)}</h2>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color:#43A047;margin:0;">✅ Joined</h3>
            <h2 style="margin:5px 0;">{joined_count} {trend_badge(joined_count, prev_joined_count)}</h2>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        balance_color = "green" if balance > 0 else "red" if balance < 0 else "gray"
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color:{balance_color};margin:0;">⚖️ Balance</h3>
            <h2 style="margin:5px 0;">{balance:,}</h2>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color:#FB8C00;margin:0;">📊 Total Activity</h3>
            <h2 style="margin:5px 0;">{total_count}</h2>
        </div>
        """, unsafe_allow_html=True)

    with col5:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color:#8E24AA;margin:0;">🔄 Conversion Rate</h3>
            <h2 style="margin:5px 0;">{conversion_rate:.1f}%</h2>
        </div>
        """, unsafe_allow_html=True)

    st.caption("Trend badges compare the selected period to an equal-length preceding period.")
    st.divider()

    # ============================================================
    # NEW: REPORTING SECTIONS - Overall, Current Month, Today
    # ============================================================
    
    # Get latest data time from the file
    latest_data_time = get_latest_data_time(employee_filtered_df)
    
    st.markdown("""
    <div class="reporting-section">
        <div class="reporting-section-header">
            📋 BHIMA Jewellery – e-Gold App Telecaller Referral Report
        </div>
    """, unsafe_allow_html=True)
    
    # ---- 1. OVERALL DETAILS ----
    st.markdown(f"""
    <div class="reporting-sub-header">
        📊 Reporting Period: Overall Details | Data Till {latest_data_time}
    </div>
    """, unsafe_allow_html=True)
    
    overall_registered_df = employee_filtered_df[employee_filtered_df["Registered Date"].notna()].copy()
    overall_joined_df = employee_filtered_df[employee_filtered_df["Joined Date"].notna()].copy()
    overall_joined_df, _ = prepare_enrolment_amount(overall_joined_df, detected_amount_column)
    
    overall_styled, overall_df = create_reporting_table(
        overall_registered_df, 
        overall_joined_df, 
        employee_dict, 
        selected_codes,
        "Overall Details"
    )
    st.dataframe(overall_styled, use_container_width=True, hide_index=True)

    render_category_split_tables(overall_registered_df, overall_joined_df, employee_dict, selected_codes, "Overall Details")
    
    st.divider()
    
    # ---- 2. CURRENT MONTH STATUS (1st to today) ----
    today = date.today()
    month_start = today.replace(day=1)
    
    st.markdown(f"""
    <div class="reporting-sub-header-current">
        📅 Reporting Period: {month_start.strftime('%B %d')}–{today.strftime('%B %d, %Y')} | Data Till {latest_data_time}
    </div>
    """, unsafe_allow_html=True)
    
    month_registered_df = employee_filtered_df[
        employee_filtered_df["Registered Date"].notna() &
        (employee_filtered_df["Registered Date"].dt.date >= month_start) &
        (employee_filtered_df["Registered Date"].dt.date <= today)
    ].copy()
    
    month_joined_df = employee_filtered_df[
        employee_filtered_df["Joined Date"].notna() &
        (employee_filtered_df["Joined Date"].dt.date >= month_start) &
        (employee_filtered_df["Joined Date"].dt.date <= today)
    ].copy()
    month_joined_df, _ = prepare_enrolment_amount(month_joined_df, detected_amount_column)
    
    month_styled, month_df = create_reporting_table(
        month_registered_df, 
        month_joined_df, 
        employee_dict, 
        selected_codes,
        "Current Month"
    )
    st.dataframe(month_styled, use_container_width=True, hide_index=True)

    render_category_split_tables(month_registered_df, month_joined_df, employee_dict, selected_codes, "Current Month")
    
    st.divider()
    
    # ---- 3. TODAY'S STATUS ----
    st.markdown(f"""
    <div class="reporting-sub-header-today">
        📆 Date: {today.strftime('%d-%m-%Y')} | Data Till {latest_data_time}
    </div>
    """, unsafe_allow_html=True)
    
    today_registered_df = employee_filtered_df[
        employee_filtered_df["Registered Date"].notna() &
        (employee_filtered_df["Registered Date"].dt.date == today)
    ].copy()
    
    today_joined_df = employee_filtered_df[
        employee_filtered_df["Joined Date"].notna() &
        (employee_filtered_df["Joined Date"].dt.date == today)
    ].copy()
    today_joined_df, _ = prepare_enrolment_amount(today_joined_df, detected_amount_column)
    
    today_styled, today_df = create_reporting_table(
        today_registered_df, 
        today_joined_df, 
        employee_dict, 
        selected_codes,
        "Today"
    )
    st.dataframe(today_styled, use_container_width=True, hide_index=True)

    render_category_split_tables(today_registered_df, today_joined_df, employee_dict, selected_codes, "Today")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # ============================================================
    # DOWNLOAD REPORTING EXCEL - SINGLE SHEET
    # ============================================================
    st.markdown("### ⬇️ Download Reporting Excel")
    st.caption("Download a single-sheet Excel report with all three reporting sections (Overall, Current Month, Today)")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_data = create_reporting_excel(overall_df, month_df, today_df)
        st.download_button(
            "📥 Download Single-Sheet Reporting Excel",
            data=excel_data,
            file_name=f"Referral_Report_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    st.divider()

    # ============================================================
    # EMPLOYEE PERFORMANCE WITH PERCENTAGES AND ENROLMENT AMOUNT
    # ============================================================
    employee_rows = []
    total_registered_all = len(registered_df)
    total_joined_all = len(joined_df)
    total_enrolment_all = enrolment_amount_sum(joined_df) or 0
    
    for code in selected_codes:
        name = employee_dict.get(code, code)
        emp_reg = registered_df[registered_df["Employee Code"] == code]
        emp_joi = joined_df[joined_df["Employee Code"] == code]
        reg_count = len(emp_reg)
        join_count = len(emp_joi)
        enrol_amount = enrolment_amount_sum(emp_joi) or 0
        
        reg_pct = (reg_count / total_registered_all * 100) if total_registered_all > 0 else 0
        join_pct = (join_count / total_joined_all * 100) if total_joined_all > 0 else 0
        
        # FIX: Cap conversion at 100% for individual employees
        if reg_count > 0:
            conversion_pct = min((join_count / reg_count * 100), 100.0)
        else:
            conversion_pct = 0
        
        enrol_pct = (enrol_amount / total_enrolment_all * 100) if total_enrolment_all > 0 else 0
        
        # Get employee category
        category = get_employee_category(name)
        
        employee_rows.append({
            "Employee Code": code, 
            "Employee Name": name,
            "Category": category,
            "Registered": reg_count, 
            "Joined": join_count,
            "Reg %": reg_pct,
            "Join %": join_pct,
            "Conversion %": conversion_pct,
            "Enrolment Amount": enrol_amount,
            "Enrolment %": enrol_pct,
            "Balance": reg_count - join_count, 
            "Total": reg_count + join_count
        })
    employee_summary = pd.DataFrame(employee_rows)

    # ============================================================
    # INSIGHTS
    # ============================================================
    display_insights(registered_df, joined_df, employee_summary)

    # ============================================================
    # TABBED LAYOUT FOR THE REST OF THE DASHBOARD
    # ============================================================
    (tab_overview, tab_reg, tab_joined, tab_employees, tab_performance,
     tab_daily, tab_monthly, tab_raw, tab_export) = st.tabs([
        "📈 Overview", "📝 Registered", "✅ Joined", "👥 Employees",
        "🏆 Performance", "📅 Day-wise", "🗓️ Month-wise", "📄 Raw Data", "⬇️ Export"
    ])

    # ---------------- OVERVIEW ----------------
    with tab_overview:
        registered_by_date = (
            registered_df.groupby(registered_df["Registered Date"].dt.date).size()
            .reset_index(name="Registered").rename(columns={"Registered Date": "Date"})
            if not registered_df.empty else pd.DataFrame(columns=["Date", "Registered"])
        )
        if not registered_by_date.empty:
            registered_by_date.columns = ["Date", "Registered"]

        joined_by_date = (
            joined_df.groupby(joined_df["Joined Date"].dt.date).size()
            .reset_index(name="Joined")
            if not joined_df.empty else pd.DataFrame(columns=["Date", "Joined"])
        )
        if not joined_by_date.empty:
            joined_by_date.columns = ["Date", "Joined"]

        if not registered_by_date.empty or not joined_by_date.empty:
            combined = pd.merge(registered_by_date, joined_by_date, on="Date", how="outer").fillna(0)
            combined = combined.sort_values("Date")
            combined["Registered"] = combined["Registered"].astype(int)
            combined["Joined"] = combined["Joined"].astype(int)
            combined["Balance"] = combined["Registered"] - combined["Joined"]
            combined["Total"] = combined["Registered"] + combined["Joined"]

            total_row = pd.DataFrame([{
                "Date": "TOTAL",
                "Registered": combined["Registered"].sum(),
                "Joined": combined["Joined"].sum(),
                "Balance": combined["Registered"].sum() - combined["Joined"].sum(),
                "Total": combined["Registered"].sum() + combined["Joined"].sum()
            }])
            combined_display = pd.concat([combined, total_row], ignore_index=True)

            st.subheader("Combined View: Registered vs Joined by Date")
            render_amount_cards([
                ("📝 Total Registered", f"{combined['Registered'].sum():,}", "#1E88E5"),
                ("✅ Total Joined", f"{combined['Joined'].sum():,}", "#43A047"),
                ("⚖️ Total Balance", f"{combined['Balance'].sum():+,}", "green" if combined['Balance'].sum() > 0 else "red" if combined['Balance'].sum() < 0 else "gray"),
                ("📊 Total Activity", f"{combined['Registered'].sum():,}", "#FB8C00"),
            ])
            st.dataframe(
                combined_display.style
                .background_gradient(subset=["Registered", "Joined", "Balance", "Total"], cmap="Oranges")
                .map(color_balance, subset=["Balance"]),
                use_container_width=True, hide_index=True
            )
        else:
            combined_display = pd.DataFrame(columns=["Date", "Registered", "Joined", "Balance", "Total"])
            st.info("No data available for combined view.")

        st.subheader("📈 Visualizations")
        c1, c2 = st.columns(2)
        with c1:
            conv_fig = create_conversion_chart(registered_count, joined_count)
            if conv_fig:
                st.plotly_chart(conv_fig, use_container_width=True)
        with c2:
            perf_fig = create_employee_performance_chart(employee_summary)
            if perf_fig:
                st.plotly_chart(perf_fig, use_container_width=True)

        registered_day = create_day_wise_summary(registered_df, "Registered Date", "Registered")
        joined_day = create_day_wise_summary(joined_df, "Joined Date", "Joined")

        trend_fig = create_trend_chart(
            registered_day[registered_day["Date"] != "OVERALL TOTAL"] if not registered_day.empty else pd.DataFrame(),
            joined_day[joined_day["Date"] != "OVERALL TOTAL"] if not joined_day.empty else pd.DataFrame()
        )
        if trend_fig:
            st.plotly_chart(trend_fig, use_container_width=True)

        st.subheader("⚖️ Balance Analysis")
        balance_fig, balance_data = create_balance_chart(registered_day, joined_day)
        if balance_fig and balance_data is not None:
            st.plotly_chart(balance_fig, use_container_width=True)
            b1, b2, b3 = st.columns(3)
            with b1:
                overall_balance = balance_data["Registered"].sum() - balance_data["Joined"].sum()
                st.metric("Overall Balance", f"{overall_balance:+,}")
            with b2:
                max_balance = balance_data["Balance"].max()
                max_balance_date = balance_data[balance_data["Balance"] == max_balance]["Date"].iloc[0]
                st.metric("Best Day Balance", f"{max_balance:+,}", f"on {max_balance_date}")
            with b3:
                min_balance = balance_data["Balance"].min()
                min_balance_date = balance_data[balance_data["Balance"] == min_balance]["Date"].iloc[0]
                st.metric("Worst Day Balance", f"{min_balance:+,}", f"on {min_balance_date}")
        else:
            balance_data = None

    # ---------------- REGISTERED ----------------
    with tab_reg:
        st.markdown("""
        <div class="registered-header">
            <h3 style="color:#1E88E5;margin:0;">📝 REGISTERED DATE — All Registrations</h3>
            <p style="margin:5px 0 0 0;color:#666;">Dates when employees were REGISTERED</p>
        </div>
        """, unsafe_allow_html=True)

        if not registered_df.empty:
            registered_employee = registered_df.groupby(
                [registered_df["Employee Code"], registered_df["Employee Name"], registered_df["Registered Date"].dt.date]
            ).size().reset_index(name="Count")
            registered_employee.columns = ["Employee Code", "Employee Name", "Registered Date", "Count"]

            render_amount_cards([
                ("📝 Total Registrations", f"{len(registered_df):,}", "#1E88E5"),
                ("📅 Unique Dates", f"{len(registered_df['Registered Date'].dt.date.unique()):,}", "#8E24AA"),
                ("👥 Employees", f"{len(registered_df['Employee Code'].unique()):,}", "#FB8C00"),
            ])
            st.dataframe(
                paginate_dataframe(registered_employee, key_prefix="reg_employee"),
                use_container_width=True, hide_index=True
            )

            registered_only = registered_df.groupby(registered_df["Registered Date"].dt.date).size().reset_index(name="Count")
            registered_only.columns = ["Date", "Registered Count"]
            registered_only["Registered Count"] = registered_only["Registered Count"].astype(int)
            total_row = pd.DataFrame([{"Date": "TOTAL", "Registered Count": registered_only["Registered Count"].sum()}])
            registered_only = pd.concat([registered_only, total_row], ignore_index=True)

            st.subheader("Day-wise Registrations")
            render_amount_cards([
                ("📝 Total Registered", f"{len(registered_df):,}", "#1E88E5"),
            ])
            st.dataframe(
                registered_only.style.background_gradient(subset=["Registered Count"], cmap="Greens"),
                use_container_width=True, hide_index=True
            )
        else:
            st.info("No registration data available.")
            registered_employee = pd.DataFrame(columns=["Employee Code", "Employee Name", "Registered Date", "Count"])
            registered_only = pd.DataFrame(columns=["Date", "Registered Count"])

    # ---------------- JOINED ----------------
    with tab_joined:
        st.markdown("""
        <div class="joined-header">
            <h3 style="color:#43A047;margin:0;">✅ JOINED DATE — All Joins</h3>
            <p style="margin:5px 0 0 0;color:#666;">Dates when employees JOINED</p>
        </div>
        """, unsafe_allow_html=True)

        if not joined_df.empty:
            joined_employee = joined_df.groupby(
                [joined_df["Employee Code"], joined_df["Employee Name"], joined_df["Joined Date"].dt.date]
            ).size().reset_index(name="Joined Count")
            joined_employee.columns = ["Employee Code", "Employee Name", "Joined Date", "Joined Count"]
            amount_by_group = (
                joined_df.assign(_JoinedDay=joined_df["Joined Date"].dt.date)
                .groupby(["Employee Code", "Employee Name", "_JoinedDay"])["Enrolment Amount"]
                .sum().reset_index().rename(columns={"_JoinedDay": "Joined Date"})
            )
            joined_employee = joined_employee.merge(amount_by_group, on=["Employee Code", "Employee Name", "Joined Date"], how="left")
            joined_employee["Enrolment Amount"] = joined_employee["Enrolment Amount"].fillna(0.0)

            render_amount_cards([
                ("✅ Total Enrolment Amount", format_enrolment_amount(enrolment_amount_sum(joined_df)), "#43A047"),
                ("📅 Unique Dates", f"{len(joined_df['Joined Date'].dt.date.unique()):,}", "#8E24AA"),
                ("👥 Employees", f"{len(joined_df['Employee Code'].unique()):,}", "#FB8C00"),
            ])
            st.dataframe(
                paginate_dataframe(joined_employee, key_prefix="join_employee"),
                use_container_width=True, hide_index=True
            )

            joined_only = joined_df.groupby(joined_df["Joined Date"].dt.date).size().reset_index(name="Joined Count")
            joined_only.columns = ["Date", "Joined Count"]
            joined_only["Joined Count"] = joined_only["Joined Count"].astype(int)
            joined_amount_by_day = joined_df.groupby(joined_df["Joined Date"].dt.date)["Enrolment Amount"].sum()
            joined_only["Enrolment Amount"] = joined_only["Date"].map(joined_amount_by_day).fillna(0.0)
            total_row = pd.DataFrame([{
                "Date": "TOTAL",
                "Joined Count": joined_only["Joined Count"].sum(),
                "Enrolment Amount": joined_only["Enrolment Amount"].sum()
            }])
            joined_only = pd.concat([joined_only, total_row], ignore_index=True)

            st.subheader("Day-wise Joins")
            render_amount_cards([
                ("✅ Total Enrolment Amount", format_enrolment_amount(enrolment_amount_sum(joined_df)), "#43A047"),
            ])
            st.dataframe(
                joined_only.style.background_gradient(subset=["Joined Count"], cmap="Blues"),
                use_container_width=True, hide_index=True
            )
        else:
            st.info("No joined data available.")
            joined_employee = pd.DataFrame(columns=["Employee Code", "Employee Name", "Joined Date", "Joined Count", "Enrolment Amount"])
            joined_only = pd.DataFrame(columns=["Date", "Joined Count", "Enrolment Amount"])

    # ---------------- EMPLOYEES ----------------
    with tab_employees:
        st.subheader("🔍 Particular Employee — Enrolment Amount")
        st.caption("Pick one employee to see how many of their referrals have enrolled (joined), within the current filters.")

        lookup_code = st.selectbox(
            "Select Employee",
            selected_codes,
            format_func=lambda x: f"{x} - {employee_dict.get(x, x)}",
            key="lookup_employee_code"
        )
        lookup_name = employee_dict.get(lookup_code, lookup_code)
        lookup_registered = len(registered_df[registered_df["Employee Code"] == lookup_code])
        lookup_joined = len(joined_df[joined_df["Employee Code"] == lookup_code])
        lookup_balance = lookup_registered - lookup_joined
        lookup_enrolment_amount = enrolment_amount_sum(joined_df[joined_df["Employee Code"] == lookup_code])
        lookup_conv = min((lookup_joined / lookup_registered * 100), 100) if lookup_registered > 0 else 0
        lookup_category = get_employee_category(lookup_name)

        lc1, lc2, lc3, lc4 = st.columns(4)
        with lc1:
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color:#1E88E5;margin:0;">📝 Registered</h3>
                <h2 style="margin:5px 0;">{lookup_registered}</h2>
                <small>{lookup_category}</small>
            </div>
            """, unsafe_allow_html=True)
        with lc2:
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color:#43A047;margin:0;">✅ Enrolment Amount</h3>
                <h2 style="margin:5px 0;">{format_enrolment_amount(lookup_enrolment_amount)}</h2>
                <small>{lookup_joined} joined • SUM of actual source enrolment amounts</small>
            </div>
            """, unsafe_allow_html=True)
        with lc3:
            lookup_balance_color = "green" if lookup_balance > 0 else "red" if lookup_balance < 0 else "gray"
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color:{lookup_balance_color};margin:0;">⚖️ Balance</h3>
                <h2 style="margin:5px 0;">{lookup_balance:+,}</h2>
            </div>
            """, unsafe_allow_html=True)
        with lc4:
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color:#8E24AA;margin:0;">🔄 Conversion Rate</h3>
                <h2 style="margin:5px 0;">{lookup_conv:.1f}%</h2>
            </div>
            """, unsafe_allow_html=True)
        st.caption(f"Showing figures for **{lookup_name}** ({lookup_code}) within the current filters.")

        with st.expander(f"📋 Enrolment records for {lookup_name}", expanded=False):
            lookup_joined_rows = joined_df[joined_df["Employee Code"] == lookup_code].copy()
            if not lookup_joined_rows.empty:
                lookup_joined_rows = lookup_joined_rows[["Employee Code", "Employee Name", "Joined Date"]].sort_values("Joined Date")
                st.dataframe(lookup_joined_rows, use_container_width=True, hide_index=True)
            else:
                st.info(f"No enrolment (joined) records found for {lookup_name} in the current filters.")

        st.divider()

        st.subheader("👥 Employee-wise Consolidated")

        employee_total = pd.DataFrame([{
            "Employee Code": "", "Employee Name": "OVERALL TOTAL", "Category": "",
            "Registered": employee_summary["Registered"].sum() if not employee_summary.empty else 0,
            "Joined": employee_summary["Joined"].sum() if not employee_summary.empty else 0,
            "Reg %": 100.0,
            "Join %": 100.0,
            "Conversion %": (employee_summary["Joined"].sum() / employee_summary["Registered"].sum() * 100) if employee_summary["Registered"].sum() > 0 else 0,
            "Enrolment Amount": employee_summary["Enrolment Amount"].sum() if not employee_summary.empty else 0,
            "Enrolment %": 100.0 if employee_summary["Enrolment Amount"].sum() > 0 else 0,
            "Balance": (employee_summary["Registered"].sum() - employee_summary["Joined"].sum()) if not employee_summary.empty else 0,
            "Total": employee_summary["Total"].sum() if not employee_summary.empty else 0
        }])
        employee_display = pd.concat([employee_summary, employee_total], ignore_index=True)

        _emp_bal_sum = employee_summary["Balance"].sum() if not employee_summary.empty else 0
        render_amount_cards([
            ("📝 Total Registered", f"{employee_summary['Registered'].sum() if not employee_summary.empty else 0:,}", "#1E88E5"),
            ("✅ Total Enrolment Amount", f"₹{format_number(employee_summary['Enrolment Amount'].sum() if not employee_summary.empty else 0, st.session_state.decimal_places)}", "#43A047"),
            ("⚖️ Total Balance", f"{_emp_bal_sum:+,}", "green" if _emp_bal_sum > 0 else "red" if _emp_bal_sum < 0 else "gray"),
            ("📊 Total Activity", f"{employee_summary['Registered'].sum() if not employee_summary.empty else 0:,}", "#FB8C00"),
        ])
        st.dataframe(
            employee_display.style
            .background_gradient(subset=["Registered", "Joined", "Total", "Enrolment Amount"], cmap="Blues")
            .map(color_balance, subset=["Balance"])
            .map(color_conversion, subset=["Conversion %"])
            .format({"Reg %": "{:.1f}%", "Join %": "{:.1f}%", "Conversion %": "{:.1f}%", 
                    "Enrolment Amount": "{:,.3f}", "Enrolment %": "{:.1f}%"}),
            use_container_width=True, hide_index=True
        )

        st.markdown("##### 🟢🔴 Inbound vs Outbound — Employee-wise")
        inbound_summary = employee_summary[employee_summary["Category"] == "Inbound"] if not employee_summary.empty else employee_summary
        outbound_summary = employee_summary[employee_summary["Category"].isin(["Outbound", "Outbound (Left)"])] if not employee_summary.empty else employee_summary
        col_in, col_out = st.columns(2)
        with col_in:
            st.markdown('<div class="employee-category-legend"><span class="inbound">🟢 Inbound Employees</span></div>', unsafe_allow_html=True)
            if not inbound_summary.empty:
                st.dataframe(
                    inbound_summary.style
                    .map(color_balance, subset=["Balance"])
                    .map(color_conversion, subset=["Conversion %"])
                    .format({"Reg %": "{:.1f}%", "Join %": "{:.1f}%", "Conversion %": "{:.1f}%",
                            "Enrolment Amount": "{:,.3f}", "Enrolment %": "{:.1f}%"}),
                    use_container_width=True, hide_index=True
                )
            else:
                st.info("No Inbound employees in the current selection.")
        with col_out:
            st.markdown('<div class="employee-category-legend"><span class="outbound">🔴 Outbound Employees</span></div>', unsafe_allow_html=True)
            if not outbound_summary.empty:
                st.dataframe(
                    outbound_summary.style
                    .map(color_balance, subset=["Balance"])
                    .map(color_conversion, subset=["Conversion %"])
                    .format({"Reg %": "{:.1f}%", "Join %": "{:.1f}%", "Conversion %": "{:.1f}%",
                            "Enrolment Amount": "{:,.3f}", "Enrolment %": "{:.1f}%"}),
                    use_container_width=True, hide_index=True
                )
            else:
                st.info("No Outbound employees in the current selection.")
        render_category_comparison(registered_df, joined_df, employee_dict, selected_codes)

        st.subheader("📋 Employee Status Particulars")
        status_rows = []
        for code in selected_codes:
            name = employee_dict.get(code, code)
            reg_count = len(registered_df[registered_df["Employee Code"] == code])
            join_count = len(joined_df[joined_df["Employee Code"] == code])
            emp_joi = joined_df[joined_df["Employee Code"] == code]
            category = get_employee_category(name)
            status_rows.append({
                "Employee": f"{code} - {name}", 
                "Category": category,
                "Registered": reg_count, 
                "Joined": join_count,
                "Enrolment Amount": enrolment_amount_sum(emp_joi),
                "Balance": reg_count - join_count, 
                "Total": reg_count + join_count
            })
        status_matrix = pd.DataFrame(status_rows)
        status_total = pd.DataFrame([{
            "Employee": "OVERALL TOTAL",
            "Category": "",
            "Registered": status_matrix["Registered"].sum(),
            "Joined": status_matrix["Joined"].sum(),
            "Enrolment Amount": status_matrix["Enrolment Amount"].sum(),
            "Balance": status_matrix["Registered"].sum() - status_matrix["Joined"].sum(),
            "Total": status_matrix["Total"].sum()
        }])
        status_display = pd.concat([status_matrix, status_total], ignore_index=True)

        _status_bal_sum = status_matrix["Balance"].sum() if not status_matrix.empty else 0
        render_amount_cards([
            ("📝 Total Registered", f"{status_matrix['Registered'].sum() if not status_matrix.empty else 0:,}", "#1E88E5"),
            ("✅ Total Enrolment Amount", f"₹{format_number(status_matrix['Enrolment Amount'].sum() if not status_matrix.empty else 0, st.session_state.decimal_places)}", "#43A047"),
            ("⚖️ Total Balance", f"{_status_bal_sum:+,}", "green" if _status_bal_sum > 0 else "red" if _status_bal_sum < 0 else "gray"),
            ("📊 Total Activity", f"{status_matrix['Registered'].sum() if not status_matrix.empty else 0:,}", "#FB8C00"),
        ])
        st.dataframe(
            status_display.style
            .background_gradient(subset=["Registered", "Joined", "Total"], cmap="Purples")
            .map(color_balance, subset=["Balance"])
            .format({"Enrolment Amount": "{:,.3f}"}),
            use_container_width=True, hide_index=True
        )

    # ---------------- PERFORMANCE TAB ----------------
    with tab_performance:
        st.subheader("🏆 Employee Performance Analysis")
        
        with st.container():
            st.markdown('<div class="filter-section">', unsafe_allow_html=True)
            st.markdown("### 📅 Performance Date Filter")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                perf_date_filter_type = st.selectbox(
                    "Date Filter Type",
                    ["All Data", "Specific Date", "Date Range"],
                    key="perf_date_filter_type"
                )
            
            with col2:
                all_perf_dates = []
                if not registered_df.empty:
                    all_perf_dates.extend(registered_df["Registered Date"].dt.date.tolist())
                if not joined_df.empty:
                    all_perf_dates.extend(joined_df["Joined Date"].dt.date.tolist())
                all_perf_dates = sorted(set(all_perf_dates)) if all_perf_dates else []
                
                if perf_date_filter_type == "Specific Date" and all_perf_dates:
                    perf_selected_date = st.date_input(
                        "Select Date",
                        value=max(all_perf_dates[-1], date.today()) if all_perf_dates else date.today(),
                        key="perf_selected_date"
                    )
                elif perf_date_filter_type == "Date Range" and all_perf_dates:
                    perf_date_range = st.date_input(
                        "Select Date Range",
                        value=(min(all_perf_dates), max(max(all_perf_dates), date.today())),
                        key="perf_date_range"
                    )
                else:
                    st.info("No dates available for filtering")
            
            with col3:
                perf_employee_filter = st.multiselect(
                    "Filter by Employee",
                    options=sorted(employee_dict.values()),
                    default=sorted(employee_dict.values()),
                    key="perf_employee_filter"
                )
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        perf_registered_df = registered_df.copy()
        perf_joined_df = joined_df.copy()
        
        perf_registered_df = filter_by_employees(perf_registered_df, perf_employee_filter)
        perf_joined_df = filter_by_employees(perf_joined_df, perf_employee_filter)
        
        if perf_date_filter_type == "Specific Date" and all_perf_dates:
            if 'perf_selected_date' in locals():
                perf_registered_df = perf_registered_df[perf_registered_df["Registered Date"].dt.date == perf_selected_date]
                perf_joined_df = perf_joined_df[perf_joined_df["Joined Date"].dt.date == perf_selected_date]
        elif perf_date_filter_type == "Date Range" and all_perf_dates:
            if 'perf_date_range' in locals() and isinstance(perf_date_range, tuple) and len(perf_date_range) == 2:
                start_dt, end_dt = perf_date_range
                perf_registered_df = perf_registered_df[
                    (perf_registered_df["Registered Date"].dt.date >= start_dt) &
                    (perf_registered_df["Registered Date"].dt.date <= end_dt)
                ]
                perf_joined_df = perf_joined_df[
                    (perf_joined_df["Joined Date"].dt.date >= start_dt) &
                    (perf_joined_df["Joined Date"].dt.date <= end_dt)
                ]
        
        perf_employee_rows = []
        perf_total_registered = len(perf_registered_df)
        perf_total_joined = len(perf_joined_df)
        perf_total_enrolment = enrolment_amount_sum(perf_joined_df) or 0
        
        perf_selected_codes = [code for code, name in employee_dict.items() if name in perf_employee_filter] if perf_employee_filter else selected_codes
        
        for code in perf_selected_codes:
            name = employee_dict.get(code, code)
            emp_reg = perf_registered_df[perf_registered_df["Employee Code"] == code]
            emp_joi = perf_joined_df[perf_joined_df["Employee Code"] == code]
            reg_count = len(emp_reg)
            join_count = len(emp_joi)
            enrol_amount = enrolment_amount_sum(emp_joi) or 0
            
            reg_pct = (reg_count / perf_total_registered * 100) if perf_total_registered > 0 else 0
            join_pct = (join_count / perf_total_joined * 100) if perf_total_joined > 0 else 0
            
            # FIX: Cap conversion at 100% for individual employees
            if reg_count > 0:
                conversion_pct = min((join_count / reg_count * 100), 100.0)
            else:
                conversion_pct = 0
            
            enrol_pct = (enrol_amount / perf_total_enrolment * 100) if perf_total_enrolment > 0 else 0
            
            category = get_employee_category(name)
            
            perf_employee_rows.append({
                "Employee Code": code, 
                "Employee Name": name,
                "Category": category,
                "Registered": reg_count, 
                "Joined": join_count,
                "Reg %": reg_pct,
                "Join %": join_pct,
                "Conversion %": conversion_pct,
                "Enrolment Amount": enrol_amount,
                "Enrolment %": enrol_pct,
                "Balance": reg_count - join_count, 
                "Total": reg_count + join_count
            })
        perf_employee_summary = pd.DataFrame(perf_employee_rows)
        
        if not perf_employee_summary.empty:
            st.markdown("### 📊 Performance Metrics with Rankings")
            
            perf_df = perf_employee_summary.copy()
            
            if not perf_df.empty:
                perf_df["Reg Rank"] = perf_df["Registered"].rank(method="min", ascending=False).astype(int)
                perf_df["Join Rank"] = perf_df["Joined"].rank(method="min", ascending=False).astype(int)
                perf_df["Conversion Rank"] = perf_df["Conversion %"].rank(method="min", ascending=False).astype(int)
                perf_df["Enrolment Rank"] = perf_df["Enrolment Amount"].rank(method="min", ascending=False).astype(int)
                
                def get_performance_badge(conversion_pct):
                    if conversion_pct >= 80:
                        return "⭐⭐⭐ Excellent"
                    elif conversion_pct >= 50:
                        return "⭐⭐ Good"
                    elif conversion_pct >= 25:
                        return "⭐ Average"
                    else:
                        return "⚡ Needs Improvement"
                
                perf_df["Performance"] = perf_df["Conversion %"].apply(get_performance_badge)
                
                def color_performance(val):
                    if "Excellent" in str(val):
                        return 'background-color: #d4edda; color: #155724;'
                    elif "Good" in str(val):
                        return 'background-color: #d1ecf1; color: #0c5460;'
                    elif "Average" in str(val):
                        return 'background-color: #fff3cd; color: #856404;'
                    else:
                        return 'background-color: #f8d7da; color: #721c24;'
                
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    top_reg = perf_df.loc[perf_df["Registered"].idxmax()] if not perf_df[perf_df["Registered"] > 0].empty else None
                    if top_reg is not None:
                        st.metric("🏆 Top Referrer", top_reg["Employee Name"], f"{top_reg['Registered']} registrations")
                
                with col2:
                    top_join = perf_df.loc[perf_df["Joined"].idxmax()] if not perf_df[perf_df["Joined"] > 0].empty else None
                    if top_join is not None:
                        st.metric("🎯 Top Enroller", top_join["Employee Name"], f"{top_join['Joined']} joins")
                
                with col3:
                    top_conv = perf_df.loc[perf_df["Conversion %"].idxmax()] if not perf_df[perf_df["Conversion %"] > 0].empty else None
                    if top_conv is not None:
                        st.metric("📈 Best Conversion", top_conv["Employee Name"], f"{top_conv['Conversion %']:.1f}%")
                
                with col4:
                    top_enrol = perf_df.loc[perf_df["Enrolment Amount"].idxmax()] if not perf_df[perf_df["Enrolment Amount"] > 0].empty else None
                    if top_enrol is not None:
                        st.metric("💰 Top Enrolment", top_enrol["Employee Name"], format_enrolment_amount(top_enrol["Enrolment Amount"]))
                
                with col5:
                    avg_conv = perf_df["Conversion %"].mean()
                    st.metric("📊 Avg Conversion", f"{avg_conv:.1f}%")
                
                st.divider()
                
                st.markdown("### 📋 Employee Performance Analysis")

                performance_display = perf_employee_summary[[
                    "Employee Code", "Employee Name", "Category", "Registered", "Joined",
                    "Reg %", "Join %", "Conversion %", "Enrolment Amount",
                    "Balance", "Total"
                ]].copy()

                overall_row = pd.DataFrame([{
                    "Employee Code": "OVERALL TOTAL",
                    "Employee Name": "",
                    "Category": "",
                    "Registered": perf_total_registered,
                    "Joined": perf_total_joined,
                    "Reg %": 100.0 if perf_total_registered > 0 else 0.0,
                    "Join %": 100.0 if perf_total_joined > 0 else 0.0,
                    "Conversion %": (perf_total_joined / perf_total_registered * 100) if perf_total_registered > 0 else 0.0,
                    "Enrolment Amount": perf_total_enrolment,
                    "Balance": perf_total_registered - perf_total_joined,
                    "Total": perf_total_registered + perf_total_joined
                }])

                performance_display = pd.concat([performance_display, overall_row], ignore_index=True)

                def style_overall_total(row):
                    styles = pd.Series("", index=row.index)
                    if row.get("Employee Code") == "OVERALL TOTAL":
                        styles[:] = "font-weight: bold; background-color: #fff3cd;"
                    return styles

                st.dataframe(
                    performance_display.style
                    .apply(style_overall_total, axis=1)
                    .map(color_balance, subset=["Balance"])
                    .map(color_conversion, subset=["Conversion %"])
                    .format({
                        "Registered": "{:,.0f}",
                        "Joined": "{:,.0f}",
                        "Reg %": "{:.1f}%",
                        "Join %": "{:.1f}%",
                        "Conversion %": "{:.1f}%",
                        "Enrolment Amount": format_indian_amount,
                        "Balance": "{:,.0f}",
                        "Total": "{:,.0f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )
                
                st.subheader("📊 Performance Visualizations")
                chart_col1, chart_col2 = st.columns(2)
                
                with chart_col1:
                    fig_conv = go.Figure()
                    fig_conv.add_trace(go.Bar(
                        x=perf_df["Employee Name"],
                        y=perf_df["Conversion %"],
                        name="Conversion %",
                        marker_color=perf_df["Conversion %"].apply(
                            lambda x: "#28a745" if x >= 80 else "#ffc107" if x >= 50 else "#dc3545"
                        ),
                        text=perf_df["Conversion %"].apply(lambda x: f"{x:.1f}%"),
                        textposition="outside"
                    ))
                    fig_conv.update_layout(
                        title="Conversion Rate by Employee (capped at 100%)",
                        xaxis_title="Employee",
                        yaxis_title="Conversion %",
                        height=400,
                        yaxis=dict(range=[0, max(perf_df["Conversion %"].max() * 1.2, 10)])
                    )
                    st.plotly_chart(fig_conv, use_container_width=True)
                
                with chart_col2:
                    fig_enrol = go.Figure()
                    fig_enrol.add_trace(go.Bar(
                        x=perf_df["Employee Name"],
                        y=perf_df["Enrolment Amount"],
                        name="Enrolment Amount",
                        marker_color="#43A047",
                        text=perf_df["Enrolment Amount"].apply(lambda x: format_enrolment_amount(x)),
                        textposition="outside"
                    ))
                    fig_enrol.update_layout(
                        title="Enrolment Amount by Employee",
                        xaxis_title="Employee",
                        yaxis_title="Enrolment Amount (₹)",
                        height=400
                    )
                    st.plotly_chart(fig_enrol, use_container_width=True)
                
                chart_col3, chart_col4 = st.columns(2)
                
                with chart_col3:
                    fig_comp = go.Figure()
                    fig_comp.add_trace(go.Bar(
                        name="Registered",
                        x=perf_df["Employee Name"],
                        y=perf_df["Registered"],
                        marker_color="#1E88E5"
                    ))
                    fig_comp.add_trace(go.Bar(
                        name="Joined",
                        x=perf_df["Employee Name"],
                        y=perf_df["Joined"],
                        marker_color="#43A047"
                    ))
                    fig_comp.update_layout(
                        title="Registration vs Join Performance",
                        xaxis_title="Employee",
                        yaxis_title="Count",
                        barmode="group",
                        height=400
                    )
                    st.plotly_chart(fig_comp, use_container_width=True)
                
                with chart_col4:
                    fig_enrol_pie = go.Figure()
                    fig_enrol_pie.add_trace(go.Pie(
                        labels=perf_df["Employee Name"],
                        values=perf_df["Enrolment Amount"],
                        hole=0.3,
                        marker=dict(colors=px.colors.qualitative.Set3),
                        textinfo="label+percent",
                        textposition="inside"
                    ))
                    fig_enrol_pie.update_layout(
                        title="Enrolment Amount Distribution",
                        height=400
                    )
                    st.plotly_chart(fig_enrol_pie, use_container_width=True)
                
                st.subheader("🏅 Top Performers")
                top_n = st.slider("Number of top performers to show", 1, min(10, len(perf_df)), 3, key="perf_top_n")
                
                top_by_reg = perf_df.nlargest(top_n, "Registered")[["Employee Name", "Category", "Registered", "Joined", "Conversion %", "Enrolment Amount"]]
                top_by_conv = perf_df.nlargest(top_n, "Conversion %")[["Employee Name", "Category", "Registered", "Joined", "Conversion %", "Enrolment Amount"]]
                top_by_enrol = perf_df.nlargest(top_n, "Enrolment Amount")[["Employee Name", "Category", "Registered", "Joined", "Conversion %", "Enrolment Amount"]]
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown("**📝 Top by Registrations**")
                    st.dataframe(
                        top_by_reg.style
                        .background_gradient(subset=["Registered"], cmap="Blues")
                        .format({"Conversion %": "{:.1f}%", "Enrolment Amount": "{:,.3f}"}),
                        use_container_width=True, hide_index=True
                    )
                
                with col2:
                    st.markdown("**📈 Top by Conversion Rate**")
                    st.dataframe(
                        top_by_conv.style
                        .background_gradient(subset=["Conversion %"], cmap="Greens")
                        .format({"Conversion %": "{:.1f}%", "Enrolment Amount": "{:,.3f}"}),
                        use_container_width=True, hide_index=True
                    )
                
                with col3:
                    st.markdown("**💰 Top by Enrolment Amount**")
                    st.dataframe(
                        top_by_enrol.style
                        .background_gradient(subset=["Enrolment Amount"], cmap="Oranges")
                        .format({"Conversion %": "{:.1f}%", "Enrolment Amount": "{:,.3f}"}),
                        use_container_width=True, hide_index=True
                    )
                
                with st.expander("📊 Performance Statistics Summary"):
                    stats_data = {
                        "Metric": ["Average Registrations", "Average Joins", "Average Conversion Rate", 
                                  "Best Conversion Rate", "Worst Conversion Rate", 
                                  "Total Registrations", "Total Joins", "Total Enrolment Amount",
                                  "Average Enrolment Amount", "Best Enrolment Amount", "Worst Enrolment Amount"],
                        "Value": [
                            f"{perf_df['Registered'].mean():.1f}",
                            f"{perf_df['Joined'].mean():.1f}",
                            f"{perf_df['Conversion %'].mean():.1f}%",
                            f"{perf_df['Conversion %'].max():.1f}%",
                            f"{perf_df['Conversion %'].min():.1f}%",
                            f"{perf_df['Registered'].sum():,}",
                            f"{perf_df['Joined'].sum():,}",
                            format_enrolment_amount(perf_df['Enrolment Amount'].sum()),
                            format_enrolment_amount(perf_df['Enrolment Amount'].mean()),
                            format_enrolment_amount(perf_df['Enrolment Amount'].max()),
                            format_enrolment_amount(perf_df['Enrolment Amount'].min())
                        ]
                    }
                    stats_df = pd.DataFrame(stats_data)
                    st.dataframe(stats_df, use_container_width=True, hide_index=True)
        else:
            st.info("No performance data available for the selected filters.")

    # ---------------- DAY-WISE ----------------
    with tab_daily:
        st.subheader("📅👥 Day-wise Employee Particulars")
        
        detail_summary = pd.DataFrame(columns=["Date", "Employee Code", "Employee Name", "Registered", "Joined", "Enrolment Amount", "Balance", "Total"])
        matrix_summary = pd.DataFrame()
        employee_day_data = pd.DataFrame(columns=["Date", "Employee Name", "Registered", "Joined", "Enrolment Amount", "Balance"])
        
        with st.container():
            st.markdown('<div class="filter-section">', unsafe_allow_html=True)
            col_filter1, col_filter2, col_filter3 = st.columns(3)
            
            with col_filter1:
                all_dates_available = []
                if not registered_df.empty:
                    all_dates_available.extend(registered_df["Registered Date"].dt.date.tolist())
                if not joined_df.empty:
                    all_dates_available.extend(joined_df["Joined Date"].dt.date.tolist())
                all_dates_available = sorted(set(all_dates_available))
                
                if all_dates_available:
                    date_filter_type = st.selectbox(
                        "Date Filter Type",
                        ["All Dates", "Specific Date", "Date Range"]
                    )
                else:
                    date_filter_type = "All Dates"
                    st.info("No dates available")
            
            with col_filter2:
                if date_filter_type == "Specific Date" and all_dates_available:
                    selected_date = st.date_input(
                        "Select Date",
                        value=max(all_dates_available[-1], date.today()) if all_dates_available else date.today()
                    )
                elif date_filter_type == "Date Range" and all_dates_available:
                    date_range = st.date_input(
                        "Select Date Range",
                        value=(min(all_dates_available), max(max(all_dates_available), date.today()))
                    )
            
            with col_filter3:
                day_employee_filter = st.multiselect(
                    "Filter by Employee",
                    options=sorted(employee_dict.values()),
                    default=sorted(employee_dict.values())
                )
                
                show_all_employees = st.checkbox("Show all employees (including 0)", value=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        filtered_registered_df = registered_df.copy()
        filtered_joined_df = joined_df.copy()
        
        filtered_registered_df = filter_by_employees(filtered_registered_df, day_employee_filter)
        filtered_joined_df = filter_by_employees(filtered_joined_df, day_employee_filter)
        
        if date_filter_type == "Specific Date" and all_dates_available:
            filtered_registered_df = filtered_registered_df[filtered_registered_df["Registered Date"].dt.date == selected_date]
            filtered_joined_df = filtered_joined_df[filtered_joined_df["Joined Date"].dt.date == selected_date]
        elif date_filter_type == "Date Range" and all_dates_available:
            if 'date_range' in locals() and isinstance(date_range, tuple) and len(date_range) == 2:
                start_dt, end_dt = date_range
                filtered_registered_df = filtered_registered_df[
                    (filtered_registered_df["Registered Date"].dt.date >= start_dt) &
                    (filtered_registered_df["Registered Date"].dt.date <= end_dt)
                ]
                filtered_joined_df = filtered_joined_df[
                    (filtered_joined_df["Joined Date"].dt.date >= start_dt) &
                    (filtered_joined_df["Joined Date"].dt.date <= end_dt)
                ]
        
        if show_all_employees:
            all_employees_for_day = set(day_employee_filter) if day_employee_filter else set(employee_dict.values())
            
            if date_filter_type == "Specific Date" and all_dates_available:
                all_dates_for_day = [selected_date]
            elif date_filter_type == "Date Range" and all_dates_available:
                if 'date_range' in locals() and isinstance(date_range, tuple) and len(date_range) == 2:
                    start_dt, end_dt = date_range
                    all_dates_for_day = [start_dt + timedelta(days=x) for x in range((end_dt - start_dt).days + 1)]
                else:
                    all_dates_for_day = all_dates_available
            else:
                all_dates_for_day = all_dates_available if all_dates_available else []
            
            employee_day_data = create_employee_day_summary_with_zeros(
                filtered_registered_df, 
                filtered_joined_df,
                all_employees_for_day,
                set(all_dates_for_day) if all_dates_for_day else None
            )
            
            if not employee_day_data.empty:
                st.subheader("📋 Day-wise Employee Details (All Employees)")
                _dd_bal_sum = employee_day_data["Balance"].sum()
                render_amount_cards([
                    ("📝 Total Registered", f"{employee_day_data['Registered'].sum():,}", "#1E88E5"),
                    ("✅ Total Joined", f"{employee_day_data['Joined'].sum():,}", "#43A047"),
                    ("💰 Total Enrolment Amount", f"₹{format_number(employee_day_data['Enrolment Amount'].sum(), st.session_state.decimal_places)}", "#2E7D32"),
                    ("⚖️ Total Balance", f"{_dd_bal_sum:+,}", "green" if _dd_bal_sum > 0 else "red" if _dd_bal_sum < 0 else "gray"),
                ])
                st.dataframe(
                    paginate_dataframe(employee_day_data, key_prefix="day_emp")
                    .style.format({"Enrolment Amount": "{:,.3f}"}),
                    use_container_width=True, hide_index=True
                )
                
                st.subheader("📊 Pivot View - Daily Activity by Employee")
                try:
                    pivot_reg = employee_day_data.pivot(index="Date", columns="Employee Name", values="Registered").fillna(0).astype(int)
                    pivot_join = employee_day_data.pivot(index="Date", columns="Employee Name", values="Joined").fillna(0).astype(int)
                    
                    pivot_combined = pd.DataFrame()
                    for col in pivot_reg.columns:
                        pivot_combined[f"{col} (R/J)"] = pivot_reg[col].astype(str) + " / " + pivot_join[col].astype(str)
                    
                    pivot_combined["Total R"] = pivot_reg.sum(axis=1)
                    pivot_combined["Total J"] = pivot_join.sum(axis=1)
                    pivot_combined["Balance"] = pivot_reg.sum(axis=1) - pivot_join.sum(axis=1)
                    pivot_combined.index = pivot_reg.index
                    
                    st.dataframe(pivot_combined, use_container_width=True)
                except Exception as e:
                    st.warning(f"Could not create pivot view: {e}")
            else:
                st.info("No data available for the selected filters.")
        else:
            if not filtered_registered_df.empty:
                filtered_registered_df["Registered Day"] = filtered_registered_df["Registered Date"].dt.date
            else:
                filtered_registered_df["Registered Day"] = pd.Series(dtype='object')
            
            if not filtered_joined_df.empty:
                filtered_joined_df["Joined Day"] = filtered_joined_df["Joined Date"].dt.date
            else:
                filtered_joined_df["Joined Day"] = pd.Series(dtype='object')
            
            all_activity_dates = sorted(
                set(filtered_registered_df["Registered Day"].dropna().tolist()) |
                set(filtered_joined_df["Joined Day"].dropna().tolist())
            )
            
            if all_activity_dates:
                detail_rows = []
                filtered_codes = [code for code, name in employee_dict.items() if name in day_employee_filter] if day_employee_filter else selected_codes
                
                for current_date_d in all_activity_dates:
                    for code in filtered_codes:
                        name = employee_dict.get(code, code)
                        reg_day = len(filtered_registered_df[(filtered_registered_df["Registered Day"] == current_date_d) & (filtered_registered_df["Employee Code"] == code)])
                        joi_day = len(filtered_joined_df[(filtered_joined_df["Joined Day"] == current_date_d) & (filtered_joined_df["Employee Code"] == code)])
                        if reg_day > 0 or joi_day > 0:
                            detail_rows.append({
                                "Date": current_date_d, "Employee Code": code, "Employee Name": name,
                                "Registered": reg_day, "Joined": joi_day,
                                "Enrolment Amount": float(filtered_joined_df.loc[(filtered_joined_df["Joined Day"] == current_date_d) & (filtered_joined_df["Employee Code"] == code), "Enrolment Amount"].sum()),
                                "Balance": reg_day - joi_day, "Total": reg_day + joi_day
                            })
                
                if detail_rows:
                    detail_summary = pd.DataFrame(detail_rows)
                    st.subheader("📋 Day-wise Employee Details")
                    _det_bal_sum = detail_summary["Balance"].sum()
                    render_amount_cards([
                        ("📝 Total Registered", f"{detail_summary['Registered'].sum():,}", "#1E88E5"),
                        ("✅ Total Joined", f"{detail_summary['Joined'].sum():,}", "#43A047"),
                        ("💰 Total Enrolment Amount", f"₹{format_number(detail_summary['Enrolment Amount'].sum(), st.session_state.decimal_places)}", "#2E7D32"),
                        ("⚖️ Total Balance", f"{_det_bal_sum:+,}", "green" if _det_bal_sum > 0 else "red" if _det_bal_sum < 0 else "gray"),
                    ])
                    st.dataframe(
                        paginate_dataframe(detail_summary, key_prefix="day_detail")
                        .style.format({"Enrolment Amount": "{:,.3f}"}),
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.info("ℹ️ No activity data available for the selected filters.")
                
                st.subheader("📊 Day-wise Consolidated Matrix")
                matrix_rows = []
                for current_date_d in all_activity_dates:
                    row = {"Date": current_date_d}
                    day_total = day_registered = day_joined = 0
                    day_enrolment_amount = 0.0
                    for code in filtered_codes:
                        name = employee_dict.get(code, code)
                        reg_m = len(filtered_registered_df[(filtered_registered_df["Registered Day"] == current_date_d) & (filtered_registered_df["Employee Code"] == code)])
                        joi_m = len(filtered_joined_df[(filtered_joined_df["Joined Day"] == current_date_d) & (filtered_joined_df["Employee Code"] == code)])
                        day_registered += reg_m
                        day_joined += joi_m
                        day_enrolment_amount += float(filtered_joined_df.loc[(filtered_joined_df["Joined Day"] == current_date_d) & (filtered_joined_df["Employee Code"] == code), "Enrolment Amount"].sum())
                        row[f"{code} - {name}"] = f"R: {reg_m} | J: {joi_m}"
                        day_total += reg_m + joi_m
                    if day_total > 0:
                        row["Day Registered"] = day_registered
                        row["Day Joined"] = day_joined
                        row["Day Enrolment Amount"] = day_enrolment_amount
                        row["Day Balance"] = day_registered - day_joined
                        row["Day Total"] = day_total
                        matrix_rows.append(row)
                
                if matrix_rows:
                    matrix_summary = pd.DataFrame(matrix_rows)
                    columns_order = ["Date"] + [f"{code} - {employee_dict.get(code, code)}" for code in filtered_codes] + \
                                    ["Day Registered", "Day Joined", "Day Enrolment Amount", "Day Balance", "Day Total"]
                    matrix_summary = matrix_summary[columns_order]
                    _mx_bal_sum = matrix_summary["Day Balance"].sum()
                    render_amount_cards([
                        ("📝 Total Registered", f"{matrix_summary['Day Registered'].sum():,}", "#1E88E5"),
                        ("✅ Total Joined", f"{matrix_summary['Day Joined'].sum():,}", "#43A047"),
                        ("💰 Total Enrolment Amount", f"₹{format_number(matrix_summary['Day Enrolment Amount'].sum(), st.session_state.decimal_places)}", "#2E7D32"),
                        ("⚖️ Total Balance", f"{_mx_bal_sum:+,}", "green" if _mx_bal_sum > 0 else "red" if _mx_bal_sum < 0 else "gray"),
                        ("📊 Total Activity", f"{matrix_summary['Day Registered'].sum():,}", "#FB8C00"),
                    ])
                    st.dataframe(
                        matrix_summary.style
                        .map(color_balance, subset=["Day Balance"])
                        .format({"Day Enrolment Amount": "{:,.3f}"}),
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.info("ℹ️ No consolidated data available for the selected filters.")
            else:
                st.info("ℹ️ No activity data available for the selected filters.")

    # ---------------- MONTH-WISE ----------------
    with tab_monthly:
        st.subheader("🗓️ Month-over-Month Summary")
        
        with st.container():
            st.markdown('<div class="filter-section">', unsafe_allow_html=True)
            col_filter1, col_filter2 = st.columns(2)
            
            with col_filter1:
                all_months_available = set()
                if not registered_df.empty:
                    reg_temp = registered_df.copy()
                    reg_temp["Month"] = reg_temp["Registered Date"].dt.to_period("M").astype(str)
                    all_months_available.update(reg_temp["Month"].unique())
                if not joined_df.empty:
                    join_temp = joined_df.copy()
                    join_temp["Month"] = join_temp["Joined Date"].dt.to_period("M").astype(str)
                    all_months_available.update(join_temp["Month"].unique())
                all_months_available = sorted(all_months_available)
                
                if all_months_available:
                    month_filter = st.multiselect(
                        "Select Months",
                        options=all_months_available,
                        default=all_months_available
                    )
                else:
                    month_filter = []
                    st.info("No months available")
            
            with col_filter2:
                month_employee_filter = st.multiselect(
                    "Filter by Employee (Month View)",
                    options=sorted(employee_dict.values()),
                    default=sorted(employee_dict.values())
                )
                
                show_all_employees_month = st.checkbox("Show all employees (including 0) in month view", value=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        filtered_registered_df_month = registered_df.copy()
        filtered_joined_df_month = joined_df.copy()
        
        filtered_registered_df_month = filter_by_employees(filtered_registered_df_month, month_employee_filter)
        filtered_joined_df_month = filter_by_employees(filtered_joined_df_month, month_employee_filter)
        
        month_summary = create_month_wise_summary(filtered_registered_df_month, filtered_joined_df_month)
        
        if month_filter and not month_summary.empty:
            month_summary = month_summary[month_summary["Month"].isin(month_filter)]
        
        if not month_summary.empty:
            _mo_bal_sum = month_summary["Balance"].sum()
            render_amount_cards([
                ("📝 Total Registered", f"{month_summary['Registered'].sum():,}", "#1E88E5"),
                ("✅ Total Enrolment Amount", f"₹{format_number(month_summary['Enrolment Amount'].sum(), st.session_state.decimal_places)}", "#43A047"),
                ("⚖️ Total Balance", f"{_mo_bal_sum:+,}", "green" if _mo_bal_sum > 0 else "red" if _mo_bal_sum < 0 else "gray"),
            ])
            st.dataframe(
                month_summary.style
                .background_gradient(subset=["Registered", "Joined"], cmap="Blues")
                .map(color_balance, subset=["Balance"])
                .format({"Enrolment Amount": "{:,.3f}"}),
                use_container_width=True, hide_index=True
            )
            fig = go.Figure(data=[
                go.Bar(name="Registered", x=month_summary["Month"], y=month_summary["Registered"], marker_color="#1E88E5"),
                go.Bar(name="Joined", x=month_summary["Month"], y=month_summary["Joined"], marker_color="#43A047"),
            ])
            fig.update_layout(title="Monthly Registered vs Joined", barmode="group", height=400)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No data available to build a monthly summary.")
        
        st.subheader("👥 Employee-wise Monthly Activity")
        
        if show_all_employees_month:
            all_employees_for_month = set(month_employee_filter) if month_employee_filter else set(employee_dict.values())
            all_months_for_month = set(month_filter) if month_filter else all_months_available
            
            employee_monthly = create_employee_monthly_summary_with_zeros(
                filtered_registered_df_month, 
                filtered_joined_df_month,
                all_employees_for_month,
                all_months_for_month
            )
            
            if not employee_monthly.empty:
                _em_bal_sum = employee_monthly["Balance"].sum()
                render_amount_cards([
                    ("📝 Total Registered", f"{employee_monthly['Registered'].sum():,}", "#1E88E5"),
                    ("✅ Total Enrolment Amount", f"₹{format_number(employee_monthly['Enrolment Amount'].sum(), st.session_state.decimal_places)}", "#43A047"),
                    ("⚖️ Total Balance", f"{_em_bal_sum:+,}", "green" if _em_bal_sum > 0 else "red" if _em_bal_sum < 0 else "gray"),
                ])
                st.dataframe(
                    paginate_dataframe(employee_monthly, key_prefix="month_emp")
                    .style.format({"Enrolment Amount": "{:,.3f}"}),
                    use_container_width=True, hide_index=True
                )
                
                st.subheader("📊 Pivot View - Monthly Activity by Employee")
                
                try:
                    pivot_reg = employee_monthly.pivot(index="Month", columns="Employee Name", values="Registered").fillna(0).astype(int)
                    pivot_join = employee_monthly.pivot(index="Month", columns="Employee Name", values="Joined").fillna(0).astype(int)
                    
                    pivot_combined = pd.DataFrame()
                    for col in pivot_reg.columns:
                        pivot_combined[f"{col} (R/J)"] = pivot_reg[col].astype(str) + " / " + pivot_join[col].astype(str)
                    
                    pivot_combined["Total R"] = pivot_reg.sum(axis=1)
                    pivot_combined["Total J"] = pivot_join.sum(axis=1)
                    pivot_combined["Balance"] = pivot_reg.sum(axis=1) - pivot_join.sum(axis=1)
                    
                    pivot_combined.index = pivot_reg.index
                    
                    st.dataframe(pivot_combined, use_container_width=True)
                except Exception as e:
                    st.warning(f"Could not create pivot view: {e}")
            else:
                st.info("No employee-wise monthly data available.")
        else:
            employee_monthly = create_employee_monthly_summary_with_zeros(
                filtered_registered_df_month, 
                filtered_joined_df_month
            )
            
            if month_filter and not employee_monthly.empty:
                employee_monthly = employee_monthly[employee_monthly["Month"].isin(month_filter)]
            
            if not employee_monthly.empty:
                _ema_bal_sum = employee_monthly["Balance"].sum()
                render_amount_cards([
                    ("📝 Total Registered", f"{employee_monthly['Registered'].sum():,}", "#1E88E5"),
                    ("✅ Total Enrolment Amount", f"₹{format_number(employee_monthly['Enrolment Amount'].sum(), st.session_state.decimal_places)}", "#43A047"),
                    ("⚖️ Total Balance", f"{_ema_bal_sum:+,}", "green" if _ema_bal_sum > 0 else "red" if _ema_bal_sum < 0 else "gray"),
                ])
                st.dataframe(
                    paginate_dataframe(employee_monthly, key_prefix="month_emp_active")
                    .style.format({"Enrolment Amount": "{:,.3f}"}),
                    use_container_width=True, hide_index=True
                )
                
                st.subheader("📊 Pivot View - Monthly Activity by Employee")
                
                try:
                    pivot_reg = employee_monthly.pivot(index="Month", columns="Employee Name", values="Registered").fillna(0).astype(int)
                    pivot_join = employee_monthly.pivot(index="Month", columns="Employee Name", values="Joined").fillna(0).astype(int)
                    
                    pivot_combined = pd.DataFrame()
                    for col in pivot_reg.columns:
                        pivot_combined[f"{col} (R/J)"] = pivot_reg[col].astype(str) + " / " + pivot_join[col].astype(str)
                    
                    pivot_combined["Total R"] = pivot_reg.sum(axis=1)
                    pivot_combined["Total J"] = pivot_join.sum(axis=1)
                    pivot_combined["Balance"] = pivot_reg.sum(axis=1) - pivot_join.sum(axis=1)
                    
                    pivot_combined.index = pivot_reg.index
                    
                    st.dataframe(pivot_combined, use_container_width=True)
                except Exception as e:
                    st.warning(f"Could not create pivot view: {e}")
            else:
                st.info("No employee-wise monthly data available.")

    # ---------------- RAW DATA ----------------
    with tab_raw:
        st.subheader("📄 Raw Data")
        with st.expander("📄 View Raw Registered Data", expanded=False):
            if not registered_df.empty:
                render_amount_cards([("📝 Total Rows", f"{len(registered_df):,}", "#1E88E5")])
                st.dataframe(paginate_dataframe(registered_df, key_prefix="raw_reg"), use_container_width=True, hide_index=True)
            else:
                st.info("No registered data available for the selected filters.")

        with st.expander("✅ View Raw Joined Data", expanded=False):
            if not joined_df.empty:
                render_amount_cards([("✅ Enrolment Amount", format_enrolment_amount(enrolment_amount_sum(joined_df)), "#43A047")])
                st.dataframe(
                    paginate_dataframe(joined_df, key_prefix="raw_join")
                    .style.format({"Enrolment Amount": "{:,.3f}"}),
                    use_container_width=True, hide_index=True
                )
            else:
                st.info("No joined data available for the selected filters.")

        with st.expander("📋 View All Raw Data", expanded=False):
            if not df_filtered.empty:
                render_amount_cards([("📊 Total Rows", f"{len(df_filtered):,}", "#FB8C00")])
                st.dataframe(paginate_dataframe(df_filtered, key_prefix="raw_all"), use_container_width=True, hide_index=True)
            else:
                st.info("No data available.")

        st.divider()
        st.subheader("📊 Additional Statistics")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Registration Statistics**")
            if not registered_df.empty:
                unique_days = len(registered_df["Registered Date"].dt.date.unique())
                st.write(f"Average registrations per day: **{registered_count / unique_days:.1f}**")
                st.write(f"Most active registration day: **{registered_df['Registered Date'].dt.date.mode().iloc[0]}**")
                st.write(f"Earliest registration: **{registered_df['Registered Date'].min().date()}**")
                st.write(f"Latest registration: **{registered_df['Registered Date'].max().date()}**")
            else:
                st.write("No registration data.")
        with c2:
            st.markdown("**Join Statistics**")
            if not joined_df.empty:
                unique_days = len(joined_df["Joined Date"].dt.date.unique())
                st.write(f"Average joins per day: **{joined_count / unique_days:.1f}**")
                st.write(f"Total enrolment amount: **{format_enrolment_amount(enrolment_amount_sum(joined_df))}**")
                st.write(f"Most active join day: **{joined_df['Joined Date'].dt.date.mode().iloc[0]}**")
                st.write(f"Earliest join: **{joined_df['Joined Date'].min().date()}**")
                st.write(f"Latest join: **{joined_df['Joined Date'].max().date()}**")
            else:
                st.write("No join data.")

        if balance_data is not None and not balance_data.empty:
            st.markdown("**⚖️ Balance Statistics**")
            b1, b2, b3 = st.columns(3)
            b1.metric("Total Registered", f"{balance_data['Registered'].sum():,}")
            b2.metric("Total Joined", f"{balance_data['Joined'].sum():,}")
            b3.metric("Overall Balance", f"{(balance_data['Registered'].sum() - balance_data['Joined'].sum()):+,}")

        if st.checkbox("📋 Show Data Validation Details"):
            st.write(f"✅ Total records loaded: {len(df)}")
            st.write(f"✅ Records with registered date: {len(df[df['Registered Date'].notna()])}")
            st.write(f"✅ Records with joined date: {len(df[df['Joined Date'].notna()])}")
            st.write(f"✅ Employees found: {len(df['Employee Code'].unique())}")
            
            future_reg = len(df[df["Registered Date"] > current_date])
            future_join = len(df[df["Joined Date"] > current_date])
            st.write(f"📅 Future registered dates: {future_reg}")
            st.write(f"📅 Future joined dates: {future_join}")
            
            st.write(f"✅ Duplicate Referee Phone rows flagged: {len(duplicates)}")

            anomalies = df[df["Registered Date"].notna() & df["Joined Date"].notna()]
            anomalies = anomalies[anomalies["Registered Date"] > anomalies["Joined Date"]]
            if len(anomalies) > 0:
                st.warning(f"⚠️ {len(anomalies)} record(s) have a registered date after the joined date")
                st.dataframe(anomalies, use_container_width=True, hide_index=True)
            else:
                st.success("✅ No date anomalies found (registered date always ≤ joined date)")

    # ---------------- EXPORT ----------------
    with tab_export:
        st.subheader("⬇️ Download Report")
        
        st.info(f"📊 The Excel report includes formatted tables with colors, conditional formatting, and proper number formatting.")
        st.info(f"📅 Report Date Range: **{date_range_text}**")
        
        export_format = st.selectbox(
            "Select Export Format",
            ["Enhanced Excel (with colors & formatting)", "Standard Excel", "CSV"]
        )

        all_employees_export = set(employee_dict.values())
        all_months_export = set()
        if not registered_df.empty:
            reg_temp = registered_df.copy()
            reg_temp["Month"] = reg_temp["Registered Date"].dt.to_period("M").astype(str)
            all_months_export.update(reg_temp["Month"].unique())
        if not joined_df.empty:
            join_temp = joined_df.copy()
            join_temp["Month"] = join_temp["Joined Date"].dt.to_period("M").astype(str)
            all_months_export.update(join_temp["Month"].unique())
        
        employee_monthly_export = create_employee_monthly_summary_with_zeros(
            registered_df, 
            joined_df,
            all_employees_export,
            all_months_export
        )
        
        all_dates_export = set()
        if not registered_df.empty:
            all_dates_export.update(registered_df["Registered Date"].dt.date.tolist())
        if not joined_df.empty:
            all_dates_export.update(joined_df["Joined Date"].dt.date.tolist())
        
        employee_day_export = create_employee_day_summary_with_zeros(
            registered_df, 
            joined_df,
            all_employees_export,
            all_dates_export
        )

        perf_export = perf_employee_summary.copy() if not perf_employee_summary.empty else pd.DataFrame()
        if not perf_export.empty:
            perf_export["Performance"] = perf_export["Conversion %"].apply(
                lambda x: "Excellent" if x >= 80 else "Good" if x >= 50 else "Average" if x >= 25 else "Needs Improvement"
            )

        inbound_codes_export, outbound_codes_export = split_codes_by_category(employee_dict, selected_codes)
        _, overall_inbound_df = create_reporting_table(overall_registered_df, overall_joined_df, employee_dict, inbound_codes_export, "Overall Details", show_note=False, total_label="INBOUND TOTAL")
        _, overall_outbound_df = create_reporting_table(overall_registered_df, overall_joined_df, employee_dict, outbound_codes_export, "Overall Details", show_note=False, total_label="OUTBOUND TOTAL")
        _, month_inbound_df = create_reporting_table(month_registered_df, month_joined_df, employee_dict, inbound_codes_export, "Current Month", show_note=False, total_label="INBOUND TOTAL")
        _, month_outbound_df = create_reporting_table(month_registered_df, month_joined_df, employee_dict, outbound_codes_export, "Current Month", show_note=False, total_label="OUTBOUND TOTAL")
        _, today_inbound_df = create_reporting_table(today_registered_df, today_joined_df, employee_dict, inbound_codes_export, "Today", show_note=False, total_label="INBOUND TOTAL")
        _, today_outbound_df = create_reporting_table(today_registered_df, today_joined_df, employee_dict, outbound_codes_export, "Today", show_note=False, total_label="OUTBOUND TOTAL")

        sheets = {
            "Raw Data": df_filtered,
            "Employee Summary": employee_display,
            "Performance Summary": perf_export,
            "Registered Day": registered_day if not registered_day.empty else pd.DataFrame(columns=["Date", "Registered"]),
            "Joined Day": joined_day if not joined_day.empty else pd.DataFrame(columns=["Date", "Joined"]),
            "Day Employee": detail_summary if not detail_summary.empty else pd.DataFrame(columns=["Date", "Employee Code", "Employee Name", "Registered", "Joined", "Enrolment Amount", "Balance", "Total"]),
            "Employee Matrix": matrix_summary if not matrix_summary.empty else pd.DataFrame(),
            "Status Particulars": status_display,
            "Filtered Registered": registered_df,
            "Filtered Joined": joined_df,
            "Registered Only": registered_only,
            "Joined Only": joined_only,
            "Combined View": combined_display,
            "Reg Employee Wise": registered_employee,
            "Joined Employee Wise": joined_employee,
            "Month Wise": month_summary if not month_summary.empty else pd.DataFrame(columns=["Month", "Registered", "Joined", "Enrolment Amount", "Balance"]),
            "Employee Monthly": employee_monthly_export if not employee_monthly_export.empty else pd.DataFrame(columns=["Month", "Employee Name", "Registered", "Joined", "Enrolment Amount", "Balance"]),
            "Employee Daily": employee_day_export if not employee_day_export.empty else pd.DataFrame(columns=["Date", "Employee Name", "Registered", "Joined", "Enrolment Amount", "Balance"]),
            "Overall Report": overall_df if not overall_df.empty else pd.DataFrame(columns=["Employee Code", "Employee Name", "Registered", "Joined", "Reg %", "Join %", "Conversion %", "Enrolment Amount", "Enrolment %", "Balance"]),
            "Current Month Report": month_df if not month_df.empty else pd.DataFrame(columns=["Employee Code", "Employee Name", "Registered", "Joined", "Reg %", "Join %", "Join/Reg Ratio %", "Enrolment Amount", "Enrolment %", "Balance"]),
            "Today Report": today_df if not today_df.empty else pd.DataFrame(columns=["Employee Code", "Employee Name", "Registered", "Joined", "Reg %", "Join %", "Join/Reg Ratio %", "Enrolment Amount", "Enrolment %", "Balance"]),
            "Overall - Inbound": overall_inbound_df if not overall_inbound_df.empty else pd.DataFrame(),
            "Overall - Outbound": overall_outbound_df if not overall_outbound_df.empty else pd.DataFrame(),
            "Month - Inbound": month_inbound_df if not month_inbound_df.empty else pd.DataFrame(),
            "Month - Outbound": month_outbound_df if not month_outbound_df.empty else pd.DataFrame(),
            "Today - Inbound": today_inbound_df if not today_inbound_df.empty else pd.DataFrame(),
            "Today - Outbound": today_outbound_df if not today_outbound_df.empty else pd.DataFrame(),
        }

        st.markdown("### 👀 Report Preview")
        if not perf_export.empty:
            preview_cols = [c for c in ["Employee Code", "Employee Name", "Category", "Registered", "Joined", "Reg %", "Join %", "Conversion %", "Enrolment Amount", "Enrolment %", "Balance", "Total", "Performance"] if c in perf_export.columns]
            preview_df = perf_export[preview_cols].copy()
            for pct_col in ["Reg %", "Join %", "Conversion %", "Enrolment %"]:
                if pct_col in preview_df.columns:
                    preview_df[pct_col] = preview_df[pct_col].map(lambda x: f"{x:.1f}%" if pd.notna(x) else "")
            if "Enrolment Amount" in preview_df.columns:
                preview_df["Enrolment Amount"] = preview_df["Enrolment Amount"].map(lambda x: format_enrolment_amount(x) if pd.notna(x) else "")
            st.dataframe(preview_df, use_container_width=True, hide_index=True)
        else:
            st.info("No Performance Summary data available for preview.")

        c1, c2 = st.columns(2)
        with c1:
            date_str = date_range_text.replace(" ", "_").replace("-", "_").replace("to", "_to_")
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            if export_format == "Enhanced Excel (with colors & formatting)":
                excel_data = create_styled_excel_report(sheets, date_range_text)
                file_name = f"Employee_Referral_Report_{date_str}_{timestamp}.xlsx"
            elif export_format == "Standard Excel":
                excel_data = create_enhanced_excel_report(sheets, date_range_text)
                file_name = f"Employee_Referral_Report_{date_str}_{timestamp}.xlsx"
            else:
                import zipfile
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for name, df_sheet in sheets.items():
                        if not df_sheet.empty:
                            csv_data = df_sheet.to_csv(index=False).encode('utf-8')
                            zip_file.writestr(f"{name[:50]}.csv", csv_data)
                zip_buffer.seek(0)
                
                st.download_button(
                    "📥 Download All Reports as ZIP",
                    data=zip_buffer,
                    file_name=f"Employee_Referral_Reports_{date_str}_{timestamp}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                st.stop()
            
            st.download_button(
                "📥 Download Report",
                data=excel_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.caption(f"📅 Report covers: **{date_range_text}**")
        with c2:
            if not employee_display.empty:
                csv_data = employee_display.to_csv(index=False).encode()
                st.download_button(
                    "📥 Download Employee Summary (CSV)",
                    data=csv_data,
                    file_name=f"Employee_Referral_Summary_{date_str}_{timestamp}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

        st.caption("💡 The enhanced Excel report includes:")
        st.caption("• Colored headers and borders • Conditional formatting for balance columns")
        st.caption("• Color-coded positive/negative values • Auto-adjusted column widths")
        st.caption("• Number formatting with thousands separators • Frozen header rows")
        st.caption("• Report Information sheet with date range and summary metrics")
        st.caption("• Overall / Month / Today reports each split into Inbound & Outbound sheets")
        st.caption(f"💰 Enrolment Amounts are displayed with {st.session_state.decimal_places} decimal places")

    # ============================================================
    # FOOTER
    # ============================================================
    st.divider()
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.caption("**Employee Referral Dashboard** | Dynamic Employee Management")
        st.caption(f"📅 Report Period: {date_range_text} | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        st.caption(f"💰 Decimal precision: {st.session_state.decimal_places} places | Rounding mode: {st.session_state.rounding_mode}")

    if st.sidebar.button("🔄 Refresh Dashboard"):
        st.rerun()

if __name__ == "__main__":
    main()