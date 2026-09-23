import streamlit as st
import pandas as pd
from io import BytesIO
from pandas.tseries.offsets import DateOffset
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import logging
import numpy as np
import re

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page configuration
st.set_page_config(
    page_title="Scheme Due Report",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        text-align: center;
    }
    .status-active {
        color: #2ecc71;
        font-weight: bold;
    }
    .status-completed {
        color: #3498db;
        font-weight: bold;
    }
    .status-overdue {
        color: #e74c3c;
        font-weight: bold;
    }
    .skipped-highlight {
        background-color: #ffe6e6;
        font-weight: bold;
    }
    .matched-highlight {
        background-color: #d4edda;
        font-weight: bold;
    }
    .unmatched-highlight {
        background-color: #f8d7da;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# Title
st.markdown('<h1 class="main-header">📊 Scheme Due Report</h1>', unsafe_allow_html=True)

# Initialize session state
if 'report_generated' not in st.session_state:
    st.session_state.report_generated = False
if 'report_data' not in st.session_state:
    st.session_state.report_data = None
if 'matched_data' not in st.session_state:
    st.session_state.matched_data = None
if 'unmatched_data' not in st.session_state:
    st.session_state.unmatched_data = None

def read_file(uploaded_file):
    """Read either Excel or CSV file based on file extension"""
    if uploaded_file is None:
        return None
    
    try:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension in ['xlsx', 'xls']:
            return pd.read_excel(uploaded_file)
        elif file_extension == 'csv':
            # Try to detect delimiter and encoding
            try:
                # Try comma delimiter first
                return pd.read_csv(uploaded_file, encoding='utf-8')
            except:
                try:
                    # Try with different encoding
                    uploaded_file.seek(0)
                    return pd.read_csv(uploaded_file, encoding='latin-1')
                except:
                    try:
                        # Try semicolon delimiter
                        uploaded_file.seek(0)
                        return pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
                    except:
                        # Try tab delimiter
                        uploaded_file.seek(0)
                        return pd.read_csv(uploaded_file, sep='\t', encoding='utf-8')
        else:
            st.error(f"❌ Unsupported file format: {file_extension}. Please upload Excel (.xlsx, .xls) or CSV (.csv) files.")
            return None
    except Exception as e:
        st.error(f"❌ Error reading file {uploaded_file.name}: {str(e)}")
        return None

# File uploaders with better descriptions - Accept all file types
col1, col2, col3 = st.columns(3)

with col1:
    joining_file = st.file_uploader(
        "📁 Upload Scheme Joining Details",
        type=["xlsx", "xls", "csv"],
        help="File should contain columns: Doc No, Date, Installment Amount, Customer, Scheme, Mobile No/Mobileno\n\nSupported formats: Excel (.xlsx, .xls) or CSV (.csv)"
    )

with col2:
    payment_file = st.file_uploader(
        "📁 Upload Scheme Payment Details",
        type=["xlsx", "xls", "csv"],
        help="File should contain columns: Order No, Date, Amount Received\n\nSupported formats: Excel (.xlsx, .xls) or CSV (.csv)"
    )

with col3:
    app_customer_file = st.file_uploader(
        "📁 Upload App Customer Details",
        type=["xlsx", "xls", "csv"],
        help="File should contain columns: Phone Number, Customer Name, etc.\n\nSupported formats: Excel (.xlsx, .xls) or CSV (.csv)"
    )

# Sidebar for configuration
with st.sidebar:
    st.header("⚙️ Configuration")
    
    total_installments = st.number_input(
        "Total Installments",
        min_value=1,
        max_value=36,
        value=11,
        help="Total number of installments for the scheme"
    )
    
    report_date = st.date_input(
        "Report Date",
        value=pd.Timestamp.today(),
        help="Date to use for calculating overdue status"
    )
    
    st.divider()
    
    st.header("🎨 Display Options")
    show_summary = st.checkbox("Show Summary Statistics", value=True)
    show_charts = st.checkbox("Show Charts", value=True)
    show_filters = st.checkbox("Show Filters", value=True)
    show_unmatched = st.checkbox("Show Unmatched Records", value=False)
    
    st.divider()
    
    st.header("📤 Export Options")
    export_format = st.radio(
        "Export Format",
        ["Excel", "CSV"],
        horizontal=True
    )

def validate_data(df, required_cols, file_name):
    """Validate that required columns exist in the dataframe"""
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"❌ Missing columns in {file_name}: {', '.join(missing_cols)}")
        st.info(f"Available columns: {', '.join(df.columns)}")
        return False
    return True

def format_amount(value):
    """Format amount with commas and 2 decimal places"""
    try:
        if pd.isna(value) or value == 0:
            return "0"
        return f"{value:,.0f}"
    except:
        return str(value)

def amount_series_to_numeric(series):
    """Convert formatted or numeric amount values to numbers."""
    return pd.to_numeric(
        series.astype("string").str.replace(",", "", regex=False),
        errors="coerce"
    )

def normalize_passbook(value):
    """Normalize passbook number by removing special characters"""
    if pd.isna(value):
        return ""
    # Convert to string and remove special characters (keep only alphanumeric)
    normalized = re.sub(r'[^a-zA-Z0-9]', '', str(value))
    return normalized.upper().strip()

def clean_mobile_number(value):
    """Clean mobile number - remove decimals and format properly"""
    if pd.isna(value) or value == "":
        return ""
    
    # Convert to string
    mobile_str = str(value).strip()
    
    # If it contains a decimal, split and take the integer part
    if '.' in mobile_str:
        mobile_str = mobile_str.split('.')[0]
    
    # Remove any non-digit characters
    mobile_str = re.sub(r'[^0-9]', '', mobile_str)
    
    # If it starts with 91 and has 12 digits, keep as is
    if mobile_str.startswith('91') and len(mobile_str) == 12:
        return mobile_str
    # If it's 10 digits, return as is
    elif len(mobile_str) == 10:
        return mobile_str
    # If it's more than 10 digits, take last 10
    elif len(mobile_str) > 10:
        return mobile_str[-10:]
    else:
        return mobile_str

def find_mobile_column(df):
    """Find the mobile number column regardless of spelling variations"""
    mobile_variations = [
        'Mobile No', 'Mobileno', 'MobileNo', 'Mobile Number', 
        'Mobilenumber', 'Mobile_No', 'Mobile #', 'Phone', 'Phone No',
        'Phone Number', 'Phonenumber', 'Phone_No'
    ]
    
    for col in df.columns:
        col_clean = col.strip().replace(' ', '').lower()
        for variant in mobile_variations:
            variant_clean = variant.replace(' ', '').lower()
            if col_clean == variant_clean:
                return col
    return None

def get_skipped_months_names(joining_date, installments_paid, skipped_months, total_installments):
    """Get the names of skipped months"""
    if skipped_months <= 0:
        return ""
    
    # Calculate the month where payments stopped
    current_date = pd.Timestamp.now()
    joining_date = pd.to_datetime(joining_date)
    
    # The month after the last paid installment
    next_due_month = joining_date + DateOffset(months=int(installments_paid))
    
    # Get skipped months from next_due_month to current month
    skipped_months_list = []
    current_month = next_due_month
    
    for i in range(skipped_months):
        # Ensure we don't go beyond total installments
        months_from_joining = (current_month.year - joining_date.year) * 12 + (current_month.month - joining_date.month)
        if months_from_joining >= total_installments:
            break
            
        skipped_months_list.append(current_month.strftime("%b-%y"))
        current_month = current_month + DateOffset(months=1)
    
    return ", ".join(skipped_months_list)

@st.cache_data
def process_data(joining_data, payment_data, app_customer_data, total_installments, report_date):
    """Process the data and generate the report with app customer matching"""
    
    try:
        # Create copies to avoid modifying originals
        joining = joining_data.copy()
        payment = payment_data.copy()
        app_customer = app_customer_data.copy() if app_customer_data is not None else None
        
        # Clean column names
        joining.columns = joining.columns.str.strip()
        payment.columns = payment.columns.str.strip()
        if app_customer is not None:
            app_customer.columns = app_customer.columns.str.strip()
        
        # Validate required columns
        required_joining = ["Doc No", "Date", "Installment Amount", "Customer", "Scheme"]
        required_payment = ["Order No", "Date", "Amount Received"]
        
        if not validate_data(joining, required_joining, "Joining Details"):
            return None
        if not validate_data(payment, required_payment, "Payment Details"):
            return None
        
        # CRITICAL: Find the mobile column regardless of name
        mobile_col_name = find_mobile_column(joining)
        
        if mobile_col_name:
            st.info(f"✅ Found mobile column in Joining: '{mobile_col_name}'")
            
            # Rename it to standard "Mobile No"
            joining = joining.rename(columns={mobile_col_name: "Mobile No"})
            
            # CLEAN THE MOBILE NUMBERS - Remove decimals and format properly
            joining["Mobile No"] = joining["Mobile No"].apply(clean_mobile_number)
            
            # Count non-empty mobile numbers
            non_empty = (joining["Mobile No"].astype(str).str.strip() != "").sum()
            st.info(f"📱 Found {non_empty} records with mobile numbers in Joining")
            
        else:
            st.warning("⚠️ No mobile column found in Joining. Available columns: " + ", ".join(joining.columns))
            joining["Mobile No"] = ""
        
        # Process App Customer data if provided
        app_mobile_numbers = set()
        if app_customer is not None:
            app_mobile_col = find_mobile_column(app_customer)
            if app_mobile_col:
                st.info(f"✅ Found mobile column in App Customer: '{app_mobile_col}'")
                app_customer = app_customer.rename(columns={app_mobile_col: "Phone Number"})
                app_customer["Phone Number"] = app_customer["Phone Number"].apply(clean_mobile_number)
                app_mobile_numbers = set(app_customer["Phone Number"].astype(str).str.strip())
                app_mobile_numbers = {num for num in app_mobile_numbers if num != ""}
                st.info(f"📱 Found {len(app_mobile_numbers)} unique mobile numbers in App Customer")
            else:
                st.warning("⚠️ No mobile column found in App Customer. Available columns: " + ", ".join(app_customer.columns))
        
        # Rename columns for processing
        joining = joining.rename(columns={
            "Doc No": "Passbook_Original",
            "Date": "Joining_Date",
            "Installment Amount": "Monthly_Amount"
        })
        
        payment = payment.rename(columns={
            "Order No": "Passbook_Original",
            "Date": "Payment_Date",
            "Amount Received": "Paid_Amount"
        })
        
        # Normalize Passbook numbers to match between files
        joining["Passbook"] = joining["Passbook_Original"].apply(normalize_passbook)
        payment["Passbook"] = payment["Passbook_Original"].apply(normalize_passbook)
        
        # Show warning for unmatched passbooks
        joining_passbooks = set(joining["Passbook"].unique())
        payment_passbooks = set(payment["Passbook"].unique())
        unmatched = payment_passbooks - joining_passbooks
        
        if len(unmatched) > 0:
            st.warning(f"⚠️ Found {len(unmatched)} passbooks in payment file that don't match any joining record. They will be ignored.")
            if len(unmatched) > 0:
                st.info(f"Unmatched passbooks (first 5): {list(unmatched)[:5]}")
        
        # Date formatting
        joining["Joining_Date"] = pd.to_datetime(joining["Joining_Date"], errors="coerce")
        payment["Payment_Date"] = pd.to_datetime(payment["Payment_Date"], errors="coerce")
        
        # Drop rows with invalid dates
        joining = joining.dropna(subset=["Joining_Date"])
        payment = payment.dropna(subset=["Payment_Date"])
        
        # Clean data
        joining["Monthly_Amount"] = pd.to_numeric(joining["Monthly_Amount"], errors="coerce").fillna(0)
        payment["Paid_Amount"] = pd.to_numeric(payment["Paid_Amount"], errors="coerce").fillna(0)
        
        # Remove zero amount payments
        payment = payment[payment["Paid_Amount"] > 0]
        
        # Keep original Passbook for display
        joining["Passbook_Display"] = joining["Passbook_Original"].astype(str).str.strip()
        
        # Handle duplicates - keep first occurrence
        joining = joining.drop_duplicates(subset=["Passbook"], keep="first")
        
        # OPTIMIZATION: Calculate installments paid using vectorized operations
        payment_with_joining = payment.merge(
            joining[["Passbook", "Joining_Date"]], 
            on="Passbook", 
            how="left"
        )
        
        # Drop payments without matching joining record
        payment_with_joining = payment_with_joining.dropna(subset=["Joining_Date"])
        
        if len(payment_with_joining) == 0:
            st.warning("⚠️ No matching payments found after passbook normalization. Please check the passbook numbers.")
        
        # Calculate which installment each payment belongs to
        payment_with_joining["Months_Diff"] = (
            (payment_with_joining["Payment_Date"].dt.year - payment_with_joining["Joining_Date"].dt.year) * 12 +
            (payment_with_joining["Payment_Date"].dt.month - payment_with_joining["Joining_Date"].dt.month)
        )
        
        # Payment is for installment number = months_diff + 1
        payment_with_joining["Installment_Number"] = payment_with_joining["Months_Diff"] + 1
        payment_with_joining["Installment_Number"] = payment_with_joining["Installment_Number"].clip(lower=1, upper=total_installments)
        
        # For each passbook, get the max installment number paid
        max_installment_paid = payment_with_joining.groupby("Passbook")["Installment_Number"].max().reset_index()
        max_installment_paid.columns = ["Passbook", "Max_Installment_Paid"]
        
        # Get unique installments paid
        unique_installments = payment_with_joining.groupby("Passbook")["Installment_Number"].nunique().reset_index()
        unique_installments.columns = ["Passbook", "Unique_Installments_Paid"]
        
        # Get payment summary
        payment_summary = payment_with_joining.groupby("Passbook").agg(
            Payment_Total=("Paid_Amount", "sum"),
            Last_Paid=("Payment_Date", "max")
        ).reset_index()
        
        # Now create final dataframe with all data - KEEP MOBILE NO FROM JOINING
        final = joining.merge(payment_summary, on="Passbook", how="left")
        final = final.merge(max_installment_paid, on="Passbook", how="left")
        final = final.merge(unique_installments, on="Passbook", how="left")
        
        # Fill NaN values
        final["Payment_Total"] = final["Payment_Total"].fillna(0)
        final["Last_Paid"] = final["Last_Paid"].fillna(pd.NaT)
        final["Max_Installment_Paid"] = final["Max_Installment_Paid"].fillna(0).astype(int)
        final["Unique_Installments_Paid"] = final["Unique_Installments_Paid"].fillna(0).astype(int)
        
        # Ensure Mobile No is preserved and clean
        final["Mobile No"] = final["Mobile No"].fillna("")
        final["Mobile No"] = final["Mobile No"].apply(clean_mobile_number)
        
        # Calculate Installments Paid
        final["Installments_Paid"] = final[["Max_Installment_Paid", "Unique_Installments_Paid"]].max(axis=1)
        final["Installments_Paid"] = final["Installments_Paid"].clip(lower=1)
        
        # Calculate Total Paid (include joining amount)
        final["Total_Paid"] = final["Payment_Total"] + final["Monthly_Amount"]
        
        # Calculate other fields
        final["Total_Installments"] = total_installments
        final["Scheme_Value"] = final["Monthly_Amount"] * total_installments
        
        # Check if fully paid
        final["Is_Fully_Paid"] = final["Total_Paid"] >= final["Scheme_Value"]
        
        # Remaining installments
        final["Remaining_Installments"] = np.where(
            final["Is_Fully_Paid"],
            0,
            (final["Total_Installments"] - final["Installments_Paid"]).clip(lower=0)
        )
        
        # Balance amount
        final["Balance_Amount"] = np.where(
            final["Is_Fully_Paid"],
            0,
            (final["Scheme_Value"] - final["Total_Paid"]).clip(lower=0)
        )
        
        # Calculate Final Due Date
        final["Final_Due"] = final["Joining_Date"] + DateOffset(months=total_installments - 1)
        
        # Calculate Maturity Date
        final["Maturity_Date"] = final["Joining_Date"] + DateOffset(months=total_installments) - DateOffset(days=1)
        
        # Calculate Next Due Date
        today = pd.Timestamp(report_date)
        
        def calculate_next_due(row):
            if row["Is_Fully_Paid"]:
                return pd.NaT
            
            next_due = row["Joining_Date"] + DateOffset(months=int(row["Installments_Paid"]))
            
            while next_due < today:
                next_due = next_due + DateOffset(months=1)
                months_from_joining = (next_due.year - row["Joining_Date"].year) * 12 + (next_due.month - row["Joining_Date"].month)
                if months_from_joining >= total_installments:
                    return row["Joining_Date"] + DateOffset(months=total_installments - 1)
            
            return next_due
        
        final["Next_Due"] = final.apply(calculate_next_due, axis=1)
        
        # Calculate Skipped Months
        def calculate_skipped_months(row):
            if row["Is_Fully_Paid"]:
                return 0
            
            months_from_joining = (today.year - row["Joining_Date"].year) * 12 + (today.month - row["Joining_Date"].month)
            expected_installments = months_from_joining + 1
            paid_installments = row["Installments_Paid"]
            
            skipped = expected_installments - paid_installments
            
            if today < row["Next_Due"]:
                skipped = skipped - 1
            
            return max(0, skipped)
        
        final["Skipped_Months"] = final.apply(calculate_skipped_months, axis=1)
        final["Skipped_Months"] = final["Skipped_Months"].clip(lower=0, upper=total_installments)
        final["Skipped_Months"] = np.where(final["Is_Fully_Paid"], 0, final["Skipped_Months"])
        
        # Calculate Skipped Amount
        final["Skipped_Amount"] = final["Skipped_Months"] * final["Monthly_Amount"]
        
        # NEW: Calculate Skipped Months Names
        final["Skipped_Months_Names"] = final.apply(
            lambda row: get_skipped_months_names(
                row["Joining_Date"], 
                row["Installments_Paid"], 
                row["Skipped_Months"], 
                total_installments
            ), 
            axis=1
        )
        
        # Status calculation
        conditions = [
            final["Is_Fully_Paid"] | (final["Installments_Paid"] >= total_installments),
            (final["Skipped_Months"] > 0) & ~final["Is_Fully_Paid"],
            final["Next_Due"] >= today,
        ]
        
        choices = ["Completed", "Overdue", "Active"]
        final["Status"] = np.select(conditions, choices, default="Active")
        
        # Clear completed scheme fields
        final.loc[final["Is_Fully_Paid"], "Next_Due"] = pd.NaT
        final.loc[final["Is_Fully_Paid"], "Skipped_Months"] = 0
        final.loc[final["Is_Fully_Paid"], "Skipped_Amount"] = 0
        final.loc[final["Is_Fully_Paid"], "Skipped_Months_Names"] = ""
        
        # Use Passbook_Display for display
        final["Passbook"] = final["Passbook_Display"]
        
        # CRITICAL: Match with App Customer data
        final["Is_Matched"] = False
        final["In_App_Customer"] = "No"
        
        if app_mobile_numbers:
            # Check if mobile number is in app customer
            final["Is_Matched"] = final["Mobile No"].apply(
                lambda x: str(x).strip() in app_mobile_numbers if x != "" else False
            )
            final["In_App_Customer"] = final["Is_Matched"].apply(lambda x: "Yes" if x else "No")
            
            # Split data into matched and unmatched
            matched_final = final[final["Is_Matched"] == True].copy()
            unmatched_final = final[final["Is_Matched"] == False].copy()
            
            st.info(f"✅ Found {len(matched_final)} matching records with App Customer")
            st.warning(f"⚠️ Found {len(unmatched_final)} records NOT matching with App Customer (will be excluded from main report)")
            
            # Use only matched data for the main report
            final = matched_final
        else:
            st.warning("⚠️ No App Customer data provided or no mobile numbers found. All records will be included.")
        
        # Create report with ALL columns including Mobile No and Skipped_Months_Names
        report = final[[
            "Passbook", 
            "Customer", 
            "Mobile No",
            "Scheme", 
            "Monthly_Amount",
            "Installments_Paid", 
            "Remaining_Installments",
            "Total_Paid", 
            "Balance_Amount", 
            "Scheme_Value",
            "Joining_Date", 
            "Last_Paid", 
            "Next_Due", 
            "Final_Due", 
            "Maturity_Date",
            "Skipped_Months", 
            "Skipped_Months_Names",  # NEW COLUMN
            "Skipped_Amount", 
            "Status",
            "In_App_Customer"
        ]].copy()
        
        # Debug: Check if Mobile No has data
        if len(report) > 0:
            non_empty_mobile = (report["Mobile No"].astype(str).str.strip() != "").sum()
            st.info(f"📱 Final Report: Mobile No column has data for {non_empty_mobile} out of {len(report)} records")
        
        # Format dates
        report["Joining_Date"] = pd.to_datetime(report["Joining_Date"], errors="coerce").dt.strftime("%d-%m-%Y")
        report["Last_Paid"] = pd.to_datetime(report["Last_Paid"], errors="coerce").dt.strftime("%d-%m-%Y")
        report["Final_Due"] = pd.to_datetime(report["Final_Due"], errors="coerce").dt.strftime("%d-%m-%Y")
        report["Maturity_Date"] = pd.to_datetime(report["Maturity_Date"], errors="coerce").dt.strftime("%d-%m-%Y")
        
        # For Next_Due, handle NaT separately
        report["Next_Due"] = pd.to_datetime(report["Next_Due"], errors="coerce")
        report["Next_Due"] = report["Next_Due"].apply(
            lambda x: x.strftime("%d-%m-%Y") if pd.notna(x) else ""
        )
        
        # Format amount columns with commas
        amount_columns = ["Monthly_Amount", "Total_Paid", "Balance_Amount", "Scheme_Value", "Skipped_Amount"]
        for col in amount_columns:
            report[col] = report[col].apply(format_amount)
        
        # Sort by status
        status_order = {"Overdue": 0, "Active": 1, "Completed": 2}
        report["Status_Order"] = report["Status"].map(status_order)
        report = report.sort_values(["Status_Order", "Skipped_Months"], ascending=[True, False])
        report = report.drop("Status_Order", axis=1)
        
        # Add Serial Number
        report.insert(0, "S.No", range(1, len(report) + 1))
        
        # Store unmatched data for display
        if app_mobile_numbers:
            if len(unmatched_final) > 0:
                st.session_state.unmatched_data = unmatched_final.copy()
            else:
                st.session_state.unmatched_data = None
        
        return report
        
    except Exception as e:
        st.error(f"Error processing data: {str(e)}")
        logger.error(f"Data processing error: {str(e)}")
        import traceback
        st.error(f"Details: {traceback.format_exc()}")
        return None

def create_excel(df):
    """Create Excel file with proper formatting"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
        
        # Auto-adjust column widths
        ws = writer.sheets["Report"]
        for column in ws.columns:
            max_len = 0
            for cell in column:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[column[0].column_letter].width = min(max_len + 5, 50)
        
        # Add color coding for status
        from openpyxl.styles import PatternFill
        from openpyxl.utils import get_column_letter
        
        # Find Status column
        status_col = None
        for col_idx, col_name in enumerate(df.columns, 1):
            if col_name == "Status":
                status_col = get_column_letter(col_idx)
                break
        
        if status_col:
            # Add status color coding
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            
            for row in range(2, ws.max_row + 1):
                cell = ws[f"{status_col}{row}"]
                if cell.value == "Completed":
                    cell.fill = green_fill
                elif cell.value == "Active":
                    cell.fill = yellow_fill
                elif cell.value == "Overdue":
                    cell.fill = red_fill
    
    output.seek(0)
    return output

def create_charts(df):
    """Create interactive charts for the report"""
    
    # Convert string amounts back to numeric for charts
    df_chart = df.copy()
    amount_columns = ["Monthly_Amount", "Total_Paid", "Balance_Amount", "Scheme_Value", "Skipped_Amount"]
    for col in amount_columns:
        df_chart[col] = amount_series_to_numeric(df_chart[col])
    
    # Status distribution
    status_counts = df_chart["Status"].value_counts()
    colors = {"Active": "#2ecc71", "Completed": "#3498db", "Overdue": "#e74c3c"}
    
    fig1 = px.pie(
        values=status_counts.values,
        names=status_counts.index,
        title="Scheme Status Distribution",
        color=status_counts.index,
        color_discrete_map=colors,
        hole=0.3
    )
    fig1.update_traces(textposition='inside', textinfo='percent+label')
    
    # Skipped months distribution (excluding completed)
    df_active = df_chart[df_chart["Status"] != "Completed"]
    if len(df_active) > 0:
        fig2 = px.histogram(
            df_active,
            x="Skipped_Months",
            color="Status",
            title="Distribution of Skipped Months (Active & Overdue)",
            color_discrete_map=colors,
            nbins=20
        )
        fig2.update_layout(
            xaxis_title="Skipped Months",
            yaxis_title="Number of Schemes"
        )
    else:
        fig2 = go.Figure()
        fig2.add_annotation(
            text="No Active or Overdue Schemes",
            x=0.5, y=0.5,
            showarrow=False,
            font=dict(size=20)
        )
        fig2.update_layout(
            title="Distribution of Skipped Months (Active & Overdue)",
            xaxis_title="Skipped Months",
            yaxis_title="Number of Schemes"
        )
    
    # Paid vs Skipped Amount
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(
        x=df_chart["Passbook"].head(20),
        y=df_chart["Total_Paid"].head(20),
        name="Total Paid",
        marker_color="#2ecc71"
    ))
    fig3.add_trace(go.Bar(
        x=df_chart["Passbook"].head(20),
        y=df_chart["Skipped_Amount"].head(20),
        name="Skipped Amount",
        marker_color="#e74c3c"
    ))
    fig3.update_layout(
        title="Top 20 Schemes: Paid vs Skipped Amount",
        xaxis_title="Passbook",
        yaxis_title="Amount",
        barmode="group"
    )
    
    return fig1, fig2, fig3

# Main processing logic
if joining_file and payment_file:
    
    try:
        # Read the uploaded files
        with st.spinner("🔄 Reading files..."):
            joining_data = read_file(joining_file)
            payment_data = read_file(payment_file)
            app_customer_data = read_file(app_customer_file) if app_customer_file else None
            
            if joining_data is None or payment_data is None:
                st.error("❌ Failed to read one or more files. Please check the file format.")
                st.stop()
            
            # Show file info
            st.info(f"📄 Joining file: {len(joining_data)} records")
            st.info(f"📄 Payment file: {len(payment_data)} records")
            if app_customer_data is not None:
                st.info(f"📄 App Customer file: {len(app_customer_data)} records")
            else:
                st.warning("⚠️ No App Customer file uploaded. All records will be included without matching.")
            
            # Show columns for debugging
            with st.expander("🔍 View Column Details"):
                st.write("**Joining File Columns:**", joining_data.columns.tolist())
                st.write("**Payment File Columns:**", payment_data.columns.tolist())
                if app_customer_data is not None:
                    st.write("**App Customer File Columns:**", app_customer_data.columns.tolist())
                
                # Show first few rows of data
                st.write("**Joining File - First 5 rows:**")
                st.dataframe(joining_data.head())
        
        with st.spinner("🔄 Processing data..."):
            report = process_data(joining_data, payment_data, app_customer_data, total_installments, report_date)
        
        if report is not None:
            st.session_state.report_generated = True
            st.session_state.report_data = report
            report_count = len(report)
            
            # Success message
            st.success(f"✅ Report Generated Successfully! Total records: {len(report)}")
            
            # Show unmatched records if available
            if hasattr(st.session_state, 'unmatched_data') and st.session_state.unmatched_data is not None:
                unmatched_count = len(st.session_state.unmatched_data)
                st.warning(f"⚠️ {unmatched_count} records were removed because they didn't match with App Customer data")
                
                if show_unmatched:
                    st.subheader("🚫 Unmatched Records (Excluded from Report)")
                    st.caption("These records had mobile numbers not found in App Customer data")
                    
                    unmatched_display = st.session_state.unmatched_data[[
                        "Passbook", "Customer", "Mobile No", "Scheme", "Status"
                    ]].copy()
                    unmatched_display["S.No"] = range(1, len(unmatched_display) + 1)
                    st.dataframe(unmatched_display, use_container_width=True)
            
            # Summary statistics
            if show_summary:
                st.divider()
                st.subheader("📊 Summary Statistics")
                
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    st.metric(
                        "Total Schemes (Matched)",
                        len(report),
                        delta=None,
                        help="Total number of schemes matching with App Customer"
                    )
                
                with col2:
                    active_count = len(report[report["Status"] == "Active"])
                    st.metric(
                        "Active",
                        active_count,
                        delta=f"{active_count/report_count*100:.1f}%" if report_count else "0.0%",
                        delta_color="normal"
                    )
                
                with col3:
                    completed_count = len(report[report["Status"] == "Completed"])
                    st.metric(
                        "Completed",
                        completed_count,
                        delta=f"{completed_count/report_count*100:.1f}%" if report_count else "0.0%",
                        delta_color="normal"
                    )
                
                with col4:
                    overdue_count = len(report[report["Status"] == "Overdue"])
                    st.metric(
                        "Overdue",
                        overdue_count,
                        delta=f"{overdue_count/report_count*100:.1f}%" if report_count else "0.0%",
                        delta_color="inverse"
                    )
                
                with col5:
                    # Total skipped amount (excluding completed)
                    skipped_amounts = report[report["Status"] != "Completed"]["Skipped_Amount"]
                    total_skipped = amount_series_to_numeric(skipped_amounts).sum() if len(skipped_amounts) > 0 else 0
                    st.metric(
                        "Total Skipped Amount",
                        f"₹{total_skipped:,.0f}",
                        delta=None
                    )
                
                # Additional metrics
                col6, col7, col8 = st.columns(3)
                with col6:
                    total_paid = amount_series_to_numeric(report["Total_Paid"]).mean()
                    st.metric("Average Paid per Scheme", f"₹{total_paid:,.0f}")
                
                with col7:
                    # Average skipped months (excluding completed)
                    df_active = report[report["Status"] != "Completed"]
                    avg_skipped = df_active["Skipped_Months"].mean() if len(df_active) > 0 else 0
                    st.metric("Average Skipped Months", f"{avg_skipped:.1f}")
                
                with col8:
                    total_skipped_months = report[report["Status"] != "Completed"]["Skipped_Months"].sum()
                    st.metric("Total Skipped Months", f"{total_skipped_months}")
            
            # Filters
            if show_filters:
                st.divider()
                st.subheader("🔍 Filters")
                
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    status_filter = st.multiselect(
                        "Filter by Status",
                        options=report["Status"].unique(),
                        default=report["Status"].unique()
                    )
                
                with col2:
                    scheme_filter = st.multiselect(
                        "Filter by Scheme",
                        options=report["Scheme"].unique(),
                        default=report["Scheme"].unique()
                    )
                
                with col3:
                    max_skipped_months = int(report["Skipped_Months"].max()) if report_count else 0
                    min_skipped = st.number_input(
                        "Min Skipped Months",
                        min_value=0,
                        max_value=max_skipped_months,
                        value=0
                    )
                
                with col4:
                    search_passbook = st.text_input(
                        "Search Passbook",
                        placeholder="Enter passbook number..."
                    )
                
                with col5:
                    search_mobile = st.text_input(
                        "Search Mobile",
                        placeholder="Enter mobile number..."
                    )
                
                # Apply filters
                filtered_report = report[
                    report["Status"].isin(status_filter) &
                    report["Scheme"].isin(scheme_filter) &
                    (report["Skipped_Months"] >= min_skipped)
                ]
                
                if search_passbook:
                    filtered_report = filtered_report[
                        filtered_report["Passbook"].str.contains(search_passbook, case=False)
                    ]
                
                if search_mobile:
                    filtered_report = filtered_report[
                        filtered_report["Mobile No"].str.contains(search_mobile, case=False)
                    ]
            else:
                filtered_report = report
            
            # Charts
            if show_charts and len(filtered_report) > 0:
                st.divider()
                st.subheader("📈 Visualizations")
                
                fig1, fig2, fig3 = create_charts(filtered_report)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.plotly_chart(fig1, use_container_width=True)
                with col2:
                    st.plotly_chart(fig2, use_container_width=True)
                
                st.plotly_chart(fig3, use_container_width=True)
            
            # Data table with pagination
            st.divider()
            st.subheader("📋 Report Data (Matched with App Customer)")
            
            # Pagination for large datasets
            if len(filtered_report) > 100:
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    page_size = st.selectbox(
                        "Rows per page",
                        [50, 100, 500, 1000],
                        index=1 if len(filtered_report) > 1000 else 0
                    )
                
                total_pages = (len(filtered_report) + page_size - 1) // page_size
                if total_pages > 1:
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        page = st.number_input(
                            "Page",
                            min_value=1,
                            max_value=total_pages,
                            value=1,
                            step=1
                        )
                    
                    start_idx = (page - 1) * page_size
                    end_idx = min(start_idx + page_size, len(filtered_report))
                    display_report = filtered_report.iloc[start_idx:end_idx].copy()
                    
                    # Reset S.No for displayed page
                    display_report["S.No"] = range(start_idx + 1, end_idx + 1)
                    
                    st.caption(f"Showing {start_idx+1} to {end_idx} of {len(filtered_report)} records")
                else:
                    display_report = filtered_report
            else:
                display_report = filtered_report
            
            # Apply status colors to dataframe display
            def color_status(val):
                if val == "Active":
                    return 'background-color: #c6efce; color: #006100'
                elif val == "Completed":
                    return 'background-color: #ffeb9c; color: #9c6500'
                elif val == "Overdue":
                    return 'background-color: #ffc7ce; color: #9c0000'
                return ''
            
            # Highlight skipped months
            def highlight_skipped(val):
                if isinstance(val, (int, float)) and val > 0:
                    return 'background-color: #ffe6e6; font-weight: bold'
                return ''
            
            # Display dataframe with styling
            styled_df = display_report.style.map(color_status, subset=["Status"])
            styled_df = styled_df.map(highlight_skipped, subset=["Skipped_Months"])
            
            st.dataframe(
                styled_df,
                use_container_width=True,
                height=400
            )
            
            # Export options
            st.divider()
            st.subheader("📥 Export Report")
            
            col1, col2, col3 = st.columns([1, 1, 2])
            
            with col1:
                if export_format == "Excel":
                    excel_data = create_excel(filtered_report)
                    st.download_button(
                        "📊 Download Excel Report",
                        excel_data,
                        file_name=f"Scheme_Due_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    csv_data = filtered_report.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        "📄 Download CSV Report",
                        csv_data,
                        file_name=f"Scheme_Due_Report_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
            
            with col2:
                # Export filtered vs all data
                if len(filtered_report) < len(report):
                    if export_format == "Excel":
                        excel_data = create_excel(report)
                        st.download_button(
                            "📊 Download All Data",
                            excel_data,
                            file_name=f"Scheme_Due_Report_All_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        csv_data = report.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            "📄 Download All Data",
                            csv_data,
                            file_name=f"Scheme_Due_Report_All_{datetime.now().strftime('%Y%m%d')}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
            
            with col3:
                # Summary of overdue schemes with skipped months
                overdue_report = report[report["Status"] == "Overdue"][
                    ["S.No", "Passbook", "Customer", "Mobile No", "Scheme", "Skipped_Months", "Skipped_Months_Names", "Skipped_Amount", "Next_Due"]
                ].sort_values("Skipped_Months", ascending=False)
                
                if len(overdue_report) > 0:
                    if st.button("📋 Export Overdue Summary", use_container_width=True):
                        excel_data = create_excel(overdue_report)
                        st.download_button(
                            "📊 Download Overdue Summary",
                            excel_data,
                            file_name=f"Overdue_Summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        logger.error(f"Main processing error: {str(e)}")
        import traceback
        st.error(f"Details: {traceback.format_exc()}")

else:
    # Show instructions when no files are uploaded
    st.info("👈 Please upload Scheme Joining Details, Scheme Payment Details, and App Customer files to generate the report")
    
    # Show sample data format
    with st.expander("📋 Required File Formats (Excel or CSV)"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**Joining Details File**")
            sample_joining = pd.DataFrame({
                "Srno": [1, 2, 3],
                "Date": ["01-11-2025", "01-11-2025", "01-11-2025"],
                "Doc No": ["CDGLGTF-6374", "CDGLGTF-6375", "CDGLGTF-6376"],
                "Scheme": ["GOLD TREE", "GOLD TREE", "GOLD TREE"],
                "Customer": ["V.R JANANI", "k.venkatesh babu", "k.venkatesh babu"],
                "Mobileno": ["9876543210", "8765432109", "7654321098"],
                "Installment Amount": [5000, 10000, 10000]
            })
            st.dataframe(sample_joining)
            st.caption("✅ Column can be 'Mobile No' or 'Mobileno' - the app will find it automatically!")
        
        with col2:
            st.markdown("**Payment Details File**")
            sample_payment = pd.DataFrame({
                "Order No": ["CDGLGTF6374", "CDGLGTF6375", "CDGLGTF6376"],
                "Date": ["01-11-2025", "01-11-2025", "01-11-2025"],
                "Amount Received": [5000, 10000, 10000]
            })
            st.dataframe(sample_payment)
            st.caption("Note: Order No can be without special characters")
        
        with col3:
            st.markdown("**App Customer File**")
            sample_app = pd.DataFrame({
                "Customer Name": ["V.R JANANI", "k.venkatesh babu"],
                "Phone Number": ["9876543210", "8765432109"],
                "Email": ["janani@email.com", "venkatesh@email.com"]
            })
            st.dataframe(sample_app)
            st.caption("✅ Column can be 'Phone Number' or 'Mobile No' - the app will find it automatically!")
        
        st.info("📌 **Supported File Formats:** Excel (.xlsx, .xls) and CSV (.csv)")
        st.info("📌 For CSV files: The app automatically detects delimiter (comma, semicolon, or tab) and encoding")
    
    # Footer
    st.divider()
    st.caption("💡 Tip: Only records matching with App Customer data will be included in the final report")

# Footer
st.divider()
st.caption("🔒 Data is processed locally and not stored on any server")