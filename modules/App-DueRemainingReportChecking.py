import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import warnings
warnings.filterwarnings('ignore')

# Configure Streamlit page
st.set_page_config(page_title="Scheme Payment Tracker", layout="wide")

st.title("📊 Scheme Payment Tracker")
st.markdown("Match Scheme Due Report with Saving Scheme Receipt Register by Passbook ↔ Order No")

# Initialize session state with default values
def initialize_session_state():
    """Initialize all session state variables with default values"""
    defaults = {
        'processed': False,
        'final_df': pd.DataFrame(),
        'total_records': 0,
        'matched_records': 0,
        'same_day': 0,
        'before_paid': 0,
        'after_paid': 0,
        'not_paid': 0,
        'no_due_date': 0,
        'processing_time': 0,
        'total_amount_received': 0,
        'total_overdue_amount': 0,
        'scheme_df': None,
        'receipt_df': None,
        'auto_fill_used': False
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

# Call initialization
initialize_session_state()

def parse_date(date_val):
    """Parse date from various formats"""
    if pd.isna(date_val) or date_val == '' or date_val == 'NA' or date_val == 'N/A' or date_val == 'nan':
        return pd.NaT
    
    if isinstance(date_val, pd.Timestamp):
        return date_val
    
    if isinstance(date_val, datetime):
        return pd.Timestamp(date_val)
    
    # Try different date formats
    date_formats = ['%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%d.%m.%Y', '%b %d, %Y', '%m-%d-%Y', '%m/%d/%Y']
    
    for fmt in date_formats:
        try:
            if isinstance(date_val, str):
                date_val_clean = date_val.strip()
                return pd.Timestamp(datetime.strptime(date_val_clean, fmt))
        except:
            continue
    
    try:
        parsed = pd.to_datetime(date_val)
        if not pd.isna(parsed):
            return parsed
    except:
        pass
    
    return pd.NaT

def calculate_next_due(last_paid_date):
    """
    Calculate the next due date as the 1st of the next month after Last_Paid
    If Last_Paid is 01-07-2026, Next_Due should be 01-08-2026
    """
    if pd.isna(last_paid_date):
        return pd.NaT
    
    # If it's already a datetime, add one month
    if isinstance(last_paid_date, (datetime, pd.Timestamp)):
        # Add one month and set to 1st
        year = last_paid_date.year
        month = last_paid_date.month + 1
        if month > 12:
            month = 1
            year += 1
        return pd.Timestamp(f"{year}-{month:02d}-01")
    
    return pd.NaT

def format_date_for_display(date_val):
    """Format date for display"""
    if pd.isna(date_val) or date_val == '':
        return ''
    if isinstance(date_val, (datetime, pd.Timestamp)):
        return date_val.strftime('%d-%m-%Y')
    return str(date_val)

def load_file(uploaded_file):
    """Load CSV or Excel file"""
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            return df
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")
            return None
    return None

def preprocess_scheme_data(scheme_df, auto_fill=False):
    """
    Preprocess scheme data - fill missing Next_Due dates if auto_fill is enabled
    """
    df = scheme_df.copy()
    
    if auto_fill:
        # Parse Last_Paid dates
        df['Last_Paid_parsed'] = df['Last_Paid'].apply(parse_date)
        
        # Count missing Next_Due values
        missing_count = df['Next_Due'].isna().sum() + (df['Next_Due'] == '').sum() + (df['Next_Due'] == ' ').sum()
        
        if missing_count > 0:
            st.info(f"🔄 Auto-filling {missing_count} missing Next_Due dates based on Last_Paid")
            
            # Fill missing Next_Due with calculated dates
            for idx, row in df.iterrows():
                if pd.isna(row['Next_Due']) or row['Next_Due'] == '' or row['Next_Due'] == ' ':
                    if pd.notna(row['Last_Paid_parsed']):
                        calculated_date = calculate_next_due(row['Last_Paid_parsed'])
                        df.at[idx, 'Next_Due'] = format_date_for_display(calculated_date)
                        df.at[idx, 'Next_Due_auto_filled'] = 'Yes'
                    else:
                        df.at[idx, 'Next_Due_auto_filled'] = 'No (No Last_Paid)'
            df['Next_Due_auto_filled'] = df.get('Next_Due_auto_filled', '')
            
    return df

def process_matching(scheme_df, receipt_df):
    """
    Process matching with the following logic:
    Step 1: Match Scheme_Due_Report.Passbook = SAVING SCHEME RECEIPT REGISTER.Order No
    Step 2: For matched records, compare Scheme_Due_Report.Next_Due with SAVING SCHEME RECEIPT REGISTER.Date
    Step 3: Bring payment details from SAVING SCHEME RECEIPT REGISTER
    """
    
    # CRITICAL: Create a clean copy of the original scheme dataframe
    # This preserves ALL original values including Status and In_App_Customer
    original_df = scheme_df.copy()
    
    # Parse dates
    scheme_df['Next_Due_parsed'] = scheme_df['Next_Due'].apply(parse_date)
    receipt_df['Date_parsed'] = receipt_df['Date'].apply(parse_date)
    
    # Convert Passbook and Order No to string for matching
    scheme_df['Passbook_str'] = scheme_df['Passbook'].astype(str).str.strip()
    receipt_df['Order No_str'] = receipt_df['Order No'].astype(str).str.strip()
    
    # Create results list
    results = []
    
    # Progress bar
    progress_bar = st.progress(0)
    total_records = len(scheme_df)
    
    # Process each scheme record
    for idx, scheme_row in scheme_df.iterrows():
        passbook = scheme_row['Passbook_str']
        next_due = scheme_row['Next_Due_parsed']
        
        # Start with ALL original scheme columns from the ORIGINAL dataframe
        result_row = {}
        
        # Copy ALL columns from the ORIGINAL dataframe - this preserves Status, In_App_Customer, etc.
        for col in original_df.columns:
            result_row[col] = original_df.iloc[idx][col]
        
        # Add auto-fill tracking column if it exists
        if 'Next_Due_auto_filled' in original_df.columns:
            result_row['Next_Due_Auto_Filled'] = original_df.iloc[idx].get('Next_Due_auto_filled', '')
        
        # Initialize new columns with empty values
        result_row['Customer Paid Date'] = ''
        result_row['Payment Status'] = ''  # NEW - payment matching status
        result_row['Same Day'] = ''
        result_row['After Paid'] = ''
        result_row['Before Paid'] = ''
        result_row['Note Paid'] = ''
        result_row['Cash'] = ''
        result_row['Online'] = ''
        result_row['Card Amt'] = ''
        result_row['Cheque Amt'] = ''
        result_row['Neft Amt'] = ''
        result_row['Rtgs Amt'] = ''
        result_row['Vou Ref Amt'] = ''
        result_row['Amount Received'] = ''
        
        # Step 1: Match Passbook with Order No
        matching_receipts = receipt_df[receipt_df['Order No_str'] == passbook]
        
        if len(matching_receipts) > 0:
            # Get the matching receipt (first match)
            receipt_row = matching_receipts.iloc[0]
            receipt_date = receipt_row['Date_parsed']
            
            # Add Customer Paid Date
            if pd.notna(receipt_date):
                result_row['Customer Paid Date'] = format_date_for_display(receipt_date)
            else:
                result_row['Customer Paid Date'] = 'Date Error'
            
            # Add payment details from receipt
            result_row['Cash'] = receipt_row.get('Cash', '')
            result_row['Online'] = receipt_row.get('Online', '')
            result_row['Card Amt'] = receipt_row.get('Card Amt', '')
            result_row['Cheque Amt'] = receipt_row.get('Cheque Amt', '')
            result_row['Neft Amt'] = receipt_row.get('Neft Amt', '')
            result_row['Rtgs Amt'] = receipt_row.get('Rtgs Amt', '')
            result_row['Vou Ref Amt'] = receipt_row.get('Vou Ref Amt', '')
            result_row['Amount Received'] = receipt_row.get('Amount Received', '')
            
            # Step 2: Compare dates and classify
            if pd.isna(next_due):
                # Next_Due is empty or invalid
                result_row['Payment Status'] = 'No Due Date'
                result_row['Same Day'] = ''
                result_row['After Paid'] = ''
                result_row['Before Paid'] = ''
                result_row['Note Paid'] = ''
            elif pd.notna(next_due) and pd.notna(receipt_date):
                days_diff = (receipt_date - next_due).days
                
                # Classify payment status - set the appropriate flag
                if days_diff == 0:
                    result_row['Payment Status'] = 'Same Day'
                    result_row['Same Day'] = 'Yes'
                    result_row['After Paid'] = ''
                    result_row['Before Paid'] = ''
                    result_row['Note Paid'] = ''
                elif days_diff > 0:
                    result_row['Payment Status'] = f'After Paid by {days_diff} days'
                    result_row['Same Day'] = ''
                    result_row['After Paid'] = 'Yes'
                    result_row['Before Paid'] = ''
                    result_row['Note Paid'] = ''
                else:  # days_diff < 0
                    result_row['Payment Status'] = f'Before Paid by {-days_diff} days'
                    result_row['Same Day'] = ''
                    result_row['After Paid'] = ''
                    result_row['Before Paid'] = 'Yes'
                    result_row['Note Paid'] = ''
            else:
                result_row['Payment Status'] = 'Date Error'
                result_row['Same Day'] = ''
                result_row['After Paid'] = ''
                result_row['Before Paid'] = ''
                result_row['Note Paid'] = ''
        else:
            # No matching Order No found - Mark Not Paid
            # IMPORTANT: The original Status column is NOT changed - it stays as "Overdue", "Active", or "Completed"
            result_row['Customer Paid Date'] = 'Not Found'
            result_row['Payment Status'] = 'Not Paid'  # NEW column shows "Not Paid"
            result_row['Same Day'] = ''
            result_row['After Paid'] = ''
            result_row['Before Paid'] = ''
            result_row['Note Paid'] = 'Not Paid'
        
        results.append(result_row)
        
        # Update progress
        progress_bar.progress((idx + 1) / total_records)
    
    progress_bar.empty()
    
    # Convert to DataFrame
    final_df = pd.DataFrame(results)
    
    # Define the EXACT column order for download
    # Original columns first, then new columns
    final_columns = [
        'S.No',
        'Passbook',
        'Customer',
        'Mobile No',
        'Scheme',
        'Monthly_Amount',
        'Installments_Paid',
        'Remaining_Installments',
        'Total_Paid',
        'Balance_Amount',
        'Scheme_Value',
        'Joining_Date',
        'Last_Paid',
        'Next_Due',
        'Customer Paid Date',  # NEW - moved here as requested
        'Final_Due',
        'Maturity_Date',
        'Skipped_Months',
        'Skipped_Amount',
        'Status',  # ORIGINAL Status - UNCHANGED (Overdue, Active, Completed, etc.)
        'In_App_Customer',  # ORIGINAL In_App_Customer - UNCHANGED
        'Payment Status',  # NEW - shows: Same Day, After Paid, Before Paid, Not Paid, No Due Date
        'Same Day',  # NEW - Yes/No
        'After Paid',  # NEW - Yes/No
        'Before Paid',  # NEW - Yes/No
        'Note Paid',  # NEW - Not Paid if no receipt
        'Cash',  # NEW
        'Online',  # NEW
        'Card Amt',  # NEW
        'Cheque Amt',  # NEW
        'Neft Amt',  # NEW
        'Rtgs Amt',  # NEW
        'Vou Ref Amt',  # NEW
        'Amount Received'  # NEW
    ]
    
    # Add auto-fill column if it exists
    if 'Next_Due_Auto_Filled' in final_df.columns:
        final_columns.insert(final_columns.index('Next_Due') + 1, 'Next_Due_Auto_Filled')
    
    # Ensure all columns exist
    for col in final_columns:
        if col not in final_df.columns:
            final_df[col] = ''
    
    # Reorder columns
    final_df = final_df[final_columns]
    
    return final_df

def create_pivot_table(df):
    """Create a consolidated pivot table report"""
    # Create payment status categories
    status_categories = ['Same Day', 'Before Paid', 'After Paid', 'Not Paid', 'No Due Date']
    
    # Create a clean status column for grouping
    df_clean = df.copy()
    df_clean['Payment Status Group'] = df_clean['Payment Status'].apply(
        lambda x: 'Not Paid' if x == 'Not Paid' else
                 'No Due Date' if x == 'No Due Date' else
                 'Same Day' if x == 'Same Day' else
                 'Before Paid' if 'Before Paid' in str(x) else
                 'After Paid' if 'After Paid' in str(x) else
                 'Other'
    )
    
    # Group by Payment Status
    pivot_data = []
    
    for status in status_categories:
        status_df = df_clean[df_clean['Payment Status Group'] == status]
        
        if len(status_df) > 0:
            # Count records
            record_count = len(status_df)
            
            # Sum amounts (convert to numeric, handling errors)
            cash_sum = pd.to_numeric(status_df['Cash'], errors='coerce').sum()
            online_sum = pd.to_numeric(status_df['Online'], errors='coerce').sum()
            card_sum = pd.to_numeric(status_df['Card Amt'], errors='coerce').sum()
            cheque_sum = pd.to_numeric(status_df['Cheque Amt'], errors='coerce').sum()
            neft_sum = pd.to_numeric(status_df['Neft Amt'], errors='coerce').sum()
            rtgs_sum = pd.to_numeric(status_df['Rtgs Amt'], errors='coerce').sum()
            vou_ref_sum = pd.to_numeric(status_df['Vou Ref Amt'], errors='coerce').sum()
            amount_received_sum = pd.to_numeric(status_df['Amount Received'], errors='coerce').sum()
            
            pivot_data.append({
                'Payment Status': status,
                'Records Count': record_count,
                'Same Day': len(status_df[status_df['Same Day'] == 'Yes']),
                'After Paid': len(status_df[status_df['After Paid'] == 'Yes']),
                'Before Paid': len(status_df[status_df['Before Paid'] == 'Yes']),
                'Note Paid': len(status_df[status_df['Note Paid'] == 'Not Paid']),
                'Cash': cash_sum if cash_sum > 0 else 0,
                'Online': online_sum if online_sum > 0 else 0,
                'Card Amt': card_sum if card_sum > 0 else 0,
                'Cheque Amt': cheque_sum if cheque_sum > 0 else 0,
                'Neft Amt': neft_sum if neft_sum > 0 else 0,
                'Rtgs Amt': rtgs_sum if rtgs_sum > 0 else 0,
                'Vou Ref Amt': vou_ref_sum if vou_ref_sum > 0 else 0,
                'Amount Received': amount_received_sum if amount_received_sum > 0 else 0
            })
    
    # Create Grand Total row
    grand_total = {
        'Payment Status': 'GRAND TOTAL',
        'Records Count': len(df_clean),
        'Same Day': len(df_clean[df_clean['Same Day'] == 'Yes']),
        'After Paid': len(df_clean[df_clean['After Paid'] == 'Yes']),
        'Before Paid': len(df_clean[df_clean['Before Paid'] == 'Yes']),
        'Note Paid': len(df_clean[df_clean['Note Paid'] == 'Not Paid']),
        'Cash': pd.to_numeric(df_clean['Cash'], errors='coerce').sum(),
        'Online': pd.to_numeric(df_clean['Online'], errors='coerce').sum(),
        'Card Amt': pd.to_numeric(df_clean['Card Amt'], errors='coerce').sum(),
        'Cheque Amt': pd.to_numeric(df_clean['Cheque Amt'], errors='coerce').sum(),
        'Neft Amt': pd.to_numeric(df_clean['Neft Amt'], errors='coerce').sum(),
        'Rtgs Amt': pd.to_numeric(df_clean['Rtgs Amt'], errors='coerce').sum(),
        'Vou Ref Amt': pd.to_numeric(df_clean['Vou Ref Amt'], errors='coerce').sum(),
        'Amount Received': pd.to_numeric(df_clean['Amount Received'], errors='coerce').sum()
    }
    
    pivot_df = pd.DataFrame(pivot_data)
    
    # Add Grand Total if we have data
    if len(pivot_df) > 0:
        # Format numbers in Grand Total
        grand_total_df = pd.DataFrame([grand_total])
        pivot_df = pd.concat([pivot_df, grand_total_df], ignore_index=True)
    
    return pivot_df

def create_pivot_table_without_skipped_months(df):
    """
    Create a pivot table that ONLY includes records where Skipped_Months > 0
    (i.e., filters out regular monthly payers with zero skipped installments)
    """
    # Filter out records with zero skipped months
    df_filtered = df.copy()
    
    # Convert Skipped_Months to numeric, handling errors
    df_filtered['Skipped_Months_Numeric'] = pd.to_numeric(df_filtered['Skipped_Months'], errors='coerce')
    
    # Keep only records where Skipped_Months > 0 (not regular monthly payers)
    df_filtered = df_filtered[df_filtered['Skipped_Months_Numeric'] > 0]
    
    if len(df_filtered) == 0:
        # Return empty DataFrame if no records with skipped months
        return pd.DataFrame()
    
    # Create payment status categories
    status_categories = ['Same Day', 'Before Paid', 'After Paid', 'Not Paid', 'No Due Date']
    
    # Create a clean status column for grouping
    df_filtered['Payment Status Group'] = df_filtered['Payment Status'].apply(
        lambda x: 'Not Paid' if x == 'Not Paid' else
                 'No Due Date' if x == 'No Due Date' else
                 'Same Day' if x == 'Same Day' else
                 'Before Paid' if 'Before Paid' in str(x) else
                 'After Paid' if 'After Paid' in str(x) else
                 'Other'
    )
    
    # Group by Payment Status
    pivot_data = []
    
    for status in status_categories:
        status_df = df_filtered[df_filtered['Payment Status Group'] == status]
        
        if len(status_df) > 0:
            # Count records
            record_count = len(status_df)
            
            # Sum amounts (convert to numeric, handling errors)
            cash_sum = pd.to_numeric(status_df['Cash'], errors='coerce').sum()
            online_sum = pd.to_numeric(status_df['Online'], errors='coerce').sum()
            card_sum = pd.to_numeric(status_df['Card Amt'], errors='coerce').sum()
            cheque_sum = pd.to_numeric(status_df['Cheque Amt'], errors='coerce').sum()
            neft_sum = pd.to_numeric(status_df['Neft Amt'], errors='coerce').sum()
            rtgs_sum = pd.to_numeric(status_df['Rtgs Amt'], errors='coerce').sum()
            vou_ref_sum = pd.to_numeric(status_df['Vou Ref Amt'], errors='coerce').sum()
            amount_received_sum = pd.to_numeric(status_df['Amount Received'], errors='coerce').sum()
            
            # Calculate average skipped months for this status
            avg_skipped_months = status_df['Skipped_Months_Numeric'].mean()
            
            pivot_data.append({
                'Payment Status': status,
                'Records Count': record_count,
                'Same Day': len(status_df[status_df['Same Day'] == 'Yes']),
                'After Paid': len(status_df[status_df['After Paid'] == 'Yes']),
                'Before Paid': len(status_df[status_df['Before Paid'] == 'Yes']),
                'Note Paid': len(status_df[status_df['Note Paid'] == 'Not Paid']),
                'Avg Skipped Months': round(avg_skipped_months, 2) if pd.notna(avg_skipped_months) else 0,
                'Cash': cash_sum if cash_sum > 0 else 0,
                'Online': online_sum if online_sum > 0 else 0,
                'Card Amt': card_sum if card_sum > 0 else 0,
                'Cheque Amt': cheque_sum if cheque_sum > 0 else 0,
                'Neft Amt': neft_sum if neft_sum > 0 else 0,
                'Rtgs Amt': rtgs_sum if rtgs_sum > 0 else 0,
                'Vou Ref Amt': vou_ref_sum if vou_ref_sum > 0 else 0,
                'Amount Received': amount_received_sum if amount_received_sum > 0 else 0
            })
    
    # Create Grand Total row (only for filtered data)
    grand_total = {
        'Payment Status': 'GRAND TOTAL',
        'Records Count': len(df_filtered),
        'Same Day': len(df_filtered[df_filtered['Same Day'] == 'Yes']),
        'After Paid': len(df_filtered[df_filtered['After Paid'] == 'Yes']),
        'Before Paid': len(df_filtered[df_filtered['Before Paid'] == 'Yes']),
        'Note Paid': len(df_filtered[df_filtered['Note Paid'] == 'Not Paid']),
        'Avg Skipped Months': round(df_filtered['Skipped_Months_Numeric'].mean(), 2),
        'Cash': pd.to_numeric(df_filtered['Cash'], errors='coerce').sum(),
        'Online': pd.to_numeric(df_filtered['Online'], errors='coerce').sum(),
        'Card Amt': pd.to_numeric(df_filtered['Card Amt'], errors='coerce').sum(),
        'Cheque Amt': pd.to_numeric(df_filtered['Cheque Amt'], errors='coerce').sum(),
        'Neft Amt': pd.to_numeric(df_filtered['Neft Amt'], errors='coerce').sum(),
        'Rtgs Amt': pd.to_numeric(df_filtered['Rtgs Amt'], errors='coerce').sum(),
        'Vou Ref Amt': pd.to_numeric(df_filtered['Vou Ref Amt'], errors='coerce').sum(),
        'Amount Received': pd.to_numeric(df_filtered['Amount Received'], errors='coerce').sum()
    }
    
    pivot_df = pd.DataFrame(pivot_data)
    
    # Add Grand Total if we have data
    if len(pivot_df) > 0:
        grand_total_df = pd.DataFrame([grand_total])
        pivot_df = pd.concat([pivot_df, grand_total_df], ignore_index=True)
    
    return pivot_df

def generate_email_summary(df, total_amount_received, total_overdue_amount):
    """
    Generate a professional email summary report
    """
    # Create a copy to avoid modifying original
    df_copy = df.copy()
    
    # Calculate metrics
    total_records = len(df_copy)
    
    # Payment status breakdown
    status_counts = df_copy['Payment Status'].value_counts()
    same_day = len(df_copy[df_copy['Same Day'] == 'Yes'])
    after_paid = len(df_copy[df_copy['After Paid'] == 'Yes'])
    before_paid = len(df_copy[df_copy['Before Paid'] == 'Yes'])
    not_paid = len(df_copy[df_copy['Note Paid'] == 'Not Paid'])
    no_due_date = len(df_copy[df_copy['Payment Status'] == 'No Due Date'])
    
    # Calculate percentages
    same_day_pct = (same_day / total_records * 100) if total_records > 0 else 0
    after_paid_pct = (after_paid / total_records * 100) if total_records > 0 else 0
    before_paid_pct = (before_paid / total_records * 100) if total_records > 0 else 0
    not_paid_pct = (not_paid / total_records * 100) if total_records > 0 else 0
    no_due_date_pct = (no_due_date / total_records * 100) if total_records > 0 else 0
    
    # Payment mode breakdown
    total_cash = pd.to_numeric(df_copy['Cash'], errors='coerce').sum()
    total_online = pd.to_numeric(df_copy['Online'], errors='coerce').sum()
    total_card = pd.to_numeric(df_copy['Card Amt'], errors='coerce').sum()
    total_cheque = pd.to_numeric(df_copy['Cheque Amt'], errors='coerce').sum()
    total_neft = pd.to_numeric(df_copy['Neft Amt'], errors='coerce').sum()
    total_rtgs = pd.to_numeric(df_copy['Rtgs Amt'], errors='coerce').sum()
    total_vou_ref = pd.to_numeric(df_copy['Vou Ref Amt'], errors='coerce').sum()
    
    # Status distribution (original)
    status_dist = df_copy['Status'].value_counts()
    
    # Skipped months analysis - create numeric column
    df_copy['Skipped_Months_Numeric'] = pd.to_numeric(df_copy['Skipped_Months'], errors='coerce')
    total_skipped_months = df_copy['Skipped_Months_Numeric'].sum()
    avg_skipped_months = df_copy['Skipped_Months_Numeric'].mean()
    max_skipped_months = df_copy['Skipped_Months_Numeric'].max()
    customers_with_skipped = len(df_copy[df_copy['Skipped_Months_Numeric'] > 0])
    
    # Top defaulters (highest skipped months)
    top_defaulters = df_copy.nlargest(5, 'Skipped_Months_Numeric')[['Passbook', 'Customer', 'Skipped_Months_Numeric', 'Skipped_Amount', 'Status']]
    
    # Current date
    current_date = datetime.now().strftime('%d-%B-%Y')
    
    # Build email content
    email_content = f"""
{'='*80}
                    SCHEME PAYMENT TRACKING REPORT
{'='*80}

Report Generated: {current_date}
Total Records Processed: {total_records:,}
Processing Time: {st.session_state.processing_time:.2f} seconds

{'='*80}
                        1. PAYMENT STATUS SUMMARY
{'='*80}

Payment Status                    | Count    | Percentage
----------------------------------|----------|------------
✅ Same Day Payment               | {same_day:>6,}  | {same_day_pct:>8.1f}%
📅 Before Paid                   | {before_paid:>6,}  | {before_paid_pct:>8.1f}%
⏰ After Paid                    | {after_paid:>6,}  | {after_paid_pct:>8.1f}%
❌ Not Paid                      | {not_paid:>6,}  | {not_paid_pct:>8.1f}%
⚠️ No Due Date                   | {no_due_date:>6,}  | {no_due_date_pct:>8.1f}%
----------------------------------|----------|------------
TOTAL                             | {total_records:>6,}  | {100.0:>8.1f}%

{'='*80}
                        2. FINANCIAL SUMMARY
{'='*80}

Total Amount Received: ₹{total_amount_received:,.2f}
Total Overdue Amount:  ₹{total_overdue_amount:,.2f}
Collection Efficiency: {((total_amount_received / (total_amount_received + total_overdue_amount)) * 100) if (total_amount_received + total_overdue_amount) > 0 else 0:.1f}%

Payment Mode Breakdown:
  • Cash:        ₹{total_cash:>12,.2f}
  • Online:      ₹{total_online:>12,.2f}
  • Card:        ₹{total_card:>12,.2f}
  • Cheque:      ₹{total_cheque:>12,.2f}
  • NEFT:        ₹{total_neft:>12,.2f}
  • RTGS:        ₹{total_rtgs:>12,.2f}
  • Vou Ref:     ₹{total_vou_ref:>12,.2f}
  --------------------------------------------
  TOTAL:         ₹{(total_cash + total_online + total_card + total_cheque + total_neft + total_rtgs + total_vou_ref):>12,.2f}

{'='*80}
                        3. SKIPPED MONTHS ANALYSIS
{'='*80}

Total Skipped Months:        {total_skipped_months:,.0f}
Average Skipped Months:      {avg_skipped_months:.2f}
Maximum Skipped Months:      {max_skipped_months:,.0f}
Customers with Skipped Months: {customers_with_skipped:,}

{'='*80}
                        4. TOP 5 DEFAULTING CUSTOMERS
{'='*80}

Passbook    | Customer Name          | Skipped Months | Skipped Amount | Status
------------|------------------------|----------------|----------------|---------------
"""
    
    # Add top defaulters
    for _, row in top_defaulters.iterrows():
        passbook = str(row['Passbook'])[:10]
        customer = str(row['Customer'])[:20]
        skipped_months = int(row['Skipped_Months_Numeric']) if pd.notna(row['Skipped_Months_Numeric']) else 0
        skipped_amount = row['Skipped_Amount'] if pd.notna(row['Skipped_Amount']) else 0
        status = str(row['Status'])[:10]
        email_content += f"{passbook:<10} | {customer:<22} | {skipped_months:>14} | {skipped_amount:>14,.0f} | {status:<10}\n"
    
    email_content += f"""
{'='*80}
                        5. ORIGINAL STATUS DISTRIBUTION
{'='*80}

"""
    
    # Add status distribution
    for status, count in status_dist.items():
        pct = (count / total_records * 100) if total_records > 0 else 0
        email_content += f"  • {status:<15}: {count:>6,} ({pct:>5.1f}%)\n"
    
    email_content += f"""
{'='*80}
                        6. RECOMMENDATIONS & ACTION ITEMS
{'='*80}

🔴 HIGH PRIORITY:
"""
    
    # Generate recommendations based on data
    if not_paid > 0:
        email_content += f"  • {not_paid} customers have NOT made any payment. Immediate follow-up required.\n"
    
    if customers_with_skipped > 0:
        email_content += f"  • {customers_with_skipped} customers have skipped payments. Need to investigate reasons.\n"
    
    if total_skipped_months > 12:
        email_content += f"  • High total skipped months ({total_skipped_months:.0f}). Consider implementing stricter collection policies.\n"
    
    email_content += f"""
🟡 MEDIUM PRIORITY:
"""
    
    if after_paid > 0:
        email_content += f"  • {after_paid} customers are paying after due date. Send reminders before due dates.\n"
    
    if avg_skipped_months > 1:
        email_content += f"  • Average skipped months is {avg_skipped_months:.1f}. Suggest offering flexible payment options.\n"
    
    email_content += f"""
🟢 LOW PRIORITY:
"""
    
    if before_paid > 0:
        email_content += f"  • {before_paid} customers are paying before due date. Good practice - consider rewarding.\n"
    
    if same_day > 0:
        email_content += f"  • {same_day} customers are paying on time. Maintain this positive trend.\n"
    
    email_content += f"""
{'='*80}
                        7. KEY INSIGHTS
{'='*80}

1. Payment Discipline: {same_day_pct + before_paid_pct:.1f}% of customers paid on or before due date
2. Payment Gap: {not_paid} customers ({(not_paid/total_records*100):.1f}%) need immediate attention
3. Most Used Payment Mode: """
    
    # Find most used payment mode
    payment_modes = {
        'Cash': total_cash,
        'Online': total_online,
        'Card': total_card,
        'Cheque': total_cheque,
        'NEFT': total_neft,
        'RTGS': total_rtgs,
        'Vou Ref': total_vou_ref
    }
    most_used = max(payment_modes, key=payment_modes.get) if any(payment_modes.values()) else 'N/A'
    email_content += f"{most_used} (₹{payment_modes.get(most_used, 0):,.2f})\n"
    
    # Collection efficiency
    collection_efficiency = ((total_amount_received / (total_amount_received + total_overdue_amount)) * 100) if (total_amount_received + total_overdue_amount) > 0 else 0
    email_content += f"4. Collection Efficiency: {collection_efficiency:.1f}%\n"
    
    if collection_efficiency >= 90:
        email_content += "   ✅ Excellent collection performance! Keep up the good work.\n"
    elif collection_efficiency >= 70:
        email_content += "   ⚠️ Good collection performance. Some improvement areas identified.\n"
    else:
        email_content += "   🔴 Collection efficiency needs significant improvement. Immediate action required.\n"
    
    email_content += f"""
{'='*80}
                    END OF REPORT
{'='*80}

This report is auto-generated by the Scheme Payment Tracker System.
For any queries, please contact the Finance Department.

Generated by: Scheme Payment Tracker v2.0
{'='*80}
"""
    
    return email_content

def export_to_excel_with_headings(final_df, pivot_df, pivot_skipped_df, summary_data, mode_data, status_dist, skipped_df):
    """
    Export data to Excel with proper headings for each sheet
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        # Define styles
        header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell_alignment = Alignment(horizontal='left', vertical='center')
        number_alignment = Alignment(horizontal='right', vertical='center')
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # ============ SHEET 1: DETAILED REPORT ============
        detailed_data = []
        
        # Add title row
        detailed_data.append(['SCHEME PAYMENT TRACKER - DETAILED REPORT'])
        detailed_data.append([f'Report Generated: {datetime.now().strftime("%d-%B-%Y %H:%M:%S")}'])
        detailed_data.append([f'Total Records: {len(final_df):,}'])
        detailed_data.append([])  # Empty row
        
        # Add column headers
        detailed_data.append(final_df.columns.tolist())
        
        # Add data rows
        for _, row in final_df.iterrows():
            detailed_data.append(row.tolist())
        
        # Create DataFrame for the sheet
        detailed_df = pd.DataFrame(detailed_data)
        detailed_df.to_excel(writer, sheet_name='Detailed Report', index=False, header=False)
        
        # Apply formatting
        ws = writer.sheets['Detailed Report']
        
        # Format title rows
        for row in range(1, 4):  # Rows 1-3
            for col in range(1, 4):
                cell = ws.cell(row=row, column=col)
                cell.font = Font(name='Arial', size=12, bold=True)
                cell.alignment = Alignment(horizontal='left', vertical='center')
        
        # Format headers (row 5)
        for col in range(1, len(final_df.columns) + 1):
            cell = ws.cell(row=5, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = border
            ws.column_dimensions[chr(64 + col) if col <= 26 else 'A' + chr(64 + col - 26)].width = 15
        
        # Format data cells
        for row in range(6, len(detailed_data) + 1):
            for col in range(1, len(final_df.columns) + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                # Check if it's a number column
                if col in [6, 7, 8, 9, 10, 11, 18, 19, 27, 28, 29, 30, 31, 32, 33, 34]:  # Amount columns
                    cell.alignment = number_alignment
                else:
                    cell.alignment = cell_alignment
        
        # ============ SHEET 2: PIVOT TABLE - ALL ============
        if len(pivot_df) > 0:
            pivot_data = []
            
            # Add title row
            pivot_data.append(['SCHEME PAYMENT TRACKER - PIVOT TABLE (ALL RECORDS)'])
            pivot_data.append([f'Report Generated: {datetime.now().strftime("%d-%B-%Y %H:%M:%S")}'])
            pivot_data.append([])  # Empty row
            
            # Add column headers
            pivot_data.append(pivot_df.columns.tolist())
            
            # Add data rows
            for _, row in pivot_df.iterrows():
                pivot_data.append(row.tolist())
            
            pivot_df_formatted = pd.DataFrame(pivot_data)
            pivot_df_formatted.to_excel(writer, sheet_name='Pivot Table - All', index=False, header=False)
            
            # Apply formatting
            ws = writer.sheets['Pivot Table - All']
            
            # Format title rows
            for row in range(1, 3):
                for col in range(1, 4):
                    cell = ws.cell(row=row, column=col)
                    cell.font = Font(name='Arial', size=12, bold=True)
                    cell.alignment = Alignment(horizontal='left', vertical='center')
            
            # Format headers (row 4)
            for col in range(1, len(pivot_df.columns) + 1):
                cell = ws.cell(row=4, column=col)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border
                ws.column_dimensions[chr(64 + col) if col <= 26 else 'A' + chr(64 + col - 26)].width = 18
            
            # Format data cells and highlight Grand Total
            for row in range(5, len(pivot_data) + 1):
                for col in range(1, len(pivot_df.columns) + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.border = border
                    
                    # Check if it's a number column
                    if col >= 5:  # Amount columns start from column 5
                        cell.alignment = number_alignment
                    else:
                        cell.alignment = cell_alignment
                    
                    # Highlight Grand Total row
                    if row > 4 and len(pivot_data) > 0:
                        cell_value = ws.cell(row=row, column=1).value
                        if cell_value == 'GRAND TOTAL':
                            cell.font = Font(name='Arial', size=10, bold=True)
                            cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        
        # ============ SHEET 3: PIVOT - SKIPPED ONLY ============
        if len(pivot_skipped_df) > 0:
            skipped_pivot_data = []
            
            # Add title row
            skipped_pivot_data.append(['SCHEME PAYMENT TRACKER - PIVOT (SKIPPED MONTHS ONLY)'])
            skipped_pivot_data.append([f'Report Generated: {datetime.now().strftime("%d-%B-%Y %H:%M:%S")}'])
            skipped_pivot_data.append(['Records with Skipped_Months > 0 only (Regular payers excluded)'])
            skipped_pivot_data.append([])  # Empty row
            
            # Add column headers
            skipped_pivot_data.append(pivot_skipped_df.columns.tolist())
            
            # Add data rows
            for _, row in pivot_skipped_df.iterrows():
                skipped_pivot_data.append(row.tolist())
            
            skipped_pivot_df_formatted = pd.DataFrame(skipped_pivot_data)
            skipped_pivot_df_formatted.to_excel(writer, sheet_name='Pivot - Skipped Only', index=False, header=False)
            
            # Apply formatting
            ws = writer.sheets['Pivot - Skipped Only']
            
            # Format title rows
            for row in range(1, 4):
                for col in range(1, 4):
                    cell = ws.cell(row=row, column=col)
                    cell.font = Font(name='Arial', size=12, bold=True)
                    cell.alignment = Alignment(horizontal='left', vertical='center')
            
            # Format headers (row 5)
            for col in range(1, len(pivot_skipped_df.columns) + 1):
                cell = ws.cell(row=5, column=col)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border
                ws.column_dimensions[chr(64 + col) if col <= 26 else 'A' + chr(64 + col - 26)].width = 18
            
            # Format data cells
            for row in range(6, len(skipped_pivot_data) + 1):
                for col in range(1, len(pivot_skipped_df.columns) + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.border = border
                    
                    # Check if it's a number column
                    if col >= 6:  # Amount columns start from column 6
                        cell.alignment = number_alignment
                    else:
                        cell.alignment = cell_alignment
                    
                    # Highlight Grand Total row
                    if row > 5 and len(skipped_pivot_data) > 0:
                        cell_value = ws.cell(row=row, column=1).value
                        if cell_value == 'GRAND TOTAL':
                            cell.font = Font(name='Arial', size=10, bold=True)
                            cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        
        # ============ SHEET 4: SUMMARY ============
        summary_data_with_heading = []
        summary_data_with_heading.append(['SCHEME PAYMENT TRACKER - EXECUTIVE SUMMARY'])
        summary_data_with_heading.append([f'Report Generated: {datetime.now().strftime("%d-%B-%Y %H:%M:%S")}'])
        summary_data_with_heading.append([])
        summary_data_with_heading.append(['Metric', 'Value'])
        
        # Convert summary_data to DataFrame if it's not already
        if isinstance(summary_data, pd.DataFrame):
            for _, row in summary_data.iterrows():
                summary_data_with_heading.append([row['Metric'], row['Count']])
        else:
            # If it's a dictionary
            for key, value in summary_data.items():
                summary_data_with_heading.append([key, value])
        
        # Add additional metrics
        summary_data_with_heading.append([])
        summary_data_with_heading.append(['ADDITIONAL METRICS', ''])
        summary_data_with_heading.append(['Collection Efficiency', f"{((st.session_state.total_amount_received / (st.session_state.total_amount_received + st.session_state.total_overdue_amount)) * 100) if (st.session_state.total_amount_received + st.session_state.total_overdue_amount) > 0 else 0:.1f}%"])
        summary_data_with_heading.append(['Match Rate', f"{(st.session_state.matched_records / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0:.1f}%"])
        
        summary_df_formatted = pd.DataFrame(summary_data_with_heading)
        summary_df_formatted.to_excel(writer, sheet_name='Summary', index=False, header=False)
        
        # Apply formatting
        ws = writer.sheets['Summary']
        
        # Format title rows
        for row in range(1, 3):
            for col in range(1, 3):
                cell = ws.cell(row=row, column=col)
                cell.font = Font(name='Arial', size=12, bold=True)
                cell.alignment = Alignment(horizontal='left', vertical='center')
        
        # Format headers
        for row in range(4, 5):
            for col in range(1, 3):
                cell = ws.cell(row=row, column=col)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border
        
        # Format data
        for row in range(5, len(summary_data_with_heading) + 1):
            for col in range(1, 3):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                cell.alignment = cell_alignment if col == 1 else number_alignment
                ws.column_dimensions['A'].width = 30
                ws.column_dimensions['B'].width = 20
        
        # ============ SHEET 5: PAYMENT MODES ============
        # Check if mode_data has valid data
        has_mode_data = False
        if isinstance(mode_data, dict):
            if 'Payment Mode' in mode_data and 'Total Amount' in mode_data:
                if len(mode_data['Payment Mode']) > 0 and sum(mode_data['Total Amount']) > 0:
                    has_mode_data = True
        elif isinstance(mode_data, pd.DataFrame):
            if len(mode_data) > 0 and 'Total Amount' in mode_data.columns:
                if mode_data['Total Amount'].sum() > 0:
                    has_mode_data = True
        
        if has_mode_data:
            mode_data_with_heading = []
            mode_data_with_heading.append(['SCHEME PAYMENT TRACKER - PAYMENT MODE BREAKDOWN'])
            mode_data_with_heading.append([f'Report Generated: {datetime.now().strftime("%d-%B-%Y %H:%M:%S")}'])
            mode_data_with_heading.append([])
            mode_data_with_heading.append(['Payment Mode', 'Total Amount (₹)'])
            
            if isinstance(mode_data, dict):
                # Convert dict to list
                for i in range(len(mode_data['Payment Mode'])):
                    if mode_data['Total Amount'][i] > 0:
                        mode_data_with_heading.append([mode_data['Payment Mode'][i], mode_data['Total Amount'][i]])
            else:
                # DataFrame
                for _, row in mode_data.iterrows():
                    if row['Total Amount'] > 0:
                        mode_data_with_heading.append([row['Payment Mode'], row['Total Amount']])
            
            mode_df_formatted = pd.DataFrame(mode_data_with_heading)
            mode_df_formatted.to_excel(writer, sheet_name='Payment Modes', index=False, header=False)
            
            # Apply formatting
            ws = writer.sheets['Payment Modes']
            
            # Format title rows
            for row in range(1, 3):
                for col in range(1, 3):
                    cell = ws.cell(row=row, column=col)
                    cell.font = Font(name='Arial', size=12, bold=True)
                    cell.alignment = Alignment(horizontal='left', vertical='center')
            
            # Format headers
            for row in range(4, 5):
                for col in range(1, 3):
                    cell = ws.cell(row=row, column=col)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment
                    cell.border = border
            
            # Format data
            for row in range(5, len(mode_data_with_heading) + 1):
                for col in range(1, 3):
                    cell = ws.cell(row=row, column=col)
                    cell.border = border
                    cell.alignment = cell_alignment if col == 1 else number_alignment
                    ws.column_dimensions['A'].width = 25
                    ws.column_dimensions['B'].width = 20
        
        # ============ SHEET 6: ORIGINAL STATUS DISTRIBUTION ============
        if len(status_dist) > 0:
            status_data_with_heading = []
            status_data_with_heading.append(['SCHEME PAYMENT TRACKER - ORIGINAL STATUS DISTRIBUTION'])
            status_data_with_heading.append([f'Report Generated: {datetime.now().strftime("%d-%B-%Y %H:%M:%S")}'])
            status_data_with_heading.append([])
            status_data_with_heading.append(['Original Status', 'Count', 'Percentage'])
            
            total_records_status = len(final_df)
            if isinstance(status_dist, pd.Series):
                for status, count in status_dist.items():
                    pct = (count / total_records_status * 100) if total_records_status > 0 else 0
                    status_data_with_heading.append([status, count, f"{pct:.1f}%"])
            else:
                # If it's a dict
                for status, count in status_dist.items():
                    pct = (count / total_records_status * 100) if total_records_status > 0 else 0
                    status_data_with_heading.append([status, count, f"{pct:.1f}%"])
            
            status_df_formatted = pd.DataFrame(status_data_with_heading)
            status_df_formatted.to_excel(writer, sheet_name='Original Status', index=False, header=False)
            
            # Apply formatting
            ws = writer.sheets['Original Status']
            
            # Format title rows
            for row in range(1, 3):
                for col in range(1, 4):
                    cell = ws.cell(row=row, column=col)
                    cell.font = Font(name='Arial', size=12, bold=True)
                    cell.alignment = Alignment(horizontal='left', vertical='center')
            
            # Format headers
            for row in range(4, 5):
                for col in range(1, 4):
                    cell = ws.cell(row=row, column=col)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment
                    cell.border = border
            
            # Format data
            for row in range(5, len(status_data_with_heading) + 1):
                for col in range(1, 4):
                    cell = ws.cell(row=row, column=col)
                    cell.border = border
                    cell.alignment = cell_alignment
                    ws.column_dimensions['A'].width = 25
                    ws.column_dimensions['B'].width = 15
                    ws.column_dimensions['C'].width = 15
        
        # ============ SHEET 7: RECORDS WITH SKIPPED MONTHS ============
        if len(skipped_df) > 0:
            skipped_data_with_heading = []
            skipped_data_with_heading.append(['SCHEME PAYMENT TRACKER - RECORDS WITH SKIPPED MONTHS'])
            skipped_data_with_heading.append([f'Report Generated: {datetime.now().strftime("%d-%B-%Y %H:%M:%S")}'])
            skipped_data_with_heading.append([f'Total Records with Skipped Months: {len(skipped_df):,}'])
            skipped_data_with_heading.append([])
            skipped_data_with_heading.append(skipped_df.columns.tolist())
            
            for _, row in skipped_df.iterrows():
                skipped_data_with_heading.append(row.tolist())
            
            skipped_df_formatted = pd.DataFrame(skipped_data_with_heading)
            skipped_df_formatted.to_excel(writer, sheet_name='Skipped Records', index=False, header=False)
            
            # Apply formatting
            ws = writer.sheets['Skipped Records']
            
            # Format title rows
            for row in range(1, 4):
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.font = Font(name='Arial', size=12, bold=True)
                    cell.alignment = Alignment(horizontal='left', vertical='center')
            
            # Format headers
            for row in range(5, 6):
                for col in range(1, len(skipped_df.columns) + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment
                    cell.border = border
                    ws.column_dimensions[chr(64 + col) if col <= 26 else 'A' + chr(64 + col - 26)].width = 18
            
            # Format data
            for row in range(6, len(skipped_data_with_heading) + 1):
                for col in range(1, len(skipped_df.columns) + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.border = border
                    cell.alignment = cell_alignment
        
        # ============ SHEET 8: REPORT METADATA ============
        metadata_data = [
            ['SCHEME PAYMENT TRACKER - REPORT METADATA'],
            [''],
            ['Report Details:'],
            ['Report Generated', datetime.now().strftime('%d-%B-%Y %H:%M:%S')],
            ['Processing Time', f"{st.session_state.processing_time:.2f} seconds"],
            ['Total Records Processed', st.session_state.total_records],
            ['Auto-Fill Used', 'Yes' if st.session_state.get('auto_fill_used', False) else 'No'],
            [''],
            ['File Information:'],
            ['Scheme Report Columns', ', '.join([col for col in final_df.columns[:20]])],
            ['Receipt Register Columns', 'Order No, Date, Cash, Online, Card Amt, Cheque Amt, Neft Amt, Rtgs Amt, Vou Ref Amt, Amount Received'],
            [''],
            ['Generated By:', 'Scheme Payment Tracker v2.0'],
        ]
        
        metadata_df = pd.DataFrame(metadata_data)
        metadata_df.to_excel(writer, sheet_name='Report Metadata', index=False, header=False)
        
        # Apply formatting
        ws = writer.sheets['Report Metadata']
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 50
        
        for row in range(1, len(metadata_data) + 1):
            for col in range(1, 3):
                cell = ws.cell(row=row, column=col)
                if row in [1, 3, 8, 10, 13]:
                    cell.font = Font(name='Arial', size=11, bold=True)
                cell.alignment = Alignment(horizontal='left', vertical='center')
    
    return output.getvalue()

# File upload section
st.subheader("📄 Upload Files")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Scheme Due Report**")
    scheme_file = st.file_uploader("Upload Scheme Due Report", type=['csv', 'xlsx'], key="scheme")
    
    if scheme_file:
        scheme_df = load_file(scheme_file)
        if scheme_df is not None:
            st.session_state.scheme_df = scheme_df
            st.success(f"✅ Loaded {len(scheme_df)} records")
            
            # Show original Status values
            if 'Status' in scheme_df.columns:
                st.info(f"📊 Original Status values in your Scheme Report:")
                status_counts = scheme_df['Status'].value_counts()
                for status, count in status_counts.items():
                    st.info(f"   - '{status}': {count} records")
            
            with st.expander("Preview Scheme Report (First 5 rows)"):
                st.dataframe(scheme_df.head(5), width="stretch")

with col2:
    st.markdown("**Saving Scheme Receipt Register**")
    receipt_file = st.file_uploader("Upload Saving Scheme Receipt Register", type=['csv', 'xlsx'], key="receipt")
    
    if receipt_file:
        receipt_df = load_file(receipt_file)
        if receipt_df is not None:
            st.session_state.receipt_df = receipt_df
            st.success(f"✅ Loaded {len(receipt_df)} records")
            with st.expander("Preview Receipt Register (First 5 rows)"):
                st.dataframe(receipt_df.head(5), width="stretch")

# Process files when both are uploaded
if scheme_file and receipt_file:
    st.markdown("---")
    st.subheader("🔍 Matching Settings")
    
    # Check if required columns exist
    required_scheme_cols = ['Passbook', 'Next_Due', 'Status']
    required_receipt_cols = ['Order No', 'Date']
    
    missing_scheme_cols = [col for col in required_scheme_cols if col not in scheme_df.columns]
    missing_receipt_cols = [col for col in required_receipt_cols if col not in receipt_df.columns]
    
    if missing_scheme_cols:
        st.error(f"❌ Missing columns in Scheme Report: {', '.join(missing_scheme_cols)}")
        st.info(f"Available columns: {', '.join(scheme_df.columns)}")
        st.stop()
    
    if missing_receipt_cols:
        st.error(f"❌ Missing columns in Receipt Register: {', '.join(missing_receipt_cols)}")
        st.info(f"Available columns: {', '.join(receipt_df.columns)}")
        st.stop()
    
    # Check for missing Next_Due dates and offer auto-fill option
    missing_next_due = scheme_df['Next_Due'].isna().sum() + (scheme_df['Next_Due'] == '').sum() + (scheme_df['Next_Due'] == ' ').sum()
    
    if missing_next_due > 0:
        st.warning(f"⚠️ Found {missing_next_due} records with missing Next_Due dates")
        st.info("💡 Auto-fill will calculate Next_Due as the 1st of the next month after Last_Paid")
        
        # Check if Last_Paid column exists
        if 'Last_Paid' in scheme_df.columns:
            auto_fill = st.checkbox(
                "✅ Auto-fill missing Next_Due dates using Last_Paid",
                value=True,
                help="Calculate Next_Due as 1st of next month after Last_Paid (e.g., Last_Paid=01-07-2026 → Next_Due=01-08-2026)"
            )
        else:
            st.error("❌ 'Last_Paid' column not found in Scheme Report. Cannot auto-fill Next_Due.")
            auto_fill = False
    else:
        auto_fill = False
        st.success("✅ All Next_Due dates are present!")
    
    col1, col2, col3 = st.columns([2, 2, 1])
    
    with col1:
        st.info(f"🔑 Step 1: Match **Passbook** (Scheme) ↔ **Order No** (Receipt)")
        st.info(f"📅 Step 2: Compare **Next_Due** (Scheme) ↔ **Date** (Receipt)")
        st.info(f"💰 Step 3: Bring payment details from Receipt")
    
    with col2:
        st.info(f"📊 Scheme records: {len(scheme_df)}")
        st.info(f"📊 Receipt records: {len(receipt_df)}")
        
        # Count unique values
        unique_passbooks = scheme_df['Passbook'].nunique()
        unique_order_nos = receipt_df['Order No'].nunique()
        st.info(f"📌 Unique Passbooks: {unique_passbooks}")
        st.info(f"📌 Unique Order Nos: {unique_order_nos}")
        
        # Check for empty Next_Due values
        empty_next_due = scheme_df['Next_Due'].isna().sum() + (scheme_df['Next_Due'] == '').sum()
        if empty_next_due > 0:
            st.warning(f"⚠️ {empty_next_due} records have empty Next_Due dates")
        
        # Show original Status values
        status_counts = scheme_df['Status'].value_counts()
        st.info(f"📊 Original Status values in Scheme Report:")
        for status, count in status_counts.items():
            st.info(f"   - '{status}': {count}")
    
    with col3:
        if st.button("🚀 Process Matches", type="primary", width="stretch"):
            if len(scheme_df) == 0:
                st.warning("⚠️ No scheme records found!")
                st.stop()
            
            if len(receipt_df) == 0:
                st.warning("⚠️ No receipt records found!")
                st.stop()
            
            # Preprocess scheme data - auto-fill missing Next_Due if enabled
            if auto_fill and 'Last_Paid' in scheme_df.columns:
                scheme_df = preprocess_scheme_data(scheme_df, auto_fill=True)
                st.session_state.auto_fill_used = True
            else:
                st.session_state.auto_fill_used = False
            
            # Perform matching
            start_time = datetime.now()
            
            with st.spinner(f'Processing {len(scheme_df)} scheme records...'):
                final_df = process_matching(scheme_df, receipt_df)
            
            end_time = datetime.now()
            processing_time = (end_time - start_time).total_seconds()
            
            if len(final_df) > 0:
                # Calculate statistics
                total_records = len(final_df)
                matched_records = len(final_df[
                    (final_df['Same Day'] == 'Yes') | 
                    (final_df['After Paid'] == 'Yes') | 
                    (final_df['Before Paid'] == 'Yes')
                ])
                same_day = len(final_df[final_df['Same Day'] == 'Yes'])
                before_paid = len(final_df[final_df['Before Paid'] == 'Yes'])
                after_paid = len(final_df[final_df['After Paid'] == 'Yes'])
                not_paid = len(final_df[final_df['Note Paid'] == 'Not Paid'])
                no_due_date = len(final_df[final_df['Payment Status'] == 'No Due Date'])
                
                # Calculate total amount received
                total_amount_received = pd.to_numeric(final_df['Amount Received'], errors='coerce').sum()
                
                # Calculate total overdue amount (from Balance_Amount)
                total_overdue_amount = pd.to_numeric(final_df['Balance_Amount'], errors='coerce').sum()
                
                # Store results in session state
                st.session_state.final_df = final_df
                st.session_state.total_records = total_records
                st.session_state.matched_records = matched_records
                st.session_state.same_day = same_day
                st.session_state.before_paid = before_paid
                st.session_state.after_paid = after_paid
                st.session_state.not_paid = not_paid
                st.session_state.no_due_date = no_due_date
                st.session_state.processing_time = processing_time
                st.session_state.total_amount_received = total_amount_received
                st.session_state.total_overdue_amount = total_overdue_amount
                st.session_state.processed = True
                
                st.success(f"✅ Processing complete in {processing_time:.2f} seconds!")
                st.success(f"📊 Results: {matched_records} matched | {not_paid} not paid | {no_due_date} no due date")
                
                if auto_fill:
                    st.info("🔄 Auto-fill was applied to missing Next_Due dates using Last_Paid")

# Display results
if st.session_state.processed and len(st.session_state.final_df) > 0:
    st.markdown("---")
    
    # Summary statistics - FIXED LAYOUT
    st.subheader("📊 Summary Statistics")
    
    # Row 1: Key Metrics (4 columns)
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="📋 Total Records",
            value=f"{st.session_state.total_records:,}",
            delta=None
        )
    
    with col2:
        match_pct = (st.session_state.matched_records / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.metric(
            label="✅ Matched",
            value=f"{st.session_state.matched_records:,}",
            delta=f"{match_pct:.1f}%"
        )
    
    with col3:
        not_paid_pct = (st.session_state.not_paid / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.metric(
            label="❌ Not Paid",
            value=f"{st.session_state.not_paid:,}",
            delta=f"{not_paid_pct:.1f}%",
            delta_color="inverse"
        )
    
    with col4:
        st.metric(
            label="⏱️ Processing Time",
            value=f"{st.session_state.processing_time:.1f}s",
            delta=None
        )
    
    # Row 2: Payment Timing (4 columns)
    col5, col6, col7, col8 = st.columns(4)
    
    with col5:
        same_day_pct = (st.session_state.same_day / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.metric(
            label="📅 Same Day",
            value=f"{st.session_state.same_day:,}",
            delta=f"{same_day_pct:.1f}%",
            delta_color="normal"
        )
    
    with col6:
        before_pct = (st.session_state.before_paid / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.metric(
            label="⬅️ Before Paid",
            value=f"{st.session_state.before_paid:,}",
            delta=f"{before_pct:.1f}%",
            delta_color="normal"
        )
    
    with col7:
        after_pct = (st.session_state.after_paid / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.metric(
            label="➡️ After Paid",
            value=f"{st.session_state.after_paid:,}",
            delta=f"{after_pct:.1f}%",
            delta_color="off"
        )
    
    with col8:
        no_due_pct = (st.session_state.no_due_date / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.metric(
            label="⚠️ No Due Date",
            value=f"{st.session_state.no_due_date:,}",
            delta=f"{no_due_pct:.1f}%",
            delta_color="inverse"
        )
    
    # Row 3: Financial Metrics (4 columns)
    col9, col10, col11, col12 = st.columns(4)
    
    with col9:
        st.metric(
            label="💰 Amount Received",
            value=f"₹{st.session_state.total_amount_received:,.0f}",
            help="Total amount received from all customers"
        )
    
    with col10:
        st.metric(
            label="📋 Overdue Amount",
            value=f"₹{st.session_state.total_overdue_amount:,.0f}",
            help="Total outstanding overdue amount"
        )
    
    with col11:
        total = st.session_state.total_amount_received + st.session_state.total_overdue_amount
        efficiency = (st.session_state.total_amount_received / total * 100) if total > 0 else 0
        st.metric(
            label="📈 Collection Efficiency",
            value=f"{efficiency:.1f}%",
            delta=f"{efficiency - 50:.1f}%" if efficiency > 0 else None,
            help="Percentage of total amount collected"
        )
    
    with col12:
        match_rate = (st.session_state.matched_records / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.metric(
            label="🎯 Match Rate",
            value=f"{match_rate:.1f}%",
            delta=f"{match_rate - 50:.1f}%" if match_rate > 0 else None,
            help="Percentage of records matched"
        )
    
    # Visual Progress Bars for Payment Status
    st.markdown("---")
    st.subheader("📊 Payment Status Distribution")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        pct = (st.session_state.same_day / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.markdown(f"**✅ Same Day**")
        st.progress(pct / 100, text=f"{pct:.1f}%")
        st.caption(f"{st.session_state.same_day:,} customers")
    
    with col2:
        pct = (st.session_state.before_paid / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.markdown(f"**🟢 Before Paid**")
        st.progress(pct / 100, text=f"{pct:.1f}%")
        st.caption(f"{st.session_state.before_paid:,} customers")
    
    with col3:
        pct = (st.session_state.after_paid / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.markdown(f"**🟡 After Paid**")
        st.progress(pct / 100, text=f"{pct:.1f}%")
        st.caption(f"{st.session_state.after_paid:,} customers")
    
    with col4:
        pct = (st.session_state.not_paid / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.markdown(f"**🔴 Not Paid**")
        st.progress(pct / 100, text=f"{pct:.1f}%")
        st.caption(f"{st.session_state.not_paid:,} customers")
    
    with col5:
        pct = (st.session_state.no_due_date / st.session_state.total_records * 100) if st.session_state.total_records > 0 else 0
        st.markdown(f"**⚪ No Due Date**")
        st.progress(pct / 100, text=f"{pct:.1f}%")
        st.caption(f"{st.session_state.no_due_date:,} customers")
    
    # Show auto-fill status
    if st.session_state.get('auto_fill_used', False):
        st.info("🔄 Auto-fill was used to populate missing Next_Due dates based on Last_Paid")
    
    # Email Summary Report Section
    st.markdown("---")
    st.subheader("📧 Email Summary Report")
    
    # Generate email summary
    email_content = generate_email_summary(
        st.session_state.final_df,
        st.session_state.total_amount_received,
        st.session_state.total_overdue_amount
    )
    
    # Display email summary in a text area
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.text_area(
            "📋 Email Report Content",
            email_content,
            height=600,
            help="Copy this content for your email report"
        )
    
    with col2:
        st.markdown("### 📥 Download Options")
        
        # Download as text file
        st.download_button(
            label="📥 Download Report as TXT",
            data=email_content,
            file_name=f"scheme_payment_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain",
            width="stretch"
        )
        
        st.markdown("---")
        
        # Copy to clipboard button (via download)
        st.info("💡 Click download to save the report as a text file.")
        st.info("📧 You can paste this content directly into an email.")
    
    # Expandable section with formatted preview
    with st.expander("📄 Formatted Report Preview"):
        st.markdown("```")
        st.text(email_content)
        st.markdown("```")
    
    # Display Pivot Table Reports
    st.markdown("---")
    st.subheader("📊 Pivot Table Reports")
    
    # Create two tabs for the pivot tables
    tab1, tab2 = st.tabs(["📊 All Records", "📊 Without Zero Skipped Months (Regular Payers Excluded)"])
    
    with tab1:
        st.info("💡 Consolidated summary by Payment Status for ALL records")
        
        # Create and display pivot table
        pivot_df = create_pivot_table(st.session_state.final_df)
        
        if len(pivot_df) > 0:
            # Format the pivot table for display
            display_pivot = pivot_df.copy()
            
            # Format currency columns
            currency_cols = ['Cash', 'Online', 'Card Amt', 'Cheque Amt', 'Neft Amt', 'Rtgs Amt', 'Vou Ref Amt', 'Amount Received']
            for col in currency_cols:
                if col in display_pivot.columns:
                    display_pivot[col] = display_pivot[col].apply(lambda x: f"₹{x:,.2f}" if pd.notna(x) and x > 0 else "₹0.00")
            
            # Highlight Grand Total row
            def highlight_grand_total(row):
                if row['Payment Status'] == 'GRAND TOTAL':
                    return ['background-color: #f0f0f0; font-weight: bold'] * len(row)
                return [''] * len(row)
            
            st.dataframe(
                display_pivot.style.apply(highlight_grand_total, axis=1),
                width="stretch",
                height=400
            )
            
            # Download pivot table
            pivot_csv = pivot_df.to_csv(index=False)
            st.download_button(
                label="📥 Download Pivot Table (All Records) as CSV",
                data=pivot_csv,
                file_name=f"pivot_table_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                width="stretch"
            )
            
            # Visualize pivot data
            st.subheader("📊 Pivot Table Visualizations")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Records by Payment Status**")
                chart_data = pivot_df[pivot_df['Payment Status'] != 'GRAND TOTAL'][['Payment Status', 'Records Count']]
                if len(chart_data) > 0:
                    st.bar_chart(chart_data.set_index('Payment Status'))
            
            with col2:
                st.markdown("**Amount Received by Payment Status**")
                chart_data_amount = pivot_df[pivot_df['Payment Status'] != 'GRAND TOTAL'][['Payment Status', 'Amount Received']]
                if len(chart_data_amount) > 0:
                    st.bar_chart(chart_data_amount.set_index('Payment Status'))
            
            # Additional visualization - Payment Mode Distribution
            st.subheader("💰 Payment Mode Distribution (Total Amounts)")
            
            # Get totals from Grand Total row
            grand_total_row = pivot_df[pivot_df['Payment Status'] == 'GRAND TOTAL']
            if len(grand_total_row) > 0:
                mode_data = {
                    'Payment Mode': ['Cash', 'Online', 'Card', 'Cheque', 'NEFT', 'RTGS', 'Vou Ref'],
                    'Total Amount': [
                        grand_total_row['Cash'].iloc[0],
                        grand_total_row['Online'].iloc[0],
                        grand_total_row['Card Amt'].iloc[0],
                        grand_total_row['Cheque Amt'].iloc[0],
                        grand_total_row['Neft Amt'].iloc[0],
                        grand_total_row['Rtgs Amt'].iloc[0],
                        grand_total_row['Vou Ref Amt'].iloc[0]
                    ]
                }
                mode_df = pd.DataFrame(mode_data)
                mode_df = mode_df[mode_df['Total Amount'] > 0]  # Only show modes with amounts
                
                if len(mode_df) > 0:
                    st.bar_chart(mode_df.set_index('Payment Mode'))
                else:
                    st.info("No payment amounts found in the data")
    
    with tab2:
        st.info("💡 **Excludes regular monthly payers** - Only shows records with Skipped_Months > 0")
        
        # Create pivot table without skipped months
        pivot_skipped_df = create_pivot_table_without_skipped_months(st.session_state.final_df)
        
        if len(pivot_skipped_df) > 0:
            # Show count of excluded records
            total_all = len(st.session_state.final_df)
            total_with_skipped = len(st.session_state.final_df[
                pd.to_numeric(st.session_state.final_df['Skipped_Months'], errors='coerce') > 0
            ])
            total_regular = total_all - total_with_skipped
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Records (All)", total_all)
            with col2:
                st.metric("Records with Skipped Months", total_with_skipped)
            with col3:
                st.metric("Regular Payers (Excluded)", total_regular)
            
            st.markdown("---")
            
            # Format the pivot table for display
            display_pivot_skipped = pivot_skipped_df.copy()
            
            # Format currency columns
            currency_cols = ['Cash', 'Online', 'Card Amt', 'Cheque Amt', 'Neft Amt', 'Rtgs Amt', 'Vou Ref Amt', 'Amount Received']
            for col in currency_cols:
                if col in display_pivot_skipped.columns:
                    display_pivot_skipped[col] = display_pivot_skipped[col].apply(lambda x: f"₹{x:,.2f}" if pd.notna(x) and x > 0 else "₹0.00")
            
            # Highlight Grand Total row
            def highlight_grand_total(row):
                if row['Payment Status'] == 'GRAND TOTAL':
                    return ['background-color: #f0f0f0; font-weight: bold'] * len(row)
                return [''] * len(row)
            
            st.dataframe(
                display_pivot_skipped.style.apply(highlight_grand_total, axis=1),
                width="stretch",
                height=400
            )
            
            # Download pivot table
            pivot_skipped_csv = pivot_skipped_df.to_csv(index=False)
            st.download_button(
                label="📥 Download Pivot Table (Without Zero Skipped Months) as CSV",
                data=pivot_skipped_csv,
                file_name=f"pivot_table_skipped_only_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                width="stretch"
            )
            
            # Visualize pivot data
            st.subheader("📊 Pivot Table Visualizations (Skipped Months Only)")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Records by Payment Status**")
                chart_data = pivot_skipped_df[pivot_skipped_df['Payment Status'] != 'GRAND TOTAL'][['Payment Status', 'Records Count']]
                if len(chart_data) > 0:
                    st.bar_chart(chart_data.set_index('Payment Status'))
            
            with col2:
                st.markdown("**Average Skipped Months by Payment Status**")
                chart_data_skipped = pivot_skipped_df[pivot_skipped_df['Payment Status'] != 'GRAND TOTAL'][['Payment Status', 'Avg Skipped Months']]
                if len(chart_data_skipped) > 0:
                    st.bar_chart(chart_data_skipped.set_index('Payment Status'))
        else:
            st.info("ℹ️ No records found with Skipped_Months > 0. All records are regular monthly payers.")
    
    # Display the final report
    st.subheader("📋 Detailed Report")
    
    # Show column description
    st.info("💡 **Important**: The original 'Status' column (Overdue, Active, Completed, etc.) is PRESERVED and NEVER modified.")
    st.info("💡 **New Columns**: Customer Paid Date | Payment Status | Same Day | After Paid | Before Paid | Note Paid | Cash | Online | Card Amt | Cheque Amt | Neft Amt | Rtgs Amt | Vou Ref Amt | Amount Received")
    
    # Show dataframe
    st.dataframe(st.session_state.final_df, width="stretch", height=500)
    
    # Download options
    st.subheader("💾 Download Options")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Download as CSV
        csv = st.session_state.final_df.to_csv(index=False)
        st.download_button(
            label="📥 Download Detailed Report as CSV",
            data=csv,
            file_name=f"scheme_payment_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            width="stretch"
        )
    
    with col2:
        # Download as Excel with proper headings
        st.info("📝 Generating Excel with proper headings...")
        
        # Prepare data for Excel export
        pivot_df = create_pivot_table(st.session_state.final_df)
        pivot_skipped_df = create_pivot_table_without_skipped_months(st.session_state.final_df)
        
        # Summary data
        summary_data = {
            'Metric': ['Total Records', 'Matched Records', 'Not Paid', 
                      'Same Day Payments', 'Before Paid', 'After Paid', 
                      'No Due Date', 'Processing Time (seconds)'],
            'Count': [
                st.session_state.total_records, 
                st.session_state.matched_records,
                st.session_state.not_paid,
                st.session_state.same_day,
                st.session_state.before_paid,
                st.session_state.after_paid,
                st.session_state.no_due_date,
                f"{st.session_state.processing_time:.2f}"
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        
        # Payment Modes data
        mode_data = {'Payment Mode': [], 'Total Amount': []}
        if len(pivot_df) > 0:
            grand_total_row = pivot_df[pivot_df['Payment Status'] == 'GRAND TOTAL']
            if len(grand_total_row) > 0:
                mode_data = {
                    'Payment Mode': ['Cash', 'Online', 'Card', 'Cheque', 'NEFT', 'RTGS', 'Vou Ref'],
                    'Total Amount': [
                        grand_total_row['Cash'].iloc[0],
                        grand_total_row['Online'].iloc[0],
                        grand_total_row['Card Amt'].iloc[0],
                        grand_total_row['Cheque Amt'].iloc[0],
                        grand_total_row['Neft Amt'].iloc[0],
                        grand_total_row['Rtgs Amt'].iloc[0],
                        grand_total_row['Vou Ref Amt'].iloc[0]
                    ]
                }
        
        # Original Status distribution
        status_dist = st.session_state.final_df['Status'].value_counts()
        
        # Skipped Months distribution
        skipped_df = st.session_state.final_df[['Passbook', 'Customer', 'Skipped_Months', 'Payment Status']].copy()
        skipped_df['Skipped_Months'] = pd.to_numeric(skipped_df['Skipped_Months'], errors='coerce')
        skipped_df = skipped_df[skipped_df['Skipped_Months'] > 0].sort_values('Skipped_Months', ascending=False)
        
        # Generate Excel with proper headings
        excel_data = export_to_excel_with_headings(
            st.session_state.final_df,
            pivot_df,
            pivot_skipped_df,
            summary_df,
            mode_data,
            status_dist,
            skipped_df
        )
        
        st.download_button(
            label="📥 Download Complete Report (Excel with Headings)",
            data=excel_data,
            file_name=f"scheme_payment_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch"
        )
        
        st.success("✅ Excel report generated with proper headings for all sheets!")
    
    with col3:
        # Filter options
        filter_type = st.selectbox(
            "Filter by",
            ["All", "Same Day", "After Paid", "Before Paid", "Not Paid", "No Due Date", "Has Skipped Months"]
        )
        
        if filter_type == "Same Day":
            filtered_df = st.session_state.final_df[st.session_state.final_df['Same Day'] == 'Yes']
        elif filter_type == "After Paid":
            filtered_df = st.session_state.final_df[st.session_state.final_df['After Paid'] == 'Yes']
        elif filter_type == "Before Paid":
            filtered_df = st.session_state.final_df[st.session_state.final_df['Before Paid'] == 'Yes']
        elif filter_type == "Not Paid":
            filtered_df = st.session_state.final_df[st.session_state.final_df['Note Paid'] == 'Not Paid']
        elif filter_type == "No Due Date":
            filtered_df = st.session_state.final_df[st.session_state.final_df['Payment Status'] == 'No Due Date']
        elif filter_type == "Has Skipped Months":
            filtered_df = st.session_state.final_df[
                pd.to_numeric(st.session_state.final_df['Skipped_Months'], errors='coerce') > 0
            ]
        else:
            filtered_df = st.session_state.final_df
        
        if filter_type != "All":
            st.info(f"Showing {len(filtered_df)} records")
            st.dataframe(filtered_df, width="stretch", height=300)
    
    # Search functionality
    st.subheader("🔎 Search Records")
    search_term = st.text_input("Search by Passbook or Customer Name")
    if search_term:
        search_result = st.session_state.final_df[
            st.session_state.final_df['Passbook'].astype(str).str.contains(search_term, case=False, na=False) |
            st.session_state.final_df['Customer'].astype(str).str.contains(search_term, case=False, na=False)
        ]
        if len(search_result) > 0:
            st.dataframe(search_result, width="stretch")
        else:
            st.warning("No records found")

# Help section
with st.expander("ℹ️ How to use this tool"):
    st.markdown("""
    ### What This Tool Does:
    
    **1. Keeps Your Original Scheme Report INTACT**
    - All original columns are preserved
    - The original 'Status' column (Overdue, Active, Completed, etc.) is NEVER modified
    - The original 'In_App_Customer' column is NEVER modified
    
    **2. Auto-Fill Missing Next_Due Dates**
    - If Next_Due is missing, auto-fill using Last_Paid
    - Calculates: Next_Due = 1st of next month after Last_Paid
    - Example: Last_Paid = 01-07-2026 → Next_Due = 01-08-2026
    
    **3. Excel Export with Proper Headings**
    - All 8 sheets have professional headers
    - Formatted with colors and borders
    - Grand Total rows highlighted
    - Number columns right-aligned
    - Report metadata included
    
    **4. Email Summary Report**
    - Professional formatted report ready for email
    - Includes payment status breakdown with percentages
    - Financial summary with payment mode analysis
    - Skipped months analysis and top defaulters
    - Actionable recommendations and key insights
    
    **5. Two Pivot Table Reports**
    - **All Records**: Consolidated summary for ALL records
    - **Without Zero Skipped Months**: Only shows records where Skipped_Months > 0 (excludes regular monthly payers)
    - Both include Grand Total rows and visualizations
    
    ---
    
    ### Excel Export Sheets:
    
    **1. Detailed Report** - Full detailed data with all columns
    **2. Pivot Table - All** - Consolidated summary for all records
    **3. Pivot - Skipped Only** - Only records with Skipped_Months > 0
    **4. Summary** - Key metrics and statistics
    **5. Payment Modes** - Payment mode breakdown
    **6. Original Status** - Distribution of original Status values
    **7. Skipped Records** - Records with skipped months
    **8. Report Metadata** - Report generation details
    """)

st.markdown("---")
st.caption("Built with ❤️ using Streamlit | Original Scheme Report + New Columns Added")