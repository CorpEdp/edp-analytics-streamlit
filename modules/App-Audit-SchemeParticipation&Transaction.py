# app.py
# Customer Scheme Verification Tool
# Run with: streamlit run app.py

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import datetime
from difflib import get_close_matches
import plotly.express as px
import plotly.graph_objects as go
import time
import warnings
warnings.filterwarnings('ignore')

# ============================================================================
# PANDAS CONFIGURATION - Fix for large dataframes
# ============================================================================
# Increase maximum cells for Pandas Styler
pd.set_option("styler.render.max_elements", 1000000)  # Allow up to 1 million cells

# Increase display limits for better performance
pd.set_option('display.max_columns', 50)
pd.set_option('display.max_rows', 100)
pd.set_option('display.width', None)

# ============================================================================
# DATE FORMATTING FUNCTIONS
# ============================================================================

def format_date_columns(df, date_columns=None):
    """Convert datetime columns to date-only format (YYYY-MM-DD)"""
    if df is None or df.empty:
        return df
    
    df_copy = df.copy()
    
    # If no columns specified, find all datetime columns
    if date_columns is None:
        date_columns = df_copy.select_dtypes(include=['datetime64']).columns.tolist()
    
    # Also check for object columns that might contain dates
    for col in df_copy.columns:
        if col not in date_columns and df_copy[col].dtype == 'object':
            try:
                # Try to convert to datetime and check if it works
                test_convert = pd.to_datetime(df_copy[col], errors='coerce')
                if not test_convert.isna().all():
                    date_columns.append(col)
            except:
                pass
    
    for col in date_columns:
        if col in df_copy.columns:
            try:
                # Convert to datetime then extract date
                df_copy[col] = pd.to_datetime(df_copy[col]).dt.date
            except:
                pass
    
    return df_copy

def format_dates_in_dataframe(df):
    """Format all date columns in a dataframe for display"""
    if df is None or df.empty:
        return df
    
    df_copy = df.copy()
    
    # Handle datetime64 columns
    datetime_cols = df_copy.select_dtypes(include=['datetime64']).columns.tolist()
    for col in datetime_cols:
        try:
            df_copy[col] = df_copy[col].dt.date
        except:
            pass
    
    # Handle object columns that might contain dates
    for col in df_copy.columns:
        if df_copy[col].dtype == 'object':
            try:
                # Check if column contains date-like strings
                sample = df_copy[col].dropna().iloc[0] if not df_copy[col].dropna().empty else None
                if sample and isinstance(sample, (str, datetime.datetime, datetime.date)):
                    # Try to convert to date
                    converted = pd.to_datetime(df_copy[col], errors='coerce')
                    if not converted.isna().all():
                        df_copy[col] = converted.dt.date
            except:
                pass
    
    return df_copy

def format_date_for_display(date_val):
    """Format a single date value for display"""
    if date_val is None or pd.isna(date_val):
        return ''
    if isinstance(date_val, (datetime.date, datetime.datetime)):
        return date_val.strftime('%Y-%m-%d')
    try:
        dt = pd.to_datetime(date_val)
        return dt.strftime('%Y-%m-%d')
    except:
        return str(date_val)

# Set page configuration
st.set_page_config(
    page_title="Customer Scheme Verification",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better UI
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #F3F4F6;
        padding: 1rem;
        border-radius: 0.5rem;
        text-align: center;
    }
    .match {
        color: #10B981;
        font-weight: bold;
    }
    .mismatch {
        color: #EF4444;
        font-weight: bold;
    }
    .stMetric {
        background-color: #F9FAFB;
        padding: 0.5rem;
        border-radius: 0.5rem;
    }
    .success-box {
        background-color: #D1FAE5;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #10B981;
    }
    .warning-box {
        background-color: #FEF3C7;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #F59E0B;
    }
    .error-box {
        background-color: #FEE2E2;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #EF4444;
    }
    .stButton > button {
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# ============================================================================
# CONFIGURATION AND SETTINGS
# ============================================================================

def create_config_sidebar():
    """Create sidebar with configurable settings"""
    with st.sidebar:
        st.header("⚙️ Comparison Settings")
        
        # Tolerance settings
        st.subheader("📏 Tolerance Values")
        amount_tolerance = st.number_input(
            "Amount Tolerance (₹)", 
            min_value=0.0, 
            value=0.01,
            step=0.01,
            format="%.2f",
            help="Maximum allowed difference in amount values"
        )
        
        weight_tolerance = st.number_input(
            "Weight Tolerance (g)", 
            min_value=0.0, 
            value=0.0001,
            step=0.0001,
            format="%.4f",
            help="Maximum allowed difference in weight values"
        )
        
        # Comparison options
        st.subheader("🔍 Comparison Options")
        compare_dates = st.checkbox("Compare Dates", value=False, 
                                   help="Compare date fields between participation and transaction records")
        fuzzy_matching = st.checkbox("Use Fuzzy Column Matching", value=True,
                                    help="Automatically match column names that are similar")
        
        # Display options
        st.subheader("📊 Display Options")
        show_matched = st.checkbox("Show Matched Records", value=True)
        records_to_show = st.slider("Records to Preview", 10, 100, 20)
        
        # Advanced options
        st.subheader("🔧 Advanced Options")
        chunk_size = st.number_input(
            "Processing Chunk Size",
            min_value=1000,
            max_value=50000,
            value=10000,
            step=1000,
            help="Number of records to process at once for large files"
        )
        
        # Email notification option
        st.subheader("📧 Email Notifications")
        enable_email = st.checkbox("Enable Email Notifications", value=False)
        if enable_email:
            recipient_email = st.text_input("Recipient Email")
            email_trigger = st.selectbox(
                "Send notification when",
                ["All mismatches found", "Mismatch > 10%", "Mismatch > 20%", "Mismatch > 50%"]
            )
        else:
            recipient_email = None
            email_trigger = None
        
        return {
            'amount_tolerance': amount_tolerance,
            'weight_tolerance': weight_tolerance,
            'compare_dates': compare_dates,
            'fuzzy_matching': fuzzy_matching,
            'show_matched': show_matched,
            'records_to_show': records_to_show,
            'chunk_size': chunk_size,
            'enable_email': enable_email,
            'recipient_email': recipient_email,
            'email_trigger': email_trigger
        }

# ============================================================================
# DATA PROCESSING FUNCTIONS
# ============================================================================

def smart_column_mapping(df, target_columns, cutoff=0.8):
    """Intelligently map columns using fuzzy matching"""
    mapping = {}
    df_columns = [col.strip() for col in df.columns]
    
    for col in df_columns:
        matches = get_close_matches(col, target_columns, n=1, cutoff=cutoff)
        if matches:
            mapping[col] = matches[0]
    
    return mapping

def standardize_columns(df, file_type, use_fuzzy=True):
    """
    Standardize column names for both participation and transaction files
    """
    # Remove leading/trailing spaces
    df.columns = df.columns.str.strip()
    
    # Define column mappings
    participation_mapping = {
        'Scheme Name': 'Scheme Name',
        'Date of joining': 'Date of joining',
        'Customer Name': 'Customer Name',
        'Phone': 'Phone',
        'Passbook Number': 'Passbook Number',
        'Maturity Date': 'Maturity Date',
        'Closed Date': 'Closed Date',
        'Last receipt date': 'Last receipt date',
        'Saved Amount': 'Saved Amount',
        'Benefit Amount': 'Benefit Amount',
        'Overall installment': 'Overall installment',
        'Saved Weight': 'Saved Weight',
        'Benefit Weight': 'Benefit Weight',
        'Status': 'Status_Participation',  # Rename to indicate it's participation status
        'Is Renewal': 'Is Renewal'
    }
    
    transaction_mapping = {
        'Scheme Participation Id': 'Scheme Participation Id',
        'Paid Date': 'Paid Date',
        'Status': 'Status_Transaction',  # Rename to indicate it's transaction status
        'Saved Amount': 'Saved Amount',
        'Reward Amount': 'Reward Amount',
        'Transaction Reference': 'Transaction Reference',
        'Installment number': 'Installment number',
        'Metal Type': 'Metal Type',
        'Metal Rate': 'Metal Rate',
        'Saved Metal Weight': 'Saved Metal Weight',
        'Rewards Metal Weight': 'Rewards Metal Weight',
        'Benefit Metal Amount': 'Benefit Metal Amount',
        'Benefit Metal Weight': 'Benefit Metal Weight',
        'Benefit Metal Percentage': 'Benefit Metal Percentage',
        'Receipt ID': 'Receipt ID',
        'Customer Name': 'Customer Name',
        'Customer Phone Number': 'Customer Phone Number',
        'Passbook number': 'Passbook number',
        'Scheme Name': 'Scheme Name'
    }
    
    # Standardize column names
    if file_type == 'participation':
        target_cols = list(participation_mapping.keys())
        if use_fuzzy:
            mapping = smart_column_mapping(df, target_cols)
            if mapping:
                df = df.rename(columns=mapping)
        
        # Direct mapping for remaining columns
        for col in df.columns:
            if col in participation_mapping:
                df.rename(columns={col: participation_mapping[col]}, inplace=True)
    else:  # transaction
        target_cols = list(transaction_mapping.keys())
        if use_fuzzy:
            mapping = smart_column_mapping(df, target_cols)
            if mapping:
                df = df.rename(columns=mapping)
        
        for col in df.columns:
            if col in transaction_mapping:
                df.rename(columns={col: transaction_mapping[col]}, inplace=True)
    
    return df

def validate_file_structure(df, file_type):
    """Validate that file has minimum required structure"""
    if file_type == 'participation':
        min_cols = ['Scheme Name', 'Customer Name', 'Phone', 'Passbook Number']
        required_cols = ['Scheme Name', 'Customer Name', 'Phone', 'Passbook Number', 
                        'Saved Amount', 'Benefit Amount']
    else:
        min_cols = ['Scheme Name', 'Customer Name', 'Customer Phone Number', 'Passbook number']
        required_cols = ['Scheme Name', 'Customer Name', 'Customer Phone Number', 
                        'Passbook number', 'Saved Amount']
    
    missing_min = [col for col in min_cols if col not in df.columns]
    if missing_min:
        st.error(f"❌ Missing required columns in {file_type} file: {missing_min}")
        st.info("Available columns: " + ", ".join(df.columns))
        return False
    
    missing_req = [col for col in required_cols if col not in df.columns]
    if missing_req:
        st.warning(f"⚠️ Some recommended columns missing: {missing_req}")
    
    return True

def clean_data(df, file_type):
    """
    Clean and prepare data for comparison
    """
    # Remove any leading/trailing spaces from string columns
    for col in df.select_dtypes(include=['object']):
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace(['nan', 'None', 'NULL', ''], np.nan)
    
    # Convert numeric columns
    numeric_cols = ['Saved Amount', 'Benefit Amount', 'Saved Weight', 'Benefit Weight']
    if file_type == 'transaction':
        numeric_cols.extend(['Benefit Metal Amount', 'Benefit Metal Weight', 'Saved Metal Weight'])
    
    for col in numeric_cols:
        if col in df.columns:
            # Remove currency symbols and commas
            df[col] = df[col].astype(str).str.replace('₹', '').str.replace(',', '').str.strip()
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Parse dates - store as datetime64 for processing
    date_columns = ['Date of joining', 'Maturity Date', 'Closed Date', 'Last receipt date', 'Paid Date']
    for col in date_columns:
        if col in df.columns:
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            except:
                pass
    
    return df

def process_large_file(df, file_type, chunk_size=10000, progress_placeholder=None):
    """Process large files with progress tracking"""
    total_rows = len(df)
    
    if total_rows <= chunk_size:
        return clean_data(df, file_type)
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    processed_dfs = []
    
    for i in range(0, total_rows, chunk_size):
        chunk = df.iloc[i:i+chunk_size]
        chunk = clean_data(chunk, file_type)
        processed_dfs.append(chunk)
        
        progress = min((i + chunk_size) / total_rows, 1.0)
        progress_bar.progress(progress)
        status_text.text(f"🔄 Processing {min(i+chunk_size, total_rows):,} of {total_rows:,} records...")
    
    progress_bar.empty()
    status_text.empty()
    
    return pd.concat(processed_dfs, ignore_index=True)

def process_participation(df):
    """
    Process participation file - ensure required columns exist
    """
    required_cols = ['Scheme Name', 'Customer Name', 'Phone', 'Passbook Number']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        st.error(f"❌ Missing required columns in Participation file: {missing_cols}")
        return None
    
    # Standardize phone numbers
    if 'Phone' in df.columns:
        df['Phone'] = df['Phone'].astype(str).str.replace('\+91', '').str.replace('-', '').str.replace(' ', '').str.strip()
    
    # ===== FIX: CONVERT DATES TO DATE-ONLY =====
    date_cols = ['Date of joining', 'Maturity Date', 'Closed Date', 'Last receipt date']
    for col in date_cols:
        if col in df.columns:
            try:
                df[col] = pd.to_datetime(df[col]).dt.date
            except:
                pass
    
    # ===== FIX: ENSURE INDEX IS RESET =====
    df = df.reset_index(drop=True)
    
    return df

def process_transactions(df):
    """
    Process transaction file and aggregate by customer scheme
    """
    required_cols = ['Scheme Name', 'Customer Name', 'Customer Phone Number', 'Passbook number']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        st.error(f"❌ Missing required columns in Transaction file: {missing_cols}")
        return None
    
    # Standardize phone and passbook columns
    df['Phone'] = df['Customer Phone Number'].astype(str).str.replace('\+91', '').str.replace('-', '').str.replace(' ', '').str.strip()
    df['Passbook Number'] = df['Passbook number'].astype(str).str.strip()
    
    # ===== FIX: CONVERT PAID DATE TO DATETIME FOR AGGREGATION =====
    if 'Paid Date' in df.columns:
        try:
            df['Paid Date'] = pd.to_datetime(df['Paid Date'], errors='coerce')
        except Exception as e:
            st.warning(f"Could not convert Paid Date: {str(e)}")
    
    # Ensure numeric columns are properly typed before aggregation
    numeric_cols = ['Saved Amount', 'Benefit Metal Amount', 'Saved Metal Weight', 'Benefit Metal Weight']
    for col in numeric_cols:
        if col in df.columns:
            # Clean the column first
            df[col] = df[col].astype(str).str.replace('₹', '').str.replace(',', '').str.strip()
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Build aggregation dictionary
    agg_dict = {}
    
    # Numeric columns - use 'sum'
    numeric_sum_cols = ['Saved Amount', 'Benefit Metal Amount', 'Saved Metal Weight', 'Benefit Metal Weight']
    for col in numeric_sum_cols:
        if col in df.columns:
            agg_dict[col] = 'sum'
    
    # Installment number - convert to numeric then use 'max'
    if 'Installment number' in df.columns:
        df['Installment number'] = pd.to_numeric(df['Installment number'], errors='coerce')
        agg_dict['Installment number'] = 'max'
    
    # Paid Date - use 'max' (works with datetime64)
    if 'Paid Date' in df.columns:
        agg_dict['Paid Date'] = 'max'
    
    # String columns - use 'last' (works with strings)
    string_cols = ['Status_Transaction', 'Transaction Reference', 'Receipt ID', 'Metal Type']
    for col in string_cols:
        if col in df.columns:
            agg_dict[col] = 'last'
    
    # Only perform aggregation if there are columns to aggregate
    if not agg_dict:
        st.error("❌ No columns available for aggregation in transaction data")
        return None
    
    # Perform grouping with error handling
    try:
        grouped = df.groupby(['Scheme Name', 'Customer Name', 'Phone', 'Passbook Number']).agg(agg_dict).reset_index()
    except Exception as e:
        st.error(f"❌ Error during aggregation: {str(e)}")
        st.info("Please check that numeric columns contain valid numbers and date columns contain valid dates.")
        return None
    
    # Convert Paid Date back to date-only format after aggregation
    if 'Paid Date' in grouped.columns:
        try:
            grouped['Paid Date'] = pd.to_datetime(grouped['Paid Date']).dt.date
        except:
            pass
    
    # Rename columns to match participation format
    rename_map = {
        'Benefit Metal Amount': 'Benefit Amount',
        'Saved Metal Weight': 'Saved Weight',
        'Benefit Metal Weight': 'Benefit Weight'
    }
    grouped.rename(columns=rename_map, inplace=True)
    
    # ===== FIX: ENSURE INDEX IS RESET =====
    grouped = grouped.reset_index(drop=True)
    
    return grouped

# ============================================================================
# COMPARISON FUNCTIONS
# ============================================================================

def detailed_comparison(participation, transaction, config):
    """Enhanced comparison with detailed mismatch categories"""
    # Merge on key fields
    merged = participation.merge(
        transaction,
        on=['Scheme Name', 'Customer Name', 'Phone', 'Passbook Number'],
        how='outer',
        suffixes=('_Participation', '_Transaction'),
        indicator=True
    )
    
    # ===== FIX: CORRECTLY PRESERVE DATE COLUMNS WITH DATE-ONLY FORMAT =====
    date_cols = ['Date of joining', 'Paid Date', 'Maturity Date', 'Closed Date', 'Last receipt date']
    for col in date_cols:
        # From participation - use reindex to align
        if col in participation.columns:
            # Ensure dates are in date-only format
            participation_series = participation[col].reindex(merged.index, fill_value=pd.NA)
            merged[f'{col}_Participation'] = participation_series
        else:
            merged[f'{col}_Participation'] = pd.NA
        
        # From transaction - use reindex to align
        if col in transaction.columns:
            transaction_series = transaction[col].reindex(merged.index, fill_value=pd.NA)
            merged[f'{col}_Transaction'] = transaction_series
        else:
            merged[f'{col}_Transaction'] = pd.NA
    
    # Initialize comparison columns
    merged['Amount Match'] = 'NO'
    merged['Weight Match'] = 'NO'
    merged['Benefit Match'] = 'NO'
    merged['Date Match'] = 'NO'
    merged['Final Result'] = 'MISMATCH'
    merged['Mismatch Category'] = 'No Match Found'
    merged['Match Score'] = 0.0
    
    # Compare only where both exist
    both_exist = merged['_merge'] == 'both'
    
    if both_exist.any():
        # Compare Amount
        merged.loc[both_exist, 'Amount Match'] = np.where(
            np.isclose(
                merged.loc[both_exist, 'Saved Amount_Participation'],
                merged.loc[both_exist, 'Saved Amount_Transaction'],
                rtol=config['amount_tolerance'] / 100 if config['amount_tolerance'] > 0 else 0.01
            ),
            'YES', 'NO'
        )
        
        # Calculate amount difference
        merged.loc[both_exist, 'Amount Diff'] = np.abs(
            merged.loc[both_exist, 'Saved Amount_Participation'] - 
            merged.loc[both_exist, 'Saved Amount_Transaction']
        )
        
        # Compare Weight
        merged.loc[both_exist, 'Weight Match'] = np.where(
            np.isclose(
                merged.loc[both_exist, 'Saved Weight_Participation'],
                merged.loc[both_exist, 'Saved Weight_Transaction'],
                rtol=config['weight_tolerance']
            ),
            'YES', 'NO'
        )
        
        # Compare Benefit Amount
        merged.loc[both_exist, 'Benefit Match'] = np.where(
            np.isclose(
                merged.loc[both_exist, 'Benefit Amount_Participation'],
                merged.loc[both_exist, 'Benefit Amount_Transaction'],
                rtol=config['amount_tolerance'] / 100 if config['amount_tolerance'] > 0 else 0.01
            ),
            'YES', 'NO'
        )
        
        # Compare Dates if enabled
        if config['compare_dates']:
            if 'Paid Date_Transaction' in merged.columns and 'Date of joining_Participation' in merged.columns:
                # Convert to datetime if needed
                paid_date = pd.to_datetime(merged.loc[both_exist, 'Paid Date_Transaction'], errors='coerce')
                join_date = pd.to_datetime(merged.loc[both_exist, 'Date of joining_Participation'], errors='coerce')
                
                # Calculate date difference
                date_diff = (paid_date - join_date).dt.days
                merged.loc[both_exist, 'Date Match'] = np.where(
                    date_diff.abs() <= 30,
                    'YES', 'NO'
                )
        
        # Calculate match score (percentage of fields matched)
        match_fields = ['Amount Match', 'Weight Match', 'Benefit Match']
        if config['compare_dates']:
            match_fields.append('Date Match')
        
        for idx in merged[both_exist].index:
            matches = sum([1 for field in match_fields if merged.loc[idx, field] == 'YES'])
            total = len(match_fields)
            merged.loc[idx, 'Match Score'] = (matches / total * 100) if total > 0 else 0
        
        # Determine final result (Status is NOT compared as it has different business meanings)
        merged.loc[both_exist, 'Final Result'] = np.where(
            (merged.loc[both_exist, 'Amount Match'] == 'YES') &
            (merged.loc[both_exist, 'Weight Match'] == 'YES') &
            (merged.loc[both_exist, 'Benefit Match'] == 'YES'),
            'MATCH', 'MISMATCH'
        )
        
        # Categorize mismatches
        conditions = [
            (merged['Amount Match'] == 'NO') & (merged['Weight Match'] == 'YES') & (merged['Benefit Match'] == 'YES'),
            (merged['Amount Match'] == 'YES') & (merged['Weight Match'] == 'NO') & (merged['Benefit Match'] == 'YES'),
            (merged['Amount Match'] == 'YES') & (merged['Weight Match'] == 'YES') & (merged['Benefit Match'] == 'NO'),
            (merged['_merge'] == 'left_only') | (merged['_merge'] == 'right_only')
        ]
        
        categories = ['Amount Mismatch', 'Weight Mismatch', 'Benefit Mismatch', 'Record Missing']
        
        for condition, category in zip(conditions, categories):
            merged.loc[both_exist & condition, 'Mismatch Category'] = category
        
        # Multiple mismatches
        multiple = (merged['Amount Match'] == 'NO') & (merged['Weight Match'] == 'NO')
        merged.loc[both_exist & multiple, 'Mismatch Category'] = 'Multiple Mismatches'
        
        # Perfect match
        merged.loc[both_exist & (merged['Final Result'] == 'MATCH'), 'Mismatch Category'] = 'Perfect Match'
    
    # ===== FIX: CREATE CONSOLIDATED DATE COLUMNS WITH DATE-ONLY FORMAT =====
    # Use the existing columns that were already added
    if 'Date of joining_Participation' in merged.columns:
        merged['Date of joining'] = merged['Date of joining_Participation']
    elif 'Date of joining' in merged.columns:
        merged['Date of joining'] = merged['Date of joining']
    
    if 'Paid Date_Transaction' in merged.columns:
        merged['Paid Date'] = merged['Paid Date_Transaction']
    elif 'Paid Date' in merged.columns:
        merged['Paid Date'] = merged['Paid Date']
    
    # Ensure dates are in date-only format in the final merged dataframe
    date_cols_final = ['Date of joining', 'Paid Date', 'Maturity Date', 'Closed Date', 'Last receipt date',
                       'Date of joining_Participation', 'Paid Date_Transaction', 
                       'Maturity Date_Participation', 'Last receipt date_Participation']
    for col in date_cols_final:
        if col in merged.columns:
            try:
                # If it's already date object, keep it; otherwise convert
                if not merged[col].empty:
                    sample = merged[col].iloc[0]
                    if not isinstance(sample, (datetime.date, type(None))):
                        merged[col] = pd.to_datetime(merged[col], errors='coerce').dt.date
            except:
                pass
    
    return merged

# ============================================================================
# VISUALIZATION FUNCTIONS
# ============================================================================

def create_visualizations(comparison_data, config):
    """Create visualizations for the comparison results"""
    col1, col2 = st.columns(2)
    
    with col1:
        # Match/Mismatch ratio
        status_counts = comparison_data['Final Result'].value_counts()
        fig = px.pie(
            values=status_counts.values,
            names=status_counts.index,
            title="📊 Match vs Mismatch Distribution",
            color=status_counts.index,
            color_discrete_map={'MATCH': '#10B981', 'MISMATCH': '#EF4444'},
            hole=0.4
        )
        fig.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig, width="stretch")
    
    with col2:
        # Mismatch category distribution (only for mismatches)
        mismatch_data = comparison_data[comparison_data['Final Result'] == 'MISMATCH']
        if not mismatch_data.empty and 'Mismatch Category' in mismatch_data.columns:
            category_counts = mismatch_data['Mismatch Category'].value_counts()
            fig = px.bar(
                x=category_counts.index,
                y=category_counts.values,
                color=category_counts.values,
                title="🔍 Mismatch Categories",
                labels={'x': 'Category', 'y': 'Count'},
                color_continuous_scale='Reds'
            )
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, width="stretch")
    
    # Second row - additional metrics
    col3, col4 = st.columns(2)
    
    with col3:
        # Match score distribution
        if 'Match Score' in comparison_data.columns and not comparison_data[comparison_data['_merge'] == 'both'].empty:
            match_scores = comparison_data[comparison_data['_merge'] == 'both']['Match Score']
            fig = px.histogram(
                match_scores,
                title="📈 Match Score Distribution",
                labels={'value': 'Match Score (%)', 'count': 'Number of Records'},
                nbins=20,
                color_discrete_sequence=['#3B82F6']
            )
            fig.update_layout(bargap=0.1)
            st.plotly_chart(fig, width="stretch")
    
    with col4:
        # Amount difference distribution
        if 'Amount Diff' in comparison_data.columns:
            diff_data = comparison_data[comparison_data['_merge'] == 'both']['Amount Diff']
            if not diff_data.empty and diff_data.max() > 0:
                fig = px.box(
                    diff_data,
                    title="💰 Amount Differences Distribution",
                    labels={'value': 'Amount Difference (₹)'},
                    color_discrete_sequence=['#8B5CF6']
                )
                st.plotly_chart(fig, width="stretch")

# ============================================================================
# REPORT GENERATION FUNCTIONS
# ============================================================================

def generate_excel_report(comparison_data, missing_trans, extra_trans, config):
    """
    Generate Excel report with multiple sheets
    """
    output = BytesIO()
    
    # Helper function to format dates for Excel
    def format_dates_for_excel(df):
        if df is None or df.empty:
            return df
        
        df_copy = df.copy()
        for col in df_copy.columns:
            # Check if column contains date objects
            if not df_copy[col].empty:
                sample = df_copy[col].iloc[0]
                if isinstance(sample, (datetime.date, datetime.datetime)):
                    df_copy[col] = df_copy[col].astype(str)
                elif df_copy[col].dtype == 'datetime64[ns]':
                    df_copy[col] = df_copy[col].dt.date.astype(str)
        return df_copy
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Summary sheet
        total_participation = len(comparison_data[comparison_data['_merge'].isin(['both', 'left_only'])])
        total_transaction = len(comparison_data[comparison_data['_merge'].isin(['both', 'right_only'])])
        matched_count = len(comparison_data[comparison_data['Final Result'] == 'MATCH'])
        mismatched_count = len(comparison_data[comparison_data['Final Result'] == 'MISMATCH'])
        
        summary = pd.DataFrame({
            'Metric': [
                'Total Records in Participation',
                'Total Records in Transaction',
                'Matched Records',
                'Mismatched Records',
                'Missing Transactions (in Participation only)',
                'Extra Transactions (in Transaction only)',
                'Match Rate (%)',
                'Mismatch Rate (%)'
            ],
            'Count': [
                total_participation,
                total_transaction,
                matched_count,
                mismatched_count,
                len(missing_trans),
                len(extra_trans),
                f"{(matched_count/total_participation*100):.2f}%" if total_participation > 0 else "0%",
                f"{(mismatched_count/total_participation*100):.2f}%" if total_participation > 0 else "0%"
            ]
        })
        summary.to_excel(writer, sheet_name='Summary', index=False)
        
        # Configuration used
        config_df = pd.DataFrame({
            'Setting': ['Amount Tolerance', 'Weight Tolerance', 'Compare Dates', 'Fuzzy Matching'],
            'Value': [
                config['amount_tolerance'],
                config['weight_tolerance'],
                config['compare_dates'],
                config['fuzzy_matching']
            ]
        })
        config_df.to_excel(writer, sheet_name='Settings', index=False)
        
        # Define columns to include for matched records
        match_cols = ['Scheme Name', 'Customer Name', 'Phone', 'Passbook Number', 
                     'Date of joining', 'Paid Date',
                     'Saved Amount_Participation', 'Saved Amount_Transaction',
                     'Saved Weight_Participation', 'Saved Weight_Transaction',
                     'Benefit Amount_Participation', 'Benefit Amount_Transaction',
                     'Amount Match', 'Weight Match', 'Benefit Match', 
                     'Match Score', 'Final Result']
        
        # Matched records
        matched = comparison_data[comparison_data['Final Result'] == 'MATCH']
        if not matched.empty:
            # Include only relevant columns
            matched_cols = [col for col in match_cols if col in matched.columns]
            matched_formatted = format_dates_for_excel(matched[match_cols])
            matched_formatted.to_excel(writer, sheet_name='Matched', index=False)
        
        # Mismatched records
        mismatched = comparison_data[comparison_data['Final Result'] == 'MISMATCH']
        if not mismatched.empty:
            mismatched_cols = [col for col in match_cols if col in mismatched.columns]
            mismatched_formatted = format_dates_for_excel(mismatched[mismatched_cols])
            mismatched_formatted.to_excel(writer, sheet_name='Mismatched', index=False)
        
        # Missing Transactions (in participation but not in transaction)
        if not missing_trans.empty:
            # Include date columns
            missing_cols = ['Scheme Name', 'Customer Name', 'Phone', 'Passbook Number', 
                           'Date of joining', 'Date of joining_Participation']
            if 'Status_Participation' in missing_trans.columns:
                missing_cols.append('Status_Participation')
            missing_cols = [col for col in missing_cols if col in missing_trans.columns]
            missing_formatted = format_dates_for_excel(missing_trans[missing_cols])
            missing_formatted.to_excel(writer, sheet_name='Missing Transactions', index=False)
        
        # Extra Transactions (in transaction but not in participation)
        if not extra_trans.empty:
            extra_cols = ['Scheme Name', 'Customer Name', 'Phone', 'Passbook Number', 
                         'Paid Date', 'Paid Date_Transaction']
            if 'Status_Transaction' in extra_trans.columns:
                extra_cols.append('Status_Transaction')
            extra_cols = [col for col in extra_cols if col in extra_trans.columns]
            extra_formatted = format_dates_for_excel(extra_trans[extra_cols])
            extra_formatted.to_excel(writer, sheet_name='Extra Transactions', index=False)
        
        # All Records - include all columns with formatted dates
        all_records_formatted = format_dates_for_excel(comparison_data)
        all_records_formatted.to_excel(writer, sheet_name='All Records', index=False)
    
    return output.getvalue()

# ============================================================================
# EMAIL NOTIFICATION (Optional)
# ============================================================================

def send_notification_email(recipient, summary_stats, config):
    """Send email notification with summary (placeholder implementation)"""
    # This is a placeholder - implement actual email sending using smtplib
    # or use a service like SendGrid, Amazon SES, etc.
    
    if config['enable_email'] and recipient:
        # Check if notification should be sent
        mismatched_pct = summary_stats.get('mismatch_rate', 0)
        trigger_pct = {
            "All mismatches found": 0,
            "Mismatch > 10%": 10,
            "Mismatch > 20%": 20,
            "Mismatch > 50%": 50
        }.get(config['email_trigger'], 0)
        
        if mismatched_pct >= trigger_pct:
            # Would send email here
            # For now, just show a message
            st.info(f"📧 Email notification would be sent to {recipient} (mismatch rate: {mismatched_pct:.1f}%)")
            return True
    
    return False

# ============================================================================
# DISPLAY FUNCTIONS
# ============================================================================

def display_results(comparison, missing_trans, extra_trans, participation_df, config):
    """Display verification results"""
    
    # Calculate statistics
    total_participation = len(comparison[comparison['_merge'].isin(['both', 'left_only'])])
    total_transaction = len(comparison[comparison['_merge'].isin(['both', 'right_only'])])
    matched_count = len(comparison[comparison['Final Result'] == 'MATCH'])
    mismatched_count = len(comparison[comparison['Final Result'] == 'MISMATCH'])
    
    # Display statistics
    st.markdown("---")
    st.subheader("📊 Verification Results")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "📄 Total Participation Records",
            f"{total_participation:,}"
        )
    
    with col2:
        st.metric(
            "📊 Total Transaction Records",
            f"{total_transaction:,}"
        )
    
    with col3:
        if matched_count > 0:
            st.markdown(f"""
            <div class="success-box">
                <strong>✅ Matched Records</strong><br>
                {matched_count:,}
            </div>
            """, unsafe_allow_html=True)
        else:
            st.metric("✅ Matched Records", "0")
    
    with col4:
        if mismatched_count > 0:
            st.markdown(f"""
            <div class="error-box">
                <strong>❌ Mismatched Records</strong><br>
                {mismatched_count:,}
            </div>
            """, unsafe_allow_html=True)
        else:
            st.metric("❌ Mismatched Records", "0")
    
    # Additional metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "🔴 Missing Transactions",
            f"{len(missing_trans):,}"
        )
    
    with col2:
        st.metric(
            "🔵 Extra Transactions",
            f"{len(extra_trans):,}"
        )
    
    with col3:
        match_rate = (matched_count / total_participation * 100) if total_participation > 0 else 0
        delta_color = "normal" if match_rate >= 90 else "inverse"
        st.metric(
            "📈 Match Rate",
            f"{match_rate:.2f}%",
            delta=f"{match_rate - 100:.1f}%" if match_rate < 100 else "Perfect!",
            delta_color=delta_color
        )
    
    with col4:
        mismatch_rate = (mismatched_count / total_participation * 100) if total_participation > 0 else 0
        st.metric(
            "⚠️ Mismatch Rate",
            f"{mismatch_rate:.2f}%"
        )
    
    # Send email notification if enabled
    if config['enable_email'] and config['recipient_email']:
        summary_stats = {
            'total_participation': total_participation,
            'total_transaction': total_transaction,
            'matched_count': matched_count,
            'mismatched_count': mismatched_count,
            'match_rate': match_rate,
            'mismatch_rate': mismatch_rate
        }
        send_notification_email(config['recipient_email'], summary_stats, config)
    
    # Show preview
    st.markdown("---")
    st.subheader("📋 Comparison Preview")
    
    # ===== FIX: INCLUDE DATE COLUMNS IN PREVIEW WITH DATE-ONLY FORMAT =====
    preview_cols = ['Scheme Name', 'Customer Name', 'Phone', 'Passbook Number', 
                   'Date of joining', 'Paid Date',
                   'Final Result', 'Amount Match', 'Weight Match', 'Benefit Match', 
                   'Match Score', 'Mismatch Category']
    
    # Add status columns if they exist (but note they're not compared)
    if 'Status_Participation' in comparison.columns:
        preview_cols.append('Status_Participation')
    if 'Status_Transaction' in comparison.columns:
        preview_cols.append('Status_Transaction')
    
    # Also check for date columns with suffixes (backup)
    if 'Date of joining' not in comparison.columns:
        if 'Date of joining_Participation' in comparison.columns:
            preview_cols.append('Date of joining_Participation')
    if 'Paid Date' not in comparison.columns:
        if 'Paid Date_Transaction' in comparison.columns:
            preview_cols.append('Paid Date_Transaction')
    
    preview_cols = [col for col in preview_cols if col in comparison.columns]
    
    # Filter based on display settings
    display_data = comparison
    if not config['show_matched']:
        display_data = comparison[comparison['Final Result'] == 'MISMATCH']
    
    # Show top N records
    display_data = display_data.head(config['records_to_show'])
    
    # Format dates for display
    display_data_formatted = format_dates_in_dataframe(display_data)
    
    # Check if we can apply styling
    total_cells = len(display_data_formatted) * len(preview_cols)
    max_styler_cells = pd.get_option("styler.render.max_elements")
    
    try:
        if total_cells <= max_styler_cells and len(display_data_formatted) <= 100:
            # Apply styling if within limits
            def color_result(val):
                if val == 'MATCH':
                    return 'background-color: #D1FAE5; color: #065F46'
                elif val == 'MISMATCH':
                    return 'background-color: #FEE2E2; color: #991B1B'
                return ''
            
            styled_df = display_data_formatted[preview_cols].style.map(color_result, subset=['Final Result'])
            st.dataframe(styled_df, width="stretch", height=400)
        else:
            # Show without styling for large datasets
            st.dataframe(display_data_formatted[preview_cols], width="stretch", height=400)
    except Exception as e:
        # Fallback to unstyled view
        st.dataframe(display_data_formatted[preview_cols], width="stretch", height=400)
    
    return {
        'total_participation': total_participation,
        'total_transaction': total_transaction,
        'matched_count': matched_count,
        'mismatched_count': mismatched_count,
        'match_rate': match_rate,
        'mismatch_rate': mismatch_rate
    }

def display_mismatch_details(comparison):
    """Display detailed mismatch information with efficient rendering"""
    mismatched_count = len(comparison[comparison['Final Result'] == 'MISMATCH'])
    
    if mismatched_count > 0:
        st.markdown("---")
        
        # Show warning with count
        if mismatched_count > 1000:
            st.warning(f"⚠️ Found {mismatched_count:,} mismatched records. For performance, only summary is shown below. Download the full report for details.")
        else:
            st.warning(f"⚠️ Found {mismatched_count} mismatched records. Check the downloaded report for details.")
        
        with st.expander(f"🔍 Show Mismatch Details ({mismatched_count:,} records)"):
            mismatched_df = comparison[comparison['Final Result'] == 'MISMATCH'].copy()
            
            # ===== FIX: INCLUDE DATE COLUMNS IN MISMATCH DETAILS =====
            mismatch_cols = ['Scheme Name', 'Customer Name', 'Phone', 'Passbook Number',
                            'Date of joining', 'Paid Date',
                            'Saved Amount_Participation', 'Saved Amount_Transaction',
                            'Saved Weight_Participation', 'Saved Weight_Transaction',
                            'Benefit Amount_Participation', 'Benefit Amount_Transaction',
                            'Amount Match', 'Weight Match', 'Benefit Match', 
                            'Mismatch Category']
            
            # Add status columns if they exist (for reference only, not compared)
            if 'Status_Participation' in mismatched_df.columns:
                mismatch_cols.append('Status_Participation')
            if 'Status_Transaction' in mismatched_df.columns:
                mismatch_cols.append('Status_Transaction')
            
            # Also check for date columns with suffixes
            if 'Date of joining' not in mismatched_df.columns and 'Date of joining_Participation' in mismatched_df.columns:
                mismatch_cols.append('Date of joining_Participation')
            if 'Paid Date' not in mismatched_df.columns and 'Paid Date_Transaction' in mismatched_df.columns:
                mismatch_cols.append('Paid Date_Transaction')
            
            mismatch_cols = [col for col in mismatch_cols if col in mismatched_df.columns]
            
            if not mismatched_df.empty:
                # Tabbed view for better organization
                tab1, tab2, tab3 = st.tabs(["📊 Summary", "📋 Sample Records", "📈 Category Breakdown"])
                
                with tab1:
                    # Summary statistics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Mismatches", f"{len(mismatched_df):,}")
                    with col2:
                        unique_customers = mismatched_df['Customer Name'].nunique()
                        st.metric("Unique Customers", f"{unique_customers:,}")
                    with col3:
                        unique_schemes = mismatched_df['Scheme Name'].nunique()
                        st.metric("Unique Schemes", f"{unique_schemes:,}")
                    
                    # Show mismatch category summary
                    st.subheader("Mismatch Category Summary")
                    category_summary = mismatched_df['Mismatch Category'].value_counts().reset_index()
                    category_summary.columns = ['Category', 'Count']
                    st.dataframe(category_summary, width="stretch")
                
                with tab2:
                    # Show sample of records (first 100) with formatted dates
                    sample_size = min(100, len(mismatched_df))
                    st.caption(f"Showing first {sample_size} mismatched records (out of {len(mismatched_df):,})")
                    
                    # Format dates for display
                    mismatched_sample = format_dates_in_dataframe(mismatched_df[mismatch_cols].head(100))
                    
                    # Display without styling to avoid limit issues
                    st.dataframe(
                        mismatched_sample,
                        width="stretch",
                        height=400
                    )
                    
                    # If there are more records, show message
                    if len(mismatched_df) > 100:
                        st.info(f"📝 Showing only first 100 records. Download the Excel report to see all {len(mismatched_df):,} mismatched records.")
                
                with tab3:
                    # Category breakdown with chart
                    category_counts = mismatched_df['Mismatch Category'].value_counts()
                    
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        fig = px.bar(
                            x=category_counts.index,
                            y=category_counts.values,
                            title="Mismatch Categories Distribution",
                            labels={'x': 'Category', 'y': 'Count'},
                            color=category_counts.values,
                            color_continuous_scale='Reds'
                        )
                        fig.update_layout(showlegend=False)
                        st.plotly_chart(fig, width="stretch")
                    
                    with col2:
                        # Show as table
                        category_df = category_counts.reset_index()
                        category_df.columns = ['Category', 'Count']
                        st.dataframe(category_df, width="stretch")

def create_download_section(comparison, missing_trans, extra_trans):
    """Create download buttons for reports"""
    st.markdown("---")
    st.subheader("📥 Download Reports")
    
    col1, col2, col3 = st.columns(3)
    
    # Generate Excel report
    excel_data = generate_excel_report(comparison, missing_trans, extra_trans, st.session_state.get('config', {}))
    
    with col1:
        st.download_button(
            label="📊 Download Excel Report",
            data=excel_data,
            file_name=f"Scheme_Verification_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch"
        )
    
    with col2:
        # Format dates before CSV export
        comparison_formatted = format_dates_in_dataframe(comparison)
        csv_data = comparison_formatted.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📄 Download CSV Report",
            data=csv_data,
            file_name=f"Scheme_Verification_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            width="stretch"
        )
    
    with col3:
        # Download only mismatched records with formatted dates
        mismatched = comparison[comparison['Final Result'] == 'MISMATCH']
        if not mismatched.empty:
            mismatched_formatted = format_dates_in_dataframe(mismatched)
            csv_mismatch = mismatched_formatted.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📉 Download Mismatched Records",
                data=csv_mismatch,
                file_name=f"Mismatched_Records_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                width="stretch"
            )

# ============================================================================
# CACHE FUNCTIONS FOR PERFORMANCE
# ============================================================================

@st.cache_data(ttl=3600)
def load_and_process_file(file_content, file_name, file_type, use_fuzzy=True, chunk_size=10000):
    """Cached file processing for better performance"""
    if file_name.endswith('.csv'):
        df = pd.read_csv(BytesIO(file_content))
    else:
        df = pd.read_excel(BytesIO(file_content))
    
    df = standardize_columns(df, file_type, use_fuzzy)
    df = process_large_file(df, file_type, chunk_size)
    return df

# ============================================================================
# MAIN APPLICATION
# ============================================================================

def main():
    # Initialize config in session state
    if 'config' not in st.session_state:
        st.session_state.config = {}
    
    st.markdown('<h1 class="main-header">🏦 Customer Scheme Verification Tool</h1>', unsafe_allow_html=True)
    
    st.markdown("""
    <div style='text-align: center; margin-bottom: 2rem;'>
    <p style='color: #6B7280;'>Upload both Participation and Transaction files to verify customer scheme details.</p>
    <p style='color: #6B7280; font-size: 0.9rem;'>
    <strong>Note:</strong> Status fields are displayed for reference but are <strong>not compared</strong> as they represent different business concepts 
    (Scheme Status vs Transaction Status).<br>
    <strong>Date Format:</strong> All dates are displayed in <strong>YYYY-MM-DD</strong> format without time component.
    </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Get configuration from sidebar
    config = create_config_sidebar()
    st.session_state.config = config
    
    # Main file upload section
    col1, col2 = st.columns(2)
    
    with col1:
        participation_file = st.file_uploader(
            "📄 Customer Scheme Participation",
            type=['csv', 'xlsx'],
            help="Upload the participation file with customer scheme details"
        )
        if participation_file:
            st.success(f"✅ File loaded: {participation_file.name}")
    
    with col2:
        transaction_file = st.file_uploader(
            "📊 Customer Scheme Transaction",
            type=['csv', 'xlsx'],
            help="Upload the transaction file with payment details"
        )
        if transaction_file:
            st.success(f"✅ File loaded: {transaction_file.name}")
    
    # Process files when both are uploaded
    if participation_file and transaction_file:
        try:
            # Create progress tracking
            progress_text = st.empty()
            progress_bar = st.progress(0)
            
            progress_text.text("📂 Loading and processing files...")
            progress_bar.progress(10)
            
            # Read and process files with caching
            participation_df = load_and_process_file(
                participation_file.getvalue(), 
                participation_file.name, 
                'participation',
                config['fuzzy_matching'],
                config['chunk_size']
            )
            
            progress_bar.progress(40)
            progress_text.text("📂 Processing transaction file...")
            
            transaction_df = load_and_process_file(
                transaction_file.getvalue(), 
                transaction_file.name, 
                'transaction',
                config['fuzzy_matching'],
                config['chunk_size']
            )
            
            progress_bar.progress(60)
            
            # Validate file structures
            if not validate_file_structure(participation_df, 'participation'):
                st.stop()
            if not validate_file_structure(transaction_df, 'transaction'):
                st.stop()
            
            # Process participation file
            progress_text.text("🔄 Processing participation data...")
            participation_df = process_participation(participation_df)
            if participation_df is None:
                st.stop()
            
            progress_bar.progress(70)
            
            # Process transaction file
            progress_text.text("🔄 Processing transaction data...")
            transaction_grouped = process_transactions(transaction_df)
            if transaction_grouped is None:
                st.stop()
            
            progress_bar.progress(85)
            
            # Compare records
            progress_text.text("🔍 Comparing records...")
            comparison = detailed_comparison(
                participation_df, 
                transaction_grouped, 
                config
            )
            
            # Extract missing and extra records
            missing_trans = comparison[comparison['_merge'] == 'left_only']
            extra_trans = comparison[comparison['_merge'] == 'right_only']
            
            progress_bar.progress(100)
            progress_text.text("✅ Processing complete!")
            
            # Clear progress indicators after a moment
            time.sleep(0.5)
            progress_text.empty()
            progress_bar.empty()
            
            # Display results
            stats = display_results(comparison, missing_trans, extra_trans, participation_df, config)
            
            # Create visualizations
            st.markdown("---")
            create_visualizations(comparison, config)
            
            # Show mismatch details
            display_mismatch_details(comparison)
            
            # Download section
            create_download_section(comparison, missing_trans, extra_trans)
            
            # Success message if all matched
            if stats['matched_count'] == stats['total_participation'] and len(missing_trans) == 0 and len(extra_trans) == 0:
                st.balloons()
                st.markdown("""
                <div class="success-box" style='text-align: center;'>
                    <h3>🎉 Perfect Match!</h3>
                    <p>All records from both files match exactly!</p>
                </div>
                """, unsafe_allow_html=True)
        
        except Exception as e:
            st.error(f"❌ Error processing files: {str(e)}")
            st.exception(e)
            
    else:
        st.info("👆 Please upload both participation and transaction files to begin verification.")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #6B7280; font-size: 0.8rem;'>
    Customer Scheme Verification Tool v2.0 | Built with Streamlit & Plotly
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()