# app_daily_scheme.py - Daily Scheme Payment Frequency Tracker
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import base64
import plotly.express as px
import plotly.graph_objects as go
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Daily Scheme - Payment Frequency Tracker",
    page_icon="📒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CUSTOM CSS FOR PROFESSIONAL UI
# ============================================================
def inject_custom_css():
    st.markdown("""
    <style>
        /* Main container styling */
        .main {
            padding: 0rem 1rem;
        }
        
        /* Metric cards */
        .metric-card {
            background: white;
            border-radius: 12px;
            padding: 1.2rem 1.5rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            border: 1px solid #f0f2f6;
            transition: all 0.2s ease;
            height: 100%;
        }
        .metric-card:hover {
            box-shadow: 0 4px 16px rgba(0,0,0,0.10);
            transform: translateY(-2px);
        }
        .metric-label {
            font-size: 0.8rem;
            font-weight: 600;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 0.3rem;
        }
        .metric-value {
            font-size: 1.6rem;
            font-weight: 700;
            color: #111827;
        }
        .metric-sub {
            font-size: 0.8rem;
            color: #6b7280;
            margin-top: 0.2rem;
        }
        
        /* Custom header */
        .app-header {
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
            padding: 1.5rem 2rem;
            border-radius: 16px;
            margin-bottom: 1.5rem;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
        }
        .app-title {
            font-size: 2.2rem;
            font-weight: 700;
            color: white;
            letter-spacing: -0.5px;
        }
        .app-subtitle {
            font-size: 0.95rem;
            color: rgba(255,255,255,0.7);
            margin-top: 0.3rem;
        }
        
        /* Sidebar styling */
        .css-1d391kg {
            background-color: #f8fafc;
        }
        .css-1d391kg .css-1v3fvcr {
            background-color: #f8fafc;
        }
        
        /* Card containers */
        .card-container {
            background: white;
            border-radius: 12px;
            padding: 1.5rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
            border: 1px solid #f0f2f6;
            margin-bottom: 1rem;
        }
        
        /* Button styling */
        .stButton > button {
            border-radius: 8px;
            font-weight: 600;
            transition: all 0.2s ease;
            border: none;
        }
        .stButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }
        
        /* Tab styling */
        .stTabs [data-baseweb="tab-list"] {
            gap: 2px;
            background-color: #f8fafc;
            border-radius: 10px;
            padding: 4px;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 8px;
            padding: 0.5rem 1.2rem;
            font-weight: 500;
            color: #6b7280;
        }
        .stTabs [aria-selected="true"] {
            background-color: white !important;
            color: #1a1a2e !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        
        /* Dataframe styling */
        .dataframe-container {
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid #f0f2f6;
        }
        .dataframe-container table {
            font-size: 0.9rem;
        }
        
        /* Status badges */
        .badge-daily {
            background: #d1fae5;
            color: #065f46;
            padding: 0.2rem 0.7rem;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
        }
        .badge-weekly {
            background: #dbeafe;
            color: #1e40af;
            padding: 0.2rem 0.7rem;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
        }
        .badge-monthly {
            background: #fef3c7;
            color: #92400e;
            padding: 0.2rem 0.7rem;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
        }
        .badge-irregular {
            background: #fee2e2;
            color: #991b1b;
            padding: 0.2rem 0.7rem;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
        }
        .badge-trial {
            background: #e5e7eb;
            color: #374151;
            padding: 0.2rem 0.7rem;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
        }
        
        /* Footer */
        .footer {
            text-align: center;
            padding: 1.5rem;
            color: #9ca3af;
            font-size: 0.85rem;
            border-top: 1px solid #f0f2f6;
            margin-top: 2rem;
        }
        
        /* Metric grid */
        .metric-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 1rem;
            margin-bottom: 1.5rem;
        }
        
        /* Responsive adjustments */
        @media (max-width: 768px) {
            .app-title {
                font-size: 1.5rem;
            }
            .metric-value {
                font-size: 1.2rem;
            }
            .metric-grid {
                grid-template-columns: repeat(2, 1fr);
            }
        }
        
        /* Custom scrollbar */
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f1f1;
            border-radius: 3px;
        }
        ::-webkit-scrollbar-thumb {
            background: #c1c7cd;
            border-radius: 3px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #a0a7ae;
        }
        
        /* Divider */
        .custom-divider {
            height: 2px;
            background: linear-gradient(to right, transparent, #e5e7eb, transparent);
            margin: 1.5rem 0;
        }
        
        /* Download button styling */
        .download-btn {
            display: inline-block;
            padding: 0.4rem 1rem;
            background: #f3f4f6;
            color: #374151;
            border-radius: 8px;
            text-decoration: none;
            font-size: 0.85rem;
            font-weight: 500;
            transition: all 0.2s ease;
            border: 1px solid #e5e7eb;
            margin: 0.2rem 0;
        }
        .download-btn:hover {
            background: #e5e7eb;
            transform: translateY(-1px);
        }
        .download-btn-primary {
            background: #1a1a2e;
            color: white;
            border-color: #1a1a2e;
        }
        .download-btn-primary:hover {
            background: #2d2d4a;
            color: white;
        }
        
        /* Section headers */
        .section-header {
            font-size: 1.3rem;
            font-weight: 600;
            color: #1a1a2e;
            margin: 1.5rem 0 1rem 0;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid #f0f2f6;
        }
        
        /* Info boxes */
        .info-box {
            background: #f8fafc;
            border-left: 4px solid #0f3460;
            padding: 0.8rem 1.2rem;
            border-radius: 8px;
            margin: 0.5rem 0;
        }
        
        /* Highlight boxes */
        .highlight-box {
            background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%);
            border-left: 4px solid #f59e0b;
            padding: 0.8rem 1.2rem;
            border-radius: 8px;
            margin: 0.5rem 0;
        }
    </style>
    """, unsafe_allow_html=True)

inject_custom_css()

# ============================================================
# SESSION STATE
# ============================================================
def init_session_state():
    if "data" not in st.session_state:
        st.session_state.data = None
    if "raw_data" not in st.session_state:
        st.session_state.raw_data = None
    if "processed_data" not in st.session_state:
        st.session_state.processed_data = None
    if "passbook_summary" not in st.session_state:
        st.session_state.passbook_summary = None
    if "category_summary" not in st.session_state:
        st.session_state.category_summary = None
    if "last_payment_category_summary" not in st.session_state:
        st.session_state.last_payment_category_summary = None
    if "multi_passbook_customers" not in st.session_state:
        st.session_state.multi_passbook_customers = None
    if "single_passbook_customers" not in st.session_state:
        st.session_state.single_passbook_customers = None
    if "summary_cache" not in st.session_state:
        st.session_state.summary_cache = {}
    if "slab_config" not in st.session_state:
        st.session_state.slab_config = {
            "gold": {"slab1_max": 10, "slab2_min": 11, "slab2_max": 50, "slab3_min": 51},
            "silver": {"slab1_max": 10, "slab2_min": 11, "slab2_max": 50, "slab3_min": 51}
        }
    if "page_size" not in st.session_state:
        st.session_state.page_size = 50

init_session_state()

# ============================================================
# CONSTANTS
# ============================================================
SLAB_ORDER = ["Category A", "Category B", "Category C", "Category D"]

# Updated Customer Summary Columns with all desired fields
CUSTOMER_SUMMARY_COLUMNS = [
    "Customer Name",
    "Customer Phone Number",
    "Passbook number",
    "Total Saved",
    "Payment Behavior",
    "Avg Ticket",
    "Skipped Days",
    "Overall Payment Period",
    "Latest Installment",
    "Avg Days Between Payments",
    "Category Label (Avg Days Payments)",
    "Consistency Score",
    "Max Gap Days",
    "Last Payment Date",
    "Days Since Last Payment To Today",
    "Category Label (Last Payment)",
    "Transaction Count",
    "Total Value",
    "Payment Status",
    "Rank"
]

# Multi-passbook customer columns (Sheet 2)
MULTI_PASSBOOK_COLUMNS = [
    "Customer Name",
    "Customer Phone Number",
    "Number of Passbooks",
    "Passbook Numbers",
    "Total Saved (All Passbooks)",
    "Avg Ticket (All Passbooks)",
    "Payment Behavior",
    "Skipped Days (Total)",
    "Overall Payment Period (Avg)",
    "Avg Days Between Payments (Avg)",
    "Category Label (Avg Days Payments)",
    "Consistency Score (Avg)",
    "Max Gap Days (Max)",
    "Last Payment Date (Latest)",
    "Days Since Last Payment To Today",
    "Category Label (Last Payment)",
    "Total Transactions (All Passbooks)",
    "Payment Status"
]

# Single passbook customer columns (Sheet 3)
SINGLE_PASSBOOK_COLUMNS = [
    "Customer Name",
    "Customer Phone Number",
    "Passbook number",
    "Total Saved",
    "Payment Behavior",
    "Avg Ticket",
    "Skipped Days",
    "Overall Payment Period",
    "Latest Installment",
    "Avg Days Between Payments",
    "Category Label (Avg Days Payments)",
    "Consistency Score",
    "Max Gap Days",
    "Last Payment Date",
    "Days Since Last Payment To Today",
    "Category Label (Last Payment)",
    "Transaction Count",
    "Total Value",
    "Payment Status",
    "Rank"
]

# Define the columns we want for Category Summary
CATEGORY_SUMMARY_COLUMNS = [
    "Metal Type", "Category", "Description", 
    "Unique Passbooks", "Total Transactions", 
    "Total Saved", "Avg Per Ticket"
]

# Define the columns we want for Last Payment Category Summary
LAST_PAYMENT_CATEGORY_SUMMARY_COLUMNS = [
    "Metal Type", "Category", "Description",
    "Unique Passbooks", "Total Transactions",
    "Total Saved", "Avg Per Ticket"
]

def build_slab_labels(slab_config):
    """Build Category label/description text from the actual configured
    thresholds. These labels are used for display purposes only and 
    should match the category assignment logic."""
    slab1_max = slab_config["gold"]["slab1_max"]
    slab2_max = slab_config["gold"]["slab2_max"]

    categories = {
        "Category A": f"≤ {slab1_max} days - Active Daily",
        "Category B": f"{slab1_max + 1}–{slab2_max} days - Gap",
        "Category C": f"≥ {slab2_max + 1} days - Irregular",
        "Category D": "One-Time - Trial"
    }
    
    descriptions = {
        "Category A": f"≤ {slab1_max} days - Active Daily",
        "Category B": f"{slab1_max + 1}–{slab2_max} days - Gap",
        "Category C": f"≥ {slab2_max + 1} days - Irregular",
        "Category D": "One-Time - Trial"
    }
    return categories, descriptions

def get_category_for_days(days, slab_config, is_one_time=False):
    """Determine category based on days using slab configuration"""
    if is_one_time or days == 0:
        return "Category D"
    
    slab1_max = slab_config["gold"]["slab1_max"]
    slab2_max = slab_config["gold"]["slab2_max"]
    
    if days <= slab1_max:
        return "Category A"
    elif days <= slab2_max:
        return "Category B"
    else:
        return "Category C"

def get_category_label_for_days(days, slab_config, is_one_time=False):
    """Get the full category label based on days using slab configuration"""
    category = get_category_for_days(days, slab_config, is_one_time)
    categories, _ = build_slab_labels(slab_config)
    return categories.get(category, category)

# ============================================================
# COLUMN MAPPING
# ============================================================
def map_columns_fast(df):
    """Fast column mapping using lowercase comparison"""
    df = df.copy()
    lower_cols = {col.lower(): col for col in df.columns}
    
    column_mapping = {
        "id": "Id",
        "passbook number": "Passbook number",
        "passbook_no": "Passbook number",
        "account number": "Passbook number",
        "account_no": "Passbook number",
        "metal type": "Metal Type",
        "metal": "Metal Type",
        "metal_type": "Metal Type",
        "installment number": "Installment number",
        "installment": "Installment number",
        "installment_no": "Installment number",
        "paid date": "Paid Date",
        "payment date": "Paid Date",
        "date": "Paid Date",
        "saved amount": "Saved Amount",
        "amount": "Saved Amount",
        "saved_amount": "Saved Amount",
        "reward amount": "Reward Amount",
        "reward_amount": "Reward Amount",
        "bonus": "Reward Amount",
        "customer name": "Customer Name",
        "customer": "Customer Name",
        "name": "Customer Name",
        "customer_name": "Customer Name",
        "customer phone number": "Customer Phone Number",
        "phone": "Customer Phone Number",
        "mobile": "Customer Phone Number",
        "phone number": "Customer Phone Number"
    }
    
    rename_dict = {}
    for lower_key, std_key in column_mapping.items():
        if lower_key in lower_cols:
            rename_dict[lower_cols[lower_key]] = std_key
    
    if rename_dict:
        df = df.rename(columns=rename_dict)
    
    return df

# ============================================================
# ENHANCED DATA PROCESSING WITH FREQUENCY TRACKING
# ============================================================
@st.cache_data(ttl=3600)
def process_daily_scheme_data(df, slab_config):
    """Process daily scheme data with payment frequency tracking"""
    df = df.copy()
    
    # Map columns
    df = map_columns_fast(df)
    
    # Fill missing columns with defaults
    default_cols = {
        "Id": 0,
        "Passbook number": "Unknown",
        "Metal Type": "Gold",
        "Installment number": 1,
        "Saved Amount": 0,
        "Reward Amount": 0,
        "Customer Name": "Not Available",
        "Customer Phone Number": "Not Available",
        "Paid Date": pd.NaT
    }
    
    for col, default in default_cols.items():
        if col not in df.columns:
            df[col] = default
    
    # ============================================================
    # Convert date column with proper dd-mm-yyyy format
    # ============================================================
    if "Paid Date" in df.columns:
        # First try parsing as dd-mm-yyyy (dayfirst=True)
        df["Paid Date"] = pd.to_datetime(df["Paid Date"], errors="coerce", dayfirst=True)
        
        # If still having issues with some dates, try explicit format
        if df["Paid Date"].isna().any():
            mask = df["Paid Date"].isna()
            try:
                df.loc[mask, "Paid Date"] = pd.to_datetime(
                    df.loc[mask, "Paid Date"], 
                    format="%d-%m-%Y", 
                    errors="coerce"
                )
            except:
                pass
    
    # Convert numeric columns
    numeric_cols = {
        "Installment number": int,
        "Saved Amount": float,
        "Reward Amount": float
    }
    
    for col, dtype in numeric_cols.items():
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        if dtype == int:
            df[col] = df[col].astype(int)
    
    # Clean string columns
    df["Metal Type"] = df["Metal Type"].fillna("Gold").astype(str).str.strip().str.title()
    df.loc[~df["Metal Type"].isin(["Gold", "Silver"]), "Metal Type"] = "Gold"
    
    df["Passbook number"] = df["Passbook number"].fillna("Unknown").astype(str).str.strip()
    invalid = ["", "nan", "None", "NaN", "null", "NULL"]
    df.loc[df["Passbook number"].str.lower().isin(invalid), "Passbook number"] = "Unknown"
    
    df["Customer Name"] = df["Customer Name"].fillna("Not Available").astype(str).str.strip()
    df.loc[df["Customer Name"].str.lower().isin(["", "nan", "none", "null"]), "Customer Name"] = "Not Available"
    
    df["Customer Phone Number"] = df["Customer Phone Number"].fillna("Not Available").astype(str).str.strip()
    df.loc[df["Customer Phone Number"].str.lower().isin(["", "nan", "none", "null"]), "Customer Phone Number"] = "Not Available"
    
    # Calculate transaction count per passbook
    df["Passbook Txn Count"] = df.groupby("Passbook number")["Passbook number"].transform("size")
    
    # ============================================================
    # PAYMENT FREQUENCY ANALYSIS - FIXED OFF-BY-ONE ERROR
    # ============================================================
    
    # Sort by passbook and date for proper gap calculation
    df = df.sort_values(["Passbook number", "Paid Date"])
    
    # Calculate gaps between consecutive payments for each passbook
    df["Previous Payment Date"] = df.groupby("Passbook number")["Paid Date"].shift(1)
    
    # FIXED: Subtract 1 to get days BETWEEN payments (excluding the payment day itself)
    # 20-04-2026 to 28-04-2026 = 7 days (21,22,23,24,25,26,27)
    df["Days Since Last Payment"] = (df["Paid Date"] - df["Previous Payment Date"]).dt.days - 1
    
    # For first payment of each passbook, set Days Since Last Payment to 0
    df.loc[df.groupby("Passbook number").cumcount() == 0, "Days Since Last Payment"] = 0
    
    # Handle any negative values (consecutive days should be 0)
    df["Days Since Last Payment"] = df["Days Since Last Payment"].fillna(0).astype(int)
    df.loc[df["Days Since Last Payment"] < 0, "Days Since Last Payment"] = 0
    
    # Calculate average and max gaps using only valid gaps (>0)
    def calculate_gap_metrics(group):
        """Calculate average and max gaps using only valid gaps (>0)"""
        gaps = group["Days Since Last Payment"]
        # Filter out 0 gaps (first payment) and any invalid values
        valid_gaps = gaps[(gaps > 0) & (gaps.notna())]
        
        if len(valid_gaps) == 0:
            return pd.Series({
                'Avg Days Between Payments': 0,
                'Max Gap Days': 0,
                'Gap Count': 0
            })
        
        # Convert to int for proper calculation
        valid_gaps_int = valid_gaps.astype(int)
        
        return pd.Series({
            'Avg Days Between Payments': valid_gaps_int.mean(),
            'Max Gap Days': valid_gaps_int.max(),
            'Gap Count': len(valid_gaps_int)
        })
    
    # Calculate gap metrics per passbook using only valid gaps
    gap_metrics = df.groupby("Passbook number").apply(calculate_gap_metrics)
    
    # Map back to dataframe
    df["Avg Days Between Payments"] = df["Passbook number"].map(gap_metrics["Avg Days Between Payments"])
    df["Max Gap Days"] = df["Passbook number"].map(gap_metrics["Max Gap Days"])
    
    # Round Avg Days Between Payments to whole number
    df["Avg Days Between Payments"] = df["Avg Days Between Payments"].round(0).astype(int)
    
    # For passbooks with only 1 transaction or no valid gaps, set values to 0
    df.loc[df["Passbook Txn Count"] <= 1, "Avg Days Between Payments"] = 0
    df.loc[df["Passbook Txn Count"] <= 1, "Max Gap Days"] = 0
    
    # Calculate payment streak
    df["Payment Streak"] = df.groupby("Passbook number").cumcount() + 1
    
    # Get last payment date for each passbook
    last_payment_date = df.groupby("Passbook number")["Paid Date"].max()
    df["Last Payment Date"] = df["Passbook number"].map(last_payment_date)
    
    # ============================================================
    # Calculate unique payment days per passbook
    # ============================================================
    def get_unique_payment_days(group):
        """Get the number of unique dates with payments"""
        unique_dates = group["Paid Date"].dt.date.nunique()
        return unique_dates
    
    unique_payment_days = df.groupby("Passbook number").apply(get_unique_payment_days)
    df["Unique Payment Days"] = df["Passbook number"].map(unique_payment_days)
    
    # ============================================================
    # Calculate days since last payment to today
    # ============================================================
    today = datetime.now().date()
    
    def get_days_since_last(pb):
        last_date = last_payment_date[pb]
        if pd.isna(last_date):
            return 0
        if hasattr(last_date, 'date'):
            last_date = last_date.date()
        days_diff = (today - last_date).days
        return days_diff if days_diff >= 0 else 0
    
    df["Days Since Last Payment To Today"] = df["Passbook number"].map(get_days_since_last)
    
    # ============================================================
    # Category Label based on Days Since Last Payment To Today
    # ============================================================
    def get_last_payment_category_label(row):
        """Get full category label based on Days Since Last Payment To Today"""
        days = row["Days Since Last Payment To Today"]
        is_one_time = row["Passbook Txn Count"] <= 1
        return get_category_label_for_days(days, slab_config, is_one_time)
    
    df["Category Label (Last Payment)"] = df.apply(get_last_payment_category_label, axis=1)
    
    # ============================================================
    # Consistency Score (based on historical gaps)
    # ============================================================
    def calculate_consistency(group):
        """Calculate consistency score based on payment gaps"""
        gaps = group["Days Since Last Payment"]
        valid_gaps = gaps[gaps > 0]
        
        if len(valid_gaps) == 0:
            return 0
        
        avg_gap = valid_gaps.mean()
        # Higher score for smaller gaps
        score = max(0, 100 - (avg_gap - 1) * 5)
        return round(score, 1)
    
    consistency_scores = df.groupby("Passbook number").apply(calculate_consistency)
    df["Consistency Score"] = df["Passbook number"].map(consistency_scores)
    df.loc[df["Passbook Txn Count"] <= 1, "Consistency Score"] = 0
    
    # ============================================================
    # SLAB CATEGORIZATION (Based on Avg Days Between Payments - Historical)
    # ============================================================
    
    config_gold = slab_config["gold"]
    config_silver = slab_config["silver"]
    
    is_gold = df["Metal Type"] == "Gold"
    is_silver = df["Metal Type"] == "Silver"
    is_one_time = df["Passbook Txn Count"] <= 1
    
    # Initialize with default
    df["Selected Slab"] = "Category C"
    
    # --- Use Avg Days Between Payments for categorization ---
    avg_gap = df["Avg Days Between Payments"]
    
    # Category A: Average gap is very short (Daily payers)
    mask_a_gold = is_gold & (avg_gap <= config_gold["slab1_max"]) & ~is_one_time & (avg_gap > 0)
    mask_a_silver = is_silver & (avg_gap <= config_silver["slab1_max"]) & ~is_one_time & (avg_gap > 0)
    df.loc[mask_a_gold | mask_a_silver, "Selected Slab"] = "Category A"
    
    # Category B: Average gap is moderate (Weekly/Monthly)
    mask_b_gold = is_gold & (avg_gap > config_gold["slab1_max"]) & (avg_gap <= config_gold["slab2_max"]) & ~is_one_time
    mask_b_silver = is_silver & (avg_gap > config_silver["slab1_max"]) & (avg_gap <= config_silver["slab2_max"]) & ~is_one_time
    df.loc[mask_b_gold | mask_b_silver, "Selected Slab"] = "Category B"
    
    # Category C: Average gap is long (Irregular)
    mask_c_gold = is_gold & (avg_gap > config_gold["slab2_max"]) & ~is_one_time
    mask_c_silver = is_silver & (avg_gap > config_silver["slab2_max"]) & ~is_one_time
    df.loc[mask_c_gold | mask_c_silver, "Selected Slab"] = "Category C"
    
    # Category D: One-time
    df.loc[is_one_time, "Selected Slab"] = "Category D"
    
    # Add label (built from the ACTUAL configured thresholds)
    dynamic_categories, dynamic_descriptions = build_slab_labels(slab_config)
    df["Category Label"] = df["Selected Slab"].map(dynamic_categories)
    df["Category Description"] = df["Selected Slab"].map(dynamic_descriptions)
    
    return df

# ============================================================
# CUSTOMER SUMMARY WITH PAYMENT FREQUENCY - FIXED
# ============================================================
def get_customer_summary_with_frequency(df):
    """Get customer summary with payment frequency metrics"""
    if df is None or df.empty:
        return None
    
    summary_data = {}
    
    for metal_type in ["Gold", "Silver"]:
        metal_df = df[df["Metal Type"] == metal_type]
        if metal_df.empty:
            continue
        
        # Calculate overall payment period (first to last payment)
        def get_payment_period(group):
            if len(group) <= 1:
                return 0  # No period for single payment
            first_date = group["Paid Date"].min()
            last_date = group["Paid Date"].max()
            return (last_date - first_date).days
        
        payment_periods = metal_df.groupby("Passbook number").apply(get_payment_period)
        
        # Main aggregation with frequency metrics
        passbook_summary = metal_df.groupby("Passbook number", as_index=False).agg({
            "Saved Amount": "sum",
            "Installment number": "max",
            "Customer Name": "first",
            "Customer Phone Number": "first",
            "Passbook Txn Count": "first",
            "Avg Days Between Payments": "first",
            "Max Gap Days": "first",
            "Days Since Last Payment To Today": "first",
            "Category Label (Last Payment)": "first",
            "Consistency Score": "first",
            "Selected Slab": "first",
            "Category Label": "first",
            "Category Description": "first",
            "Last Payment Date": "first",
            "Unique Payment Days": "first"
        })
        
        passbook_summary.columns = [
            "Passbook number",
            "Total Saved",
            "Latest Installment",
            "Customer Name",
            "Customer Phone Number",
            "Transaction Count",
            "Avg Days Between Payments",
            "Max Gap Days",
            "Days Since Last Payment To Today",
            "Category Label (Last Payment)",
            "Consistency Score",
            "Selected Slab",
            "Category Label",
            "Category Description",
            "Last Payment Date",
            "Unique Payment Days"
        ]
        
        # Add new columns for download
        passbook_summary["Overall Payment Period"] = passbook_summary["Passbook number"].map(payment_periods)
        
        # FIXED: Skipped Days calculation with proper handling for one-time payers
        # For one-time payers, skipped days should be 0 (they haven't skipped any days)
        # For multi-payment customers: Overall Payment Period - Unique Payment Days
        passbook_summary["Skipped Days"] = np.where(
            passbook_summary["Transaction Count"] <= 1,
            0,  # No skipped days for one-time customers
            passbook_summary["Overall Payment Period"] - passbook_summary["Unique Payment Days"]
        )
        
        # Ensure no negative skipped days (possible due to data issues)
        passbook_summary["Skipped Days"] = passbook_summary["Skipped Days"].clip(lower=0)
        
        # Rename Category Label to Category Label (Avg Days Payments) for download
        if "Category Label" in passbook_summary.columns:
            passbook_summary["Category Label (Avg Days Payments)"] = passbook_summary["Category Label"]
        
        # Calculate Total Value and Avg Ticket
        passbook_summary["Total Value"] = passbook_summary["Total Saved"]
        passbook_summary["Avg Ticket"] = np.where(
            passbook_summary["Transaction Count"] > 0,
            passbook_summary["Total Saved"] / passbook_summary["Transaction Count"],
            0
        ).round(0)
        
        # ============================================================
        # Payment Status based on Avg Days Between Payments (Historical)
        # ============================================================
        def get_payment_status(row):
            """Determine payment status based on historical payment behavior (Avg Days Between Payments)"""
            if row["Transaction Count"] <= 1:
                return "Trial"
            
            avg_gap = row["Avg Days Between Payments"]
            if pd.isna(avg_gap) or avg_gap == 0:
                return "Trial"
            elif avg_gap <= 2:
                return "🌟 Daily (Excellent)"
            elif avg_gap <= 7:
                return "✅ Weekly (Good)"
            elif avg_gap <= 15:
                return "⚠️ Bi-weekly (Moderate)"
            elif avg_gap <= 30:
                return "⚠️ Monthly (Low)"
            else:
                return "🔴 Irregular (At Risk)"
        
        # ============================================================
        # Payment Behavior column - shows unique payment days
        # ============================================================
        def get_payment_behavior(row):
            """Determine payment behavior based on unique payment days count"""
            if row["Transaction Count"] <= 1:
                return "One-time"
            
            unique_days = row["Unique Payment Days"]
            if pd.isna(unique_days) or unique_days == 0:
                return "No data"
            
            return f"{int(unique_days)} days"
        
        passbook_summary["Payment Status"] = passbook_summary.apply(get_payment_status, axis=1)
        passbook_summary["Payment Behavior"] = passbook_summary.apply(get_payment_behavior, axis=1)
        
        # Sort and rank
        passbook_summary = passbook_summary.sort_values("Total Saved", ascending=False).reset_index(drop=True)
        passbook_summary["Rank"] = range(1, len(passbook_summary) + 1)
        
        # Format Last Payment Date
        passbook_summary["Last Payment Date"] = pd.to_datetime(passbook_summary["Last Payment Date"]).dt.strftime("%d-%m-%Y")
        
        summary_data[metal_type] = passbook_summary
    
    return summary_data

# ============================================================
# MULTI-PASSBOOK CUSTOMER ANALYSIS (SHEET 2)
# ============================================================
def get_multi_passbook_customers(passbook_summary):
    """Identify customers with multiple passbooks (Sheet 2)"""
    if passbook_summary is None:
        return None
    
    all_customers = []
    
    for metal_type in ["Gold", "Silver"]:
        if metal_type not in passbook_summary:
            continue
        
        df = passbook_summary[metal_type].copy()
        
        # Group by Customer Phone Number
        grouped = df.groupby("Customer Phone Number")
        
        for phone, group in grouped:
            if len(group) > 1:  # Multiple passbooks for same phone
                # Get customer name (use the most common name)
                customer_name = group["Customer Name"].mode().iloc[0] if not group["Customer Name"].mode().empty else group["Customer Name"].iloc[0]
                
                # Get all passbook numbers
                passbook_numbers = group["Passbook number"].tolist()
                
                # Calculate aggregates
                total_saved = group["Total Saved"].sum()
                total_transactions = group["Transaction Count"].sum()
                avg_ticket = total_saved / total_transactions if total_transactions > 0 else 0
                
                # Average of payment metrics
                avg_avg_days = group["Avg Days Between Payments"].mean()
                avg_consistency = group["Consistency Score"].mean()
                max_gap = group["Max Gap Days"].max()
                
                # Latest last payment date
                latest_payment_date = group["Last Payment Date"].max()
                
                # Days since last payment (use the most recent)
                days_since_last = group["Days Since Last Payment To Today"].min()
                
                # Category labels (use most common or most recent)
                category_avg = group["Category Label (Avg Days Payments)"].mode().iloc[0] if not group["Category Label (Avg Days Payments)"].mode().empty else "N/A"
                category_last = group["Category Label (Last Payment)"].mode().iloc[0] if not group["Category Label (Last Payment)"].mode().empty else "N/A"
                
                # Payment behavior
                payment_behaviors = group["Payment Behavior"].tolist()
                payment_behavior = ", ".join(payment_behaviors)
                
                # Payment status (use the most severe)
                status_priority = {
                    "🔴 Irregular (At Risk)": 5,
                    "⚠️ Monthly (Low)": 4,
                    "⚠️ Bi-weekly (Moderate)": 3,
                    "✅ Weekly (Good)": 2,
                    "🌟 Daily (Excellent)": 1,
                    "Trial": 0
                }
                
                payment_status = "Trial"
                max_priority = -1
                for status in group["Payment Status"].unique():
                    priority = status_priority.get(status, 0)
                    if priority > max_priority:
                        max_priority = priority
                        payment_status = status
                
                # Overall Payment Period (average)
                overall_period = group["Overall Payment Period"].mean()
                
                # Skipped Days (sum)
                skipped_days = group["Skipped Days"].sum()
                
                all_customers.append({
                    "Customer Name": customer_name,
                    "Customer Phone Number": phone,
                    "Number of Passbooks": len(group),
                    "Passbook Numbers": ", ".join(passbook_numbers),
                    "Total Saved (All Passbooks)": total_saved,
                    "Avg Ticket (All Passbooks)": avg_ticket,
                    "Payment Behavior": payment_behavior,
                    "Skipped Days (Total)": skipped_days,
                    "Overall Payment Period (Avg)": overall_period,
                    "Avg Days Between Payments (Avg)": avg_avg_days,
                    "Category Label (Avg Days Payments)": category_avg,
                    "Consistency Score (Avg)": avg_consistency,
                    "Max Gap Days (Max)": max_gap,
                    "Last Payment Date (Latest)": latest_payment_date,
                    "Days Since Last Payment To Today": days_since_last,
                    "Category Label (Last Payment)": category_last,
                    "Total Transactions (All Passbooks)": total_transactions,
                    "Payment Status": payment_status
                })
    
    if all_customers:
        result_df = pd.DataFrame(all_customers)
        result_df = result_df.sort_values("Total Saved (All Passbooks)", ascending=False).reset_index(drop=True)
        return result_df
    else:
        return pd.DataFrame()  # Return empty DataFrame if no multi-passbook customers

# ============================================================
# SINGLE PASSBOOK CUSTOMER ANALYSIS (SHEET 3)
# ============================================================
def get_single_passbook_customers(passbook_summary):
    """Identify customers with single passbook (Sheet 3)"""
    if passbook_summary is None:
        return None
    
    all_customers = []
    
    for metal_type in ["Gold", "Silver"]:
        if metal_type not in passbook_summary:
            continue
        
        df = passbook_summary[metal_type].copy()
        
        # Group by Customer Phone Number
        grouped = df.groupby("Customer Phone Number")
        
        for phone, group in grouped:
            if len(group) == 1:  # Single passbook for this phone
                customer = group.iloc[0].to_dict()
                all_customers.append(customer)
    
    if all_customers:
        result_df = pd.DataFrame(all_customers)
        result_df = result_df.sort_values("Total Saved", ascending=False).reset_index(drop=True)
        result_df["Rank"] = range(1, len(result_df) + 1)
        return result_df
    else:
        return pd.DataFrame()  # Return empty DataFrame if no single-passbook customers

# ============================================================
# ACCURATE CATEGORY SUMMARY - UPDATED with Avg Per Ticket
# ============================================================
def get_accurate_category_summary(df):
    """Get category summary with no duplicate counting"""
    if df is None or df.empty:
        return None
    
    summary = []
    
    for metal_type in ["Gold", "Silver"]:
        metal_df = df[df["Metal Type"] == metal_type]
        if metal_df.empty:
            continue
        
        for slab in SLAB_ORDER:
            slab_df = metal_df[metal_df["Selected Slab"] == slab]
            if slab_df.empty:
                continue
            
            slab_passbooks = slab_df["Passbook number"].unique()
            unique_count = len(slab_passbooks)
            total_transactions = len(slab_df)
            total_saved = slab_df["Saved Amount"].sum()
            
            # Calculate Avg Per Ticket = Total Saved / Total Transactions
            avg_per_ticket = total_saved / total_transactions if total_transactions > 0 else 0
            
            slab_description = slab_df["Category Label"].iloc[0] if "Category Label" in slab_df.columns else slab

            summary.append({
                "Metal Type": metal_type,
                "Category": slab,
                "Description": slab_description,
                "Unique Passbooks": unique_count,
                "Total Transactions": total_transactions,
                "Total Saved": total_saved,
                "Avg Per Ticket": avg_per_ticket
            })
    
    return pd.DataFrame(summary)

# ============================================================
# LAST PAYMENT CATEGORY SUMMARY - UPDATED with Avg Per Ticket
# ============================================================
def get_last_payment_category_summary(df, slab_config):
    """Get category summary based on Last Payment Category with proper labels"""
    if df is None or df.empty:
        return None
    
    summary = []
    
    # Get the descriptions from slab config
    categories, _ = build_slab_labels(slab_config)
    
    for metal_type in ["Gold", "Silver"]:
        metal_df = df[df["Metal Type"] == metal_type]
        if metal_df.empty:
            continue
        
        # Group by the actual category labels (the descriptive ones)
        for cat_label in SLAB_ORDER:
            # Get the full descriptive label for this category
            full_label = categories.get(cat_label, cat_label)
            
            # Filter data where Category Label (Last Payment) matches the full label
            cat_df = metal_df[metal_df["Category Label (Last Payment)"] == full_label]
            if cat_df.empty:
                continue
            
            cat_passbooks = cat_df["Passbook number"].unique()
            unique_count = len(cat_passbooks)
            total_transactions = len(cat_df)
            total_saved = cat_df["Saved Amount"].sum()
            
            # Calculate Avg Per Ticket = Total Saved / Total Transactions
            avg_per_ticket = total_saved / total_transactions if total_transactions > 0 else 0
            
            # Get description - use the same format as Historical Average
            desc = categories.get(cat_label, cat_label)

            summary.append({
                "Metal Type": metal_type,
                "Category": cat_label,
                "Description": desc,
                "Unique Passbooks": unique_count,
                "Total Transactions": total_transactions,
                "Total Saved": total_saved,
                "Avg Per Ticket": avg_per_ticket
            })
    
    return pd.DataFrame(summary)

# ============================================================
# SAMPLE DAILY SCHEME DATA
# ============================================================
def create_daily_scheme_data():
    """Generate sample daily scheme data with payment patterns"""
    np.random.seed(42)
    
    # Create 50 passbooks with different payment patterns
    passbooks = [f"PB{str(i).zfill(4)}" for i in range(1, 51)]
    
    data = []
    
    # Create some customers with multiple passbooks
    multi_customer_phones = {
        "+917000000001": ["PB0001", "PB0010", "PB0020"],
        "+917000000002": ["PB0005", "PB0015"],
        "+917000000003": ["PB0025", "PB0030", "PB0035", "PB0040"]
    }
    
    for pb in passbooks:
        # Check if this passbook belongs to a multi-passbook customer
        phone = None
        for multi_phone, pbs in multi_customer_phones.items():
            if pb in pbs:
                phone = multi_phone
                break
        
        if phone is None:
            # Regular customer with single passbook
            phone = f"+91{np.random.randint(7000000000, 9999999999)}"
        
        # Randomly assign payment pattern
        pattern = np.random.choice(["daily", "weekly", "monthly", "one_time", "irregular"], 
                                   p=[0.3, 0.3, 0.2, 0.1, 0.1])
        
        if pattern == "daily":
            # Daily payments for 30-60 days
            days = np.random.randint(30, 61)
            installments = list(range(1, days + 1))
            end_date = datetime.now().date() - timedelta(days=np.random.randint(0, 7))
            start_date = end_date - timedelta(days=days)
            dates = pd.date_range(start_date, periods=days, freq="D")
            
        elif pattern == "weekly":
            # Weekly payments for 30-90 days
            days = np.random.randint(30, 91)
            installments = list(range(1, days + 1, 7))
            end_date = datetime.now().date() - timedelta(days=np.random.randint(7, 30))
            start_date = end_date - timedelta(days=days)
            dates = pd.date_range(start_date, periods=len(installments), freq="7D")
            
        elif pattern == "monthly":
            # Monthly payments for 90-180 days
            days = np.random.randint(90, 181)
            installments = list(range(1, days + 1, 30))
            end_date = datetime.now().date() - timedelta(days=np.random.randint(30, 60))
            start_date = end_date - timedelta(days=days)
            dates = pd.date_range(start_date, periods=len(installments), freq="30D")
            
        elif pattern == "one_time":
            # One-time payment (long ago)
            installments = [1]
            dates = [datetime.now().date() - timedelta(days=np.random.randint(30, 180))]
            
        else:  # irregular
            # Irregular payments with random gaps
            num_payments = np.random.randint(5, 20)
            installments = sorted(np.random.choice(range(1, 91), num_payments, replace=False))
            end_date = datetime.now().date() - timedelta(days=np.random.randint(60, 200))
            dates = pd.date_range(end_date - timedelta(days=90), periods=num_payments, freq="D") + pd.Timedelta(days=np.random.randint(0, 90, num_payments))
        
        metal_type = np.random.choice(["Gold", "Silver"], p=[0.7, 0.3])
        amount = np.random.uniform(1000, 50000, len(installments)).round(2)
        reward = np.random.uniform(50, 5000, len(installments)).round(2)
        
        # Customer name based on phone
        customer_name = f"Customer_{phone[-4:]}"
        
        for i, (inst, date, amt, rew) in enumerate(zip(installments, dates, amount, reward)):
            data.append({
                "Id": len(data) + 1,
                "Passbook number": pb,
                "Metal Type": metal_type,
                "Installment number": inst,
                "Paid Date": date,
                "Saved Amount": amt,
                "Reward Amount": rew,
                "Customer Name": customer_name,
                "Customer Phone Number": phone,
                "Status": "Completed"
            })
    
    return pd.DataFrame(data)

# ============================================================
# UI HELPERS
# ============================================================
def format_large_number(num):
    """Format large numbers for display"""
    if num >= 1e7:
        return f"{num/1e7:.1f}Cr"
    elif num >= 1e5:
        return f"{num/1e5:.1f}L"
    elif num >= 1e3:
        return f"{num/1e3:.1f}K"
    else:
        return str(int(num))

def render_metrics(df):
    """Render metrics with professional cards"""
    total_customers = df["Passbook number"].nunique()
    gold_customers = df[df["Metal Type"] == "Gold"]["Passbook number"].nunique()
    silver_customers = df[df["Metal Type"] == "Silver"]["Passbook number"].nunique()
    total_saved = df["Saved Amount"].sum()
    total_transactions = len(df)
    avg_consistency = df["Consistency Score"].mean()
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">📒 Total Passbooks</div>
            <div class="metric-value">{total_customers:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">🥇 Gold</div>
            <div class="metric-value">{gold_customers:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">🥈 Silver</div>
            <div class="metric-value">{silver_customers:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">💰 Total Saved</div>
            <div class="metric-value">₹{format_large_number(total_saved)}</div>
        </div>
        """, unsafe_allow_html=True)
    with col5:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">📝 Transactions</div>
            <div class="metric-value">{total_transactions:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with col6:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">⭐ Avg Consistency</div>
            <div class="metric-value">{avg_consistency:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)

def render_paginated_dataframe(df, page_size=50, key_suffix=""):
    """Render dataframe with pagination for large datasets"""
    if df is None or df.empty:
        st.info("No data to display")
        return
    
    total_rows = len(df)
    total_pages = (total_rows + page_size - 1) // page_size
    
    if total_pages == 1:
        st.dataframe(df, width="stretch", height=400)
        return
    
    col1, col2, col3 = st.columns([2, 3, 2])
    with col1:
        page = st.number_input(
            "Page", 
            min_value=1, 
            max_value=total_pages, 
            value=1,
            key=f"page_num_{key_suffix}"
        )
    with col2:
        st.write(f"Showing page {page} of {total_pages} ({total_rows:,} rows)")
    with col3:
        page_size = st.selectbox(
            "Rows per page",
            [20, 50, 100, 200],
            index=1,
            key=f"page_size_{key_suffix}"
        )
    
    start_idx = (page - 1) * page_size
    end_idx = min(start_idx + page_size, total_rows)
    df_page = df.iloc[start_idx:end_idx]
    
    st.dataframe(df_page, width="stretch", height=400)
    progress = (page / total_pages) * 100
    st.progress(progress / 100)

# ------------------------------------------------------------
# PERF FIX: get_download_link - now accepts filtered dataframe
# ------------------------------------------------------------
@st.cache_data(ttl=3600, show_spinner=False, hash_funcs={pd.DataFrame: id})
def get_download_link(df, filename, file_format="csv"):
    """Generate download link with filtered columns"""
    if df is None or df.empty:
        return "No data available"
    
    df_export = df.copy()
    
    # Round Avg Days Between Payments to whole number
    if "Avg Days Between Payments" in df_export.columns:
        df_export["Avg Days Between Payments"] = df_export["Avg Days Between Payments"].round(0).astype(int)
    
    if "Max Gap Days" in df_export.columns:
        df_export["Max Gap Days"] = df_export["Max Gap Days"].fillna(0).astype(int)
    
    if "Days Since Last Payment To Today" in df_export.columns:
        df_export["Days Since Last Payment To Today"] = df_export["Days Since Last Payment To Today"].fillna(0).astype(int)
    
    if "Consistency Score" in df_export.columns:
        df_export["Consistency Score"] = df_export["Consistency Score"].round(1)
    
    # Updated to include "Avg Per Ticket"
    for col in ["Total Saved", "Total Value", "Avg Ticket", "Avg Per Ticket"]:
        if col in df_export.columns:
            df_export[col] = df_export[col].round(0)
    
    if "Last Payment Date" in df_export.columns:
        df_export["Last Payment Date"] = pd.to_datetime(df_export["Last Payment Date"]).dt.strftime("%d-%m-%Y")
    
    if "Unique Payment Days" in df_export.columns:
        df_export["Unique Payment Days"] = df_export["Unique Payment Days"].fillna(0).astype(int)
    
    if file_format == "csv":
        csv_data = df_export.to_csv(index=False)
        b64 = base64.b64encode(csv_data.encode()).decode()
        return f'<a href="data:file/csv;base64,{b64}" download="{filename}.csv" class="download-btn">📥 Download {filename}.csv</a>'
    else:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_export.to_excel(writer, sheet_name="Sheet1", index=False)
        excel_data = output.getvalue()
        b64 = base64.b64encode(excel_data).decode()
        return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}.xlsx" class="download-btn">📥 Download {filename}.xlsx</a>'

# ============================================================
# MAIN APPLICATION
# ============================================================
def main():
    # Custom header
    st.markdown("""
    <div class="app-header">
        <div class="app-title">📒 Daily Scheme Payment Frequency Tracker</div>
        <div class="app-subtitle">Track payment behavior, consistency scores, and customer categories</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        
        st.markdown("#### Category Limits (Days)")
        col1, col2 = st.columns(2)
        with col1:
            cat_a = st.number_input("Category A (≤ days)", 1, 30, 10, key="cat_a_large")
        with col2:
            cat_b = st.number_input("Category B (≤ days)", 8, 50, 50, key="cat_b_large")
        
        if st.button("🔄 Update Limits", type="primary", width="stretch"):
            if cat_b > cat_a:
                st.session_state.slab_config = {
                    "gold": {"slab1_max": cat_a, "slab2_min": cat_a+1, "slab2_max": cat_b, "slab3_min": cat_b+1},
                    "silver": {"slab1_max": cat_a, "slab2_min": cat_a+1, "slab2_max": cat_b, "slab3_min": cat_b+1}
                }
                if st.session_state.raw_data is not None:
                    process_daily_scheme_data.clear()
                    with st.spinner("Re-processing with new category limits..."):
                        reprocessed = process_daily_scheme_data(
                            st.session_state.raw_data, st.session_state.slab_config
                        )
                        passbook_summary = get_customer_summary_with_frequency(reprocessed)
                        st.session_state.passbook_summary = passbook_summary
                        st.session_state.category_summary = get_accurate_category_summary(reprocessed)
                        st.session_state.last_payment_category_summary = get_last_payment_category_summary(
                            reprocessed, st.session_state.slab_config
                        )
                        st.session_state.multi_passbook_customers = get_multi_passbook_customers(passbook_summary)
                        st.session_state.single_passbook_customers = get_single_passbook_customers(passbook_summary)
                    st.session_state.data = reprocessed
                    st.session_state.processed_data = reprocessed
                    st.success("✅ Updated!")
                    st.rerun()
                else:
                    st.warning(
                        "⚠️ No data loaded. Please upload data first, "
                        "then update the limits."
                    )
            else:
                st.error("Category B limit must be greater than Category A limit")
        
        st.markdown("---")
        
        st.markdown("### 📤 Upload Data")
        uploaded_file = st.file_uploader("Choose CSV or Excel", type=["csv", "xlsx", "xls"])
        
        if uploaded_file:
            try:
                with st.spinner("Loading file..."):
                    if uploaded_file.name.lower().endswith(".csv"):
                        df = pd.read_csv(uploaded_file, low_memory=False)
                    else:
                        df = pd.read_excel(uploaded_file)
                
                st.success(f"✅ {len(df):,} transactions loaded")
                
                with st.expander("Preview (first 5 rows)"):
                    st.dataframe(df.head(5), width="stretch")
                
                if st.button("🔄 Process Data", type="primary", width="stretch"):
                    with st.spinner(f"Processing {len(df):,} transactions..."):
                        processed = process_daily_scheme_data(df, st.session_state.slab_config)
                        st.session_state.raw_data = df
                        st.session_state.data = processed
                        st.session_state.processed_data = processed
                        passbook_summary = get_customer_summary_with_frequency(processed)
                        st.session_state.passbook_summary = passbook_summary
                        st.session_state.category_summary = get_accurate_category_summary(processed)
                        st.session_state.last_payment_category_summary = get_last_payment_category_summary(
                            processed, st.session_state.slab_config
                        )
                        st.session_state.multi_passbook_customers = get_multi_passbook_customers(passbook_summary)
                        st.session_state.single_passbook_customers = get_single_passbook_customers(passbook_summary)
                    st.success("✅ Data processed successfully!")
                    st.rerun()
                    
            except Exception as e:
                st.error(f"Error: {str(e)}")
        
        st.markdown("---")
        
        if st.button("📋 Load Sample Data", width="stretch"):
            sample = create_daily_scheme_data()
            processed = process_daily_scheme_data(sample, st.session_state.slab_config)
            st.session_state.raw_data = sample
            st.session_state.data = processed
            st.session_state.processed_data = processed
            passbook_summary = get_customer_summary_with_frequency(processed)
            st.session_state.passbook_summary = passbook_summary
            st.session_state.category_summary = get_accurate_category_summary(processed)
            st.session_state.last_payment_category_summary = get_last_payment_category_summary(
                processed, st.session_state.slab_config
            )
            st.session_state.multi_passbook_customers = get_multi_passbook_customers(passbook_summary)
            st.session_state.single_passbook_customers = get_single_passbook_customers(passbook_summary)
            st.success("✅ Sample data loaded!")
            st.rerun()
        
        st.markdown("---")
        st.caption("💡 Each Passbook = 1 Unique Customer")
        st.caption("📅 Tracks payment frequency for daily scheme")
    
    # Main content
    if st.session_state.data is None:
        st.info("👈 Upload data or load sample to begin")
        return
    
    df = st.session_state.data
    
    if st.session_state.passbook_summary is None or st.session_state.category_summary is None or st.session_state.last_payment_category_summary is None:
        with st.spinner("Generating summaries..."):
            passbook_summary = get_customer_summary_with_frequency(df)
            st.session_state.passbook_summary = passbook_summary
            st.session_state.category_summary = get_accurate_category_summary(df)
            st.session_state.last_payment_category_summary = get_last_payment_category_summary(
                df, st.session_state.slab_config
            )
            st.session_state.multi_passbook_customers = get_multi_passbook_customers(passbook_summary)
            st.session_state.single_passbook_customers = get_single_passbook_customers(passbook_summary)
    
    passbook_summary = st.session_state.passbook_summary
    category_summary = st.session_state.category_summary
    last_payment_category_summary = st.session_state.last_payment_category_summary
    multi_passbook_customers = st.session_state.multi_passbook_customers
    single_passbook_customers = st.session_state.single_passbook_customers
    
    render_metrics(df)
    
    # ============================================================
    # TABS - FIXED: Exactly 7 tabs to match unpacking
    # ============================================================
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📊 Dashboard", 
        "📒 Passbooks", 
        "👥 Multi-Passbook",
        "👤 Single Passbook",
        "📋 Category Analysis",
        "📊 Charts",
        "📤 Export"
    ])
    
    # ============================================================
    # TAB 1: DASHBOARD
    # ============================================================
    with tab1:
        st.markdown('<div class="section-header">📊 Payment Frequency Dashboard</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            if passbook_summary:
                status_data = []
                for metal_type in ["Gold", "Silver"]:
                    if metal_type in passbook_summary:
                        status_counts = passbook_summary[metal_type]["Payment Status"].value_counts()
                        for status, count in status_counts.items():
                            status_data.append({
                                "Metal": metal_type,
                                "Status": status,
                                "Count": count
                            })
                
                if status_data:
                    status_df = pd.DataFrame(status_data)
                    fig = px.sunburst(
                        status_df,
                        path=["Metal", "Status"],
                        values="Count",
                        title="Payment Status Distribution",
                        color="Status",
                        color_discrete_sequence=px.colors.qualitative.Set3
                    )
                    fig.update_traces(textinfo="label+percent parent")
                    fig.update_layout(
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif")
                    )
                    st.plotly_chart(fig, width="stretch")
        
        with col2:
            if category_summary is not None and not category_summary.empty:
                cat_data = category_summary.groupby("Category")["Unique Passbooks"].sum().reset_index()
                if not cat_data.empty:
                    fig = px.pie(
                        cat_data,
                        values="Unique Passbooks",
                        names="Category",
                        title="Passbook Distribution by Payment Frequency (Historical Avg)",
                        color_discrete_sequence=px.colors.qualitative.Set3
                    )
                    fig.update_traces(textposition="inside", textinfo="percent+label")
                    fig.update_layout(
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif")
                    )
                    st.plotly_chart(fig, width="stretch")
        
        # Customer Distribution by Passbook Count
        st.markdown('<div class="section-header">👥 Customer Distribution</div>', unsafe_allow_html=True)
        
        if multi_passbook_customers is not None and single_passbook_customers is not None:
            col1, col2, col3 = st.columns(3)
            
            total_unique_phones = 0
            if single_passbook_customers is not None:
                total_unique_phones += len(single_passbook_customers)
            if multi_passbook_customers is not None:
                total_unique_phones += len(multi_passbook_customers)
            
            total_passbooks = 0
            if passbook_summary:
                for metal in ["Gold", "Silver"]:
                    if metal in passbook_summary:
                        total_passbooks += len(passbook_summary[metal])
            
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">📱 Unique Customers (Phone)</div>
                    <div class="metric-value">{total_unique_phones:,}</div>
                </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">📒 Total Passbooks</div>
                    <div class="metric-value">{total_passbooks:,}</div>
                </div>
                """, unsafe_allow_html=True)
            with col3:
                multi_count = len(multi_passbook_customers) if multi_passbook_customers is not None else 0
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">👥 Multi-Passbook Customers</div>
                    <div class="metric-value">{multi_count:,}</div>
                    <div class="metric-sub">{multi_count/total_unique_phones*100:.1f}% of customers</div>
                </div>
                """, unsafe_allow_html=True)
    
    # ============================================================
    # TAB 2: PASSBOOKS (Originally "Customers")
    # ============================================================
    with tab2:
        st.markdown('<div class="section-header">📒 Passbook Payment Frequency</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="info-box">Tracking {df["Passbook number"].nunique():,} unique passbooks</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            metal_filter = st.multiselect(
                "Metal Type",
                ["Gold", "Silver"],
                default=["Gold", "Silver"],
                key="metal_filter_freq"
            )
        
        with col2:
            status_filter = st.multiselect(
                "Payment Status",
                ["🌟 Daily (Excellent)", "✅ Weekly (Good)", "⚠️ Bi-weekly (Moderate)", 
                 "⚠️ Monthly (Low)", "🔴 Irregular (At Risk)", "Trial"],
                default=["🌟 Daily (Excellent)", "✅ Weekly (Good)", "⚠️ Bi-weekly (Moderate)", 
                        "⚠️ Monthly (Low)", "🔴 Irregular (At Risk)", "Trial"],
                key="status_filter_freq"
            )
        
        for metal_type in metal_filter:
            if metal_type not in passbook_summary:
                continue
            
            df_display = passbook_summary[metal_type].copy()
            df_display = df_display[df_display["Payment Status"].isin(status_filter)]
            
            if df_display.empty:
                continue
            
            st.markdown(f"### {'🥇' if metal_type == 'Gold' else '🥈'} {metal_type} Passbooks ({len(df_display):,})")
            
            display_cols = [c for c in CUSTOMER_SUMMARY_COLUMNS if c in df_display.columns]
            df_formatted = df_display[display_cols].copy()
            
            for col in ["Transaction Count"]:
                if col in df_formatted:
                    df_formatted[col] = df_formatted[col].apply(lambda x: f"{int(x):,}")
            
            for col in ["Total Saved", "Total Value"]:
                if col in df_formatted:
                    df_formatted[col] = df_formatted[col].apply(lambda x: f"₹{int(x):,}")
            
            for col in ["Avg Ticket"]:
                if col in df_formatted:
                    df_formatted[col] = df_formatted[col].apply(lambda x: f"₹{int(x):,}")
            
            if "Avg Days Between Payments" in df_formatted:
                df_formatted["Avg Days Between Payments"] = df_formatted["Avg Days Between Payments"].apply(
                    lambda x: f"{int(x)}" if not pd.isna(x) else "N/A"
                )
            
            if "Max Gap Days" in df_formatted:
                df_formatted["Max Gap Days"] = df_formatted["Max Gap Days"].apply(
                    lambda x: f"{int(x)}" if not pd.isna(x) else "N/A"
                )
            
            if "Days Since Last Payment To Today" in df_formatted:
                df_formatted["Days Since Last Payment To Today"] = df_formatted["Days Since Last Payment To Today"].apply(
                    lambda x: f"{int(x)} days" if not pd.isna(x) else "N/A"
                )
            
            if "Consistency Score" in df_formatted:
                df_formatted["Consistency Score"] = df_formatted["Consistency Score"].apply(
                    lambda x: f"{x:.1f}%" if not pd.isna(x) else "N/A"
                )
            
            render_paginated_dataframe(df_formatted, page_size=st.session_state.page_size, key_suffix=metal_type)
            
            col1, col2 = st.columns(2)
            with col1:
                download_cols = [c for c in CUSTOMER_SUMMARY_COLUMNS if c in df_display.columns]
                download_df = df_display[download_cols].copy()
                ordered_cols = [c for c in CUSTOMER_SUMMARY_COLUMNS if c in download_df.columns]
                if len(ordered_cols) > 0:
                    download_df = download_df[ordered_cols]
                st.markdown(
                    get_download_link(download_df, f"{metal_type}_passbooks_frequency"),
                    unsafe_allow_html=True
                )
            with col2:
                st.markdown(
                    get_download_link(download_df, f"{metal_type}_passbooks_frequency", "excel"),
                    unsafe_allow_html=True
                )
            st.markdown("---")
    
    # ============================================================
    # TAB 3: MULTI-PASSBOOK CUSTOMERS (SHEET 2)
    # ============================================================
    with tab3:
        st.markdown('<div class="section-header">👥 Multi-Passbook Customers</div>', unsafe_allow_html=True)
        st.markdown('<div class="highlight-box">📌 Customers with multiple passbooks (same phone number has >1 passbook)</div>', unsafe_allow_html=True)
        
        if multi_passbook_customers is not None and not multi_passbook_customers.empty:
            st.markdown(f'<div class="info-box">Found {len(multi_passbook_customers):,} customers with multiple passbooks</div>', unsafe_allow_html=True)
            
            # Display the dataframe
            display_cols = [c for c in MULTI_PASSBOOK_COLUMNS if c in multi_passbook_customers.columns]
            df_display = multi_passbook_customers[display_cols].copy()
            
            # Format for display
            for col in ["Total Saved (All Passbooks)", "Avg Ticket (All Passbooks)"]:
                if col in df_display:
                    df_display[col] = df_display[col].apply(lambda x: f"₹{int(x):,}")
            
            if "Total Transactions (All Passbooks)" in df_display:
                df_display["Total Transactions (All Passbooks)"] = df_display["Total Transactions (All Passbooks)"].apply(lambda x: f"{int(x):,}")
            
            if "Avg Days Between Payments (Avg)" in df_display:
                df_display["Avg Days Between Payments (Avg)"] = df_display["Avg Days Between Payments (Avg)"].apply(
                    lambda x: f"{int(x)}" if not pd.isna(x) else "N/A"
                )
            
            if "Consistency Score (Avg)" in df_display:
                df_display["Consistency Score (Avg)"] = df_display["Consistency Score (Avg)"].apply(
                    lambda x: f"{x:.1f}%" if not pd.isna(x) else "N/A"
                )
            
            if "Days Since Last Payment To Today" in df_display:
                df_display["Days Since Last Payment To Today"] = df_display["Days Since Last Payment To Today"].apply(
                    lambda x: f"{int(x)} days" if not pd.isna(x) else "N/A"
                )
            
            render_paginated_dataframe(df_display, page_size=st.session_state.page_size, key_suffix="multi")
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(
                    get_download_link(multi_passbook_customers, "multi_passbook_customers"),
                    unsafe_allow_html=True
                )
            with col2:
                st.markdown(
                    get_download_link(multi_passbook_customers, "multi_passbook_customers", "excel"),
                    unsafe_allow_html=True
                )
            
            # Additional insights
            st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
            st.markdown("#### 📊 Multi-Passbook Insights")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                total_multi_saved = multi_passbook_customers["Total Saved (All Passbooks)"].sum()
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">💰 Total Saved</div>
                    <div class="metric-value">₹{format_large_number(total_multi_saved)}</div>
                </div>
                """, unsafe_allow_html=True)
            with col2:
                avg_passbooks = multi_passbook_customers["Number of Passbooks"].mean()
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">📒 Avg Passbooks per Customer</div>
                    <div class="metric-value">{avg_passbooks:.1f}</div>
                </div>
                """, unsafe_allow_html=True)
            with col3:
                max_passbooks = multi_passbook_customers["Number of Passbooks"].max()
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">📒 Max Passbooks</div>
                    <div class="metric-value">{max_passbooks}</div>
                </div>
                """, unsafe_allow_html=True)
            
            # Distribution of number of passbooks
            passbook_dist = multi_passbook_customers["Number of Passbooks"].value_counts().sort_index()
            fig = px.bar(
                x=passbook_dist.index,
                y=passbook_dist.values,
                title="Distribution of Passbooks per Customer",
                labels={"x": "Number of Passbooks", "y": "Number of Customers"},
                color_discrete_sequence=["#0f3460"]
            )
            fig.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Inter, sans-serif")
            )
            st.plotly_chart(fig, width="stretch")
            
        else:
            st.info("No customers with multiple passbooks found. All customers have exactly 1 passbook.")
    
    # ============================================================
    # TAB 4: SINGLE PASSBOOK CUSTOMERS (SHEET 3)
    # ============================================================
    with tab4:
        st.markdown('<div class="section-header">👤 Single Passbook Customers</div>', unsafe_allow_html=True)
        st.markdown('<div class="highlight-box">📌 Customers with exactly 1 passbook (unique phone number = 1 passbook)</div>', unsafe_allow_html=True)
        
        if single_passbook_customers is not None and not single_passbook_customers.empty:
            st.markdown(f'<div class="info-box">Found {len(single_passbook_customers):,} customers with single passbook</div>', unsafe_allow_html=True)
            
            # Display the dataframe
            display_cols = [c for c in SINGLE_PASSBOOK_COLUMNS if c in single_passbook_customers.columns]
            df_display = single_passbook_customers[display_cols].copy()
            
            # Format for display
            for col in ["Transaction Count"]:
                if col in df_display:
                    df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}")
            
            for col in ["Total Saved", "Total Value"]:
                if col in df_display:
                    df_display[col] = df_display[col].apply(lambda x: f"₹{int(x):,}")
            
            for col in ["Avg Ticket"]:
                if col in df_display:
                    df_display[col] = df_display[col].apply(lambda x: f"₹{int(x):,}")
            
            if "Avg Days Between Payments" in df_display:
                df_display["Avg Days Between Payments"] = df_display["Avg Days Between Payments"].apply(
                    lambda x: f"{int(x)}" if not pd.isna(x) else "N/A"
                )
            
            if "Max Gap Days" in df_display:
                df_display["Max Gap Days"] = df_display["Max Gap Days"].apply(
                    lambda x: f"{int(x)}" if not pd.isna(x) else "N/A"
                )
            
            if "Days Since Last Payment To Today" in df_display:
                df_display["Days Since Last Payment To Today"] = df_display["Days Since Last Payment To Today"].apply(
                    lambda x: f"{int(x)} days" if not pd.isna(x) else "N/A"
                )
            
            if "Consistency Score" in df_display:
                df_display["Consistency Score"] = df_display["Consistency Score"].apply(
                    lambda x: f"{x:.1f}%" if not pd.isna(x) else "N/A"
                )
            
            render_paginated_dataframe(df_display, page_size=st.session_state.page_size, key_suffix="single")
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(
                    get_download_link(single_passbook_customers, "single_passbook_customers"),
                    unsafe_allow_html=True
                )
            with col2:
                st.markdown(
                    get_download_link(single_passbook_customers, "single_passbook_customers", "excel"),
                    unsafe_allow_html=True
                )
            
            # Additional insights
            st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
            st.markdown("#### 📊 Single Passbook Insights")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                total_single_saved = single_passbook_customers["Total Saved"].sum()
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">💰 Total Saved</div>
                    <div class="metric-value">₹{format_large_number(total_single_saved)}</div>
                </div>
                """, unsafe_allow_html=True)
            with col2:
                avg_single_saved = single_passbook_customers["Total Saved"].mean()
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">💳 Avg Saved per Customer</div>
                    <div class="metric-value">₹{format_large_number(avg_single_saved)}</div>
                </div>
                """, unsafe_allow_html=True)
            with col3:
                avg_consistency = single_passbook_customers["Consistency Score"].mean()
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">⭐ Avg Consistency</div>
                    <div class="metric-value">{avg_consistency:.1f}%</div>
                </div>
                """, unsafe_allow_html=True)
            
        else:
            st.info("No single passbook customers found. All customers have multiple passbooks.")
    
    # ============================================================
    # TAB 5: CATEGORY ANALYSIS
    # ============================================================
    with tab5:
        st.markdown('<div class="section-header">📋 Category Analysis</div>', unsafe_allow_html=True)
        
        # Create sub-tabs for different category views
        sub_tab1, sub_tab2 = st.tabs([
            "📊 Avg Category",
            "📊 Last Payment Category"
        ])
        
        # SUB-TAB 1: Historical Avg Category
        with sub_tab1:
            st.markdown("#### Average Payment Frequency")
            st.caption("Based on Avg Days Between Payments (payment behavior)")
            
            if category_summary is not None and not category_summary.empty:
                total_unique = df["Passbook number"].nunique()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Total Unique Passbooks</div>
                        <div class="metric-value">{total_unique:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col2:
                    total_unique_cats = category_summary["Unique Passbooks"].sum()
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Unique Across Categories</div>
                        <div class="metric-value">{total_unique_cats:,}</div>
                        <div class="metric-sub">✅ Matches total</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col3:
                    slab_nunique = df.groupby("Passbook number")["Selected Slab"].transform("nunique")
                    multi_count = df.loc[slab_nunique > 1, "Passbook number"].nunique()
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Multi-Category</div>
                        <div class="metric-value">{multi_count:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                display_cols = [c for c in CATEGORY_SUMMARY_COLUMNS if c in category_summary.columns]
                df_formatted = category_summary[display_cols].copy()
                
                for col in ["Unique Passbooks", "Total Transactions"]:
                    if col in df_formatted:
                        df_formatted[col] = df_formatted[col].apply(lambda x: f"{int(x):,}")
                
                for col in ["Total Saved", "Avg Per Ticket"]:
                    if col in df_formatted:
                        df_formatted[col] = df_formatted[col].apply(lambda x: f"₹{int(x):,}")
                
                st.dataframe(df_formatted, width="stretch")
                
                col1, col2 = st.columns(2)
                with col1:
                    download_df = category_summary[display_cols].copy()
                    st.markdown(
                        get_download_link(download_df, "category_summary_historical"),
                        unsafe_allow_html=True
                    )
                with col2:
                    st.markdown(
                        get_download_link(download_df, "category_summary_historical", "excel"),
                        unsafe_allow_html=True
                    )
                
                st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
                
                st.markdown("#### 📊 Avg Category Distribution")
                
                if not category_summary.empty:
                    fig = px.bar(
                        category_summary,
                        x="Category",
                        y="Unique Passbooks",
                        color="Metal Type",
                        title="Passbook Distribution by Average Payment Frequency",
                        barmode="group",
                        color_discrete_sequence=["#FFD700", "#C0C0C0"],
                        text="Unique Passbooks"
                    )
                    fig.update_traces(textposition="outside")
                    fig.update_layout(
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif")
                    )
                    st.plotly_chart(fig, width="stretch")
        
        # SUB-TAB 2: Last Payment Category
        with sub_tab2:
            st.markdown("#### Last Payment Category")
            st.caption("Based on Days Since Last Payment To Today (current status)")
            
            if last_payment_category_summary is not None and not last_payment_category_summary.empty:
                total_unique = df["Passbook number"].nunique()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Total Unique Passbooks</div>
                        <div class="metric-value">{total_unique:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col2:
                    total_unique_cats = last_payment_category_summary["Unique Passbooks"].sum()
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Unique Across Categories</div>
                        <div class="metric-value">{total_unique_cats:,}</div>
                        <div class="metric-sub">✅ Matches total</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col3:
                    lp_nunique = df.groupby("Passbook number")["Category Label (Last Payment)"].transform("nunique")
                    multi_count = df.loc[lp_nunique > 1, "Passbook number"].nunique()
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Multi-Category</div>
                        <div class="metric-value">{multi_count:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                display_cols = [c for c in LAST_PAYMENT_CATEGORY_SUMMARY_COLUMNS if c in last_payment_category_summary.columns]
                df_formatted = last_payment_category_summary[display_cols].copy()
                
                for col in ["Unique Passbooks", "Total Transactions"]:
                    if col in df_formatted:
                        df_formatted[col] = df_formatted[col].apply(lambda x: f"{int(x):,}")
                
                for col in ["Total Saved", "Avg Per Ticket"]:
                    if col in df_formatted:
                        df_formatted[col] = df_formatted[col].apply(lambda x: f"₹{int(x):,}")
                
                st.dataframe(df_formatted, width="stretch")
                
                col1, col2 = st.columns(2)
                with col1:
                    download_df = last_payment_category_summary[display_cols].copy()
                    st.markdown(
                        get_download_link(download_df, "category_summary_last_payment"),
                        unsafe_allow_html=True
                    )
                with col2:
                    st.markdown(
                        get_download_link(download_df, "category_summary_last_payment", "excel"),
                        unsafe_allow_html=True
                    )
                
                st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
                
                st.markdown("#### 📊 Last Payment Category Distribution")
                
                if not last_payment_category_summary.empty:
                    fig = px.bar(
                        last_payment_category_summary,
                        x="Category",
                        y="Unique Passbooks",
                        color="Metal Type",
                        title="Passbook Distribution by Last Payment Category",
                        barmode="group",
                        color_discrete_sequence=["#FFD700", "#C0C0C0"],
                        text="Unique Passbooks"
                    )
                    fig.update_traces(textposition="outside")
                    fig.update_layout(
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif")
                    )
                    st.plotly_chart(fig, width="stretch")
    
    # ============================================================
    # TAB 6: CHARTS
    # ============================================================
    with tab6:
        st.markdown('<div class="section-header">📊 Visual Analytics</div>', unsafe_allow_html=True)
        
        st.markdown("#### ⭐ Payment Consistency Distribution")
        
        if passbook_summary:
            all_customers = pd.DataFrame()
            for metal_type in ["Gold", "Silver"]:
                if metal_type in passbook_summary:
                    temp = passbook_summary[metal_type].copy()
                    temp["Metal"] = metal_type
                    all_customers = pd.concat([all_customers, temp], ignore_index=True)
            
            if not all_customers.empty:
                fig = px.histogram(
                    all_customers,
                    x="Consistency Score",
                    color="Metal",
                    title="Payment Consistency Score Distribution",
                    nbins=20,
                    color_discrete_sequence=["#FFD700", "#C0C0C0"]
                )
                fig.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(family="Inter, sans-serif"),
                    bargap=0.1
                )
                st.plotly_chart(fig, width="stretch")
        
        if multi_passbook_customers is not None and not multi_passbook_customers.empty:
            st.markdown("#### 📊 Multi-Passbook Customer Analysis")
            
            fig = px.scatter(
                multi_passbook_customers,
                x="Number of Passbooks",
                y="Total Saved (All Passbooks)",
                title="Passbooks vs Total Saved (Multi-Passbook Customers)",
                labels={"Number of Passbooks": "Number of Passbooks", "Total Saved (All Passbooks)": "Total Saved (₹)"},
                color="Payment Status",
                size="Total Transactions (All Passbooks)",
                hover_data=["Customer Name"]
            )
            fig.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Inter, sans-serif")
            )
            st.plotly_chart(fig, width="stretch")
        
        if category_summary is not None and not category_summary.empty:
            st.markdown("#### 📊 Passbooks by Average Payment Frequency")
            fig = px.bar(
                category_summary,
                x="Category",
                y="Unique Passbooks",
                color="Metal Type",
                title="Passbooks by Average Payment Frequency",
                barmode="group",
                color_discrete_sequence=["#FFD700", "#C0C0C0"],
                text="Unique Passbooks"
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Inter, sans-serif")
            )
            st.plotly_chart(fig, width="stretch")
        
        if passbook_summary:
            st.markdown("#### 💰 Total Saved by Payment Status")
            status_amount = []
            for metal_type in ["Gold", "Silver"]:
                if metal_type in passbook_summary:
                    temp = passbook_summary[metal_type].groupby("Payment Status")["Total Saved"].sum().reset_index()
                    temp["Metal"] = metal_type
                    status_amount.append(temp)
            
            if status_amount:
                status_amount_df = pd.concat(status_amount, ignore_index=True)
                fig = px.bar(
                    status_amount_df,
                    x="Payment Status",
                    y="Total Saved",
                    color="Metal",
                    title="Total Saved by Payment Status",
                    color_discrete_sequence=["#FFD700", "#C0C0C0"]
                )
                fig.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(family="Inter, sans-serif")
                )
                st.plotly_chart(fig, width="stretch")
    
    # ============================================================
    # TAB 7: EXPORT
    # ============================================================
    with tab7:
        st.markdown('<div class="section-header">📤 Export Reports</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📒 Passbook Reports")
            for metal_type in ["Gold", "Silver"]:
                if metal_type in passbook_summary:
                    download_cols = [c for c in CUSTOMER_SUMMARY_COLUMNS if c in passbook_summary[metal_type].columns]
                    df_export = passbook_summary[metal_type][download_cols].copy()
                    ordered_cols = [c for c in CUSTOMER_SUMMARY_COLUMNS if c in df_export.columns]
                    if len(ordered_cols) > 0:
                        df_export = df_export[ordered_cols]
                    
                    st.markdown(
                        get_download_link(df_export, f"{metal_type}_passbook_summary"),
                        unsafe_allow_html=True
                    )
                    st.markdown(
                        get_download_link(df_export, f"{metal_type}_passbook_summary", "excel"),
                        unsafe_allow_html=True
                    )
                    st.markdown("---")
            
            st.markdown("#### 👥 Multi-Passbook Customers (Sheet 2)")
            if multi_passbook_customers is not None and not multi_passbook_customers.empty:
                st.markdown(
                    get_download_link(multi_passbook_customers, "multi_passbook_customers"),
                    unsafe_allow_html=True
                )
                st.markdown(
                    get_download_link(multi_passbook_customers, "multi_passbook_customers", "excel"),
                    unsafe_allow_html=True
                )
            else:
                st.info("No multi-passbook customers found")
            
            st.markdown("#### 👤 Single Passbook Customers (Sheet 3)")
            if single_passbook_customers is not None and not single_passbook_customers.empty:
                st.markdown(
                    get_download_link(single_passbook_customers, "single_passbook_customers"),
                    unsafe_allow_html=True
                )
                st.markdown(
                    get_download_link(single_passbook_customers, "single_passbook_customers", "excel"),
                    unsafe_allow_html=True
                )
            else:
                st.info("No single passbook customers found")
        
        with col2:
            st.markdown("#### 📋 Category Reports")
            
            st.markdown("**Historical Avg Category**")
            if category_summary is not None:
                display_cols = [c for c in CATEGORY_SUMMARY_COLUMNS if c in category_summary.columns]
                df_export = category_summary[display_cols].copy()
                st.markdown(
                    get_download_link(df_export, "category_historical"),
                    unsafe_allow_html=True
                )
                st.markdown(
                    get_download_link(df_export, "category_historical", "excel"),
                    unsafe_allow_html=True
                )
            
            st.markdown("**Last Payment Category**")
            if last_payment_category_summary is not None:
                display_cols = [c for c in LAST_PAYMENT_CATEGORY_SUMMARY_COLUMNS if c in last_payment_category_summary.columns]
                df_export = last_payment_category_summary[display_cols].copy()
                st.markdown(
                    get_download_link(df_export, "category_last_payment"),
                    unsafe_allow_html=True
                )
                st.markdown(
                    get_download_link(df_export, "category_last_payment", "excel"),
                    unsafe_allow_html=True
                )
            
            st.markdown("#### 📊 Transaction Data")
            transaction_cols = [
                "Passbook number", "Metal Type", "Installment number", 
                "Paid Date", "Saved Amount", "Customer Name", "Customer Phone Number",
                "Avg Days Between Payments", "Max Gap Days", 
                "Days Since Last Payment To Today", "Category Label (Last Payment)",
                "Consistency Score",
                "Selected Slab", "Category Label"
            ]
            transaction_cols = [c for c in transaction_cols if c in df.columns]
            df_export = df[transaction_cols].copy()
            st.markdown(
                get_download_link(df_export, "all_transactions"),
                unsafe_allow_html=True
            )
            st.markdown(
                get_download_link(df_export, "all_transactions", "excel"),
                unsafe_allow_html=True
            )
        
        st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
        st.markdown("#### 📦 Complete Excel Report (All Sheets)")
        
        if st.button("📥 Download Complete Excel Report", type="primary", width="stretch"):
            with st.spinner("Generating complete report..."):
                output = io.BytesIO()
                
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    # Sheet 1: Transactions
                    transaction_cols = [
                        "Passbook number", "Metal Type", "Installment number", 
                        "Paid Date", "Saved Amount", "Customer Name", "Customer Phone Number",
                        "Avg Days Between Payments", "Max Gap Days", 
                        "Days Since Last Payment To Today", "Category Label (Last Payment)",
                        "Consistency Score",
                        "Selected Slab", "Category Label"
                    ]
                    transaction_cols = [c for c in transaction_cols if c in df.columns]
                    df_export = df[transaction_cols].copy()
                    if "Saved Amount" in df_export.columns:
                        df_export["Saved Amount"] = df_export["Saved Amount"].round(0)
                    df_export.to_excel(writer, sheet_name="Transactions", index=False)
                    
                    # Sheet 2: Multi-Passbook Customers
                    if multi_passbook_customers is not None and not multi_passbook_customers.empty:
                        multi_passbook_customers.to_excel(writer, sheet_name="Multi-Passbook Customers", index=False)
                    
                    # Sheet 3: Single Passbook Customers
                    if single_passbook_customers is not None and not single_passbook_customers.empty:
                        single_passbook_customers.to_excel(writer, sheet_name="Single Passbook Customers", index=False)
                    
                    # Sheet 4-5: Gold/Silver Passbook Summary
                    for metal_type in ["Gold", "Silver"]:
                        if metal_type in passbook_summary:
                            download_cols = [c for c in CUSTOMER_SUMMARY_COLUMNS if c in passbook_summary[metal_type].columns]
                            summary_export = passbook_summary[metal_type][download_cols].copy()
                            ordered_cols = [c for c in CUSTOMER_SUMMARY_COLUMNS if c in summary_export.columns]
                            if len(ordered_cols) > 0:
                                summary_export = summary_export[ordered_cols]
                            summary_export.to_excel(
                                writer, sheet_name=f"{metal_type}_Passbooks", index=False
                            )
                    
                    # Sheet 6: Category Historical
                    if category_summary is not None:
                        display_cols = [c for c in CATEGORY_SUMMARY_COLUMNS if c in category_summary.columns]
                        cat_export = category_summary[display_cols].copy()
                        cat_export.to_excel(
                            writer, sheet_name="Category_Historical", index=False
                        )
                    
                    # Sheet 7: Category Last Payment
                    if last_payment_category_summary is not None:
                        display_cols = [c for c in LAST_PAYMENT_CATEGORY_SUMMARY_COLUMNS if c in last_payment_category_summary.columns]
                        lp_export = last_payment_category_summary[display_cols].copy()
                        lp_export.to_excel(
                            writer, sheet_name="Category_Last_Payment", index=False
                        )
                    
                    # Sheet 8: Metrics
                    metrics_df = pd.DataFrame({
                        "Metric": [
                            "Total Passbooks",
                            "Total Unique Customers (Phone)",
                            "Multi-Passbook Customers",
                            "Single Passbook Customers",
                            "Gold Passbooks",
                            "Silver Passbooks",
                            "Total Saved",
                            "Total Transactions",
                            "Avg Consistency Score"
                        ],
                        "Value": [
                            df["Passbook number"].nunique(),
                            len(single_passbook_customers) + len(multi_passbook_customers) if single_passbook_customers is not None and multi_passbook_customers is not None else 0,
                            len(multi_passbook_customers) if multi_passbook_customers is not None else 0,
                            len(single_passbook_customers) if single_passbook_customers is not None else 0,
                            df[df["Metal Type"] == "Gold"]["Passbook number"].nunique(),
                            df[df["Metal Type"] == "Silver"]["Passbook number"].nunique(),
                            df["Saved Amount"].sum(),
                            len(df),
                            f"{df['Consistency Score'].mean():.1f}%"
                        ]
                    })
                    metrics_df.to_excel(writer, sheet_name="Metrics", index=False)
                
                excel_data = output.getvalue()
                b64 = base64.b64encode(excel_data).decode()
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                href = f'''
                <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" 
                   download="daily_scheme_complete_report_{timestamp}.xlsx"
                   class="download-btn download-btn-primary" style="font-size:1.1rem;padding:0.8rem 2rem;">
                   📥 Download Complete Excel Report (All Sheets)
                </a>
                '''
                st.markdown(href, unsafe_allow_html=True)
                st.success("✅ Complete report generated successfully!")
    
    # Footer
    st.markdown("""
    <div class="footer">
        📅 Daily Scheme Payment Frequency Tracker | Tracking Passbooks and Customer Phone Numbers
    </div>
    """, unsafe_allow_html=True)

# ============================================================
# RUN APPLICATION
# ============================================================
if __name__ == "__main__":
    main()