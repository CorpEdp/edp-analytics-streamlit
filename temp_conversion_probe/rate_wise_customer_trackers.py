"""
Auto-converted from: AppRateWiseCustomerTrackers.py
Review this file before using — the converter does a best-effort wrap;
double-check indentation around any unusual control flow (loops, if/else
blocks that span large sections, etc.).
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime, timedelta
def find_column(df, column_name):
    if column_name in df.columns:
        return column_name
    if column_name in ALTERNATIVE_NAMES:
        for alt in ALTERNATIVE_NAMES[column_name]:
            if alt in df.columns:
                return alt
    for col in df.columns:
        if col.lower().strip() == column_name.lower().strip():
            return col
    return None
def round_rates(df):
    df = df.copy()
    rate_cols = ['Metal Rate', 'Rolling_Avg_Rate', 'Join Rate', 'Avg Rate', 'Min Rate', 'Max Rate', 'Avg_Hike_Rate']
    for col in rate_cols:
        if col in df.columns:
            df[col] = df[col].round(0).astype(int)
    return df
def format_dates(df):
    """Remove time portion from dates"""
    df = df.copy()
    date_cols = ['Paid Date', 'Join Date', 'Last Payment Date']
    for col in date_cols:
        if col in df.columns:
            # Ensure datetime first
            if not pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = pd.to_datetime(df[col], errors='coerce')
            # Then convert to date
            df[col] = df[col].dt.date
    return df
def ensure_datetime(df):
    """Ensure all date columns are datetime type"""
    df = df.copy()
    date_cols = ['Paid Date', 'Join Date', 'Last Payment Date']
    for col in date_cols:
        if col in df.columns:
            if not pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = pd.to_datetime(df[col], errors='coerce')
    return df
def detect_rate_patterns(df, rate_change_threshold=0.05):
    df = df.copy()
    
    # Ensure dates are datetime
    if 'Paid Date' in df.columns:
        df['Paid Date'] = pd.to_datetime(df['Paid Date'], errors='coerce')
    
    df = df.sort_values(['Metal Type', 'Paid Date'])
    
    for metal in df['Metal Type'].unique():
        mask = df['Metal Type'] == metal
        if mask.sum() == 0:
            continue
        
        metal_indices = df[mask].index
        daily_rates = df.loc[metal_indices].groupby('Paid Date')['Metal Rate'].mean().reset_index()
        daily_rates['Rate Change'] = daily_rates['Metal Rate'].pct_change()
        daily_rates['Rolling Avg'] = daily_rates['Metal Rate'].rolling(window=7, min_periods=1).mean()
        
        # Round rates
        daily_rates['Metal Rate'] = daily_rates['Metal Rate'].round(0).astype(int)
        daily_rates['Rolling Avg'] = daily_rates['Rolling Avg'].round(0).astype(int)
        
        # Detect movements
        daily_rates['Is_Hike'] = daily_rates['Rate Change'] > rate_change_threshold
        daily_rates['Is_Drop'] = daily_rates['Rate Change'] < -rate_change_threshold
        
        # Classify periods
        median_rate = daily_rates['Metal Rate'].median()
        daily_rates['Rate_Period'] = np.where(
            daily_rates['Metal Rate'] <= median_rate,
            '🔴 LOW',
            '🟢 HIGH'
        )
        
        # Detect drops from high to low
        daily_rates['Is_Significant_Drop'] = (
            (daily_rates['Rate Change'] < -rate_change_threshold) & 
            (daily_rates['Rate_Period'].shift(1) == '🟢 HIGH') &
            (daily_rates['Rate_Period'] == '🔴 LOW')
        )
        
        # Create mapping dictionaries
        hike_map = daily_rates.set_index('Paid Date')['Is_Hike'].to_dict()
        drop_map = daily_rates.set_index('Paid Date')['Is_Drop'].to_dict()
        sig_drop_map = daily_rates.set_index('Paid Date')['Is_Significant_Drop'].to_dict()
        period_map = daily_rates.set_index('Paid Date')['Rate_Period'].to_dict()
        rolling_map = daily_rates.set_index('Paid Date')['Rolling Avg'].to_dict()
        change_map = daily_rates.set_index('Paid Date')['Rate Change'].to_dict()
        
        # Apply mappings
        df.loc[metal_indices, 'Is_Hike'] = df.loc[metal_indices, 'Paid Date'].map(hike_map)
        df.loc[metal_indices, 'Is_Drop'] = df.loc[metal_indices, 'Paid Date'].map(drop_map)
        df.loc[metal_indices, 'Is_Significant_Drop'] = df.loc[metal_indices, 'Paid Date'].map(sig_drop_map)
        df.loc[metal_indices, 'Rate_Period'] = df.loc[metal_indices, 'Paid Date'].map(period_map)
        df.loc[metal_indices, 'Rolling_Avg_Rate'] = df.loc[metal_indices, 'Paid Date'].map(rolling_map)
        df.loc[metal_indices, 'Rate_Change_Pct'] = df.loc[metal_indices, 'Paid Date'].map(change_map)
        
        # Detect trend
        df.loc[metal_indices, 'Rate_Trend'] = np.where(
            df.loc[metal_indices, 'Rolling_Avg_Rate'].notna(),
            np.where(
                df.loc[metal_indices, 'Metal Rate'] > df.loc[metal_indices, 'Rolling_Avg_Rate'],
                '📈 Increasing',
                '📉 Decreasing'
            ),
            'Stable'
        )
    
    df = round_rates(df)
    df = ensure_datetime(df)
    df = format_dates(df)
    return df
def classify_customer_segments(customer_analysis, df):
    """Classify customers into segments for campaign targeting"""
    
    customer_analysis = customer_analysis.copy()
    
    # Ensure dates are datetime for calculation
    if 'Last Payment Date' in customer_analysis.columns:
        customer_analysis['Last Payment Date'] = pd.to_datetime(customer_analysis['Last Payment Date'], errors='coerce')
    
    # Get last payment date for each customer
    last_payment = df.groupby(['Customer ID', 'Customer Name'])['Paid Date'].max().reset_index()
    last_payment.columns = ['Customer ID', 'Customer Name', 'Last Payment Date']
    customer_analysis = customer_analysis.merge(last_payment, on=['Customer ID', 'Customer Name'], how='left')
    
    # Ensure max date is datetime
    max_date = pd.to_datetime(df['Paid Date'].max())
    
    # Calculate days since last payment
    customer_analysis['Days Since Last Payment'] = (max_date - pd.to_datetime(customer_analysis['Last Payment Date'])).dt.days
    
    # Calculate payment consistency (in days, rounded)
    payment_dates = df.groupby(['Customer ID', 'Customer Name'])['Paid Date'].apply(list).reset_index()
    payment_dates.columns = ['Customer ID', 'Customer Name', 'Payment Dates']
    
    def calc_consistency(dates):
        if len(dates) < 3:
            return 0  # Not enough data for consistency
        # Ensure dates are datetime
        dates = [pd.to_datetime(d) for d in dates]
        # Calculate gaps between payments in days
        gaps = [(dates[i+1] - dates[i]).days for i in range(len(dates)-1)]
        # Return standard deviation of gaps, rounded to nearest whole number
        return round(np.std(gaps)) if len(gaps) > 0 else 0
    
    payment_dates['Payment_Consistency_Days'] = payment_dates['Payment Dates'].apply(calc_consistency)
    customer_analysis = customer_analysis.merge(
        payment_dates[['Customer ID', 'Customer Name', 'Payment_Consistency_Days']], 
        on=['Customer ID', 'Customer Name'], 
        how='left'
    )
    customer_analysis['Payment_Consistency_Days'] = customer_analysis['Payment_Consistency_Days'].fillna(0).astype(int)
    
    # Classify segments
    def get_segment(row):
        # Check if active (paid within last 30 days)
        is_active = row['Days Since Last Payment'] <= 30
        
        # Check payment performance
        if row['Transactions'] == 0:
            return ('⚪', 'Inactive / Reactivation', 'segment-inactive')
        
        # High Potential: Many transactions + high amount + consistent
        if (row['Transactions'] > 10 and 
            row['Total Amount'] > 100000 and 
            row['Payment_Consistency_Days'] < 30 and 
            row['Payment_Consistency_Days'] > 0):
            return ('⭐', 'High Potential Customer', 'segment-potential')
        
        # Rate Sensitive - Low Rate: Joined and paid mostly at LOW rates
        if row['Join Rate Period'] == '🔴 LOW' and row['Low_Rate_Pct'] > 60:
            if is_active:
                return ('🟢', 'Rate Sensitive - Low Rate', 'segment-sensitive')
            else:
                return ('⚪', 'Inactive / Reactivation', 'segment-inactive')
        
        # Rate Sensitive - Drop: Joined or paid during significant drops
        if row['Drops_Experienced'] > 0 and row['Drop_Response'] > 0:
            if is_active:
                return ('🔵', 'Rate Sensitive - Drop', 'segment-drop')
            else:
                return ('⚪', 'Inactive / Reactivation', 'segment-inactive')
        
        # High Rate Active: Paid during HIGH rates
        if row['High_Rate_Pct'] > 60 and is_active:
            return ('🟠', 'High Rate Active', 'segment-high')
        
        # Regular Payment Circle: Consistent payments
        if (row['Payment_Consistency_Days'] < 15 and 
            row['Payment_Consistency_Days'] > 0 and
            row['Transactions'] > 5 and 
            is_active):
            return ('🟣', 'Regular Payment Circle', 'segment-regular')
        
        # Payment Slowing: Used to pay, but slowing down
        if row['Transactions'] > 3 and row['Days Since Last Payment'] > 15 and row['Days Since Last Payment'] <= 60:
            return ('🟡', 'Payment Slowing', 'segment-slowing')
        
        # Payment Dropped: Stopped paying
        if row['Transactions'] > 3 and row['Days Since Last Payment'] > 60:
            return ('🔴', 'Payment Dropped', 'segment-dropped')
        
        # Default
        if is_active:
            return ('🟣', 'Regular Payment Circle', 'segment-regular')
        else:
            return ('⚪', 'Inactive / Reactivation', 'segment-inactive')
    
    # Add rate percentages
    customer_analysis['Low_Rate_Pct'] = customer_analysis.get('Low_Rate_Pct', 0)
    customer_analysis['High_Rate_Pct'] = customer_analysis.get('High_Rate_Pct', 0)
    customer_analysis['Drop_Response'] = customer_analysis.get('Drop_Response', 0)
    
    # Apply segmentation
    segments = customer_analysis.apply(get_segment, axis=1, result_type='expand')
    customer_analysis['Segment_Icon'] = segments[0]
    customer_analysis['Segment_Name'] = segments[1]
    customer_analysis['Segment_Class'] = segments[2]
    
    # Format dates
    customer_analysis = ensure_datetime(customer_analysis)
    customer_analysis = format_dates(customer_analysis)
    
    return customer_analysis
def analyze_customer_behavior(df):
    """Comprehensive customer behavior analysis"""
    
    # Ensure df has datetime for Paid Date
    df = df.copy()
    if 'Paid Date' in df.columns:
        df['Paid Date'] = pd.to_datetime(df['Paid Date'], errors='coerce')
    
    customer_analysis = df.groupby(['Customer ID', 'Customer Name']).agg({
        'Id': 'count',
        'Metal Rate': ['mean', 'min', 'max'],
        'Saved Amount': 'sum',
        'Installment number': ['min', 'max', 'count'],
        'Is_Hike': 'sum',
        'Is_Drop': 'sum',
        'Is_Significant_Drop': 'sum'
    }).reset_index()
    
    # Flatten columns
    customer_analysis.columns = [
        'Customer ID', 'Customer Name', 'Transactions', 
        'Avg Rate', 'Min Rate', 'Max Rate',
        'Total Amount', 'Min Inst', 'Max Inst', 'Total Installments',
        'Hikes_Experienced', 'Drops_Experienced', 'Significant_Drops'
    ]
    
    # Round rates
    for col in ['Avg Rate', 'Min Rate', 'Max Rate']:
        customer_analysis[col] = customer_analysis[col].round(0).astype(int)
    
    # First transaction (Joining)
    first_trans = df.groupby(['Customer ID', 'Customer Name']).agg({
        'Paid Date': 'min',
        'Metal Rate': 'first',
        'Rate_Period': 'first'
    }).reset_index()
    first_trans.columns = ['Customer ID', 'Customer Name', 'Join Date', 'Join Rate', 'Join Rate Period']
    first_trans['Join Rate'] = first_trans['Join Rate'].round(0).astype(int)
    
    customer_analysis = customer_analysis.merge(first_trans, on=['Customer ID', 'Customer Name'], how='left')
    
    # Count transactions by rate period
    rate_pivot = df.pivot_table(
        index=['Customer ID', 'Customer Name'],
        columns='Rate_Period',
        values='Id',
        aggfunc='count',
        fill_value=0
    ).reset_index()
    
    # Handle rate period columns
    if '🔴 LOW' in rate_pivot.columns:
        customer_analysis = customer_analysis.merge(
            rate_pivot[['Customer ID', 'Customer Name', '🔴 LOW']], 
            on=['Customer ID', 'Customer Name'], 
            how='left'
        )
        customer_analysis['Low_Rate_Pct'] = (
            customer_analysis['🔴 LOW'] / customer_analysis['Transactions'] * 100
        ).fillna(0).round(1)
    else:
        customer_analysis['Low_Rate_Pct'] = 0
    
    if '🟢 HIGH' in rate_pivot.columns:
        customer_analysis = customer_analysis.merge(
            rate_pivot[['Customer ID', 'Customer Name', '🟢 HIGH']], 
            on=['Customer ID', 'Customer Name'], 
            how='left'
        )
        customer_analysis['High_Rate_Pct'] = (
            customer_analysis['🟢 HIGH'] / customer_analysis['Transactions'] * 100
        ).fillna(0).round(1)
    else:
        customer_analysis['High_Rate_Pct'] = 0
    
    # Response to drops
    drop_response = df[df['Is_Significant_Drop'] == True].groupby(['Customer ID', 'Customer Name']).agg({
        'Id': 'count'
    }).reset_index()
    drop_response.columns = ['Customer ID', 'Customer Name', 'Drop_Response']
    customer_analysis = customer_analysis.merge(drop_response, on=['Customer ID', 'Customer Name'], how='left')
    customer_analysis['Drop_Response'] = customer_analysis['Drop_Response'].fillna(0).astype(int)
    
    # Payment consistency (average days between payments, rounded)
    # Use a safer approach that handles datetime properly
    def calc_avg_gap(group):
        dates = group['Paid Date'].dropna().sort_values()
        if len(dates) < 2:
            return 0
        gaps = dates.diff().dt.days
        return round(gaps.mean()) if not gaps.isna().all() else 0
    
    payment_gaps = df.groupby(['Customer ID', 'Customer Name']).apply(calc_avg_gap).reset_index(name='Avg_Payment_Gap')
    customer_analysis = customer_analysis.merge(
        payment_gaps[['Customer ID', 'Customer Name', 'Avg_Payment_Gap']], 
        on=['Customer ID', 'Customer Name'], 
        how='left'
    )
    customer_analysis['Avg_Payment_Gap'] = customer_analysis['Avg_Payment_Gap'].fillna(0).astype(int)
    
    # Format dates
    customer_analysis = ensure_datetime(customer_analysis)
    customer_analysis = format_dates(customer_analysis)
    
    return customer_analysis
def generate_campaign_insights(customer_analysis, df):
    """Generate campaign recommendations based on segments"""
    
    insights = []
    
    # Segment counts
    segment_counts = customer_analysis['Segment_Name'].value_counts()
    
    # Rate Drop Impact
    drop_customers = customer_analysis[customer_analysis['Segment_Name'] == 'Rate Sensitive - Drop']
    if len(drop_customers) > 0:
        avg_response = drop_customers['Drop_Response'].mean()
        insights.append({
            'type': 'drop_opportunity',
            'title': '🎯 Rate Drop Campaign Opportunity',
            'description': f'{len(drop_customers)} customers historically respond to rate drops',
            'detail': f'Average response during drops: {avg_response:.1f} transactions',
            'segment': 'Rate Sensitive - Drop',
            'segment_class': 'segment-drop'
        })
    
    # Low Rate Customers
    low_rate_customers = customer_analysis[customer_analysis['Segment_Name'] == 'Rate Sensitive - Low Rate']
    if len(low_rate_customers) > 0:
        insights.append({
            'type': 'low_rate_opportunity',
            'title': '📉 Low Rate Campaign Opportunity',
            'description': f'{len(low_rate_customers)} customers prefer LOW rates',
            'detail': f'Average rate: ₹{low_rate_customers["Join Rate"].mean():,.0f}',
            'segment': 'Rate Sensitive - Low Rate',
            'segment_class': 'segment-sensitive'
        })
    
    # High Potential Customers
    high_potential = customer_analysis[customer_analysis['Segment_Name'] == 'High Potential Customer']
    if len(high_potential) > 0:
        total_value = high_potential['Total Amount'].sum()
        insights.append({
            'type': 'high_potential',
            'title': '⭐ High Potential Customers',
            'description': f'{len(high_potential)} high-value active customers',
            'detail': f'Total value: ₹{total_value:,.0f}',
            'segment': 'High Potential Customer',
            'segment_class': 'segment-potential'
        })
    
    # Regular Payment Circle
    regular_payers = customer_analysis[customer_analysis['Segment_Name'] == 'Regular Payment Circle']
    if len(regular_payers) > 0:
        avg_consistency = regular_payers['Payment_Consistency_Days'].mean()
        insights.append({
            'type': 'regular_payers',
            'title': '🔄 Regular Payment Circle',
            'description': f'{len(regular_payers)} consistent payers',
            'detail': f'Avg consistency: {avg_consistency:.0f} days variation',
            'segment': 'Regular Payment Circle',
            'segment_class': 'segment-regular'
        })
    
    # Payment Slowing
    slowing = customer_analysis[customer_analysis['Segment_Name'] == 'Payment Slowing']
    if len(slowing) > 0:
        insights.append({
            'type': 'slowing',
            'title': '🟡 Retention Campaign - Payment Slowing',
            'description': f'{len(slowing)} customers are slowing down payments',
            'detail': f'Avg days since last payment: {slowing["Days Since Last Payment"].mean():.0f} days',
            'segment': 'Payment Slowing',
            'segment_class': 'segment-slowing'
        })
    
    # Inactive/Reactivation
    inactive = customer_analysis[customer_analysis['Segment_Name'] == 'Inactive / Reactivation']
    if len(inactive) > 0:
        insights.append({
            'type': 'reactivation',
            'title': '⚪ Reactivation Campaign',
            'description': f'{len(inactive)} inactive customers need reactivation',
            'detail': f'Avg days inactive: {inactive["Days Since Last Payment"].mean():.0f} days',
            'segment': 'Inactive / Reactivation',
            'segment_class': 'segment-inactive'
        })
    
    return insights
def create_campaign_report():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Customer segments
        filtered_customers.to_excel(writer, sheet_name='Customer Segments', index=False)
        
        # Segment summary
        segment_summary = filtered_customers['Segment_Name'].value_counts().reset_index()
        segment_summary.columns = ['Segment', 'Count']
        segment_summary.to_excel(writer, sheet_name='Segment Summary', index=False)
        
        # Transactions
        filtered_df.to_excel(writer, sheet_name='Transactions', index=False)
        
        # Campaign insights
        if insights:
            insights_df = pd.DataFrame(insights)
            insights_df.to_excel(writer, sheet_name='Campaign Insights', index=False)
        
        # High value customers
        high_value = filtered_customers.nlargest(50, 'Total Amount')
        if not high_value.empty:
            high_value.to_excel(writer, sheet_name='High Value Customers', index=False)
    
    output.seek(0)
    return output

LABEL = "Rate Wise Customer Trackers"


def run():
    st.markdown("""
        <style>
        .main-title { font-size: 28px; font-weight: 700; margin-bottom: 5px; color: #1a1a2e; }
        .subtitle { font-size: 14px; color: #666; margin-bottom: 20px; }
        .segment-box { 
            padding: 15px; 
            border-radius: 10px; 
            margin: 5px 0;
            border-left: 4px solid;
        }
        .segment-sensitive { background: #d4edda; border-color: #28a745; }
        .segment-drop { background: #cce5ff; border-color: #007bff; }
        .segment-high { background: #f8d7da; border-color: #dc3545; }
        .segment-regular { background: #d6d8db; border-color: #6c757d; }
        .segment-slowing { background: #fff3cd; border-color: #ffc107; }
        .segment-dropped { background: #f5c6cb; border-color: #dc3545; }
        .segment-inactive { background: #e2e3e5; border-color: #6c757d; }
        .segment-potential { background: #cce5ff; border-color: #0056b3; }
        </style>
    """, unsafe_allow_html=True)
    EXPECTED_COLUMNS = [
        "Id", "Customer Name", "Customer Phone Number", 
        "Metal Type", "Metal Rate", "Saved Metal Weight", 
        "Saved Amount", "Scheme Name", "Paid Date",
        "Installment number", "Scheme Participation Id",
        "Status", "Reward Amount"
    ]
    ALTERNATIVE_NAMES = {
        'Customer Phone Number': ['Phone', 'Mobile', 'Contact', 'Customer Phone', 'Phone Number'],
        'Customer Name': ['Name', 'Customer', 'Client Name'],
        'Metal Rate': ['Rate', 'Gold Rate', 'Silver Rate', 'Current Rate'],
        'Saved Amount': ['Amount', 'Payment', 'Paid Amount'],
        'Scheme Name': ['Scheme', 'Plan Name'],
        'Paid Date': ['Date', 'Payment Date', 'Transaction Date'],
        'Installment number': ['Installment', 'Installment No', 'Inst No', 'Payment Number']
    }
    st.markdown('<div class="main-title">🎯 E-Gold Scheme Customer Campaign Analysis</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Identify customer segments based on rate movement response • Generate actionable campaign targets</div>', unsafe_allow_html=True)
    st.sidebar.header("📂 Upload Data")
    uploaded_file = st.sidebar.file_uploader("Upload Excel or CSV", type=["xlsx", "xls", "csv"])
    rate_change_threshold = st.sidebar.slider(
        "Rate Change Threshold (%)",
        min_value=1,
        max_value=20,
        value=5,
        help="Percentage change to detect rate hikes/drops"
    ) / 100
    if uploaded_file is None:
        st.info("👆 Upload your transaction file to start")

        st.markdown("""
        ### 🎯 What This Tool Does:

        1. **Analyzes Rate Movement** → LOW → HIKE → DROP patterns
        2. **Identifies Customer Segments** → Based on response to rate changes
        3. **Generates Campaign Insights** → Who to target and when
        4. **Provides Actionable Lists** → Directly usable for campaigns

        ### 📊 Customer Segments:
        - 🟢 Rate Sensitive - Low Rate
        - 🔵 Rate Sensitive - Drop  
        - 🟠 High Rate Active
        - 🟣 Regular Payment Circle
        - 🟡 Payment Slowing
        - 🔴 Payment Dropped
        - ⚪ Inactive / Reactivation
        - ⭐ High Potential Customer
        """)

        st.markdown("### 📋 Required Columns:")
        cols = st.columns(3)
        for i, col in enumerate(EXPECTED_COLUMNS):
            with cols[i % 3]:
                st.caption(f"• {col}")
        st.stop()
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ Error: {e}")
        st.stop()
    df_raw.columns = df_raw.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
    column_mapping = {}
    for expected_col in EXPECTED_COLUMNS:
        found_col = find_column(df_raw, expected_col)
        if found_col:
            column_mapping[expected_col] = found_col
    min_required = ['Customer Name', 'Metal Rate', 'Paid Date', 'Installment number']
    missing_min = [col for col in min_required if col not in column_mapping]
    if missing_min:
        st.error(f"❌ Missing critical columns: {missing_min}")
        st.write("### Available columns:", list(df_raw.columns))
        st.stop()
    df = df_raw.copy()
    for expected, found in column_mapping.items():
        if expected != found:
            df[expected] = df[found]
    numeric_cols = ['Metal Rate', 'Saved Metal Weight', 'Saved Amount', 'Installment number']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '').str.replace('₹', '').str.replace('$', '')
            df[col] = pd.to_numeric(df[col], errors='coerce')
    df['Metal Rate'] = df['Metal Rate'].round(0).astype(int)
    text_cols = ['Customer Name', 'Customer Phone Number', 'Metal Type', 'Scheme Name']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].fillna('').astype(str).str.strip()
    if 'Metal Type' in df.columns:
        df['Metal Type'] = df['Metal Type'].str.upper().str.strip()
        df['Metal Type'] = df['Metal Type'].replace({
            'GOLD': 'Gold', 'G': 'Gold', 'AU': 'Gold',
            'SILVER': 'Silver', 'S': 'Silver', 'AG': 'Silver'
        })
    else:
        df['Metal Type'] = 'Gold'
    df['Paid Date'] = pd.to_datetime(df['Paid Date'], errors='coerce')
    df = df[df['Metal Rate'] > 0]
    df = df[df['Paid Date'].notna()]
    if df.empty:
        st.warning("⚠️ No valid data found")
        st.stop()
    if 'Customer Phone Number' in df.columns and df['Customer Phone Number'].notna().any():
        df['Customer ID'] = df['Customer Phone Number'].astype(str).str.strip()
        df['Customer ID'] = df['Customer ID'].replace('', 'Unknown')
    else:
        df['Customer ID'] = df['Customer Name'].astype(str).str.strip()
    df = detect_rate_patterns(df, rate_change_threshold)
    customer_analysis = analyze_customer_behavior(df)
    customer_analysis = classify_customer_segments(customer_analysis, df)
    rate_cols = ['Avg Rate', 'Min Rate', 'Max Rate', 'Join Rate']
    for col in rate_cols:
        if col in customer_analysis.columns:
            customer_analysis[col] = customer_analysis[col].round(0).astype(int)
    st.sidebar.markdown("---")
    st.sidebar.header("🔎 Filters")
    metal_options = sorted(df['Metal Type'].unique().tolist())
    selected_metals = st.sidebar.multiselect("Metal Type", options=metal_options, default=metal_options)
    segment_options = sorted(customer_analysis['Segment_Name'].unique().tolist())
    selected_segments = st.sidebar.multiselect("Customer Segment", options=segment_options, default=segment_options)
    active_status = st.sidebar.selectbox("Payment Status", ["All", "Active Only (Paid in last 30 days)", "Inactive Only"])
    filtered_df = df.copy()
    if selected_metals:
        filtered_df = filtered_df[filtered_df['Metal Type'].isin(selected_metals)]
    filtered_customers = customer_analysis.copy()
    if selected_segments:
        filtered_customers = filtered_customers[filtered_customers['Segment_Name'].isin(selected_segments)]
    if active_status == "Active Only (Paid in last 30 days)":
        filtered_customers = filtered_customers[filtered_customers['Days Since Last Payment'] <= 30]
    elif active_status == "Inactive Only":
        filtered_customers = filtered_customers[filtered_customers['Days Since Last Payment'] > 30]
    filtered_ids = filtered_customers['Customer ID'].tolist()
    filtered_df = filtered_df[filtered_df['Customer ID'].isin(filtered_ids)]
    st.markdown("---")
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total Customers", f"{len(filtered_customers):,}")
    with col2:
        st.metric("Total Transactions", f"{len(filtered_df):,}")
    with col3:
        active_count = len(filtered_customers[filtered_customers['Days Since Last Payment'] <= 30])
        st.metric("Active Customers", f"{active_count:,}")
    with col4:
        avg_amount = filtered_customers['Total Amount'].mean()
        st.metric("Avg Customer Value", f"₹{avg_amount:,.0f}")
    with col5:
        segments_count = len(filtered_customers['Segment_Name'].unique())
        st.metric("Segments Found", f"{segments_count}")
    st.markdown("---")
    st.subheader("📊 Customer Segment Distribution")
    segment_counts = filtered_customers['Segment_Name'].value_counts().reset_index()
    segment_counts.columns = ['Segment', 'Count']
    col1, col2 = st.columns([2, 1])
    with col1:
        st.bar_chart(segment_counts.set_index('Segment'))
    with col2:
        st.dataframe(
            segment_counts,
            use_container_width=True,
            hide_index=True
        )
    st.markdown("---")
    st.subheader("🎯 Campaign Insights & Opportunities")
    insights = generate_campaign_insights(filtered_customers, filtered_df)
    if insights:
        for insight in insights:
            with st.container():
                st.markdown(f"""
                <div class="segment-box {insight.get('segment_class', '')}">
                    <strong>{insight['title']}</strong><br>
                    {insight['description']}<br>
                    <small>{insight['detail']}</small>
                </div>
                """, unsafe_allow_html=True)
    else:
        st.info("No campaign insights generated. Need more data.")
    st.markdown("---")
    st.subheader("👤 Customer Segments - Campaign Target List")
    for segment in sorted(filtered_customers['Segment_Name'].unique()):
        segment_customers = filtered_customers[filtered_customers['Segment_Name'] == segment]

        with st.expander(f"{segment} ({len(segment_customers)} customers)"):
            if not segment_customers.empty:
                # Show segment details
                display_cols = ['Customer Name', 'Transactions', 'Total Amount', 'Avg Rate', 'Join Rate', 
                               'Days Since Last Payment', 'Payment_Consistency_Days']
                available_cols = [col for col in display_cols if col in segment_customers.columns]

                # Format amounts and dates
                display_data = segment_customers[available_cols].copy()
                if 'Total Amount' in display_data.columns:
                    display_data['Total Amount'] = display_data['Total Amount'].apply(lambda x: f"₹{x:,.0f}")
                if 'Avg Rate' in display_data.columns:
                    display_data['Avg Rate'] = display_data['Avg Rate'].apply(lambda x: f"₹{x:,.0f}")
                if 'Join Rate' in display_data.columns:
                    display_data['Join Rate'] = display_data['Join Rate'].apply(lambda x: f"₹{x:,.0f}")
                if 'Payment_Consistency_Days' in display_data.columns:
                    display_data['Payment_Consistency_Days'] = display_data['Payment_Consistency_Days'].astype(int)

                st.dataframe(
                    display_data,
                    use_container_width=True,
                    hide_index=True
                )

                # Export button for this segment
                csv = segment_customers.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Export {segment} List",
                    data=csv,
                    file_name=f"{segment.replace(' ', '_')}_campaign_list.csv",
                    mime="text/csv",
                    key=f"export_{segment}"
                )
    st.markdown("---")
    st.subheader("📈 Rate Movement Timeline")
    for metal in selected_metals:
        metal_data = filtered_df[filtered_df['Metal Type'] == metal].copy()
        if metal_data.empty:
            continue

        daily = metal_data.groupby('Paid Date').agg({
            'Metal Rate': 'mean',
            'Is_Hike': 'max',
            'Is_Drop': 'max',
            'Is_Significant_Drop': 'max'
        }).reset_index()

        daily['Metal Rate'] = daily['Metal Rate'].round(0).astype(int)
        daily['Paid Date'] = pd.to_datetime(daily['Paid Date']).dt.date

        st.markdown(f"### {metal} Rate Movement")
        st.line_chart(daily.set_index('Paid Date')['Metal Rate'])

        # Show key events
        st.caption(f"📈 Hikes: {daily['Is_Hike'].sum()} | 📉 Drops: {daily['Is_Drop'].sum()} | 🔽 Significant Drops: {daily['Is_Significant_Drop'].sum()}")
    st.markdown("---")
    st.subheader("📋 Transaction Details")
    display_cols = ['Paid Date', 'Customer Name', 'Metal Type', 'Metal Rate', 'Rate_Period',
                    'Is_Hike', 'Is_Drop', 'Is_Significant_Drop', 'Installment number', 'Saved Amount']
    display_cols = [col for col in display_cols if col in filtered_df.columns]
    if display_cols:
        transaction_view = filtered_df[display_cols].sort_values('Paid Date', ascending=False).head(500)

        # Format display
        if 'Metal Rate' in transaction_view.columns:
            transaction_view['Metal Rate'] = transaction_view['Metal Rate'].round(0).astype(int)
        if 'Paid Date' in transaction_view.columns:
            transaction_view['Paid Date'] = pd.to_datetime(transaction_view['Paid Date']).dt.date

        st.dataframe(
            transaction_view,
            use_container_width=True,
            hide_index=True
        )
    st.markdown("---")
    st.subheader("📥 Export Reports")
    excel_file = create_campaign_report()
    st.download_button(
        label="📊 Download Complete Campaign Report",
        data=excel_file,
        file_name="egold_campaign_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown("---")
    st.caption("🎯 E-Gold Scheme Customer Campaign Analysis | Identify segments → Generate insights → Target campaigns")
