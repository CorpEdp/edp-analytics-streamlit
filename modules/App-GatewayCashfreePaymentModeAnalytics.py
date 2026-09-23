import streamlit as st
import pandas as pd
from io import BytesIO
import re

# -----------------------------
# Page Configuration
# -----------------------------
st.set_page_config(
    page_title="Gateway Report Analytics",
    page_icon="💳",
    layout="wide"
)

st.title("💳 Gateway Report Analytics")

# -----------------------------
# Parent SubType Logic
# -----------------------------
def get_parent_subtype(value):
    if pd.isna(value):
        return "Unknown"

    value = str(value).strip()

    parts = [x.strip() for x in value.split("-")]

    # Example:
    # MASTER - Credit Card - AXIS
    # VISA - Debit Card - SBI
    if len(parts) >= 3:
        return f"{parts[0]} - {parts[1]}"

    # Example:
    # UPI - GPay
    elif len(parts) == 2:
        if parts[0].upper() == "UPI":
            return "UPI"
        return f"{parts[0]} - {parts[1]}"

    return parts[0]

# -----------------------------
# Updated Payment Type Logic with Auto-Detection
# -----------------------------
def get_payment_type(value):
    if pd.isna(value):
        return "Unknown"

    value = str(value).strip()
    upper_value = value.upper()

    # -----------------------------
    # AMEX Cards
    # -----------------------------
    if "AMEX" in upper_value:
        if "DEBIT" in upper_value:
            return "AMEX Debit Card"
        return "AMEX Credit Card"

    # -----------------------------
    # Credit Cards
    # -----------------------------
    elif any(card in value for card in [
        "MASTER - Credit Card",
        "MASTER - Corporate Credit Card",
        "MASTER - Premium Credit Card",
        "MasterCard - Credit Card",
        "VISA - Credit Card",
        "VISA - Corporate Credit Card",
        "VISA - Premium Credit Card",
        "Rupay - Credit Card"
    ]):
        return "Credit Card"

    # -----------------------------
    # Debit Cards
    # -----------------------------
    elif any(card in value for card in [
        "MASTER - Debit Card",
        "MasterCard - Debit",
        "VISA - Debit Card",
        "Rupay - Debit Card"
    ]):
        return "Debit Card"

    # -----------------------------
    # Prepaid Cards
    # -----------------------------
    elif "PREPAID" in upper_value:
        return "Prepaid Card"

    # -----------------------------
    # UPI
    # -----------------------------
    elif "UPI" in upper_value:
        return "UPI"

    # -----------------------------
    # Wallets
    # -----------------------------
    elif any(wallet.upper() in upper_value for wallet in [
        "Amazon Pay",
        "MobiKwik",
        "PhonePe",
        "Wallet"
    ]):
        return "Wallet"

    # -----------------------------
    # 🔥 AUTO-DETECT ANY BANK (FUTURE-PROOF)
    # If it contains "BANK" and hasn't been classified yet, it's Net Banking
    # -----------------------------
    elif "BANK" in upper_value:
        return "Net Banking"

    # -----------------------------
    # Net Banking (explicit check)
    # -----------------------------
    elif "NET BANKING" in upper_value or "NETBANKING" in upper_value:
        return "Net Banking"

    # -----------------------------
    # EMI
    # -----------------------------
    elif "EMI" in upper_value:
        return "EMI"

    # -----------------------------
    # Default - return as-is if nothing matches
    # -----------------------------
    return value

# -----------------------------
# Upload File
# -----------------------------
uploaded_file = st.file_uploader(
    "Upload Gateway Report",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:

    try:

        # -----------------------------
        # Read File
        # -----------------------------
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.success("✅ File Uploaded Successfully")

        st.subheader("Raw Data Preview")
        st.dataframe(df.head(), width="stretch")

        # -----------------------------
        # Check if dataframe is empty
        # -----------------------------
        if df.empty:
            st.warning("The uploaded file contains no data.")
            st.stop()

        # -----------------------------
        # Required Columns Check
        # -----------------------------
        required_columns = [
            'Order Id',
            'Payment Mode SubType',
            'Amount',
            'Service Charge',
            'ST/GST',
            'Settlement Amount',
            'Transaction Time'
        ]

        missing_columns = [
            col for col in required_columns
            if col not in df.columns
        ]

        if missing_columns:
            st.error(
                f"Missing Required Columns: {', '.join(missing_columns)}"
            )
            st.stop()

        # -----------------------------
        # Convert Date
        # -----------------------------
        df['Transaction Time'] = pd.to_datetime(
            df['Transaction Time'],
            errors='coerce'
        )

        # Remove rows with invalid dates
        df = df.dropna(subset=['Transaction Time'])

        df['Month'] = df['Transaction Time'].dt.strftime('%Y-%m')

        # -----------------------------
        # Date Range Filter (Optional)
        # -----------------------------
        st.sidebar.header("Filters")
        
        min_date = df['Transaction Time'].min().date()
        max_date = df['Transaction Time'].max().date()
        
        date_range = st.sidebar.date_input(
            "Select Date Range",
            value=[min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )
        
        if len(date_range) == 2:
            start_date = pd.Timestamp(date_range[0])
            end_date = pd.Timestamp(date_range[1])
            mask = (df['Transaction Time'].dt.date >= date_range[0]) & \
                   (df['Transaction Time'].dt.date <= date_range[1])
            df_filtered = df[mask].copy()
        else:
            df_filtered = df.copy()
        
        st.sidebar.info(f"Showing data from {date_range[0]} to {date_range[1]}" if len(date_range) == 2 else "Showing all data")

        # -----------------------------
        # Convert Numeric Columns
        # -----------------------------
        numeric_cols = [
            'Amount',
            'Service Charge',
            'ST/GST',
            'Settlement Amount'
        ]

        for col in numeric_cols:
            df_filtered[col] = pd.to_numeric(
                df_filtered[col],
                errors='coerce'
            ).fillna(0)

        # -----------------------------
        # Parent SubType
        # -----------------------------
        df_filtered['Parent SubType'] = (
            df_filtered['Payment Mode SubType']
            .apply(get_parent_subtype)
        )

        # -----------------------------
        # Payment Type
        # -----------------------------
        df_filtered['Payment Type'] = (
            df_filtered['Payment Mode SubType']
            .apply(get_payment_type)
        )

        # -----------------------------
        # 🔍 Classification Summary (Debug View)
        # Shows how each Payment Mode SubType was classified
        # -----------------------------
        st.subheader("🔍 Payment Type Classification Summary")
        classification_check = df_filtered.groupby('Payment Type').agg(
            Count=('Payment Type', 'count'),
            Sample_Values=('Payment Mode SubType', lambda x: list(x.head(3)))
        ).reset_index()
        
        st.dataframe(classification_check, width="stretch")

        # -----------------------------
        # Unique SubType Summary
        # -----------------------------
        unique_subtypes = (
            df_filtered.groupby('Payment Mode SubType')
            .agg(
                Count=('Order Id', 'nunique'),
                Amount=('Amount', 'sum'),
                Service_Charge=('Service Charge', 'sum'),
                ST_GST=('ST/GST', 'sum'),
                Settlement_Amount=('Settlement Amount', 'sum')
            )
            .reset_index()
            .sort_values(
                by='Amount',
                ascending=False
            )
        )

        # -----------------------------
        # Month Wise Payment Mode SubType
        # -----------------------------
        subtype_summary = (
            df_filtered.groupby(
                ['Month', 'Payment Mode SubType'],
                dropna=False
            )
            .agg(
                Count=('Order Id', 'nunique'),
                Amount=('Amount', 'sum'),
                Service_Charge=('Service Charge', 'sum'),
                ST_GST=('ST/GST', 'sum'),
                Settlement_Amount=('Settlement Amount', 'sum')
            )
            .reset_index()
            .sort_values(
                ['Month', 'Amount'],
                ascending=[True, False]
            )
        )

        # -----------------------------
        # Month Wise Parent SubType
        # -----------------------------
        parent_summary = (
            df_filtered.groupby(
                ['Month', 'Parent SubType'],
                dropna=False
            )
            .agg(
                Count=('Order Id', 'nunique'),
                Amount=('Amount', 'sum'),
                Service_Charge=('Service Charge', 'sum'),
                ST_GST=('ST/GST', 'sum'),
                Settlement_Amount=('Settlement Amount', 'sum')
            )
            .reset_index()
            .sort_values(
                ['Month', 'Amount'],
                ascending=[True, False]
            )
        )

        # -----------------------------
        # Month Wise Payment Type (ENHANCED with Avg Service Charge & %)
        # -----------------------------
        payment_type_summary = (
            df_filtered.groupby(
                ['Month', 'Payment Type'],
                dropna=False
            )
            .agg(
                Count=('Order Id', 'nunique'),
                Amount=('Amount', 'sum'),
                Service_Charge=('Service Charge', 'sum'),
                ST_GST=('ST/GST', 'sum'),
                Settlement_Amount=('Settlement Amount', 'sum'),
                Avg_Service_Charge=('Service Charge', 'mean')
            )
            .reset_index()
        )
        
        # Calculate Service Charge Percentage
        payment_type_summary['Service_Charge_Percentage'] = (
            (payment_type_summary['Service_Charge'] / payment_type_summary['Amount']) * 100
        ).round(2)
        
        # Format the columns for display
        payment_type_summary_display = payment_type_summary.copy()
        payment_type_summary_display['Avg_Service_Charge'] = payment_type_summary_display['Avg_Service_Charge'].round(2)
        payment_type_summary_display['Service_Charge_Percentage'] = payment_type_summary_display['Service_Charge_Percentage'].astype(str) + '%'
        
        # Sort by Month and Amount
        payment_type_summary = payment_type_summary.sort_values(
            ['Month', 'Amount'],
            ascending=[True, False]
        )
        
        payment_type_summary_display = payment_type_summary_display.sort_values(
            ['Month', 'Amount'],
            ascending=[True, False]
        )

        # -----------------------------
        # Month Wise Payment Type Analytics
        # -----------------------------
        # Get unique months and sort them
        unique_months = sorted(df_filtered['Month'].unique())
        
        # Create a list to store each month's table
        month_tables = []
        
        for month in unique_months:
            month_data = payment_type_summary[payment_type_summary['Month'] == month].copy()
            
            # Select and rename columns for display
            month_table = month_data[[
                'Payment Type', 'Count', 'Amount', 
                'Service_Charge', 'ST_GST', 'Settlement_Amount',
                'Avg_Service_Charge', 'Service_Charge_Percentage'
            ]].copy()
            
            # Format Amount columns with commas and rupee symbol
            month_table['Amount'] = month_table['Amount'].apply(lambda x: f"₹{x:,.0f}")
            month_table['Service_Charge'] = month_table['Service_Charge'].apply(lambda x: f"₹{x:,.0f}")
            month_table['ST_GST'] = month_table['ST_GST'].apply(lambda x: f"₹{x:,.0f}")
            month_table['Settlement_Amount'] = month_table['Settlement_Amount'].apply(lambda x: f"₹{x:,.0f}")
            month_table['Avg_Service_Charge'] = month_table['Avg_Service_Charge'].apply(lambda x: f"₹{x:,.2f}")
            
            # Format Count with commas
            month_table['Count'] = month_table['Count'].apply(lambda x: f"{x:,}")
            
            # Add Grand Total row
            grand_total = pd.DataFrame({
                'Payment Type': ['Grand Total'],
                'Count': [f"{month_data['Count'].sum():,}"],
                'Amount': [f"₹{month_data['Amount'].sum():,.0f}"],
                'Service_Charge': [f"₹{month_data['Service_Charge'].sum():,.0f}"],
                'ST_GST': [f"₹{month_data['ST_GST'].sum():,.0f}"],
                'Settlement_Amount': [f"₹{month_data['Settlement_Amount'].sum():,.0f}"],
                'Avg_Service_Charge': [f"₹{(month_data['Service_Charge'].sum() / month_data['Count'].sum()):,.2f}" if month_data['Count'].sum() > 0 else "₹0.00"],
                'Service_Charge_Percentage': [f"{(month_data['Service_Charge'].sum() / month_data['Amount'].sum() * 100):.2f}%" if month_data['Amount'].sum() > 0 else "0%"]
            })
            
            month_table = pd.concat([month_table, grand_total], ignore_index=True)
            
            month_tables.append({
                'month': month,
                'table': month_table,
                'data': month_data
            })

        # -----------------------------
        # Grand Total
        # -----------------------------
        grand_total = pd.DataFrame({
            'Total Transactions': [df_filtered['Order Id'].nunique()],
            'Amount': [df_filtered['Amount'].sum()],
            'Service Charge': [df_filtered['Service Charge'].sum()],
            'ST/GST': [df_filtered['ST/GST'].sum()],
            'Settlement Amount': [df_filtered['Settlement Amount'].sum()],
            'Avg Service Charge': [df_filtered['Service Charge'].mean()],
            'Service Charge %': [(df_filtered['Service Charge'].sum() / df_filtered['Amount'].sum() * 100) if df_filtered['Amount'].sum() > 0 else 0]
        })

        # -----------------------------
        # Dashboard Metrics
        # -----------------------------
        col1, col2, col3, col4 = st.columns(4)

        col1.metric(
            "Transactions",
            f"{df_filtered['Order Id'].nunique():,}"
        )

        col2.metric(
            "Amount",
            f"₹{df_filtered['Amount'].sum():,.2f}"
        )

        col3.metric(
            "Service Charge",
            f"₹{df_filtered['Service Charge'].sum():,.2f}"
        )

        col4.metric(
            "Settlement",
            f"₹{df_filtered['Settlement Amount'].sum():,.2f}"
        )

        # -----------------------------
        # Display Reports
        # -----------------------------
        st.subheader("📋 Payment Mode SubType Summary")
        st.dataframe(
            unique_subtypes,
            width="stretch"
        )

        st.subheader("📅 Month Wise Payment Mode SubType")
        st.dataframe(
            subtype_summary,
            width="stretch"
        )

        st.subheader("📊 Month Wise Parent SubType")
        st.dataframe(
            parent_summary,
            width="stretch"
        )

        # -----------------------------
        # Payment Method Analytics by Month
        # -----------------------------
        st.subheader("💳 Payment Method Analytics")
        
        # Create tabs for each month
        if len(month_tables) > 0:
            tab_titles = [f"{mt['month']} - {pd.to_datetime(mt['month']).strftime('%b')}" for mt in month_tables]
            tabs = st.tabs(tab_titles)
            
            for tab, month_info in zip(tabs, month_tables):
                with tab:
                    # Display month header
                    month_name = pd.to_datetime(month_info['month']).strftime('%B %Y')
                    st.markdown(f"### {month_name}")
                    
                    # Display the table
                    st.dataframe(
                        month_info['table'],
                        width="stretch",
                        hide_index=True
                    )
        
        # Also show the detailed Payment Type Summary
        st.subheader("💳 Month Wise Payment Type Summary (Detailed)")
        st.dataframe(
            payment_type_summary_display[[
                'Month', 'Payment Type', 'Count', 'Amount', 
                'Service_Charge', 'ST_GST', 'Settlement_Amount',
                'Avg_Service_Charge', 'Service_Charge_Percentage'
            ]],
            width="stretch"
        )

        st.subheader("💰 Grand Total")
        st.dataframe(
            grand_total,
            width="stretch"
        )

        # -----------------------------
        # Export Excel
        # -----------------------------
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine='openpyxl'
        ) as writer:

            df_filtered.to_excel(
                writer,
                sheet_name='Raw_Data',
                index=False
            )

            unique_subtypes.to_excel(
                writer,
                sheet_name='SubType_Summary',
                index=False
            )

            subtype_summary.to_excel(
                writer,
                sheet_name='Month_Wise_SubType',
                index=False
            )

            parent_summary.to_excel(
                writer,
                sheet_name='Month_Wise_Parent',
                index=False
            )

            payment_type_summary.to_excel(
                writer,
                sheet_name='Month_Wise_PaymentType',
                index=False
            )
            
            # Add each month's analytics to separate sheets
            for month_info in month_tables:
                sheet_name = f"{month_info['month']}_Payment_Analytics"
                # Limit sheet name to 31 characters (Excel limitation)
                sheet_name = sheet_name[:31]
                month_info['table'].to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

            grand_total.to_excel(
                writer,
                sheet_name='Grand_Total',
                index=False
            )

        st.download_button(
            label="📥 Download Excel Report",
            data=output.getvalue(),
            file_name="Gateway_Report_Analytics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {str(e)}")
        st.info("Please check your file format and column names.")
