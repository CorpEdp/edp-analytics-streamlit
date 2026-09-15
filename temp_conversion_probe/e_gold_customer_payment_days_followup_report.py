"""
Auto-converted from: App-E-GoldCustomerPaymentDaysFollowupReport.py
Review this file before using — the converter does a best-effort wrap;
double-check indentation around any unusual control flow (loops, if/else
blocks that span large sections, etc.).
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

LABEL = "E Gold Customer Payment Days Followup Report"


def run():
    st.markdown("""
        <style>
        .main-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 2rem;
            border-radius: 15px;
            margin-bottom: 2rem;
            color: white;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .main-header h1 {
            margin: 0;
            font-size: 2.5rem;
            font-weight: 700;
        }
        .main-header p {
            margin: 0.5rem 0 0 0;
            opacity: 0.9;
            font-size: 1.1rem;
        }

        /* Enhanced Metric Cards */
        .metric-container {
            display: flex;
            gap: 1rem;
            margin: 1rem 0;
        }
        .metric-card {
            background: white;
            padding: 1.5rem;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            border-left: 5px solid #667eea;
            flex: 1;
            transition: transform 0.2s;
            height: 100%;
        }
        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.12);
        }
        .metric-card .metric-icon {
            font-size: 1.8rem;
            margin-bottom: 0.5rem;
        }
        .metric-card .metric-label {
            font-size: 0.85rem;
            color: #6b7280;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .metric-card .metric-value {
            font-size: 2rem;
            font-weight: 700;
            color: #1f2937;
            margin: 0.25rem 0;
            line-height: 1.2;
        }
        .metric-card .metric-sub {
            font-size: 0.8rem;
            color: #9ca3af;
        }
        .metric-card .metric-trend {
            display: inline-block;
            padding: 0.2rem 0.6rem;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
            margin-top: 0.25rem;
        }
        .metric-trend.up { background: #d1fae5; color: #065f46; }
        .metric-trend.down { background: #fee2e2; color: #991b1b; }
        .metric-trend.neutral { background: #e5e7eb; color: #4b5563; }

        /* Color variants */
        .metric-card.purple { border-left-color: #7c3aed; }
        .metric-card.blue { border-left-color: #3b82f6; }
        .metric-card.green { border-left-color: #10b981; }
        .metric-card.orange { border-left-color: #f59e0b; }
        .metric-card.red { border-left-color: #ef4444; }
        .metric-card.pink { border-left-color: #ec4899; }

        .info-card {
            background: #eff6ff;
            padding: 1rem 1.5rem;
            border-radius: 10px;
            border-left: 4px solid #3b82f6;
            margin: 1rem 0;
        }
        .warning-card {
            background: #fffbeb;
            padding: 1rem 1.5rem;
            border-radius: 10px;
            border-left: 4px solid #f59e0b;
            margin: 1rem 0;
        }
        .success-card {
            background: #ecfdf5;
            padding: 1rem 1.5rem;
            border-radius: 10px;
            border-left: 4px solid #10b981;
            margin: 1rem 0;
        }

        .stat-box {
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
            padding: 1.5rem;
            border-radius: 12px;
            text-align: center;
            border: 1px solid #e2e8f0;
            height: 100%;
        }
        .stat-box .stat-number {
            font-size: 2.5rem;
            font-weight: 700;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        .stat-box .stat-label {
            font-size: 0.9rem;
            color: #64748b;
            margin-top: 0.25rem;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 0.5rem;
            background: #f8fafc;
            padding: 0.5rem;
            border-radius: 12px;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 8px;
            padding: 0.5rem 1.5rem;
            font-weight: 500;
            color: #64748b;
            transition: all 0.2s;
        }
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            background: white;
            color: #1f2937;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }

        .date-range-display {
            font-size: 1.2rem !important;
            font-weight: 600 !important;
            color: #1f2937 !important;
            white-space: normal !important;
            word-break: break-word !important;
        }

        /* Configuration section styling */
        .config-section {
            background: #f8fafc;
            padding: 1.5rem;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            margin: 1rem 0;
        }

        /* Custom styling for number input */
        .threshold-value {
            font-size: 2rem;
            font-weight: 700;
            color: #667eea;
            text-align: center;
            padding: 0.5rem;
            background: white;
            border-radius: 8px;
            border: 2px solid #667eea;
        }

        /* Big number input styling */
        .big-number-input input {
            font-size: 1.5rem !important;
            font-weight: 600 !important;
            padding: 0.75rem !important;
        }
        </style>
    """, unsafe_allow_html=True)
    st.markdown("""
        <div class="main-header">
            <h1>📒 No Payment Passbook Report</h1>
            <p>Identify passbooks with no payments beyond threshold • Comprehensive Analytics</p>
        </div>
    """, unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "📂 Upload Excel or CSV File",
        type=["xlsx", "csv"],
        help="Upload a file containing payment records with 'Paid Date' and 'Passbook number' columns"
    )
    if uploaded_file:
        try:
            # Read the file
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.success(f"✅ File loaded successfully! Total records: {len(df):,}")

            # Check required columns
            required_columns = ["Paid Date", "Passbook number"]
            missing_columns = [col for col in required_columns if col not in df.columns]

            if missing_columns:
                st.error(f"❌ Missing required columns: {missing_columns}")
                st.info("Please ensure your file contains 'Paid Date' and 'Passbook number' columns")
                st.stop()

            # Data cleaning with progress bar
            with st.spinner("🔄 Processing data..."):
                progress_bar = st.progress(0)

                # Convert Passbook number to string
                df["Passbook number"] = df["Passbook number"].astype(str)
                progress_bar.progress(20)

                # Convert Paid Date to datetime
                df["Paid Date"] = pd.to_datetime(df["Paid Date"], errors="coerce")
                progress_bar.progress(40)

                # Convert numeric columns
                numeric_columns = ["Saved Amount", "Saved Metal Weight", "Benefit Metal Amount", "Benefit Metal Weight"]
                for col in numeric_columns:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
                progress_bar.progress(60)

                # Convert Installment number
                if "Installment number" in df.columns:
                    df["Installment number"] = pd.to_numeric(df["Installment number"], errors="coerce").fillna(0)
                progress_bar.progress(80)

                # Remove invalid rows
                original_count = len(df)
                df = df[df["Passbook number"].notna() & df["Passbook number"].str.strip() != ""]
                df = df[df["Paid Date"].notna()]

                if len(df) < original_count:
                    st.warning(f"⚠️ Removed {original_count - len(df):,} records with missing or invalid data")

                progress_bar.progress(100)
                progress_bar.empty()

            # Get the last date from the uploaded file
            file_last_date = df["Paid Date"].max()
            today = datetime.today()

            # ==================== ENHANCED DATA OVERVIEW ====================
            with st.expander("📊 Data Overview", expanded=True):
                # Calculate additional metrics
                total_records = len(df)
                unique_passbooks = df["Passbook number"].nunique()
                min_date = df["Paid Date"].min()
                max_date = df["Paid Date"].max()
                avg_payments = total_records / unique_passbooks if unique_passbooks > 0 else 0
                date_range_days = (max_date - min_date).days if pd.notna(min_date) and pd.notna(max_date) else 0

                # Calculate payment frequency
                payment_counts = df.groupby("Passbook number").size()
                max_payments = payment_counts.max() if len(payment_counts) > 0 else 0
                min_payments = payment_counts.min() if len(payment_counts) > 0 else 0

                # Format dates properly
                min_date_str = min_date.strftime('%b %d, %Y') if pd.notna(min_date) else 'N/A'
                max_date_str = max_date.strftime('%b %d, %Y') if pd.notna(max_date) else 'N/A'
                date_range_str = f"{min_date_str} — {max_date_str}"

                # First row: Main metrics in enhanced cards
                st.markdown("### 📈 Key Statistics")

                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    st.markdown(f"""
                        <div class="metric-card purple">
                            <div class="metric-icon">📋</div>
                            <div class="metric-label">Total Records</div>
                            <div class="metric-value">{total_records:,}</div>
                            <div class="metric-sub">Payment transactions</div>
                            <div class="metric-trend neutral">📊 Data</div>
                        </div>
                    """, unsafe_allow_html=True)

                with col2:
                    st.markdown(f"""
                        <div class="metric-card blue">
                            <div class="metric-icon">👤</div>
                            <div class="metric-label">Unique Passbooks</div>
                            <div class="metric-value">{unique_passbooks:,}</div>
                            <div class="metric-sub">Active customers</div>
                            <div class="metric-trend up">📈 {avg_payments:.1f} avg payments</div>
                        </div>
                    """, unsafe_allow_html=True)

                with col3:
                    st.markdown(f"""
                        <div class="metric-card green">
                            <div class="metric-icon">📅</div>
                            <div class="metric-label">Date Range</div>
                            <div class="metric-value date-range-display" style="font-size:1.1rem;">
                                {min_date_str} <span style="color:#9ca3af;">to</span> {max_date_str}
                            </div>
                            <div class="metric-sub">{date_range_days} days covered</div>
                            <div class="metric-trend neutral">📆 Full range</div>
                        </div>
                    """, unsafe_allow_html=True)

                with col4:
                    st.markdown(f"""
                        <div class="metric-card orange">
                            <div class="metric-icon">📈</div>
                            <div class="metric-label">Avg Payments/Passbook</div>
                            <div class="metric-value">{avg_payments:.1f}</div>
                            <div class="metric-sub">Range: {min_payments} - {max_payments}</div>
                            <div class="metric-trend up">📊 Per customer</div>
                        </div>
                    """, unsafe_allow_html=True)

                # Second row: Additional insights
                st.markdown("### 📊 Data Insights")

                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    # Payment concentration
                    top_10_percent = int(unique_passbooks * 0.1)
                    top_passbooks = df.groupby("Passbook number").size().nlargest(top_10_percent)
                    top_records = top_passbooks.sum() if len(top_passbooks) > 0 else 0
                    concentration = (top_records / total_records * 100) if total_records > 0 else 0

                    st.markdown(f"""
                        <div class="stat-box">
                            <div class="stat-number">{concentration:.1f}%</div>
                            <div class="stat-label">📊 Top 10% Passbooks</div>
                            <div style="font-size:0.8rem;color:#94a3b8;">Account for {top_records:,} records</div>
                        </div>
                    """, unsafe_allow_html=True)

                with col2:
                    # Most active period
                    df["Month"] = df["Paid Date"].dt.to_period("M")
                    monthly_counts = df.groupby("Month").size()
                    busiest_month = monthly_counts.idxmax() if len(monthly_counts) > 0 else "N/A"
                    busiest_count = monthly_counts.max() if len(monthly_counts) > 0 else 0

                    st.markdown(f"""
                        <div class="stat-box">
                            <div class="stat-number">{busiest_month}</div>
                            <div class="stat-label">📅 Busiest Month</div>
                            <div style="font-size:0.8rem;color:#94a3b8;">{busiest_count:,} transactions</div>
                        </div>
                    """, unsafe_allow_html=True)

                with col3:
                    # Data completeness
                    has_phone = "Customer Phone Number" in df.columns
                    phone_coverage = df["Customer Phone Number"].notna().sum() if has_phone else 0
                    phone_pct = (phone_coverage / total_records * 100) if total_records > 0 else 0

                    st.markdown(f"""
                        <div class="stat-box">
                            <div class="stat-number">{phone_pct:.0f}%</div>
                            <div class="stat-label">📱 Phone Number Coverage</div>
                            <div style="font-size:0.8rem;color:#94a3b8;">{phone_coverage:,} of {total_records:,} records</div>
                        </div>
                    """, unsafe_allow_html=True)

                with col4:
                    # Unique schemes
                    has_scheme = "Scheme Name" in df.columns
                    unique_schemes = df["Scheme Name"].nunique() if has_scheme else 0

                    st.markdown(f"""
                        <div class="stat-box">
                            <div class="stat-number">{unique_schemes if has_scheme else 'N/A'}</div>
                            <div class="stat-label">🏷️ Unique Schemes</div>
                            <div style="font-size:0.8rem;color:#94a3b8;">{'Available' if has_scheme else 'Not in data'}</div>
                        </div>
                    """, unsafe_allow_html=True)

                # Sample data with better display
                st.markdown("### 📋 Sample Records")
                st.dataframe(df.head(10), use_container_width=True, height=300)

            # ==================== CONFIGURATION SECTION ====================
            st.markdown("---")
            st.subheader("🎯 Configuration")

            # Get the last date from the uploaded file
            file_last_date = df["Paid Date"].max()
            today = datetime.today()

            # Configuration with only manual entry
            st.markdown('<div class="config-section">', unsafe_allow_html=True)

            # Initialize session state for threshold if not exists
            if 'days_threshold' not in st.session_state:
                st.session_state.days_threshold = 30

            col1, col2, col3 = st.columns([1, 2, 1])

            with col1:
                st.markdown("**📆 Days Threshold**")

            with col2:
                # Manual input only - big and prominent
                days_threshold = st.number_input(
                    "Enter number of days",
                    min_value=1,
                    max_value=365,
                    value=st.session_state.days_threshold,
                    step=1,
                    key="days_input",
                    help="Enter the number of days to check for no payment",
                    label_visibility="collapsed"
                )
                st.session_state.days_threshold = days_threshold

            with col3:
                # Show current value prominently
                st.markdown(f"""
                    <div style="text-align: center; padding: 0.5rem; background: #667eea; border-radius: 8px;">
                        <span style="color: white; font-size: 1.2rem; font-weight: 600;">{days_threshold} days</span>
                    </div>
                """, unsafe_allow_html=True)

            # Option to use specific end date
            use_custom_end_date = st.checkbox(
                "📌 Use custom reference date",
                value=False,
                help="By default, uses the last date in the file as the reference date"
            )

            if use_custom_end_date:
                end_date = st.date_input(
                    "📅 Reference End Date",
                    value=today.date(),
                    max_value=today.date(),
                    help="Date to calculate days from (default: today)"
                )
                end_date = pd.Timestamp(end_date)
                reference_date = end_date
            else:
                end_date_display = file_last_date.date() if hasattr(file_last_date, 'date') else file_last_date
                st.info(f"📅 Reference Date: **{end_date_display}**")
                reference_date = file_last_date

            st.markdown('</div>', unsafe_allow_html=True)

            # Display current configuration
            st.markdown(f"""
                <div class="info-card">
                    <strong>🎯 Current Configuration:</strong><br>
                    • Days Threshold: <strong>{days_threshold} days</strong> (No payment for more than {days_threshold} days)<br>
                    • Reference Date: <strong>{reference_date.strftime('%Y-%m-%d') if hasattr(reference_date, 'strftime') else reference_date}</strong>
                </div>
            """, unsafe_allow_html=True)

            # Process data
            with st.spinner("📊 Analyzing data..."):
                # Sort and process
                df_sorted = df.sort_values(["Passbook number", "Paid Date"], ascending=[True, True])

                # Calculate cumulative totals
                numeric_columns = ["Saved Amount", "Saved Metal Weight", "Benefit Metal Amount", "Benefit Metal Weight"]
                for col in numeric_columns:
                    if col in df_sorted.columns:
                        df_sorted[f"{col}_cumulative"] = df_sorted.groupby("Passbook number")[col].cumsum()

                # Get latest records
                latest_with_cumulative = df_sorted.groupby("Passbook number").last().reset_index()
                for col in numeric_columns:
                    if col in df_sorted.columns:
                        if f"{col}_cumulative" in latest_with_cumulative.columns:
                            latest_with_cumulative[col] = latest_with_cumulative[f"{col}_cumulative"]

                # Get installment count
                installment_count = df.groupby("Passbook number")["Installment number"].count().reset_index()
                installment_count.columns = ["Passbook number", "Total Installments"]

                # Calculate last payment
                last_payment_per_passbook = df.groupby("Passbook number")["Paid Date"].max().reset_index()
                last_payment_per_passbook.columns = ["Passbook number", "Last Payment Date"]

                # Calculate days since last payment
                last_payment_per_passbook["Days Since Last Payment"] = last_payment_per_passbook["Last Payment Date"].apply(
                    lambda x: (reference_date - x).days if hasattr(reference_date, 'days') else (reference_date - x).days
                )

                # Find inactive passbooks
                inactive_passbooks = last_payment_per_passbook[
                    last_payment_per_passbook["Days Since Last Payment"] > days_threshold
                ]

                # Get inactive records
                inactive_passbook_numbers = inactive_passbooks["Passbook number"].unique()
                inactive_records = latest_with_cumulative[latest_with_cumulative["Passbook number"].isin(inactive_passbook_numbers)]

                # Merge with installment count
                inactive_records = inactive_records.merge(installment_count, on="Passbook number", how="left")

                # Create installment slabs
                def get_installment_slab(count):
                    if count <= 0:
                        return "0"
                    elif count == 1:
                        return "1"
                    elif count == 2:
                        return "2"
                    elif count == 3:
                        return "3"
                    elif count == 4:
                        return "4"
                    elif count == 5:
                        return "5"
                    elif count == 6:
                        return "6"
                    elif count == 7:
                        return "7"
                    elif count == 8:
                        return "8"
                    elif count == 9:
                        return "9"
                    elif count == 10:
                        return "10"
                    else:
                        return ">10"

                inactive_records["Installment Slab"] = inactive_records["Total Installments"].apply(get_installment_slab)

                # Create slab summary
                slab_summary = inactive_records.groupby("Installment Slab").agg({
                    "Passbook number": "count",
                    "Saved Amount": "sum",
                    "Saved Metal Weight": "sum",
                    "Benefit Metal Amount": "sum",
                    "Benefit Metal Weight": "sum"
                }).reset_index()
                slab_summary.columns = ["Installment Slab", "Number of Passbooks", "Total Saved Amount", "Total Saved Metal Weight", "Total Benefit Metal Amount", "Total Benefit Metal Weight"]

                total_passbooks = slab_summary["Number of Passbooks"].sum()
                slab_summary["Percentage"] = (slab_summary["Number of Passbooks"] / total_passbooks * 100).round(1) if total_passbooks > 0 else 0

                # Sort slabs
                slab_order = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", ">10"]
                slab_summary["Installment Slab"] = pd.Categorical(slab_summary["Installment Slab"], categories=slab_order, ordered=True)
                slab_summary = slab_summary.sort_values("Installment Slab").reset_index(drop=True)

            # Display main metrics
            st.markdown("---")
            st.subheader("📊 Key Metrics")

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f"""
                    <div class="metric-card purple">
                        <div style="font-size:0.85rem;color:#6b7280;font-weight:500;text-transform:uppercase;letter-spacing:0.5px;">Total Unique Passbooks</div>
                        <div style="font-size:2rem;font-weight:700;color:#1f2937;">{df['Passbook number'].nunique():,}</div>
                    </div>
                """, unsafe_allow_html=True)

            with col2:
                st.markdown(f"""
                    <div class="metric-card red">
                        <div style="font-size:0.85rem;color:#6b7280;font-weight:500;text-transform:uppercase;letter-spacing:0.5px;">No Payment Passbooks</div>
                        <div style="font-size:2rem;font-weight:700;color:#dc3545;">{len(inactive_passbooks):,}</div>
                    </div>
                """, unsafe_allow_html=True)

            with col3:
                percentage = (len(inactive_passbooks) / df["Passbook number"].nunique() * 100) if df["Passbook number"].nunique() > 0 else 0
                st.markdown(f"""
                    <div class="metric-card orange">
                        <div style="font-size:0.85rem;color:#6b7280;font-weight:500;text-transform:uppercase;letter-spacing:0.5px;">No Payment %</div>
                        <div style="font-size:2rem;font-weight:700;color:#f59e0b;">{percentage:.1f}%</div>
                    </div>
                """, unsafe_allow_html=True)

            with col4:
                active_count = df["Passbook number"].nunique() - len(inactive_passbooks)
                st.markdown(f"""
                    <div class="metric-card green">
                        <div style="font-size:0.85rem;color:#6b7280;font-weight:500;text-transform:uppercase;letter-spacing:0.5px;">Active Passbooks</div>
                        <div style="font-size:2rem;font-weight:700;color:#10b981;">{active_count:,}</div>
                    </div>
                """, unsafe_allow_html=True)

            # Show information about the report
            if len(inactive_passbooks) > 0:
                min_days = inactive_passbooks["Days Since Last Payment"].min()
                max_days = inactive_passbooks["Days Since Last Payment"].max()
                avg_days = inactive_passbooks["Days Since Last Payment"].mean()

                st.markdown(f"""
                    <div class="warning-card">
                        <strong>📋 No Payment Report Details:</strong><br>
                        • Finding passbooks with <strong>NO payment for more than {days_threshold} days</strong><br>
                        • Reference Date: <strong>{reference_date.strftime('%Y-%m-%d') if hasattr(reference_date, 'strftime') else reference_date}</strong><br>
                        • Total No Payment passbooks: <strong>{len(inactive_passbooks):,}</strong><br>
                        • Days without payment range: <strong>{min_days} to {max_days} days</strong> (Average: {avg_days:.0f} days)<br>
                        • Values shown are <strong>CUMULATIVE TOTALS</strong> (sum of all transactions)
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                    <div class="success-card">
                        <strong>🎉 All passbooks have made payments within the last {days_threshold} days!</strong><br>
                        No passbooks found without payment.
                    </div>
                """, unsafe_allow_html=True)

            # Create tabs for better organization
            if len(inactive_passbooks) > 0:
                tab1, tab2, tab3, tab4 = st.tabs(["📊 Installment Analysis", "📋 Passbook Report", "📱 Campaign Report", "📈 Visual Analytics"])

                with tab1:
                    st.subheader("📊 Installment-wise Distribution of No Payment Passbooks")

                    # Display the slab summary with better formatting
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        # Format percentage for display
                        slab_display = slab_summary.copy()
                        slab_display["Percentage"] = slab_display["Percentage"].astype(str) + "%"
                        st.dataframe(slab_display, use_container_width=True)
                    with col2:
                        # Pie chart for installment distribution
                        fig = px.pie(
                            slab_summary, 
                            values="Number of Passbooks", 
                            names="Installment Slab",
                            title="Distribution by Installment Slab",
                            color_discrete_sequence=px.colors.qualitative.Set3
                        )
                        fig.update_traces(textposition='inside', textinfo='percent+label')
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)

                    # Bar chart
                    fig = px.bar(
                        slab_summary,
                        x="Installment Slab",
                        y="Number of Passbooks",
                        title="Number of Passbooks by Installment Slab",
                        color="Installment Slab",
                        color_discrete_sequence=px.colors.qualitative.Set2
                    )
                    fig.update_layout(showlegend=False, height=400)
                    st.plotly_chart(fig, use_container_width=True)

                with tab2:
                    st.subheader("📋 No Payment Passbooks Report")

                    # Search filter with better UI
                    search_term = st.text_input(
                        "🔍 Search Passbook Number",
                        placeholder="Enter passbook number to search...",
                        help="Filter the results by passbook number"
                    )

                    # Apply search filter
                    filtered_report = inactive_records.copy()
                    if search_term:
                        filtered_report = filtered_report[
                            filtered_report["Passbook number"].str.contains(search_term, case=False, na=False)
                        ]
                        if len(filtered_report) == 0:
                            st.info("No matching passbook found")

                    # Display the report with better styling
                    st.dataframe(
                        filtered_report,
                        use_container_width=True,
                        height=500
                    )

                    # Download buttons
                    st.subheader("📥 Download Report")
                    col1, col2 = st.columns(2)

                    with col1:
                        # Prepare export data - keep only essential columns
                        export_df = inactive_records.copy()

                        # Define column order - Customer Name before Customer Phone Number
                        keep_columns = ["Passbook number", "Customer Name", "Customer Phone Number", 
                                      "Total Installments", "Installment Slab", "Saved Amount", 
                                      "Saved Metal Weight", "Benefit Metal Amount", "Benefit Metal Weight"]

                        # Only keep columns that exist and maintain order
                        export_columns = [col for col in keep_columns if col in export_df.columns]
                        export_df = export_df[export_columns]

                        # Add days without payment and last payment date
                        days_mapping = last_payment_per_passbook.set_index("Passbook number")["Days Since Last Payment"].to_dict()
                        export_df["Days Without Payment"] = export_df["Passbook number"].map(days_mapping)
                        last_date_mapping = last_payment_per_passbook.set_index("Passbook number")["Last Payment Date"].to_dict()
                        export_df["Last Date Payment"] = export_df["Passbook number"].map(last_date_mapping)

                        # Reorder columns to ensure Customer Name comes before Phone Number
                        final_columns = ["Sno", "Passbook number", "Customer Name", "Customer Phone Number", 
                                       "Total Installments", "Installment Slab", "Days Without Payment", 
                                       "Last Date Payment", "Saved Amount", "Saved Metal Weight", 
                                       "Benefit Metal Amount", "Benefit Metal Weight"]

                        # Add Sno
                        export_df.insert(0, "Sno", range(1, len(export_df) + 1))

                        # Select only columns that exist in the dataframe
                        available_columns = [col for col in final_columns if col in export_df.columns]
                        export_df = export_df[available_columns]

                        # Excel download
                        output_excel = BytesIO()
                        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
                            export_df.to_excel(writer, index=False, sheet_name="No Payment Passbooks")

                            # Installment Slab Summary
                            slab_summary_export = slab_summary.copy()
                            slab_summary_export["Percentage"] = slab_summary_export["Percentage"].astype(str) + "%"
                            slab_summary_export.to_excel(writer, index=False, sheet_name="Installment Slab Summary")

                            # Summary
                            summary_data = {
                                "Metric": [
                                    "Report Generated", "Report Type", "Days Threshold", "Reference Date",
                                    "Total Unique Passbooks", "No Payment Passbooks", "No Payment Percentage"
                                ],
                                "Value": [
                                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    f"No Payment Passbooks (> {days_threshold} days)",
                                    days_threshold,
                                    reference_date.strftime("%Y-%m-%d") if hasattr(reference_date, 'strftime') else str(reference_date),
                                    df["Passbook number"].nunique(),
                                    len(inactive_passbooks),
                                    f"{percentage:.1f}%"
                                ]
                            }
                            pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name="Summary")

                        output_excel.seek(0)
                        st.download_button(
                            label="⬇ Download Excel Report",
                            data=output_excel,
                            file_name=f"No_Payment_Passbooks_More_Than_{days_threshold}_Days.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                    with col2:
                        csv_data = export_df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="⬇ Download CSV Report",
                            data=csv_data,
                            file_name=f"No_Payment_Passbooks_More_Than_{days_threshold}_Days.csv",
                            mime="text/csv",
                            use_container_width=True
                        )

                with tab3:
                    if "Customer Phone Number" in df.columns:
                        st.subheader("📱 Campaign Report (Phone Number Wise)")
                        st.info("📋 This report groups passbooks by Phone Number for campaign targeting")

                        # Create campaign report
                        days_mapping = last_payment_per_passbook.set_index("Passbook number")["Days Since Last Payment"].to_dict()
                        inactive_records["Days Without Payment"] = inactive_records["Passbook number"].map(days_mapping)

                        if "Customer Phone Number" in inactive_records.columns:
                            campaign_data = inactive_records[inactive_records["Customer Phone Number"].notna() & 
                                                           (inactive_records["Customer Phone Number"] != "") &
                                                           (inactive_records["Customer Phone Number"] != "0")]

                            if not campaign_data.empty:
                                # Create campaign summary with Customer Name included
                                campaign_summary = campaign_data.groupby("Customer Phone Number").agg({
                                    "Passbook number": ["count", lambda x: ", ".join(x)],
                                    "Customer Name": lambda x: ", ".join(x.unique()),  # Add Customer Names
                                    "Days Without Payment": ["min", "max", "mean"],
                                    "Saved Amount": "sum",
                                    "Saved Metal Weight": "sum",
                                    "Benefit Metal Amount": "sum",
                                    "Benefit Metal Weight": "sum"
                                }).reset_index()

                                # Define column order - Customer Phone Number first, then Customer Names
                                campaign_summary.columns = [
                                    "Customer Phone Number", "Passbook Count", "Passbook Numbers",
                                    "Customer Names", "Min Days", "Max Days", "Avg Days",
                                    "Total Saved Amount", "Total Saved Metal Weight",
                                    "Total Benefit Metal Amount", "Total Benefit Metal Weight"
                                ]

                                # Reorder columns to put Customer Names before Passbook Numbers
                                campaign_summary = campaign_summary[["Customer Phone Number", "Customer Names", "Passbook Count", 
                                                                    "Passbook Numbers", "Min Days", "Max Days", "Avg Days",
                                                                    "Total Saved Amount", "Total Saved Metal Weight",
                                                                    "Total Benefit Metal Amount", "Total Benefit Metal Weight"]]

                                campaign_summary = campaign_summary.sort_values("Passbook Count", ascending=False)

                                # Display metrics
                                col1, col2, col3, col4 = st.columns(4)
                                with col1:
                                    st.metric("📱 Unique Phone Numbers", len(campaign_summary))
                                with col2:
                                    st.metric("📋 Total Passbooks", campaign_summary["Passbook Count"].sum())
                                with col3:
                                    st.metric("📊 Avg Passbooks/Phone", f"{campaign_summary['Passbook Count'].mean():.1f}")
                                with col4:
                                    st.metric("🏆 Max Passbooks/Phone", campaign_summary["Passbook Count"].max())

                                # Display campaign data
                                st.dataframe(campaign_summary, use_container_width=True)

                                # Distribution chart
                                fig = px.histogram(
                                    campaign_summary,
                                    x="Passbook Count",
                                    title="Distribution of Passbooks per Phone Number",
                                    color_discrete_sequence=["#667eea"]
                                )
                                fig.update_layout(xaxis_title="Number of Passbooks", yaxis_title="Count", height=400)
                                st.plotly_chart(fig, use_container_width=True)

                                # Download campaign report
                                st.subheader("📥 Download Campaign Report")
                                col1, col2 = st.columns(2)

                                with col1:
                                    output_campaign = BytesIO()
                                    with pd.ExcelWriter(output_campaign, engine="openpyxl") as writer:
                                        # Export campaign data with Customer Name before Phone Number
                                        campaign_export = campaign_data.copy()
                                        keep_campaign_columns = ["Customer Name", "Customer Phone Number", "Passbook number", 
                                                                "Days Without Payment", "Saved Amount", "Saved Metal Weight",
                                                                "Benefit Metal Amount", "Benefit Metal Weight"]
                                        campaign_export_columns = [col for col in keep_campaign_columns if col in campaign_export.columns]
                                        campaign_export = campaign_export[campaign_export_columns]
                                        campaign_export.to_excel(writer, index=False, sheet_name="Campaign Data")

                                        # Campaign Summary with Customer Names before Passbook Numbers
                                        campaign_summary_export = campaign_summary.copy()
                                        campaign_summary_export.to_excel(writer, index=False, sheet_name="Campaign Summary")

                                        # Apply wrap text formatting to Passbook Numbers column
                                        from openpyxl.styles import Alignment
                                        summary_worksheet = writer.sheets['Campaign Summary']
                                        if 'Passbook Numbers' in campaign_summary_export.columns:
                                            col_idx = campaign_summary_export.columns.get_loc('Passbook Numbers') + 1
                                            for row in range(2, len(campaign_summary_export) + 2):
                                                cell = summary_worksheet.cell(row=row, column=col_idx)
                                                cell.alignment = Alignment(wrap_text=True, vertical='top')
                                            summary_worksheet.column_dimensions[chr(64 + col_idx)].width = 50

                                        # Installment Slab Summary
                                        slab_summary_export = slab_summary.copy()
                                        slab_summary_export["Percentage"] = slab_summary_export["Percentage"].astype(str) + "%"
                                        slab_summary_export.to_excel(writer, index=False, sheet_name="Installment Slab Summary")

                                    output_campaign.seek(0)
                                    st.download_button(
                                        label="⬇ Download Campaign Excel",
                                        data=output_campaign,
                                        file_name=f"Campaign_Report_No_Payment_More_Than_{days_threshold}_Days.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True
                                    )

                                with col2:
                                    # CSV with Customer Name before Phone Number
                                    campaign_csv_export = campaign_data.copy()
                                    keep_csv_columns = ["Customer Name", "Customer Phone Number", "Passbook number", 
                                                       "Days Without Payment", "Saved Amount", "Saved Metal Weight",
                                                       "Benefit Metal Amount", "Benefit Metal Weight"]
                                    csv_columns = [col for col in keep_csv_columns if col in campaign_csv_export.columns]
                                    campaign_csv_export = campaign_csv_export[csv_columns]
                                    campaign_csv = campaign_csv_export.to_csv(index=False).encode('utf-8')
                                    st.download_button(
                                        label="⬇ Download Campaign CSV",
                                        data=campaign_csv,
                                        file_name=f"Campaign_Report_No_Payment_More_Than_{days_threshold}_Days.csv",
                                        mime="text/csv",
                                        use_container_width=True
                                    )
                            else:
                                st.warning("⚠️ No valid phone numbers found for campaign report")
                        else:
                            st.warning("⚠️ 'Customer Phone Number' column not found")
                    else:
                        st.info("📱 'Customer Phone Number' column not available. Campaign report requires this column.")

                with tab4:
                    st.subheader("📈 Visual Analytics")

                    # Days without payment distribution
                    st.subheader("Distribution of Days Without Payment")
                    fig = px.histogram(
                        inactive_passbooks,
                        x="Days Since Last Payment",
                        nbins=20,
                        title="Distribution of Days Without Payment",
                        color_discrete_sequence=["#764ba2"]
                    )
                    fig.update_layout(xaxis_title="Days Without Payment", yaxis_title="Number of Passbooks", height=400)
                    st.plotly_chart(fig, use_container_width=True)

                    # Last payment date distribution
                    st.subheader("Last Payment Date Distribution")
                    last_payment_dist = inactive_passbooks["Last Payment Date"].dt.date.value_counts().sort_index().reset_index()
                    last_payment_dist.columns = ["Date", "Count"]
                    fig = px.bar(
                        last_payment_dist,
                        x="Date",
                        y="Count",
                        title="Last Payment Date Distribution",
                        color_discrete_sequence=["#667eea"]
                    )
                    fig.update_layout(xaxis_title="Date", yaxis_title="Number of Passbooks", height=400)
                    st.plotly_chart(fig, use_container_width=True)

                    # Top 10 most inactive
                    st.subheader("Top 10 Passbooks - Longest Without Payment")
                    top_inactive = inactive_passbooks.nlargest(10, "Days Since Last Payment")
                    top_inactive_details = top_inactive.merge(
                        inactive_records[["Passbook number", "Customer Name"] if "Customer Name" in inactive_records.columns else ["Passbook number"]],
                        on="Passbook number",
                        how="left"
                    )

                    fig = px.bar(
                        top_inactive_details,
                        x="Passbook number",
                        y="Days Since Last Payment",
                        title="Top 10 Passbooks with Longest No Payment Period",
                        color="Days Since Last Payment",
                        color_continuous_scale="Reds"
                    )
                    fig.update_layout(xaxis_title="Passbook Number", yaxis_title="Days Without Payment", height=400)
                    st.plotly_chart(fig, use_container_width=True)

        except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
            st.info("Please check your file format and try again.")

    else:
        # Show instructions when no file is uploaded
        st.info("👆 Please upload an Excel or CSV file to generate the report")

        with st.expander("📖 How to Use This App", expanded=True):
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("""
                    ### 🚀 Getting Started
                    1. **Upload** your Excel or CSV file
                    2. **Set** the days threshold using the number input
                    3. **Choose** reference date (file's last date or custom)
                    4. **View** passbooks with NO payment
                    5. **Download** reports in Excel or CSV

                    ### 📋 Required Columns
                    - **Paid Date**: Date of payment
                    - **Passbook number**: Unique identifier
                    - **Customer Phone Number**: For campaign report
                    - **Customer Name**: For better identification
                """)

            with col2:
                st.markdown("""
                    ### 📊 What This Report Does
                    - Finds passbooks where **last payment > N days ago**
                    - Shows **unique passbook details**
                    - **CUMULATIVE TOTALS** for all transactions
                    - **Installment Slab Summary**
                    - **Campaign Report** by Phone Number

                    ### 📥 Download Options
                    1. **No Payment Report**: Detailed passbook data
                    2. **Campaign Report**: Phone number-wise grouping
                    3. **Both include**: Installment Slab Summary
                """)

            st.markdown("---")
            st.markdown("""
                ### 💡 Pro Tips
                - **Enter** the days threshold directly in the number field
                - **Search** for specific passbooks using the search bar
                - **Campaign Report** is perfect for digital marketing targeting
                - **Visual Analytics** tab shows insights through charts
                - **All reports** include cumulative totals, not just latest payment
            """)
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.caption(f"📅 Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    with col2:
        st.caption("📒 No Payment Passbook Report")
    with col3:
        st.caption("💡 Powered by Streamlit")
