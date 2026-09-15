"""
Auto-converted from: ERP-BhimaJewelleryCustomerPaymentLogimaxCashfree.py
Review this file before using — the converter does a best-effort wrap;
double-check indentation around any unusual control flow (loops, if/else
blocks that span large sections, etc.).
"""

import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO
from datetime import datetime
import os
import numpy as np
def load_data(report_types, selected_branch):
    """Load data from Excel files - handles Cashfree and Logimax separately"""
    
    if not isinstance(report_types, list):
        report_types = [report_types]
    
    cashfree_frames = []
    logimax_frames = []
    branches = ["Madurai", "Salem"]
    column_info = []
    debug_info = []
    
    debug_info.append(f"Base Path: {BASE_PATH}")
    debug_info.append(f"Path exists: {Path(BASE_PATH).exists()}")
    
    if not Path(BASE_PATH).exists():
        debug_info.append(f"❌ ERROR: Base path does not exist: {BASE_PATH}")
        return pd.DataFrame(), pd.DataFrame(), debug_info, []
    
    for branch in branches:
        if selected_branch != "All" and selected_branch != branch:
            continue
            
        folder = Path(BASE_PATH) / branch
        debug_info.append(f"\n📁 Checking branch: {branch}")
        debug_info.append(f"   Folder path: {folder}")
        
        if not folder.exists():
            debug_info.append(f"   ❌ Folder not found: {folder}")
            continue
        
        # List ALL Excel files for debugging
        all_excel_files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))
        debug_info.append(f"   📊 Found {len(all_excel_files)} Excel files in {branch}")
        
        if all_excel_files:
            debug_info.append(f"   📄 All files: {[f.name for f in all_excel_files]}")
        
        # Load Cashfree files
        if "Cashfree" in report_types or "All" in report_types:
            cashfree_patterns = [
                "*Cashfree*.xlsx", "*cashfree*.xlsx", 
                "*CF*.xlsx", "*cf*.xlsx",
                "*Cashfree*.xls", "*cashfree*.xls"
            ]
            cashfree_files = []
            for pattern in cashfree_patterns:
                cashfree_files.extend(folder.glob(pattern))
            cashfree_files = list(set(cashfree_files))
            
            debug_info.append(f"   🔍 Found {len(cashfree_files)} Cashfree files in {branch}")
            if cashfree_files:
                debug_info.append(f"   📄 Cashfree files: {[f.name for f in cashfree_files]}")
            
            for file in cashfree_files:
                try:
                    debug_info.append(f"   📖 Reading Cashfree: {file.name}")
                    df = pd.read_excel(file)
                    df["Branch"] = branch
                    df["Source_File"] = file.name
                    df["Report_Type"] = "Cashfree"
                    cashfree_frames.append(df)
                    
                    column_info.append({
                        'Branch': branch,
                        'File': file.name,
                        'Report_Type': 'Cashfree',
                        'Columns': list(df.columns),
                        'Rows': len(df)
                    })
                    debug_info.append(f"   ✅ Loaded Cashfree: {file.name} ({len(df):,} rows)")
                except Exception as ex:
                    debug_info.append(f"   ❌ Error reading Cashfree {file.name}: {ex}")
        
        # Load Logimax files
        if "Logimax" in report_types or "All" in report_types:
            logimax_patterns = [
                "*Logimax*.xlsx", "*logimax*.xlsx",
                "*LG*.xlsx", "*lg*.xlsx",
                "*Logimax*.xls", "*logimax*.xls",
                "*padam*.xlsx", "*padam*.xls"
            ]
            logimax_files = []
            for pattern in logimax_patterns:
                logimax_files.extend(folder.glob(pattern))
            logimax_files = list(set(logimax_files))
            
            debug_info.append(f"   🔍 Found {len(logimax_files)} Logimax files in {branch}")
            if logimax_files:
                debug_info.append(f"   📄 Logimax files: {[f.name for f in logimax_files]}")
            
            for file in logimax_files:
                try:
                    debug_info.append(f"   📖 Reading Logimax: {file.name}")
                    df = pd.read_excel(file)
                    df["Branch"] = branch
                    df["Source_File"] = file.name
                    df["Report_Type"] = "Logimax"
                    logimax_frames.append(df)
                    
                    column_info.append({
                        'Branch': branch,
                        'File': file.name,
                        'Report_Type': 'Logimax',
                        'Columns': list(df.columns),
                        'Rows': len(df)
                    })
                    debug_info.append(f"   ✅ Loaded Logimax: {file.name} ({len(df):,} rows)")
                except Exception as ex:
                    debug_info.append(f"   ❌ Error reading Logimax {file.name}: {ex}")
    
    # Combine frames
    cashfree_df = pd.DataFrame()
    logimax_df = pd.DataFrame()
    
    if cashfree_frames:
        cashfree_df = pd.concat(cashfree_frames, ignore_index=True)
        debug_info.append(f"\n✅ Total Cashfree records: {len(cashfree_df):,}")
    
    if logimax_frames:
        logimax_df = pd.concat(logimax_frames, ignore_index=True)
        debug_info.append(f"✅ Total Logimax records: {len(logimax_df):,}")
    
    return cashfree_df, logimax_df, debug_info, column_info
def process_cashfree_data(df):
    """Process Cashfree DataFrame"""
    if df.empty:
        return df
    
    # Find date column
    date_col = None
    for col in df.columns:
        if 'transaction' in col.lower() or 'date' in col.lower() or 'time' in col.lower():
            date_col = col
            break
    
    if date_col:
        try:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            df["Date"] = df[date_col].dt.date
            df["Month"] = df[date_col].dt.strftime("%B")
        except:
            df["Date"] = datetime.now().date()
            df["Month"] = datetime.now().strftime("%B")
    else:
        df["Date"] = datetime.now().date()
        df["Month"] = datetime.now().strftime("%B")
    
    return df
def process_logimax_data(df):
    """Process Logimax DataFrame"""
    if df.empty:
        return df
    
    # Find date column - Logimax typically uses 'date_payment'
    date_col = None
    for col in df.columns:
        if 'date_payment' in col.lower() or 'date' in col.lower():
            date_col = col
            break
    
    if date_col:
        try:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            df["Date"] = df[date_col].dt.date
            df["Month"] = df[date_col].dt.strftime("%B")
        except:
            df["Date"] = datetime.now().date()
            df["Month"] = datetime.now().strftime("%B")
    else:
        df["Date"] = datetime.now().date()
        df["Month"] = datetime.now().strftime("%B")
    
    return df
def apply_filters(df, from_date, to_date, search_text, customer_search):
    """Apply filters to DataFrame"""
    if df.empty:
        return df
    
    filtered = df.copy()
    
    # Apply date filter
    if "Date" in filtered.columns:
        filtered = filtered[
            (filtered["Date"] >= from_date) & 
            (filtered["Date"] <= to_date)
        ]
    
    # Apply customer search
    if customer_search:
        mask = filtered.astype(str).apply(
            lambda x: x.str.contains(customer_search, case=False, na=False)
        ).any(axis=1)
        filtered = filtered[mask]
    
    # Apply general search
    if search_text:
        mask = filtered.astype(str).apply(
            lambda x: x.str.contains(search_text, case=False, na=False)
        ).any(axis=1)
        filtered = filtered[mask]
    
    return filtered
def create_excel_with_sheets(cashfree_df, logimax_df):
    """Create Excel with separate sheets - Logimax first, then Cashfree"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Logimax sheet first
        if not logimax_df.empty:
            export_df = logimax_df.copy()
            # Remove internal columns
            for col in ['Branch', 'Source_File', 'Report_Type', 'Date', 'Month']:
                if col in export_df.columns:
                    export_df = export_df.drop(col, axis=1)
            export_df.to_excel(writer, index=False, sheet_name="Logimax")
        
        # Cashfree sheet second
        if not cashfree_df.empty:
            export_df = cashfree_df.copy()
            # Remove internal columns
            for col in ['Branch', 'Source_File', 'Report_Type', 'Date', 'Month']:
                if col in export_df.columns:
                    export_df = export_df.drop(col, axis=1)
            export_df.to_excel(writer, index=False, sheet_name="Cashfree")
        
        # Summary sheet
        summary_data = []
        if not logimax_df.empty:
            amount_col = None
            for col in logimax_df.columns:
                if 'payment_amount' in col.lower() or 'amount' in col.lower():
                    amount_col = col
                    break
            total = pd.to_numeric(logimax_df[amount_col], errors="coerce").sum() if amount_col else 0
            summary_data.append({
                "Report Type": "Logimax",
                "Records": len(logimax_df),
                "Total Amount": total
            })
        
        if not cashfree_df.empty:
            amount_col = None
            for col in cashfree_df.columns:
                if 'amount' in col.lower() or 'settlement' in col.lower():
                    amount_col = col
                    break
            total = pd.to_numeric(cashfree_df[amount_col], errors="coerce").sum() if amount_col else 0
            summary_data.append({
                "Report Type": "Cashfree",
                "Records": len(cashfree_df),
                "Total Amount": total
            })
        
        if summary_data:
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, index=False, sheet_name="Summary")
    
    return output.getvalue()
def generate_filename(branch, customer_search=None):
    """Generate filename based on filters"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    
    parts = []
    if branch and branch != "All":
        parts.append(branch)
    if customer_search:
        phone_number = ''.join(filter(str.isdigit, customer_search))
        if len(phone_number) >= 10:
            phone_number = phone_number[-10:]
            parts.append(phone_number)
        else:
            clean_name = ''.join(e for e in customer_search if e.isalnum())[:20]
            if clean_name:
                parts.append(clean_name)
    
    if parts:
        return f"Payment_Report_{'_'.join(parts)}_{timestamp}"
    else:
        return f"Payment_Report_{timestamp}"

LABEL = "Bhima Jewellery Customer Payment Logimax Cashfree"


def run():
    BASE_PATH = r"D:\EDP\EDP REPORT's\Logimax-Padam Payment Details\Documents"
    st.markdown("""
    <style>
        /* Professional color scheme */
        :root {
            --primary: #1a1a2e;
            --secondary: #16213e;
            --accent: #c9a84c;
            --accent-light: #e8d5a3;
            --accent-dark: #a8892e;
            --white: #ffffff;
            --light-gray: #f8f9fa;
            --border-radius: 12px;
            --box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        }

        /* Metric cards */
        .metric-card {
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            padding: 20px 25px;
            border-radius: var(--border-radius);
            border-left: 4px solid var(--accent);
            box-shadow: var(--box-shadow);
            transition: transform 0.2s ease;
            height: 100%;
        }
        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 30px rgba(0,0,0,0.12);
        }
        .metric-label {
            color: #a0a0b8;
            font-size: 13px;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }
        .metric-value {
            color: var(--white);
            font-size: 28px;
            font-weight: 700;
            margin-bottom: 4px;
        }
        .metric-sub {
            color: var(--accent-light);
            font-size: 14px;
            font-weight: 400;
        }

        /* Section headers */
        .section-header {
            background: linear-gradient(135deg, #1a1a2e, #16213e);
            padding: 15px 20px;
            border-radius: var(--border-radius);
            border-left: 4px solid var(--accent);
            margin: 25px 0 20px 0;
        }
        .section-header h3 {
            color: var(--white);
            margin: 0;
            font-size: 20px;
            font-weight: 600;
        }
        .section-header span {
            color: var(--accent);
            font-size: 20px;
            margin-right: 10px;
        }

        /* Info cards */
        .info-card {
            background: var(--white);
            padding: 20px;
            border-radius: var(--border-radius);
            border: 1px solid #e0e0e0;
            box-shadow: var(--box-shadow);
            margin-bottom: 20px;
        }
        .info-card h4 {
            color: var(--primary);
            margin: 0 0 10px 0;
            font-weight: 600;
        }

        /* Status indicators */
        .status-badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 500;
        }
        .status-success {
            background: #d4edda;
            color: #155724;
        }
        .status-warning {
            background: #fff3cd;
            color: #856404;
        }
        .status-info {
            background: #d1ecf1;
            color: #0c5460;
        }

        /* Dataframe styling */
        .stDataFrame {
            border-radius: var(--border-radius);
            overflow: hidden;
            box-shadow: var(--box-shadow);
        }

        /* Button styling */
        .stButton > button {
            background: linear-gradient(135deg, var(--accent), var(--accent-dark)) !important;
            color: var(--white) !important;
            border: none !important;
            border-radius: 8px !important;
            padding: 10px 25px !important;
            font-weight: 600 !important;
            transition: all 0.3s ease !important;
        }
        .stButton > button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 15px rgba(201, 168, 76, 0.3);
        }

        /* Sidebar styling */
        .css-1d391kg {
            background-color: #f8f9fa;
        }

        /* Tabs styling */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
            background-color: #f1f3f5;
            border-radius: 10px;
            padding: 5px;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 8px;
            padding: 8px 20px;
            font-weight: 500;
            color: #6c757d;
            transition: all 0.2s ease;
        }
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, var(--accent), var(--accent-dark)) !important;
            color: var(--white) !important;
        }

        /* Download buttons */
        .download-section {
            background: var(--white);
            padding: 25px;
            border-radius: var(--border-radius);
            border: 1px solid #e0e0e0;
            box-shadow: var(--box-shadow);
            margin: 20px 0;
        }

        /* Footer */
        .footer {
            background: linear-gradient(135deg, #1a1a2e, #16213e);
            color: var(--white);
            padding: 20px 30px;
            border-radius: var(--border-radius);
            margin-top: 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
        }
    </style>
    """, unsafe_allow_html=True)
    current_time = datetime.now().strftime('%d-%b-%Y %I:%M %p')
    st.markdown(f"""
    <div style="
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 30px 35px;
        border-radius: 15px;
        margin-bottom: 30px;
        border-left: 6px solid #c9a84c;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
    ">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
            <div>
                <div style="display: flex; align-items: center; gap: 20px;">
                    <div style="
                        background: linear-gradient(135deg, #c9a84c, #a8892e);
                        width: 50px;
                        height: 50px;
                        border-radius: 12px;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        font-size: 28px;
                    ">
                        💎
                    </div>
                    <div>
                        <h1 style="color: #ffffff; margin: 0; font-size: 32px; font-weight: 700; letter-spacing: 1px;">
                            Bhima Jewellery
                        </h1>
                        <h2 style="color: #c9a84c; margin: 4px 0 0 0; font-size: 18px; font-weight: 400; letter-spacing: 0.5px;">
                            Customer Payment Details Dashboard
                        </h2>
                    </div>
                </div>
            </div>
            <div style="display: flex; gap: 15px; align-items: center; flex-wrap: wrap;">
                <span style="
                    background: rgba(201, 168, 76, 0.2);
                    color: #c9a84c;
                    padding: 6px 18px;
                    border-radius: 25px;
                    font-size: 13px;
                    border: 1px solid rgba(201, 168, 76, 0.3);
                    font-weight: 500;
                ">
                    💳 Cashfree
                </span>
                <span style="
                    background: rgba(201, 168, 76, 0.2);
                    color: #c9a84c;
                    padding: 6px 18px;
                    border-radius: 25px;
                    font-size: 13px;
                    border: 1px solid rgba(201, 168, 76, 0.3);
                    font-weight: 500;
                ">
                    📦 Logimax
                </span>
                <span style="
                    background: rgba(255, 255, 255, 0.1);
                    color: #ffffff;
                    padding: 6px 18px;
                    border-radius: 25px;
                    font-size: 12px;
                    border: 1px solid rgba(255, 255, 255, 0.1);
                ">
                    ⚡ Live • {current_time}
                </span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    with st.sidebar:

        st.markdown("""
        <div style="
            background: linear-gradient(135deg, #1a1a2e, #0f3460);
            padding: 20px;
            border-radius: 12px;
            margin-bottom: 25px;
            border-left: 4px solid #c9a84c;
        ">
            <div style="color: #c9a84c; font-size: 13px; font-weight: 300; letter-spacing: 1px; text-transform: uppercase;">
                BHIMA JEWELLERY
            </div>
            <div style="color: white; font-size: 18px; font-weight: 600; margin-top: 5px;">
                🔍 Payment Filters
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Allow multiple report types
        report_type = st.selectbox(
            "📊 Report Type",
            ["All", "Cashfree", "Logimax"]
        )

        branch = st.selectbox(
            "🏢 Branch",
            ["All", "Madurai", "Salem"]
        )

        # Search for specific customer
        st.divider()
        st.markdown("### 👤 Customer Search")
        customer_search = st.text_input(
            "Search Customer",
            placeholder="Enter customer name or phone number...",
            key="customer_search"
        )

        # Add option to show debug info
        show_debug = st.checkbox("🔧 Show Debug Info", value=False)

        if st.button("🔄 Load Data", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        st.divider()

        # Show current path in sidebar
        st.write("📁 **Current Path:**")
        st.code(str(BASE_PATH), language="text")

        # Sidebar stats placeholder
        stats_placeholder = st.empty()
    if report_type == "All":
        report_types = ["Cashfree", "Logimax"]
    else:
        report_types = [report_type]
    cashfree_df, logimax_df, debug_info, column_info = load_data(report_types, branch)
    if show_debug:
        with st.sidebar.expander("🐛 Debug Information", expanded=True):
            st.write("### File System Debug")
            for line in debug_info:
                st.text(line)

            st.write("\n### Data Summary")
            st.write(f"Cashfree records: {len(cashfree_df):,}")
            st.write(f"Logimax records: {len(logimax_df):,}")
    with st.expander("📋 Excel File Column Structure", expanded=False):
        if column_info:
            for info in column_info:
                st.markdown(f"""
                <div style="
                    background: #f8f9fa;
                    padding: 15px;
                    border-radius: 8px;
                    margin-bottom: 10px;
                    border-left: 3px solid #c9a84c;
                ">
                    <strong>{info['Branch']}</strong> - {info['Report_Type']} - {info['File']} 
                    <span style="color: #6c757d; font-size: 13px;">({info['Rows']:,} rows, {len(info['Columns'])} columns)</span>
                </div>
                """, unsafe_allow_html=True)

                col_data = {
                    "S.No": range(1, len(info['Columns']) + 1),
                    "Column Name": info['Columns']
                }
                col_df = pd.DataFrame(col_data)
                st.dataframe(col_df, use_container_width=True, height=min(300, len(info['Columns']) * 35 + 40))
        else:
            st.info("No files loaded yet. Click 'Load Data' to see column structure.")
    if cashfree_df.empty and logimax_df.empty:
        st.warning("⚠️ No data found. Please check the file path or try different filters.")

        with st.expander("💡 Troubleshooting Suggestions", expanded=True):
            st.write("""
            **Common Issues:**
            1. **File path**: Make sure the path exists
            2. **File naming**: Files should contain 'Cashfree' or 'Logimax' in their names
            3. **File format**: Files should be .xlsx or .xls format
            4. **Permissions**: Ensure read permissions for the folder
            """)

            if Path(BASE_PATH).exists():
                st.write("**📁 Folder Structure:**")
                for branch_folder in ["Madurai", "Salem"]:
                    folder_path = Path(BASE_PATH) / branch_folder
                    if folder_path.exists():
                        st.write(f"✅ **{branch_folder}/** exists")
                        files = list(folder_path.glob("*.xlsx")) + list(folder_path.glob("*.xls"))
                        if files:
                            st.write(f"   Found {len(files)} Excel files")
                            for f in files:
                                st.write(f"   - {f.name}")
                        else:
                            st.write(f"   ⚠️ No Excel files found in {branch_folder}")
                    else:
                        st.write(f"❌ **{branch_folder}/** does not exist")

        st.stop()
    cashfree_df = process_cashfree_data(cashfree_df)
    logimax_df = process_logimax_data(logimax_df)
    st.markdown("""
    <div class="section-header">
        <h3><span>📅</span> Date Range Selection</h3>
    </div>
    """, unsafe_allow_html=True)
    all_dates = []
    if not cashfree_df.empty and "Date" in cashfree_df.columns:
        all_dates.extend(cashfree_df["Date"].dropna().tolist())
    if not logimax_df.empty and "Date" in logimax_df.columns:
        all_dates.extend(logimax_df["Date"].dropna().tolist())
    if all_dates:
        min_date = min(all_dates)
        max_date = max(all_dates)
    else:
        min_date = datetime.now().date()
        max_date = datetime.now().date()
    col1, col2 = st.columns(2)
    with col1:
        from_date = st.date_input("From Date", value=min_date)
    with col2:
        to_date = st.date_input("To Date", value=max_date)
    search_text = st.text_input(
        "🔎 Search Records",
        placeholder="Type to search in all columns..."
    )
    if from_date > to_date:
        st.error("⚠️ From Date cannot be later than To Date!")
        st.stop()
    filtered_cashfree = apply_filters(cashfree_df, from_date, to_date, search_text, customer_search)
    filtered_logimax = apply_filters(logimax_df, from_date, to_date, search_text, customer_search)
    if filtered_cashfree.empty and filtered_logimax.empty:
        st.warning(f"No data found for the selected criteria")
        st.stop()
    st.markdown("""
    <div class="section-header">
        <h3><span>📊</span> Dashboard Overview</h3>
    </div>
    """, unsafe_allow_html=True)
    total_records = len(filtered_cashfree) + len(filtered_logimax)
    if customer_search:
        st.success(f"🎯 Showing results for: **{customer_search}**")
    else:
        st.info(f"📌 Showing all data • Cashfree: {len(filtered_cashfree):,} records • Logimax: {len(filtered_logimax):,} records")
    col1, col2, col3, col4 = st.columns(4)
    total_amount = 0
    if not filtered_cashfree.empty:
        amount_col = None
        for col in filtered_cashfree.columns:
            if 'amount' in col.lower() or 'settlement' in col.lower():
                amount_col = col
                break
        if amount_col:
            total_amount += pd.to_numeric(filtered_cashfree[amount_col], errors="coerce").sum()
    if not filtered_logimax.empty:
        amount_col = None
        for col in filtered_logimax.columns:
            if 'payment_amount' in col.lower() or 'amount' in col.lower():
                amount_col = col
                break
        if amount_col:
            total_amount += pd.to_numeric(filtered_logimax[amount_col], errors="coerce").sum()
    unique_customers = set()
    if not filtered_cashfree.empty:
        for col in filtered_cashfree.columns:
            if 'customer' in col.lower() and 'name' in col.lower():
                unique_customers.update(filtered_cashfree[col].dropna().unique())
                break
    if not filtered_logimax.empty:
        for col in filtered_logimax.columns:
            if 'firstname' in col.lower() or ('customer' in col.lower() and 'name' in col.lower()):
                unique_customers.update(filtered_logimax[col].dropna().unique())
                break
    report_types_available = []
    if not filtered_cashfree.empty:
        report_types_available.append("Cashfree")
    if not filtered_logimax.empty:
        report_types_available.append("Logimax")
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">📋 Total Records</div>
            <div class="metric-value">{total_records:,}</div>
            <div class="metric-sub">All report types</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">💰 Total Amount</div>
            <div class="metric-value">₹{total_amount:,.2f}</div>
            <div class="metric-sub">Combined total</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">👤 Unique Customers</div>
            <div class="metric-value">{len(unique_customers):,}</div>
            <div class="metric-sub">Across all reports</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">📊 Report Types</div>
            <div class="metric-value">{len(report_types_available)}</div>
            <div class="metric-sub">{' • '.join(report_types_available)}</div>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("""
    <div class="section-header">
        <h3><span>📈</span> Report Type Distribution</h3>
    </div>
    """, unsafe_allow_html=True)
    dist_data = []
    if not filtered_cashfree.empty:
        amount_col = None
        for col in filtered_cashfree.columns:
            if 'amount' in col.lower() or 'settlement' in col.lower():
                amount_col = col
                break
        total = pd.to_numeric(filtered_cashfree[amount_col], errors="coerce").sum() if amount_col else 0
        dist_data.append({
            "Report Type": "Cashfree",
            "Records": len(filtered_cashfree),
            "Percentage": f"{(len(filtered_cashfree)/total_records)*100:.1f}%",
            "Total Amount": f"₹{total:,.2f}"
        })
    if not filtered_logimax.empty:
        amount_col = None
        for col in filtered_logimax.columns:
            if 'payment_amount' in col.lower() or 'amount' in col.lower():
                amount_col = col
                break
        total = pd.to_numeric(filtered_logimax[amount_col], errors="coerce").sum() if amount_col else 0
        dist_data.append({
            "Report Type": "Logimax",
            "Records": len(filtered_logimax),
            "Percentage": f"{(len(filtered_logimax)/total_records)*100:.1f}%",
            "Total Amount": f"₹{total:,.2f}"
        })
    if dist_data:
        dist_df = pd.DataFrame(dist_data)

        # Style the dataframe
        st.dataframe(
            dist_df,
            use_container_width=True,
            column_config={
                "Report Type": st.column_config.TextColumn("Report Type", width="medium"),
                "Records": st.column_config.NumberColumn("Records", width="small"),
                "Percentage": st.column_config.TextColumn("Percentage", width="small"),
                "Total Amount": st.column_config.TextColumn("Total Amount", width="medium"),
            }
        )
    st.markdown("""
    <div class="section-header">
        <h3><span>📥</span> Export Options</h3>
    </div>
    """, unsafe_allow_html=True)
    if customer_search:
        st.success(f"🎯 Exporting data for customer: **{customer_search}**")
    else:
        st.info("📌 Exporting all data")
    excel_data = create_excel_with_sheets(filtered_cashfree, filtered_logimax)
    csv_data_dict = {}
    if not filtered_cashfree.empty:
        export_df = filtered_cashfree.copy()
        for col in ['Branch', 'Source_File', 'Report_Type', 'Date', 'Month']:
            if col in export_df.columns:
                export_df = export_df.drop(col, axis=1)
        csv_data_dict["Cashfree"] = export_df.to_csv(index=False).encode('utf-8')
    if not filtered_logimax.empty:
        export_df = filtered_logimax.copy()
        for col in ['Branch', 'Source_File', 'Report_Type', 'Date', 'Month']:
            if col in export_df.columns:
                export_df = export_df.drop(col, axis=1)
        csv_data_dict["Logimax"] = export_df.to_csv(index=False).encode('utf-8')
    base_filename = generate_filename(branch, customer_search)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button(
            "📊 Excel (Multi-Sheet)",
            excel_data,
            file_name=f"{base_filename}.xlsx",
            use_container_width=True,
            help="Excel with Logimax first, then Cashfree sheets + Summary"
        )
        st.caption("📌 Sheets: Logimax → Cashfree → Summary")
    with col2:
        # Allow downloading individual CSV files
        if len(csv_data_dict) > 1:
            selected_report = st.selectbox(
                "Select CSV to download",
                list(csv_data_dict.keys())
            )
            st.download_button(
                f"📋 CSV - {selected_report}",
                csv_data_dict[selected_report],
                file_name=f"{base_filename}_{selected_report}.csv",
                mime="text/csv",
                use_container_width=True
            )
        elif len(csv_data_dict) == 1:
            st.download_button(
                "📋 CSV",
                list(csv_data_dict.values())[0],
                file_name=f"{base_filename}.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.warning("No data to export as CSV")
    with col3:
        # Combine both dataframes for JSON export
        combined_df = pd.DataFrame()
        if not filtered_cashfree.empty:
            cf = filtered_cashfree.copy()
            cf["Report_Type"] = "Cashfree"
            combined_df = pd.concat([combined_df, cf], ignore_index=True)
        if not filtered_logimax.empty:
            lg = filtered_logimax.copy()
            lg["Report_Type"] = "Logimax"
            combined_df = pd.concat([combined_df, lg], ignore_index=True)

        if not combined_df.empty:
            for col in ['Branch', 'Source_File', 'Date', 'Month']:
                if col in combined_df.columns:
                    combined_df = combined_df.drop(col, axis=1)
            json_data = combined_df.to_json(orient='records', date_format='iso').encode('utf-8')
            st.download_button(
                "🔄 JSON",
                json_data,
                file_name=f"{base_filename}.json",
                mime="application/json",
                use_container_width=True
            )
    st.caption("📌 Note: Exported files do not include internal columns (Branch, Source_File, Date, Month)")
    st.markdown("""
    <div class="section-header">
        <h3><span>📋</span> Data View</h3>
    </div>
    """, unsafe_allow_html=True)
    tabs = []
    if not filtered_logimax.empty:
        tabs.append("📦 Logimax")
    if not filtered_cashfree.empty:
        tabs.append("💳 Cashfree")
    if len(tabs) > 1:
        tabs.insert(0, "📊 All Data")
    if len(tabs) > 1:
        tab_objects = st.tabs(tabs)

        for i, tab_name in enumerate(tabs):
            with tab_objects[i]:
                if tab_name == "📊 All Data":
                    # Combine both dataframes for display
                    display_df = pd.DataFrame()
                    if not filtered_logimax.empty:
                        lg = filtered_logimax.copy()
                        lg["Report_Type"] = "Logimax"
                        display_df = pd.concat([display_df, lg], ignore_index=True)
                    if not filtered_cashfree.empty:
                        cf = filtered_cashfree.copy()
                        cf["Report_Type"] = "Cashfree"
                        display_df = pd.concat([display_df, cf], ignore_index=True)
                    if not display_df.empty:
                        display_df.insert(0, "S.No", range(1, len(display_df)+1))
                        st.dataframe(display_df, use_container_width=True, height=400)
                        st.caption(f"📌 Total: {len(display_df)} records")
                elif tab_name == "📦 Logimax":
                    display_df = filtered_logimax.copy()
                    display_df.insert(0, "S.No", range(1, len(display_df)+1))
                    st.dataframe(display_df, use_container_width=True, height=400)
                    st.caption(f"📌 {len(display_df)} records for Logimax")
                elif tab_name == "💳 Cashfree":
                    display_df = filtered_cashfree.copy()
                    display_df.insert(0, "S.No", range(1, len(display_df)+1))
                    st.dataframe(display_df, use_container_width=True, height=400)
                    st.caption(f"📌 {len(display_df)} records for Cashfree")
    else:
        # Only one report type
        if not filtered_logimax.empty:
            display_df = filtered_logimax.copy()
            display_df.insert(0, "S.No", range(1, len(display_df)+1))
            st.dataframe(display_df, use_container_width=True, height=400)
        elif not filtered_cashfree.empty:
            display_df = filtered_cashfree.copy()
            display_df.insert(0, "S.No", range(1, len(display_df)+1))
            st.dataframe(display_df, use_container_width=True, height=400)
    st.markdown("""
    <div class="section-header">
        <h3><span>📞</span> Customer Directory</h3>
    </div>
    """, unsafe_allow_html=True)
    customer_data = []
    if not filtered_logimax.empty:
        name_col = None
        phone_col = None
        for col in filtered_logimax.columns:
            if 'firstname' in col.lower():
                name_col = col
            if 'mobile' in col.lower() or 'phone' in col.lower():
                phone_col = col
        if name_col:
            for _, row in filtered_logimax.iterrows():
                customer_data.append({
                    "Customer Name": row[name_col] if pd.notna(row[name_col]) else "",
                    "Phone": row[phone_col] if phone_col and pd.notna(row[phone_col]) else "",
                    "Report Type": "Logimax"
                })
    if not filtered_cashfree.empty:
        name_col = None
        phone_col = None
        for col in filtered_cashfree.columns:
            if 'customer name' in col.lower():
                name_col = col
            if 'customer phone' in col.lower() or 'phone' in col.lower():
                phone_col = col
        if name_col:
            for _, row in filtered_cashfree.iterrows():
                customer_data.append({
                    "Customer Name": row[name_col] if pd.notna(row[name_col]) else "",
                    "Phone": row[phone_col] if phone_col and pd.notna(row[phone_col]) else "",
                    "Report Type": "Cashfree"
                })
    if customer_data:
        customers_df = pd.DataFrame(customer_data)
        customers_df = customers_df.drop_duplicates(subset=["Customer Name", "Phone"])
        st.dataframe(customers_df, use_container_width=True, height=300)
        st.caption(f"📌 Total unique customers: {len(customers_df)}")
    with stats_placeholder.container():
        st.markdown("---")
        st.markdown("### 📊 Data Summary")
        st.write(f"**Period:** {from_date.strftime('%d-%b')} - {to_date.strftime('%d-%b-%Y')}")
        st.write(f"**Branch:** {branch}")
        st.write(f"**Report Type:** {report_type}")
        st.write(f"**Total Records:** {total_records:,}")

        if not filtered_cashfree.empty:
            st.write(f"**Cashfree:** {len(filtered_cashfree):,} records")
        if not filtered_logimax.empty:
            st.write(f"**Logimax:** {len(filtered_logimax):,} records")

        if customer_search:
            st.write(f"**Customer Search:** {customer_search}")
    footer_time = datetime.now().strftime('%d-%b-%Y %I:%M %p')
    st.markdown(f"""
    <div class="footer">
        <div>
            <span style="color: #c9a84c; font-weight: 600;">💎 Bhima Jewellery</span>
            <span style="color: #888; margin-left: 15px;">|</span>
            <span style="color: #aaa; font-size: 13px;">Customer Payment Details Dashboard</span>
            <span style="color: #888; margin-left: 15px;">|</span>
            <span style="color: #aaa; font-size: 13px;">Cashfree & Logimax</span>
        </div>
        <div style="color: #888; font-size: 13px;">
            <span>📊 {total_records:,} records</span>
            <span style="margin-left: 20px;">🕒 {footer_time}</span>
            <span style="margin-left: 20px; color: #28a745;">● Online</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
