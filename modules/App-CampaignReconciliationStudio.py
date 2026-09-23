import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime, time, timedelta
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# -----------------------------
# PAGE CONFIG - ADVANCED
# -----------------------------
st.set_page_config(
    page_title="Campaign Reconciliation Studio",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for premium UI
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    /* Main background with premium gradient */
    .stApp {
        background: linear-gradient(135deg, #f8f9fc 0%, #e9ecef 100%);
    }
    
    /* Main container padding */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }
    
    /* Premium glass morphism cards */
    .glass-card {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(10px);
        border-radius: 24px;
        padding: 24px;
        border: 1px solid rgba(255, 255, 255, 0.5);
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.08);
        transition: all 0.3s ease;
    }
    
    .glass-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 48px rgba(0, 0, 0, 0.12);
    }
    
    /* Animated gradient text */
    .gradient-text {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 50%, #667eea 100%);
        background-size: 200% auto;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 3rem;
        font-weight: 800;
        animation: shine 3s linear infinite;
    }
    
    @keyframes shine {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    
    /* Premium metric cards */
    .metric-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        border-radius: 20px;
        padding: 20px;
        color: white;
        text-align: center;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
        position: relative;
        overflow: hidden;
    }
    
    .metric-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
        transition: left 0.5s;
    }
    
    .metric-card:hover::before {
        left: 100%;
    }
    
    .metric-card:hover {
        transform: translateY(-5px) scale(1.02);
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.15);
    }
    
    .metric-card h3 {
        font-size: 0.9rem;
        font-weight: 500;
        margin-bottom: 10px;
        opacity: 0.9;
    }
    
    .metric-card h2 {
        font-size: 2rem;
        font-weight: 700;
        margin: 10px 0;
    }
    
    .metric-card small {
        font-size: 0.8rem;
        opacity: 0.8;
    }
    
    /* Progress bar styling */
    .stProgress > div > div {
        background: linear-gradient(90deg, #1e3c72, #2a5298, #667eea);
        border-radius: 10px;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 12px 32px;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(30, 60, 114, 0.3);
    }
    
    .stButton > button:active {
        transform: translateY(0);
    }
    
    /* Dataframe styling */
    .dataframe {
        border-radius: 16px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border: none;
    }
    
    .dataframe thead th {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        font-weight: 600;
        padding: 12px;
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: rgba(255,255,255,0.7);
        backdrop-filter: blur(10px);
        border-radius: 16px;
        padding: 6px;
        margin-bottom: 20px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 12px;
        padding: 10px 24px;
        transition: all 0.3s ease;
        font-weight: 500;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background-color: rgba(30, 60, 114, 0.1);
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(135deg, rgba(30, 60, 114, 0.98) 0%, rgba(42, 82, 152, 0.98) 100%);
        backdrop-filter: blur(10px);
        border-right: none;
    }
    
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stCheckbox label {
        color: rgba(255,255,255,0.9) !important;
    }
    
    [data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] {
        background-color: rgba(255,255,255,0.1);
        border-radius: 10px;
    }
    
    /* Expander styling */
    .streamlit-expanderHeader {
        background: linear-gradient(135deg, #f8f9fc 0%, #e9ecef 100%);
        border-radius: 12px;
        font-weight: 600;
    }
    
    /* Info/Warning/Success boxes */
    .stAlert {
        border-radius: 12px;
        border-left: 4px solid;
    }
    
    /* File uploader styling */
    [data-testid="stFileUploader"] {
        background-color: rgba(255,255,255,0.9);
        border-radius: 12px;
        padding: 10px;
    }
    
    /* Number input styling */
    .stNumberInput input {
        border-radius: 10px;
    }
    
    /* Slider styling */
    .stSlider div[data-baseweb="slider"] {
        margin-top: 10px;
    }
    
    /* Custom scrollbar */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: #f1f1f1;
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #1e3c72, #2a5298);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: #1e3c72;
    }
    
    /* Animations */
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    .fade-in-up {
        animation: fadeInUp 0.6s ease-out;
    }
    
    /* Badge styling */
    .badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
        background: linear-gradient(135deg, #1e3c72, #2a5298);
        color: white;
    }
    
    /* Feature card styling */
    .feature-card {
        background: white;
        border-radius: 20px;
        padding: 30px 20px;
        text-align: center;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    
    .feature-card:hover {
        transform: translateY(-8px);
        box-shadow: 0 12px 24px rgba(0,0,0,0.1);
    }
    
    .feature-icon {
        font-size: 3rem;
        margin-bottom: 15px;
    }
    
    /* Divider styling */
    hr {
        margin: 2rem 0;
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, #ccc, transparent);
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------
# HELPER FUNCTIONS (Enhanced)
# -----------------------------

@st.cache_data
def read_file(file):
    """Read CSV or Excel file with caching"""
    if file is None:
        return None
    
    try:
        if file.name.lower().endswith(".csv"):
            return pd.read_csv(file, encoding='utf-8')
        else:
            return pd.read_excel(file)
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
        return None

def clean_key_column(df, col):
    """Clean key columns for matching"""
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    return df

def normalize_phone(series):
    """
    Extract the rightmost 10 digits from phone numbers.
    Always takes the last 10 digits (right to left) after removing all non-digit characters.
    Examples:
        "919876543210" -> "9876543210"
        "+1-234-567-8901" -> "2345678901" (last 10 digits)
        "9876543210" -> "9876543210"
        "12345" -> "12345" (or empty if less than 10 digits? We'll keep original but validation will catch)
    """
    # First, remove all non-digit characters
    cleaned = series.astype(str).str.replace(r"\D", "", regex=True)
    
    # Then extract the rightmost 10 characters (last 10 digits)
    # If length is less than 10, keep as is (will be filtered out later by validation)
    normalized = cleaned.str[-10:]
    
    return normalized

def parse_datetime_auto(series):
    """Auto-detect and parse datetime with multiple format attempts"""
    try:
        parsed = pd.to_datetime(series, errors='coerce')
        if parsed.notna().sum() > len(series) * 0.5:
            return parsed
    except:
        pass
    
    formats = [
        '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%d-%m-%Y %H:%M:%S',
        '%d/%m/%Y %H:%M:%S', '%Y-%m-%d %I:%M:%S %p', '%d-%m-%Y %I:%M:%S %p',
        '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%Y',
        '%m/%d/%Y %H:%M:%S', '%m-%d-%Y %H:%M:%S', '%b %d, %Y %I:%M:%S %p',
        '%B %d, %Y %I:%M:%S %p'
    ]
    
    for fmt in formats:
        try:
            parsed = pd.to_datetime(series, format=fmt, errors='coerce')
            if parsed.notna().sum() > len(series) * 0.5:
                return parsed
        except:
            continue
    
    return pd.to_datetime(series, errors='coerce')

def get_date_range_from_column(df, datetime_col):
    """Auto-detect min and max dates from column"""
    if datetime_col not in df.columns:
        return None, None
    
    parsed_dates = parse_datetime_auto(df[datetime_col])
    valid_dates = parsed_dates.dropna()
    
    if len(valid_dates) == 0:
        return None, None
    
    return valid_dates.min(), valid_dates.max()

def filter_by_datetime_auto(df, datetime_col, from_datetime, to_datetime):
    """Filter dataframe by datetime range with auto detection"""
    if datetime_col not in df.columns:
        return df
    
    original_count = len(df)
    df[datetime_col] = parse_datetime_auto(df[datetime_col])
    df = df.dropna(subset=[datetime_col])
    
    mask = (df[datetime_col] >= pd.Timestamp(from_datetime)) & \
           (df[datetime_col] <= pd.Timestamp(to_datetime))
    
    filtered_df = df[mask]
    
    st.info(f"📅 Date filter applied: {len(filtered_df):,} records remain (filtered from {original_count:,})")
    
    return filtered_df

def create_visualizations(stats):
    """Create advanced visualizations for the dashboard"""
    if stats is None or not isinstance(stats, dict):
        fig_funnel = go.Figure()
        fig_funnel.update_layout(title="No data available")
        fig_gauge = go.Figure()
        fig_gauge.update_layout(title="No data available")
        fig_donut = go.Figure()
        fig_donut.update_layout(title="No data available")
        fig_time = None
        return fig_funnel, fig_gauge, fig_donut, fig_time
    
    # 1. Match Funnel Chart
    funnel_data = {
        'Stage': ['External Records', 'After Date Filter', 'Order ID Match', 'Phone Match', 'Final Success'],
        'Count': [
            stats.get('External_Total', 0),
            stats.get('External_Filtered', 0),
            stats.get('Step1_Matches', 0),
            stats.get('Step2_Matches', 0),
            stats.get('Final_Matches', 0)
        ]
    }
    funnel_df = pd.DataFrame(funnel_data)
    
    fig_funnel = px.funnel(
        funnel_df, 
        x='Count', 
        y='Stage',
        title='<b>📊 Reconciliation Funnel Analysis</b>',
        color_discrete_sequence=['#1e3c72', '#2a5298', '#4169e1', '#667eea', '#8b9dc3']
    )
    fig_funnel.update_layout(
        height=450,
        showlegend=False,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Inter, sans-serif", size=12),
        title_font_size=16
    )
    
    # 2. Match Rate Gauge
    match_rate = (stats.get('Step2_Matches', 0) / stats.get('External_Filtered', 1)) * 100 if stats.get('External_Filtered', 0) > 0 else 0
    
    fig_gauge = go.Figure(go.Indicator(
        mode = "gauge+number+delta",
        value = match_rate,
        title = {'text': "<b>Overall Match Rate</b>", 'font': {'size': 20, 'family': "Inter"}},
        delta = {'reference': 50, 'increasing': {'color': "green"}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': "darkblue"},
            'bar': {'color': "#1e3c72"},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "#e0e0e0",
            'steps': [
                {'range': [0, 30], 'color': '#ff6b6b'},
                {'range': [30, 70], 'color': '#ffd93d'},
                {'range': [70, 100], 'color': '#6bcb77'}
            ],
            'threshold': {
                'line': {'color': "red", 'width': 4},
                'thickness': 0.75,
                'value': 90
            }
        }
    ))
    fig_gauge.update_layout(height=350, margin=dict(t=50, b=20, l=20, r=20))
    
    # 3. Donut Chart
    matched_count = stats.get('Step2_Matches', 0)
    unmatched_count = stats.get('External_Filtered', 0) - matched_count
    
    fig_donut = go.Figure(data=[go.Pie(
        labels=['Matched Successfully', 'Unmatched'],
        values=[matched_count, max(0, unmatched_count)],
        hole=.6,
        marker_colors=['#2a5298', '#ff6b6b'],
        textinfo='label+percent',
        textposition='auto',
        textfont=dict(size=12, family="Inter")
    )])
    fig_donut.update_layout(
        title="<b>Match Success Distribution</b>",
        height=350,
        annotations=[dict(text=f"{match_rate:.1f}%", x=0.5, y=0.5, font_size=24, font_family="Inter", showarrow=False)]
    )
    
    # 4. Time Series Analysis
    time_series_data = stats.get('TimeSeriesData')
    if time_series_data is not None and not time_series_data.empty:
        fig_time = px.line(
            time_series_data,
            x='Date',
            y='Count',
            title='<b>📈 Match Trends Over Time</b>',
            color_discrete_sequence=['#1e3c72']
        )
        fig_time.update_layout(
            height=350,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Inter, sans-serif", size=12)
        )
        fig_time.update_traces(line=dict(width=3), mode='lines+markers', marker=dict(size=6))
    else:
        fig_time = None
    
    return fig_funnel, fig_gauge, fig_donut, fig_time

def to_excel_advanced(matched_internal=None, final_matched=None, unmatched_external=None, 
                      stats=None, campaign_data=None, phone_analysis=None):
    """Export results to Excel with multiple sheets and formatting"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if final_matched is not None and not final_matched.empty:
            final_matched.to_excel(writer, index=False, sheet_name="Final_Matched")
            
        if matched_internal is not None and not matched_internal.empty:
            matched_internal.to_excel(writer, index=False, sheet_name="Internal_Matched")
        
        if unmatched_external is not None and not unmatched_external.empty:
            unmatched_external.to_excel(writer, index=False, sheet_name="Unmatched_External")
        
        if stats is not None:
            stats_df = pd.DataFrame(list(stats.items()) if isinstance(stats, dict) else [], columns=["Metric", "Value"])
            if not stats_df.empty:
                stats_df.to_excel(writer, index=False, sheet_name="Statistics")
    
    output.seek(0)
    return output

# -----------------------------
# ADVANCED UI COMPONENTS
# -----------------------------

class CampaignReconciliationUI:
    """Advanced UI handler for the reconciliation tool"""
    
    def __init__(self):
        self.theme_colors = {
            'primary': '#1e3c72',
            'secondary': '#2a5298',
            'success': '#6bcb77',
            'warning': '#ffd93d',
            'danger': '#ff6b6b',
            'info': '#4facfe'
        }
        
    def render_header(self):
        """Render animated header"""
        st.markdown(f"""
        <div style='text-align: center; padding: 2rem 0 1rem 0;' class='fade-in-up'>
            <h1 class='gradient-text'>🎯 Campaign Reconciliation Studio</h1>
            <p style='font-size: 1.2rem; color: #666; margin-top: 0.5rem;'>
                Enterprise-Grade 3-Step Matching Pipeline
            </p>
            <div style='margin-top: 1rem;'>
                <span class='badge'>⚡ Real-time Processing</span>
                <span class='badge'>🔒 Privacy First</span>
                <span class='badge'>📊 Advanced Analytics</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    def render_file_upload_section(self):
        """Render modern file upload section"""
        st.markdown("---")
        st.markdown("### 📁 File Upload Center")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            with st.container():
                st.markdown("""
                <div class='glass-card' style='text-align: center; min-height: 200px;'>
                    <div style='font-size: 3rem; margin-bottom: 10px;'>📊</div>
                    <h3 style='margin: 0;'>External Data</h3>
                    <p style='color: #666; font-size: 0.9rem;'>Order/Transaction Data</p>
                    <hr style='margin: 15px 0;'>
                    <small style='color: #999;'>CSV, XLSX formats supported</small>
                </div>
                """, unsafe_allow_html=True)
                external_file = st.file_uploader(
                    "Upload External File",
                    type=['csv', 'xlsx', 'xls'],
                    key="external_adv",
                    label_visibility="collapsed"
                )
                
        with col2:
            with st.container():
                st.markdown("""
                <div class='glass-card' style='text-align: center; min-height: 200px;'>
                    <div style='font-size: 3rem; margin-bottom: 10px;'>🏢</div>
                    <h3 style='margin: 0;'>Internal Data</h3>
                    <p style='color: #666; font-size: 0.9rem;'>Transaction Records</p>
                    <hr style='margin: 15px 0;'>
                    <small style='color: #999;'>CSV, XLSX formats supported</small>
                </div>
                """, unsafe_allow_html=True)
                internal_file = st.file_uploader(
                    "Upload Internal File",
                    type=['csv', 'xlsx', 'xls'],
                    key="internal_adv",
                    label_visibility="collapsed"
                )
                
        with col3:
            with st.container():
                st.markdown("""
                <div class='glass-card' style='text-align: center; min-height: 200px;'>
                    <div style='font-size: 3rem; margin-bottom: 10px;'>📞</div>
                    <h3 style='margin: 0;'>Campaign Data</h3>
                    <p style='color: #666; font-size: 0.9rem;'>Phone List</p>
                    <hr style='margin: 15px 0;'>
                    <small style='color: #999;'>CSV, XLSX formats supported</small>
                </div>
                """, unsafe_allow_html=True)
                campaign_file = st.file_uploader(
                    "Upload Campaign File",
                    type=['csv', 'xlsx', 'xls'],
                    key="campaign_adv",
                    label_visibility="collapsed"
                )
        
        return external_file, internal_file, campaign_file

# -----------------------------
# MAIN APPLICATION
# -----------------------------

# Initialize session state
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False
if 'results' not in st.session_state:
    st.session_state.results = {}
if 'final_matched' not in st.session_state:
    st.session_state.final_matched = None
if 'internal_matched' not in st.session_state:
    st.session_state.internal_matched = None
if 'unmatched_external' not in st.session_state:
    st.session_state.unmatched_external = None
if 'stats_df' not in st.session_state:
    st.session_state.stats_df = None
if 'stats_dashboard' not in st.session_state:
    st.session_state.stats_dashboard = None
if 'filtered_ext_df' not in st.session_state:
    st.session_state.filtered_ext_df = None
if 'match_rate_step1' not in st.session_state:
    st.session_state.match_rate_step1 = 0
if 'match_rate_step2' not in st.session_state:
    st.session_state.match_rate_step2 = 0
if 'fig_funnel' not in st.session_state:
    st.session_state.fig_funnel = None
if 'fig_gauge' not in st.session_state:
    st.session_state.fig_gauge = None
if 'fig_donut' not in st.session_state:
    st.session_state.fig_donut = None
if 'fig_time' not in st.session_state:
    st.session_state.fig_time = None

ui = CampaignReconciliationUI()
ui.render_header()

# Sidebar with advanced controls
with st.sidebar:
    st.markdown("## ⚙️ Control Panel")
    st.markdown("---")
    
    # File upload in sidebar also
    st.markdown("### 📂 Data Source")
    external_file = st.file_uploader(
        "1️⃣ External Data", type=['csv', 'xlsx', 'xls'], key="sidebar_ext"
    )
    internal_file = st.file_uploader(
        "2️⃣ Internal Data", type=['csv', 'xlsx', 'xls'], key="sidebar_int"
    )
    campaign_file = st.file_uploader(
        "3️⃣ Campaign Data", type=['csv', 'xlsx', 'xls'], key="sidebar_camp"
    )
    
    st.markdown("---")
    
    # Advanced settings expander
    with st.expander("🔧 Advanced Settings", expanded=False):
        st.markdown("#### Matching Options")
        case_sensitive = st.checkbox("Case Sensitive Matching", value=False)
        fuzzy_matching = st.checkbox("Enable Fuzzy Matching", value=False)
        if fuzzy_matching:
            fuzzy_threshold = st.slider("Fuzzy Match Threshold", 0.0, 1.0, 0.8, 0.05)
        
        st.markdown("#### Export Settings")
        export_format = st.selectbox("Export Format", ["Excel", "CSV", "Both"])
        include_debug_info = st.checkbox("Include Debug Information", value=False)
        
    st.markdown("---")
    
    # Quick stats panel (shows after processing)
    if st.session_state.processing_complete:
        st.markdown("### 📊 Quick Stats")
        st.metric("Match Rate", f"{st.session_state.results.get('match_rate', 0):.1f}%")
        st.metric("Total Matches", f"{st.session_state.results.get('total_matches', 0):,}")
        st.metric("Unmatched", f"{st.session_state.results.get('unmatched', 0):,}")
        
        # Add a reset button
        if st.button("🔄 Reset All Data", width="stretch"):
            for key in ['processing_complete', 'final_matched', 'internal_matched', 
                       'unmatched_external', 'stats_df', 'stats_dashboard', 
                       'filtered_ext_df', 'fig_funnel', 'fig_gauge', 'fig_donut', 'fig_time']:
                if key in st.session_state:
                    if key == 'processing_complete':
                        st.session_state[key] = False
                    else:
                        st.session_state[key] = None
            st.rerun()

# Main content area
if not external_file or not internal_file or not campaign_file:
    st.info("👋 Welcome! Please upload all three files in the sidebar to begin the reconciliation process.")
    
    # Show feature highlights
    st.markdown("---")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown("""
        <div class='feature-card fade-in-up'>
            <div class='feature-icon'>🔗</div>
            <h4>3-Step Matching</h4>
            <small style='color: #666;'>External → Internal → Campaign</small>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class='feature-card fade-in-up'>
            <div class='feature-icon'>📅</div>
            <h4>Smart Date Filter</h4>
            <small style='color: #666;'>Auto-detects range</small>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class='feature-card fade-in-up'>
            <div class='feature-icon'>📱</div>
            <h4>Phone Normalization</h4>
            <small style='color: #666;'>Rightmost 10 digits extraction</small>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown("""
        <div class='feature-card fade-in-up'>
            <div class='feature-icon'>📊</div>
            <h4>Interactive Reports</h4>
            <small style='color: #666;'>Visual analytics</small>
        </div>
        """, unsafe_allow_html=True)
    st.stop()

# Load and validate files
try:
    with st.spinner("📂 Loading files..."):
        ext_df = read_file(external_file)
        int_df = read_file(internal_file)
        camp_df = read_file(campaign_file)
        
        if ext_df is None or int_df is None or camp_df is None:
            st.error("❌ Failed to load one or more files")
            st.stop()
            
        if ext_df.empty or int_df.empty or camp_df.empty:
            st.error("❌ One or more files are empty")
            st.stop()
        
        # Display file info in a nice grid
        st.markdown("---")
        st.markdown("### 📋 Data Load Summary")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"""
            <div class='glass-card' style='text-align: center;'>
                <h3>📊 External Data</h3>
                <h2 style='color: #1e3c72;'>{len(ext_df):,}</h2>
                <small>records • {len(ext_df.columns)} columns</small>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class='glass-card' style='text-align: center;'>
                <h3>🏢 Internal Data</h3>
                <h2 style='color: #1e3c72;'>{len(int_df):,}</h2>
                <small>records • {len(int_df.columns)} columns</small>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class='glass-card' style='text-align: center;'>
                <h3>📞 Campaign Data</h3>
                <h2 style='color: #1e3c72;'>{len(camp_df):,}</h2>
                <small>records • {len(camp_df.columns)} columns</small>
            </div>
            """, unsafe_allow_html=True)
            
except Exception as e:
    st.error(f"Error loading files: {str(e)}")
    st.stop()

# Column Mapping Section
st.markdown("---")
st.markdown("### 🔧 Column Configuration")

# Initialize variables with defaults
order_id_col = None
ref_col = None
phone_col_int = None
phone_col_camp = None
use_date_filter = False
time_col = None
from_datetime = None
to_datetime = None

# Create expandable sections for each file
with st.expander("📊 External File Configuration", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        order_id_col = st.selectbox("Order ID Column", options=ext_df.columns, key="ext_order")
    with col2:
        use_date_filter = st.checkbox("Enable Date/Time Filter", value=True, key="ext_date_filter")
        
    if use_date_filter:
        time_col = st.selectbox("Transaction Time Column", options=ext_df.columns, key="ext_time")
        
        # Date range selector with better UI
        min_date, max_date = get_date_range_from_column(ext_df, time_col)
        if min_date and max_date:
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                from_date = st.date_input("From Date", min_date, key="from_date")
                from_time = st.time_input("From Time", time(0, 0), key="from_time")
            with col2:
                to_date = st.date_input("To Date", max_date, key="to_date")
                to_time = st.time_input("To Time", time(23, 59), key="to_time")
            with col3:
                st.markdown("#### Quick Select")
                if st.button("Last 7 Days", width="stretch"):
                    from_date = (datetime.now() - timedelta(days=7)).date()
                    to_date = datetime.now().date()
                    st.rerun()
                if st.button("This Month", width="stretch"):
                    from_date = datetime.now().replace(day=1).date()
                    to_date = datetime.now().date()
                    st.rerun()
            
            from_datetime = datetime.combine(from_date, from_time)
            to_datetime = datetime.combine(to_date, to_time)

with st.expander("🏢 Internal File Configuration", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        ref_col = st.selectbox("Transaction Reference Column", options=int_df.columns, key="int_ref")
    with col2:
        phone_col_int = st.selectbox("Phone Number Column", options=int_df.columns, key="int_phone")

with st.expander("📞 Campaign File Configuration", expanded=True):
    phone_col_camp = st.selectbox("Phone Number Column", options=camp_df.columns, key="camp_phone")

# Process button with animation
st.markdown("---")
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    process_button = st.button(
        "🚀 START RECONCILIATION", 
        type="primary", 
        width="stretch"
    )

# Processing logic
if process_button:
    # Validate required columns
    if not all([order_id_col, ref_col, phone_col_int, phone_col_camp]):
        st.error("❌ Please configure all required columns before processing")
        st.stop()
    
    # Progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Date Filter
        filtered_ext_df = ext_df.copy()
        if use_date_filter and time_col and from_datetime and to_datetime:
            status_text.markdown("🔄 **Step 1/3:** Applying date filter...")
            progress_bar.progress(10)
            
            filtered_ext_df = filter_by_datetime_auto(
                filtered_ext_df, time_col, from_datetime, to_datetime
            )
            
            if filtered_ext_df.empty:
                st.warning("⚠️ No records after date filter. Please adjust date range or disable filter.")
                st.stop()
        else:
            status_text.markdown("🔄 **Step 1/3:** Skipping date filter...")
            progress_bar.progress(10)
        
        # Step 2: External → Internal Matching
        status_text.markdown("🔄 **Step 2/3:** Matching External to Internal...")
        progress_bar.progress(30)
        
        # Clean and match
        filtered_ext_df = clean_key_column(filtered_ext_df, order_id_col)
        int_df_clean = int_df.copy()
        int_df_clean = clean_key_column(int_df_clean, ref_col)
        
        internal_matched = pd.merge(
            int_df_clean,
            filtered_ext_df[[order_id_col]],
            left_on=ref_col,
            right_on=order_id_col,
            how="inner"
        ).drop_duplicates()
        
        matched_order_ids = set(internal_matched[ref_col].unique())
        unmatched_external = filtered_ext_df[~filtered_ext_df[order_id_col].isin(matched_order_ids)]
        
        match_rate_step1 = (len(internal_matched) / len(filtered_ext_df)) * 100 if len(filtered_ext_df) > 0 else 0
        
        st.success(f"✅ Step 1 Complete: {len(internal_matched):,} matches ({match_rate_step1:.1f}%)")
        progress_bar.progress(60)
        
        if internal_matched.empty:
            st.error("❌ No matches found between External and Internal files!")
            st.stop()
        
        # Step 3: Internal → Campaign Matching
        status_text.markdown("🔄 **Step 3/3:** Matching Internal to Campaign...")
        progress_bar.progress(70)
        
        # Normalize phone numbers - NOW USING RIGHTMOST 10 DIGITS
        internal_matched[phone_col_int] = normalize_phone(internal_matched[phone_col_int])
        camp_df_clean = camp_df.copy()
        camp_df_clean[phone_col_camp] = normalize_phone(camp_df_clean[phone_col_camp])
        
        # Remove invalid phones (less than 10 digits after normalization)
        valid_phone_int = internal_matched[phone_col_int].str.len() == 10
        valid_phone_camp = camp_df_clean[phone_col_camp].str.len() == 10
        
        internal_matched_clean = internal_matched[valid_phone_int].copy()
        camp_df_clean = camp_df_clean[valid_phone_camp].copy()
        
        # Display phone normalization statistics
        st.info(f"📱 Phone normalization: {valid_phone_int.sum():,} valid phones in Internal ({valid_phone_int.sum()/len(internal_matched)*100:.1f}%) | {valid_phone_camp.sum():,} valid phones in Campaign ({valid_phone_camp.sum()/len(camp_df_clean)*100:.1f}%)")
        
        # Final matching
        final_matched = pd.merge(
            internal_matched_clean,
            camp_df_clean[[phone_col_camp]],
            left_on=phone_col_int,
            right_on=phone_col_camp,
            how="inner"
        ).drop_duplicates()
        
        match_rate_step2 = (len(final_matched) / len(internal_matched_clean)) * 100 if len(internal_matched_clean) > 0 else 0
        
        st.success(f"✅ Step 2 Complete: {len(final_matched):,} matches ({match_rate_step2:.1f}%)")
        progress_bar.progress(100)
        status_text.markdown("✅ **Reconciliation Complete!**")
        
        # Create time series data for visualization
        time_series_data = pd.DataFrame()
        if use_date_filter and time_col and time_col in final_matched.columns:
            try:
                time_series = final_matched.copy()
                time_series['Date'] = parse_datetime_auto(time_series[time_col]).dt.date
                time_series_data = time_series.groupby('Date').size().reset_index(name='Count')
            except:
                pass
        
        # Prepare statistics for dashboard
        stats_dashboard = {
            'External_Total': len(ext_df),
            'External_Filtered': len(filtered_ext_df),
            'Internal_Total': len(int_df),
            'Campaign_Total': len(camp_df),
            'Step1_Matches': len(internal_matched),
            'Step2_Matches': len(final_matched),
            'Final_Matches': len(final_matched),
            'Unmatched_External': len(unmatched_external),
            'TimeSeriesData': time_series_data
        }
        
        # Create stats dataframe
        stats_df = pd.DataFrame({
            "Metric": [
                "Processing Date", "External Records (Total)", "External Records (Filtered)",
                "Internal Records", "Campaign Records", "Step 1 Matches", "Step 1 Match Rate",
                "Step 2 Matches", "Step 2 Match Rate", "Unmatched Records", "Overall Success Rate"
            ],
            "Value": [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                f"{len(ext_df):,}", f"{len(filtered_ext_df):,}",
                f"{len(int_df):,}", f"{len(camp_df):,}",
                f"{len(internal_matched):,}", f"{match_rate_step1:.2f}%",
                f"{len(final_matched):,}", f"{match_rate_step2:.2f}%",
                f"{len(unmatched_external):,}",
                f"{(len(final_matched)/len(filtered_ext_df)*100):.2f}%" if len(filtered_ext_df) > 0 else "0%"
            ]
        })
        
        # Generate visualizations
        fig_funnel, fig_gauge, fig_donut, fig_time = create_visualizations(stats_dashboard)
        
        # Store ALL results in session_state to persist after download
        st.session_state.processing_complete = True
        st.session_state.final_matched = final_matched
        st.session_state.internal_matched = internal_matched
        st.session_state.unmatched_external = unmatched_external
        st.session_state.stats_df = stats_df
        st.session_state.stats_dashboard = stats_dashboard
        st.session_state.filtered_ext_df = filtered_ext_df
        st.session_state.match_rate_step1 = match_rate_step1
        st.session_state.match_rate_step2 = match_rate_step2
        st.session_state.results = {
            'match_rate': (len(final_matched) / len(filtered_ext_df)) * 100 if len(filtered_ext_df) > 0 else 0,
            'total_matches': len(final_matched),
            'unmatched': len(unmatched_external),
            'filtered_count': len(filtered_ext_df)
        }
        st.session_state.fig_funnel = fig_funnel
        st.session_state.fig_gauge = fig_gauge
        st.session_state.fig_donut = fig_donut
        st.session_state.fig_time = fig_time
        
        # Clear progress indicators
        progress_bar.empty()
        status_text.empty()
            
    except Exception as e:
        st.error(f"❌ An error occurred: {str(e)}")
        st.exception(e)
        progress_bar.empty()
        status_text.empty()
        st.stop()

# Display results if processing is complete (retrieved from session_state)
if st.session_state.processing_complete:
    # Retrieve data from session_state
    final_matched = st.session_state.final_matched
    internal_matched = st.session_state.internal_matched
    unmatched_external = st.session_state.unmatched_external
    stats_df = st.session_state.stats_df
    stats_dashboard = st.session_state.stats_dashboard
    filtered_ext_df = st.session_state.filtered_ext_df
    match_rate_step1 = st.session_state.match_rate_step1
    match_rate_step2 = st.session_state.match_rate_step2
    
    # Retrieve figures
    fig_funnel = st.session_state.fig_funnel
    fig_gauge = st.session_state.fig_gauge
    fig_donut = st.session_state.fig_donut
    fig_time = st.session_state.fig_time
    
    # If figures are None for some reason, regenerate them
    if fig_funnel is None and stats_dashboard is not None:
        fig_funnel, fig_gauge, fig_donut, fig_time = create_visualizations(stats_dashboard)
    
    # Display results in tabs
    st.markdown("---")
    st.markdown("## 📈 Reconciliation Dashboard")
    
    # Metrics row
    metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
    with metric_col1:
        st.markdown(f"""
        <div class='metric-card fade-in-up'>
            <h3>📊 External Records</h3>
            <h2>{len(filtered_ext_df) if filtered_ext_df is not None else 0:,}</h2>
            <small>After filter</small>
        </div>
        """, unsafe_allow_html=True)
    with metric_col2:
        st.markdown(f"""
        <div class='metric-card fade-in-up'>
            <h3>✅ Order ID Matches</h3>
            <h2>{len(internal_matched) if internal_matched is not None else 0:,}</h2>
            <small>{match_rate_step1:.1f}% match rate</small>
        </div>
        """, unsafe_allow_html=True)
    with metric_col3:
        st.markdown(f"""
        <div class='metric-card fade-in-up'>
            <h3>📞 Phone Matches</h3>
            <h2>{len(final_matched) if final_matched is not None else 0:,}</h2>
            <small>{match_rate_step2:.1f}% of matched</small>
        </div>
        """, unsafe_allow_html=True)
    with metric_col4:
        st.markdown(f"""
        <div class='metric-card fade-in-up'>
            <h3>⚠️ Unmatched</h3>
            <h2>{len(unmatched_external) if unmatched_external is not None else 0:,}</h2>
            <small>External records</small>
        </div>
        """, unsafe_allow_html=True)
    
    # Visualizations row
    if fig_funnel is not None:
        st.markdown("### 📊 Interactive Analytics")
        viz_col1, viz_col2 = st.columns(2)
        with viz_col1:
            st.plotly_chart(fig_funnel, width="stretch", key="funnel_display", config={'displayModeBar': False})
        with viz_col2:
            st.plotly_chart(fig_gauge, width="stretch", key="gauge_display", config={'displayModeBar': False})
        
        viz_col3, viz_col4 = st.columns(2)
        with viz_col3:
            st.plotly_chart(fig_donut, width="stretch", key="donut_display", config={'displayModeBar': False})
        with viz_col4:
            if fig_time:
                st.plotly_chart(fig_time, width="stretch", key="timeseries_display", config={'displayModeBar': False})
            else:
                st.info("Time series data not available for selected date range")
    
    # Data tables
    st.markdown("### 📋 Detailed Results")
    tab1, tab2, tab3, tab4 = st.tabs([
        "✅ Final Matched", "📋 Internal Matched", "❌ Unmatched External", "📊 Statistics"
    ])
    
    with tab1:
        if final_matched is not None and not final_matched.empty:
            st.dataframe(final_matched, width="stretch", height=400)
            col1, col2, col3 = st.columns([1, 1, 2])
            with col1:
                csv_final = final_matched.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "📥 Download CSV",
                    csv_final,
                    f"final_matched_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="download_final_matched_csv",
                    width="stretch"
                )
        else:
            st.info("No final matched records found")
    
    with tab2:
        if internal_matched is not None and not internal_matched.empty:
            st.dataframe(internal_matched, width="stretch", height=400)
            csv_internal = internal_matched.to_csv(index=False).encode('utf-8')
            st.download_button(
                "📥 Download CSV",
                csv_internal,
                f"internal_matched_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                key="download_internal_matched_csv",
                width="stretch"
            )
        else:
            st.info("No internal matched records found")
    
    with tab3:
        if unmatched_external is not None and not unmatched_external.empty:
            st.dataframe(unmatched_external, width="stretch", height=400)
            csv_unmatched = unmatched_external.to_csv(index=False).encode('utf-8')
            st.download_button(
                "📥 Download CSV",
                csv_unmatched,
                f"unmatched_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                key="download_unmatched_csv",
                width="stretch"
            )
        else:
            st.success("🎉 All external records matched successfully!")
    
    with tab4:
        if stats_df is not None and not stats_df.empty:
            st.dataframe(stats_df, width="stretch", hide_index=True)
        else:
            st.info("No statistics available")
    
    # Export complete report
    st.markdown("---")
    st.markdown("### 💾 Export Complete Report")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if stats_df is not None and not stats_df.empty and final_matched is not None:
            export_stats = dict(zip(stats_df['Metric'], stats_df['Value']))
            excel_report = to_excel_advanced(
                matched_internal=internal_matched if internal_matched is not None else pd.DataFrame(),
                final_matched=final_matched if final_matched is not None else pd.DataFrame(),
                unmatched_external=unmatched_external if unmatched_external is not None else pd.DataFrame(),
                stats=export_stats
            )
            st.download_button(
                "📊 Download Complete Excel Report",
                excel_report,
                f"reconciliation_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                key="download_complete_report"
            )

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; padding: 20px; color: #666;'>
    <small>
    🎯 Campaign Reconciliation Studio v2.0 | Enterprise Edition<br>
    🔒 Data Privacy Guaranteed | All processing occurs locally in your browser<br>
    📱 Phone numbers normalized to rightmost 10 digits
    </small>
</div>
""", unsafe_allow_html=True)    