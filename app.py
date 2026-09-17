import importlib
from pathlib import Path

import streamlit as st

from core.error_handler import show_mode_error
from core.loader import execute_legacy_script
from modes.registry import get_mode_registry

st.set_page_config(
    page_title="EDP Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(180deg, #f4f7fb 0%, #eef3f8 100%);
    }
    .sidebar-content {
        background: rgba(255,255,255,0.86);
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #edf3f9 0%, #edf2f7 100%);
        border-right: 1px solid rgba(15,23,42,0.08);
    }
    .app-header {
        background: rgba(255,255,255,0.7);
        border: 1px solid rgba(15,23,42,0.06);
        border-radius: 18px;
        padding: 1.15rem 1.25rem;
        margin-bottom: 1rem;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
    }
    .mode-hero {
        background: linear-gradient(135deg, #eff6ff 0%, #f8fafc 100%);
        border: 1px solid rgba(59,130,246,0.15);
        border-radius: 18px;
        padding: 1.2rem 1.3rem;
        margin-bottom: 1.25rem;
        box-shadow: 0 10px 28px rgba(59, 130, 246, 0.08);
    }
    .mode-badge {
        display: inline-block;
        background: rgba(59,130,246,0.12);
        color: #1d4ed8;
        border: 1px solid rgba(59,130,246,0.18);
        border-radius: 999px;
        padding: 0.3rem 0.8rem;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.04em;
        text-transform: uppercase;
    }
    .section-title {
        font-size: 2rem;
        font-weight: 800;
        margin: 0;
        color: #0f172a;
    }
    .section-subtitle {
        font-size: 1rem;
        color: #475569;
        margin-top: 0.3rem;
    }
    .stat-card {
        background: rgba(255,255,255,0.8);
        border: 1px solid rgba(148,163,184,0.20);
        border-radius: 16px;
        padding: 1rem 1.1rem;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.03);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.sidebar.title("EDP Analytics")
st.sidebar.caption("Unified reporting workspace")

mode_registry = get_mode_registry()
search = st.sidebar.text_input("Search modes", placeholder="Search reports and tools")

filtered_modes = sorted(
    [
        mode
        for mode in mode_registry
        if not search.strip()
        or search.casefold() in mode["name"].casefold()
        or search.casefold() in mode["description"].casefold()
        or search.casefold() in mode["category"].casefold()
    ],
    key=lambda mode: f"{mode['category']} / {mode['name']}".casefold(),
)

if not filtered_modes:
    st.sidebar.warning("No matching modes found.")
    st.stop()

labels = [f"{mode['category']} / {mode['name']}" for mode in filtered_modes]
selected_label = st.sidebar.selectbox("Select mode", labels)
selected_mode = filtered_modes[labels.index(selected_label)]

st.sidebar.markdown("---")
st.sidebar.markdown(f"**Mode:** {selected_mode['name']}")
st.sidebar.caption(selected_mode["description"])
st.sidebar.caption(f"Categories: {len({m['category'] for m in mode_registry})}")
st.sidebar.caption(f"Modes available: {len(mode_registry)}")

main_title = "EDP Analytics"
subtitle = "Select a report or tool from the sidebar to begin."
if selected_mode["id"] != "home":
    main_title = selected_mode["name"]
    subtitle = selected_mode["description"]

st.markdown(
    f"""
    <div class="app-header">
        <div class="section-title">{main_title}</div>
        <div class="section-subtitle">{subtitle}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if selected_mode["id"] == "home":
    st.markdown(
        """
        <div class="mode-hero">
            <div class="mode-badge">Dashboard</div>
            <h3 style="margin: 0.7rem 0 0.3rem 0; color: #0f172a;">Welcome to your analytics workspace</h3>
            <p style="margin: 0; color: #475569;">The system loads data only when a mode is selected, which keeps startup fast and prevents unnecessary file processing.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    cols = st.columns(3)
    for col, metric in zip(cols, [("Modes", str(len(mode_registry))), ("Categories", str(len({m['category'] for m in mode_registry}))), ("Status", "Ready")]):
        with col:
            st.markdown(
                f"""
                <div class="stat-card">
                    <div style="font-size: 0.78rem; color: #64748b; text-transform: uppercase; letter-spacing: 0.04em;">{metric[0]}</div>
                    <div style="font-size: 2rem; font-weight: 800; color: #0f172a; margin-top: 0.35rem;">{metric[1]}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.info("Data files are loaded only after you select a mode.")
else:
    st.markdown(
        f"""
        <div class="mode-hero">
            <div class="mode-badge">{selected_mode['category']}</div>
            <h3 style="margin: 0.7rem 0 0.3rem 0; color: #0f172a;">Selected mode</h3>
            <p style="margin: 0; color: #475569;">{selected_mode['name']}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    try:
        if selected_mode.get("mode_type") == "script":
            source_path = Path(__file__).resolve().parent / selected_mode["source"]
            execute_legacy_script(source_path)
        else:
            selected_module = importlib.import_module(selected_mode["module"])
            render_mode = getattr(selected_module, "render", None)
            run_mode = getattr(selected_module, "run", None)
            if callable(render_mode):
                render_mode()
            elif callable(run_mode):
                run_mode()
            else:
                raise AttributeError(
                    f"{selected_mode['module']} must define render() or run()"
                )
    except Exception as error:  # Keep one broken report from taking down the shell.
        show_mode_error(selected_mode, error)