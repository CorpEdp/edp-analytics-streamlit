"""
main.py — single entry point for the merged multi-app Streamlit project.

Run with:
    streamlit run main.py

How it works:
- Every .py file inside modules/ (except files starting with "_") is treated
  as one "mode" / one of your old app.py files.
- Each module must define a run() function (this is what gets called when
  the user picks that mode from the sidebar).
- Each module can optionally define a LABEL string for a friendly sidebar
  name; otherwise the filename is used.

Adding a new app later = drop a new file into modules/. No edits needed here.
"""

import importlib
import pkgutil

import streamlit as st

import modules

# set_page_config must be called exactly ONCE, and only here.
# Make sure none of your individual modules call st.set_page_config(...).
st.set_page_config(
    page_title="ERP / Scheme Toolkit",
    page_icon="📊",
    layout="wide",
)


@st.cache_resource(show_spinner=False)
def discover_modes():
    """
    Scan modules/ once per session and build {label: module} for every
    submodule that exposes a run() function. Cached so we don't re-import
    on every rerun/widget interaction.
    """
    modes = {}
    errors = []

    for _, name, is_pkg in pkgutil.iter_modules(modules.__path__):
        if is_pkg or name.startswith("_"):
            continue
        try:
            mod = importlib.import_module(f"modules.{name}")
        except Exception as e:  # noqa: BLE001 - surface any broken module, don't crash the app
            errors.append((name, str(e)))
            continue

        if not hasattr(mod, "run"):
            errors.append((name, "no run() function defined"))
            continue

        label = getattr(mod, "LABEL", name.replace("_", " ").title())
        modes[label] = mod

    return modes, errors


modes, load_errors = discover_modes()

st.sidebar.title("📊 Toolkit")
st.sidebar.caption(f"{len(modes)} tools loaded")

if load_errors:
    with st.sidebar.expander(f"⚠️ {len(load_errors)} module(s) failed to load"):
        for name, err in load_errors:
            st.write(f"**{name}**: {err}")

if not modes:
    st.error(
        "No valid modules found in modules/. "
        "Make sure each file defines a run() function."
    )
    st.stop()

choice = st.sidebar.selectbox("Choose a report / tool", sorted(modes.keys()))

st.sidebar.divider()
if st.sidebar.button("🔄 Force reload modules"):
    discover_modes.clear()
    st.rerun()

st.title(choice)
modes[choice].run()
