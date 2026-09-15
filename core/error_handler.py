import streamlit as st


def show_mode_error(mode: dict, error: Exception) -> None:
    st.error(f"Mode failed: {mode['name']}")
    st.caption(f"Source: {mode['source']}")
    st.exception(error)
