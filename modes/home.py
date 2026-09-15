import streamlit as st


def render() -> None:
    st.markdown(
        """
        <div style="padding: 1rem 0 0.25rem 0;">
            <h3 style="margin:0; color:#0f172a;">Workspace overview</h3>
        </div>
        """,
        unsafe_allow_html=True,
    )

    cols = st.columns(3)
    with cols[0]:
        st.metric("Reports", "21")
    with cols[1]:
        st.metric("Categories", "8")
    with cols[2]:
        st.metric("Status", "Ready")

    st.info("Each mode loads only when selected from the sidebar, keeping the startup fast and avoiding unnecessary processing.")
