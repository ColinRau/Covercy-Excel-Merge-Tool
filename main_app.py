import streamlit as st
import runpy

# Set global page configuration once, before any UI elements
st.set_page_config(
    page_title="Covercy Excel Merge Tool 2.0",
    page_icon="logo.png",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ------------------------------------------------------------------
# Top-level navigation tabs
# ------------------------------------------------------------------

dist_tab, contrib_tab = st.tabs(["📤 Distributions", "📥 Contributions"])

# Render the existing Distributions app inside the first tab. We execute
# the original script via runpy so no changes to its internal logic are
# required.
with dist_tab:
    runpy.run_path("streamlit_excel_merge_app.py", run_name="__main__")

# Render the placeholder Contributions flow inside the second tab.
with contrib_tab:
    from contributions_flow import run_contributions_flow

    run_contributions_flow() 