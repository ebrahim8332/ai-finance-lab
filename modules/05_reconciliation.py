import streamlit as st

def render():
    st.markdown("## Reconciliation Exception Explainer")
    st.markdown(
        "Paste a list of reconciling items. The AI categorizes each item, "
        "assigns the most likely root cause, and recommends a resolution action."
    )
    st.markdown("---")
    st.info("This module is coming soon.")
