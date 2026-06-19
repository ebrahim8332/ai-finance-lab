import streamlit as st

def render():
    st.markdown("## Journal Entry Drafting")
    st.markdown(
        "Describe a transaction in plain English. The AI produces a properly "
        "formatted double-entry journal entry with account codes, debits, credits, "
        "and a brief rationale."
    )
    st.markdown("---")
    st.info("This module is coming soon.")
