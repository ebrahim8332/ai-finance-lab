import streamlit as st

def render():
    st.markdown("## Financial Ratio Interpreter")
    st.markdown(
        "Paste or upload a balance sheet and income statement. The AI computes "
        "key ratios — current ratio, EBITDA margin, DSO, debt-to-equity — and "
        "writes a plain-English interpretation of what they mean for the business."
    )
    st.markdown("---")
    st.info("This module is coming soon.")
