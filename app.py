"""
FinanceAI Lab — main entry point.

Sidebar navigation routes to each module. Each module lives in modules/
and exposes a render() function. Add a new module by importing it here
and adding one line to MODULES and NAV_OPTIONS.
"""

import streamlit as st
import importlib
from dotenv import load_dotenv

load_dotenv()

# ── Page config — must be the first Streamlit call ────────────────────────
st.set_page_config(
    page_title="FinanceAI Lab",
    page_icon="📊",
    layout="wide",
)

# ── Minimal global CSS (pointer cursor on interactive controls) ────────────
st.markdown("""
<style>
div[data-baseweb="select"] > div,
div[data-baseweb="select"] input,
button[kind="primary"],
button[kind="secondary"],
.stButton > button,
.stDownloadButton > button,
input[type="range"] {
    cursor: pointer !important;
}
</style>
""", unsafe_allow_html=True)


# ── Module registry ────────────────────────────────────────────────────────
# Maps nav label → module import path.
# To add a module: add one entry here and create the file in modules/.
NAV_OPTIONS = [
    ("📊 Variance Commentary",               "modules.01_variance_commentary",    True),
    ("📝 Journal Entry Drafting",            "modules.02_journal_entry",          False),
    ("📋 Budget Narrative Generator",        "modules.03_budget_narrative",       False),
    ("🔍 Audit Workpaper Assistant",         "modules.04_audit_workpaper",        False),
    ("⚖️ Reconciliation Exception Explainer","modules.05_reconciliation",         False),
    ("📐 Financial Ratio Interpreter",       "modules.06_ratio_interpreter",      False),
    ("💬 Board Question Simulator",          "modules.07_board_questions",        False),
    ("💵 Cash Flow Commentary",              "modules.08_cashflow_commentary",    False),
]
# (label, module_path, is_live)


# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📊 FinanceAI Lab")
    st.caption("AI-powered tools for FP&A and accounting professionals.")
    st.divider()

    st.markdown("### Module")
    labels = [item[0] for item in NAV_OPTIONS]
    selected_label = st.radio(
        "Navigate",
        options=labels,
        label_visibility="collapsed",
    )
    st.divider()

    # Status key
    live_count = sum(1 for item in NAV_OPTIONS if item[2])
    total_count = len(NAV_OPTIONS)
    st.caption(f"{live_count} of {total_count} modules live")

    # Show active AI model if one has been locked this session
    if "locked_provider_index" in st.session_state:
        try:
            from utils.ai_client import _build_cached_providers
            providers = _build_cached_providers()
            idx = st.session_state.get("locked_provider_index", 0)
            if idx < len(providers):
                short = providers[idx].model_name.split("/")[-1]
                st.caption(f"AI model: {short}")
        except Exception:
            pass


# ── Route to selected module ───────────────────────────────────────────────
# Find the module path for the selected label
module_path = None
is_live = False
for label, path, live in NAV_OPTIONS:
    if label == selected_label:
        module_path = path
        is_live = live
        break

if not is_live:
    # Show a clean "coming soon" page for unbuilt modules
    st.markdown(f"## {selected_label}")
    st.markdown("---")
    st.info("This module is coming soon.")
else:
    try:
        mod = importlib.import_module(module_path)
        mod.render()
    except Exception as e:
        st.error(f"Error loading module: {e}")
        st.exception(e)
