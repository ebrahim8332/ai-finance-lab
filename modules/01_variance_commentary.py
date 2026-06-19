"""
Module 1: Variance Commentary

User uploads a budget vs. actual Excel file. App auto-detects company name and
period from file content. Shows data table and five finance charts automatically.
AI generates structured commentary tailored to the chosen audience.
Output is collapsible. Download as a Word document with charts embedded.
"""

import io
import re
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

import pandas as pd

from utils.ai_client import get_chain
from utils.excel_parser import (
    parse_variance_sheet, dataframe_to_text, extract_file_metadata,
    is_parse_suspicious, parse_variance_sheet_robust,
)
from utils.chart_builder import build_all_charts, fig_to_png_bytes, prepare_data
from utils.base import AllProvidersExhausted


def _safe_md(text: str) -> str:
    """Escapes dollar signs so Streamlit doesn't render them as LaTeX math."""
    return text.replace("$", r"\$")


def _build_display_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns a formatted copy of the dataframe for screen display only.
    The raw df (unmodified) goes to the AI and the chart builder.

    Rules applied:
    - Dollar / number columns (budget, actual, variance): comma-separated integers  e.g. 1,234
    - Variance % columns: multiplied by 100 and shown as  e.g. 5.7%
    - All other columns: left as-is
    """
    disp = df.copy()
    for col in disp.columns:
        col_lower = col.lower()
        if '%' in col_lower or 'pct' in col_lower or 'percent' in col_lower:
            disp[col] = disp[col].apply(
                lambda v: f"{v * 100:.1f}%"
                if pd.notna(v) and isinstance(v, (int, float))
                else (str(v) if pd.notna(v) else "")
            )
        elif any(kw in col_lower for kw in ['budget', 'actual', 'variance', '$', '($k)', 'amount']):
            disp[col] = disp[col].apply(
                lambda v: f"{v:,.0f}"
                if pd.notna(v) and isinstance(v, (int, float))
                else (str(v) if pd.notna(v) else "")
            )
    return disp


AUDIENCE_DESCRIPTIONS = {
    "CFO":        "senior finance executive who wants concise, technically precise commentary "
                  "with focus on margin drivers, working capital, and corrective actions",
    "Board":      "board of directors who want high-level narrative, key risks, and strategic "
                  "implications — minimal technical detail, maximum clarity",
    "Operations": "operational managers who need to understand what drove variances in their "
                  "areas and what actions they should take this quarter",
}

AUDIENCE_TONE = {
    "CFO":        "concise, technically precise, focused on root cause and action",
    "Board":      "strategic, narrative-driven, focused on implications and risk",
    "Operations": "practical, direct, focused on what happened and what to do next",
}


# ── Word document builder ─────────────────────────────────────────────────

def _add_bold_runs(paragraph, text: str):
    """Converts **bold** markdown markers to actual bold formatting in Word."""
    for part in re.split(r'(\*\*.*?\*\*)', text):
        if part.startswith('**') and part.endswith('**'):
            paragraph.add_run(part[2:-2]).bold = True
        else:
            paragraph.add_run(part)


def build_word_doc(
    commentary_text: str,
    company: str,
    period: str,
    audience: str,
    chart_pngs: list[tuple[str, bytes]],  # list of (chart_title, png_bytes)
) -> bytes:
    """
    Builds a Word document with charts embedded above the AI commentary.
    Returns bytes for the Streamlit download button.
    """
    doc = Document()

    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.1)
        section.right_margin  = Inches(1.1)

    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)

    # Title
    title = doc.add_paragraph()
    run = title.add_run(f'{company} — {period} Management Commentary')
    run.bold = True
    run.font.size = Pt(16)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph(f'Audience: {audience}')
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].italic = True
    sub.runs[0].font.size = Pt(10)

    # ── Charts section ────────────────────────────────────────────────────
    if chart_pngs:
        doc.add_paragraph()
        charts_heading = doc.add_paragraph()
        charts_heading.add_run('Charts').bold = True
        charts_heading.runs[0].font.size = Pt(14)

        for chart_title, png_bytes in chart_pngs:
            if png_bytes:
                doc.add_paragraph(chart_title).runs[0].italic = True
                doc.add_picture(io.BytesIO(png_bytes), width=Inches(6.0))
                doc.add_paragraph()  # spacing between charts

    # ── Commentary section ────────────────────────────────────────────────
    doc.add_paragraph()
    comm_heading = doc.add_paragraph()
    comm_heading.add_run('Management Commentary').bold = True
    comm_heading.runs[0].font.size = Pt(14)
    doc.add_paragraph()

    for line in commentary_text.splitlines():
        s = line.strip()
        if not s:
            doc.add_paragraph()
        elif s.startswith('### '):
            doc.add_heading(s[4:], level=3)
        elif s.startswith('## '):
            doc.add_heading(s[3:], level=2)
        elif s.startswith('# '):
            doc.add_heading(s[2:], level=1)
        elif s.startswith('- '):
            p = doc.add_paragraph(style='List Bullet')
            _add_bold_runs(p, s[2:])
        else:
            p = doc.add_paragraph()
            _add_bold_runs(p, s)

    doc.add_paragraph()
    footer = doc.add_paragraph('AI-generated content. Review all figures before distributing.')
    footer.runs[0].italic = True
    footer.runs[0].font.size = Pt(9)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ── Prompt builder ────────────────────────────────────────────────────────

def build_prompt(data_text: str, audience: str, period: str, company: str) -> list[dict]:
    system = f"""You are a senior FP&A analyst writing management commentary for {company}.
Your audience is a {AUDIENCE_DESCRIPTIONS[audience]}.
Tone: {AUDIENCE_TONE[audience]}.

Rules:
- Write in clear, professional English. No jargon for the sake of it.
- Do not use phrases like "leverage", "ecosystem", "transformative", or "actionable insights".
- Label every section clearly with a markdown heading (## Section Name).
- Be specific: reference actual numbers from the data.
- All figures are in USD thousands unless stated otherwise.
- Do not hallucinate. Only reference line items present in the data.
- End with a Conclusions section that adds new synthesis, not a restatement.
"""
    user = f"""Write management commentary for {company} for the period: {period}.

FINANCIAL DATA:
{data_text}

Structure the commentary as follows:
## Executive Summary
## Revenue Analysis
## Cost and Margin Analysis
## Operating Performance
## Key Risks and Watch Items
## Conclusions

Audience: {audience}. Tailor depth and language accordingly.
Write in paragraphs, not bullet points.
"""
    return [
        {"role": "system", "content": system},
        {"role": "user",   "content": user},
    ]


# ── Module renderer ───────────────────────────────────────────────────────

def render():
    st.markdown("## Variance Commentary")
    st.markdown(
        "Upload a budget vs. actual Excel file. "
        "The app reads the company name and reporting period directly from the file, "
        "generates five finance charts automatically, and writes AI commentary "
        "tailored to your chosen audience. Download everything as a Word document."
    )

    st.markdown(
        '<hr style="border: none; border-top: 3px solid #d0d0d0; margin: 18px 0 20px 0;">',
        unsafe_allow_html=True,
    )

    # ── Inputs ────────────────────────────────────────────────────────────
    col1, col2 = st.columns([2, 1])

    with col1:
        uploaded_file = st.file_uploader(
            "Upload your budget vs. actual file",
            type=["xlsx", "xls"],
            help="Works with most Excel variance reports. Data should start in row 1 "
                 "with column headers. If the preview looks wrong, move your data to "
                 "the first sheet and remove any title rows above the headers.",
        )

    # Extract metadata from file content as soon as a new file is uploaded
    if uploaded_file is not None:
        meta = extract_file_metadata(uploaded_file)
        company_from_file = meta["company"] or uploaded_file.name.rsplit(".", 1)[0].replace("_", " ").replace("-", " ").title()
        period_from_file  = meta["period"] or "Current Period"
        if st.session_state.get("_last_uploaded") != uploaded_file.name:
            st.session_state["detected_company"] = company_from_file
            st.session_state["detected_period"]  = period_from_file
            st.session_state["_last_uploaded"]   = uploaded_file.name

    with col2:
        audience = st.selectbox(
            "Commentary audience",
            options=["CFO", "Board", "Operations"],
            help="Changes the tone, depth, and focus of the commentary.",
        )
        company = st.text_input(
            "Company name",
            value=st.session_state.get("detected_company", ""),
        )
        period = st.text_input(
            "Reporting period",
            value=st.session_state.get("detected_period", "Current Period"),
        )

    st.markdown(
        '<hr style="border: none; border-top: 3px solid #d0d0d0; margin: 18px 0 20px 0;">',
        unsafe_allow_html=True,
    )

    if uploaded_file is None:
        return

    # ── Parse file ────────────────────────────────────────────────────────
    # Try the fast standard parser first.
    # If it produces bad results (unnamed columns, title-as-header), fall back
    # to AI extraction — the file is converted to text and the AI returns JSON.
    try:
        file_bytes = uploaded_file.read()
        uploaded_file.seek(0)
        df = parse_variance_sheet(uploaded_file)
    except Exception as e:
        st.error(f"Could not read the file: {e}")
        return

    if is_parse_suspicious(df):
        try:
            df = parse_variance_sheet_robust(file_bytes)
            st.caption("Complex file structure detected — data extracted using enhanced parser.")
        except Exception as e:
            st.error(
                f"Could not parse this file: {e} "
                "Try simplifying to four columns: Line Item, Budget, Actual, Variance %."
            )
            return

    # ── Data table (auto-visible, horizontal scroll enabled) ──────────────
    # _build_display_df formats numbers for readability.
    # The raw df is kept untouched for the AI and chart builder.
    st.markdown("**Data loaded from file:**")
    st.dataframe(_build_display_df(df), use_container_width=True, hide_index=True)

    # ── Column detection display ──────────────────────────────────────────
    # Shows which columns the app mapped to each role so the user can verify
    _, detected_cols = prepare_data(df)
    col_parts = []
    for role, label in [("Label", "label"), ("Budget", "budget"), ("Actual", "actual")]:
        val = detected_cols.get(label)
        if val and not val.startswith("_"):   # skip computed columns
            col_parts.append(f"**{role}:** {val}")
    if col_parts:
        st.caption("Detected columns — " + "  ·  ".join(col_parts) +
                   "  ·  *If any column is wrong, rename it in your file and re-upload.*")

    st.markdown(
        '<hr style="border: none; border-top: 3px solid #d0d0d0; margin: 18px 0 20px 0;">',
        unsafe_allow_html=True,
    )

    # ── Charts (auto-generated from data) ─────────────────────────────────
    charts = build_all_charts(df)

    if charts:
        with st.expander("📊 Charts", expanded=True):
            for i in range(0, len(charts), 2):
                cols_chart = st.columns(2)
                for j, (title, fig) in enumerate(charts[i:i+2]):
                    with cols_chart[j]:
                        with st.container(border=True):
                            st.plotly_chart(fig, use_container_width=True)

    st.markdown(
        '<hr style="border: none; border-top: 3px solid #d0d0d0; margin: 18px 0 20px 0;">',
        unsafe_allow_html=True,
    )

    # ── AI Q&A on the data ────────────────────────────────────────────────
    # The user can ask any question about the numbers. The AI answers using
    # only the data in the file — it cannot reference figures not present.
    st.markdown("#### Ask a question about your data")
    st.caption(
        "Ask anything: 'Which line item had the biggest variance?' · "
        "'What drove the drop in net income?' · "
        "'Which three items should management focus on?'"
    )

    qa_col1, qa_col2 = st.columns([4, 1])
    with qa_col1:
        question = st.text_input(
            "Your question",
            placeholder="e.g. Which expenses were most over budget?",
            label_visibility="collapsed",
            key="qa_input",
        )
    with qa_col2:
        ask_clicked = st.button("Ask AI", type="primary", use_container_width=True)

    if ask_clicked and question.strip():
        data_text = dataframe_to_text(df)
        qa_messages = [
            {
                "role": "system",
                "content": (
                    "You are a senior FP&A analyst. Answer the user's question using only "
                    "the financial data provided. Be specific — reference actual numbers. "
                    "Keep the answer concise: 3-5 sentences maximum. "
                    "Do not reference any figures not present in the data."
                ),
            },
            {
                "role": "user",
                "content": f"FINANCIAL DATA:\n{data_text}\n\nQUESTION: {question.strip()}",
            },
        ]

        with st.spinner("Thinking..."):
            try:
                chain = get_chain(st.session_state)
                answer, model_used_qa = chain.complete(qa_messages, timeout=60)
                # Append to Q&A history so multiple questions accumulate on screen
                if "qa_history" not in st.session_state:
                    st.session_state["qa_history"] = []
                st.session_state["qa_history"].append({
                    "question": question.strip(),
                    "answer":   answer,
                    "model":    model_used_qa,
                })
            except AllProvidersExhausted as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"Unexpected error: {e}")

    # Display Q&A history — newest first
    if st.session_state.get("qa_history"):
        for item in reversed(st.session_state["qa_history"]):
            with st.expander(f"💬 {item['question']}", expanded=True):
                st.markdown(_safe_md(item["answer"]))
                st.caption(f"AI response from {item['model']}")

        if st.button("Clear questions", key="clear_qa"):
            st.session_state["qa_history"] = []
            st.rerun()

    st.markdown(
        '<hr style="border: none; border-top: 3px solid #d0d0d0; margin: 18px 0 20px 0;">',
        unsafe_allow_html=True,
    )

    # ── Generate button ───────────────────────────────────────────────────
    if st.button("Generate AI Commentary", type="primary", use_container_width=True):
        data_text = dataframe_to_text(df)
        messages  = build_prompt(data_text, audience, period, company)

        with st.spinner("Analyzing variances and writing commentary..."):
            try:
                chain = get_chain(st.session_state)
                response, model_used = chain.complete(messages, timeout=120)
            except AllProvidersExhausted as e:
                st.error(str(e))
                return
            except Exception as e:
                st.error(f"Unexpected error: {e}")
                return

        st.session_state["var_commentary"]  = response
        st.session_state["var_model_used"]  = model_used
        st.session_state["var_company"]     = company
        st.session_state["var_period"]      = period
        st.session_state["var_audience"]    = audience
        st.session_state["var_charts"]      = charts  # store for Word doc

    # ── Commentary output ─────────────────────────────────────────────────
    if "var_commentary" in st.session_state:
        response     = st.session_state["var_commentary"]
        model_used   = st.session_state["var_model_used"]
        company_out  = st.session_state["var_company"]
        period_out   = st.session_state["var_period"]
        audience_out = st.session_state["var_audience"]
        saved_charts = st.session_state.get("var_charts", [])

        st.markdown(
            '<hr style="border: none; border-top: 3px solid #d0d0d0; margin: 18px 0 20px 0;">',
            unsafe_allow_html=True,
        )

        with st.expander(
            f"📋 AI Commentary — {company_out}, {period_out} ({audience_out})",
            expanded=True,
        ):
            st.markdown(_safe_md(response))
            st.caption(f"AI response from {model_used}")

        # Export charts to PNG for Word doc
        chart_pngs = [(title, fig_to_png_bytes(fig)) for title, fig in saved_charts]

        word_bytes = build_word_doc(response, company_out, period_out, audience_out, chart_pngs)
        st.download_button(
            label="⬇ Download as Word document (.docx)",
            data=word_bytes,
            file_name=f"{company_out.lower().replace(' ', '_')}_{period_out.lower().replace(' ', '_')}_commentary.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        st.caption("Review all AI-generated content before distributing. Verify figures against source data.")
