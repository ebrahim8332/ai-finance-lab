"""
Module 1: Variance Commentary

Takes a budget vs. actual Excel file and generates structured management
commentary tailored to a chosen audience: CFO, Board, or Operations.
Commentary is displayed in a collapsible expander and downloadable as a Word doc.
"""

import io
import re
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils.ai_client import get_chain
from utils.excel_parser import parse_variance_sheet, dataframe_to_text
from utils.base import AllProvidersExhausted


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


def build_word_doc(commentary_text: str, company: str, period: str, audience: str) -> bytes:
    """
    Builds a Word document from the AI commentary text.
    Returns raw bytes so Streamlit can offer it as a download without saving to disk.
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

    # Subtitle
    sub = doc.add_paragraph(f'Audience: {audience}')
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].italic = True
    sub.runs[0].font.size = Pt(10)

    doc.add_paragraph()

    # Render commentary — parse markdown headings and bold text
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

    # Footer
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
        "Upload a budget vs. actual Excel file. The AI reads the numbers, "
        "identifies key variances, and writes structured management commentary "
        "tailored to your chosen audience. What takes an analyst 30-60 minutes "
        "is produced in seconds. Download the result as a Word document."
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
            help="The file should have columns: Line Item, Budget, Actual. "
                 "Download the sample file below to see the expected format.",
        )

    with col2:
        audience = st.selectbox(
            "Commentary audience",
            options=["CFO", "Board", "Operations"],
            help="Changes the tone, depth, and focus of the commentary.",
        )
        company = st.text_input("Company name", value="Meridian Corp")
        period  = st.text_input("Reporting period", value="Q1 2026")

    # ── Sample file download ──────────────────────────────────────────────
    try:
        with open("data/sample_variance.xlsx", "rb") as f:
            st.download_button(
                label="Download sample file (Meridian Corp Q1 2026)",
                data=f,
                file_name="sample_variance.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except FileNotFoundError:
        pass

    st.markdown(
        '<hr style="border: none; border-top: 3px solid #d0d0d0; margin: 18px 0 20px 0;">',
        unsafe_allow_html=True,
    )

    if uploaded_file is None:
        st.info("Upload a file above to generate commentary.")
        return

    # ── Parse and preview ─────────────────────────────────────────────────
    try:
        df = parse_variance_sheet(uploaded_file)
    except Exception as e:
        st.error(f"Could not read the file. Make sure it is a valid Excel file. Error: {e}")
        return

    with st.expander("Preview: data loaded from your file"):
        st.dataframe(df, use_container_width=True)

    # ── Generate ──────────────────────────────────────────────────────────
    if st.button("Generate commentary", type="primary", use_container_width=True):
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

        # Store in session state so it persists across reruns
        st.session_state["var_commentary"]  = response
        st.session_state["var_model_used"]  = model_used
        st.session_state["var_company"]     = company
        st.session_state["var_period"]      = period
        st.session_state["var_audience"]    = audience

    # ── Display output (persists after generation) ─────────────────────────
    if "var_commentary" in st.session_state:
        response    = st.session_state["var_commentary"]
        model_used  = st.session_state["var_model_used"]
        company_out = st.session_state["var_company"]
        period_out  = st.session_state["var_period"]
        audience_out= st.session_state["var_audience"]

        st.markdown(
            '<hr style="border: none; border-top: 3px solid #d0d0d0; margin: 18px 0 20px 0;">',
            unsafe_allow_html=True,
        )

        # ── Collapsible AI output (open by default) ───────────────────────
        with st.expander(
            f"📋 AI Commentary — {company_out}, {period_out} ({audience_out})",
            expanded=True,
        ):
            st.markdown(response)
            st.caption(f"AI-generated content — model: {model_used}")

        # ── Download buttons ──────────────────────────────────────────────
        col_a, col_b = st.columns(2)

        with col_a:
            word_bytes = build_word_doc(response, company_out, period_out, audience_out)
            st.download_button(
                label="⬇ Download as Word document (.docx)",
                data=word_bytes,
                file_name=f"{company_out.lower().replace(' ', '_')}_{period_out.lower().replace(' ', '_')}_commentary.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        with col_b:
            st.download_button(
                label="⬇ Download as plain text (.txt)",
                data=response,
                file_name=f"{company_out.lower().replace(' ', '_')}_{period_out.lower().replace(' ', '_')}_commentary.txt",
                mime="text/plain",
                use_container_width=True,
            )

        st.caption(
            "Review all AI-generated content before distributing. "
            "Verify figures against source data."
        )
