"""
Shared Excel parsing utilities for FinanceAI Lab.

All modules that accept Excel uploads use these functions.
openpyxl reads the raw file. pandas structures it into a dataframe
(a table with rows and columns) for easy processing.
"""

import pandas as pd
from io import BytesIO


def parse_variance_sheet(uploaded_file) -> pd.DataFrame:
    """
    Reads a budget vs. actual Excel file and returns a clean dataframe.

    Expects columns: Line Item, Budget, Actual, Variance, Variance %
    Returns the dataframe so modules can format it however they need.
    """
    df = pd.read_excel(BytesIO(uploaded_file.read()), sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")  # remove completely empty rows
    return df


def parse_generic_sheet(uploaded_file, sheet_index: int = 0) -> pd.DataFrame:
    """
    Generic parser for any single-sheet Excel file.
    Returns a clean dataframe with whitespace stripped from column names.
    """
    df = pd.read_excel(BytesIO(uploaded_file.read()), sheet_name=sheet_index)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df


def extract_file_metadata(uploaded_file) -> dict:
    """
    Reads the Excel file and tries to extract company name and reporting period
    from the content — sheet name, title cell (A1), or any text in the first
    three rows. Returns a dict with 'company' and 'period' keys.
    Falls back to empty strings if nothing is found.
    """
    import re
    from io import BytesIO
    import openpyxl

    result = {"company": "", "period": ""}

    try:
        data = uploaded_file.read()
        uploaded_file.seek(0)  # reset so the file can be read again by the parser
        wb = openpyxl.load_workbook(BytesIO(data), read_only=True, data_only=True)
        ws = wb.active

        # Collect text candidates: sheet name + first 3 rows of column A
        candidates = [ws.title or ""]
        for row in ws.iter_rows(min_row=1, max_row=3, max_col=3, values_only=True):
            for cell in row:
                if cell and isinstance(cell, str):
                    candidates.append(cell)

        full_text = " | ".join(candidates)

        # Period patterns: Q1 2026, H1 2026, FY2026, Jan 2026, January 2026, YTD 2026
        period_match = re.search(
            r'(Q[1-4]\s*\d{4}|H[1-2]\s*\d{4}|FY\s*\d{4}|YTD\s*\d{4}|'
            r'(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|'
            r'Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|'
            r'Dec(?:ember)?)\s*\d{4})',
            full_text, re.IGNORECASE
        )
        if period_match:
            result["period"] = period_match.group(0).strip()

        # Company name: look for text before a dash or "Budget" or "Variance" in A1
        if candidates[1:]:  # first content row
            first_cell = candidates[1]
            # Strip period text and common finance keywords to isolate company name
            company_text = re.sub(
                r'(Q[1-4]\s*\d{4}|H[1-2]\s*\d{4}|FY\s*\d{4}|YTD\s*\d{4}|'
                r'budget|actual|variance|vs\.?|all figures.*|simulated.*)',
                '', first_cell, flags=re.IGNORECASE
            )
            # Split on common separators and take the first meaningful chunk
            for sep in ['—', '-', '|', ':']:
                if sep in company_text:
                    company_text = company_text.split(sep)[0]
                    break
            company_text = company_text.strip().strip('—-|:').strip()
            if len(company_text) > 2:
                result["company"] = company_text

    except Exception:
        pass  # if anything fails, caller uses fallback values

    return result


def dataframe_to_text(df: pd.DataFrame) -> str:
    """
    Converts a dataframe into a plain-text table string.
    This is what gets sent to the AI model — models read text, not Excel files.
    """
    return df.to_string(index=False)
