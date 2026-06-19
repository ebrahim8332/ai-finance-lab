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


def dataframe_to_text(df: pd.DataFrame) -> str:
    """
    Converts a dataframe into a plain-text table string.
    This is what gets sent to the AI model — models read text, not Excel files.
    """
    return df.to_string(index=False)
