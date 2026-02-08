"""
Excel handler module for reading/writing National IDs and syndicate data
"""

import re
import openpyxl
from openpyxl import Workbook, load_workbook
from pathlib import Path
from typing import List, Dict, Optional
import pandas as pd


def _clean_id_value(value: object) -> str:
    """Convert a cell value to a cleaned digit-only national ID string."""
    if pd.isna(value):
        return ""
    s = str(value).strip()
    # Remove common float .0 artifacts
    if s.endswith('.0'):
        s = s[:-2]
    # Handle scientific notation by using Decimal-like conversion via float->int when safe
    # But safest is to extract digits only
    digits = re.sub(r"\D", "", s)
    return digits


def _find_id_column(df: pd.DataFrame, preferred: Optional[str] = None) -> Optional[str]:
    """Find the best matching column name for national ID in the dataframe.

    Checks the preferred name first, then falls back to common English and Arabic alternatives.
    Returns the matched column name or None if not found.
    """
    if preferred and preferred in df.columns:
        return preferred

    # Common candidate names (English first, then Arabic variants)
    candidates = [
        "National ID",
        "NationalID",
        "national_id",
        "الرقم القومى",
        "الرقم القومي",
        "الرقم القومى (National ID)",
        "الرقم القومي (National ID)",
    ]

    # Normalize column names for fuzzy matching
    normalized = {col: re.sub(r"\s+", "", col).lower() for col in df.columns}

    for cand in candidates:
        cand_norm = re.sub(r"\s+", "", cand).lower()
        for col, col_norm in normalized.items():
            if cand_norm == col_norm:
                return col

    # As a last resort, look for columns containing ID-like keywords but avoid name columns
    id_keywords = ('رقم', 'قومي', 'قومى', 'id', 'national')
    name_keywords = ('name', 'اسم')
    for col in df.columns:
        low = str(col).lower()
        if any(nk in low for nk in name_keywords):
            continue
        if any(k in low for k in id_keywords):
            return col

    return None



def read_national_ids_from_excel(file_path: str, column_name: Optional[str] = None) -> List[str]:
    """
    Read national IDs from an Excel file.
    
    :param file_path: Path to the Excel file
    :param column_name: Name of the column containing national IDs
    :return: List of national ID strings
    """
    try:
        # Read as strings where possible to preserve ID formatting
        df = pd.read_excel(file_path, dtype=str)

        id_col = _find_id_column(df, column_name)
        if id_col is None:
            raise ValueError(
                f"Could not find a National ID column. Available columns: {', '.join(df.columns)}"
            )

        series = df[id_col].dropna().astype(str)

        cleaned_ids = []
        for v in series:
            cid = _clean_id_value(v)
            if cid:
                cleaned_ids.append(cid)

        return cleaned_ids
    except Exception as e:
        raise Exception(f"Error reading Excel file: {str(e)}")


def write_results_to_excel(results: List[Dict], output_path: str):
    """
    Write syndicate lookup results to an Excel file.
    
    :param results: List of result dictionaries from scraper
    :param output_path: Path where the Excel file will be saved
    """
    try:
        df = pd.DataFrame(results)
        df.to_excel(output_path, index=False, engine='openpyxl')
        return True
    except Exception as e:
        raise Exception(f"Error writing to Excel file: {str(e)}")


def append_syndicate_to_excel(file_path: str, output_path: str = None,
                               id_column: Optional[str] = None,
                               syndicate_column: str = "Syndicate",
                               name_column: str = "Name"):
    """
    Read an Excel file with national IDs, look up syndicates and names, and write results.
    
    :param file_path: Input Excel file path
    :param output_path: Output Excel file path (if None, overwrites input)
    :param id_column: Column name for national IDs
    :param syndicate_column: Column name for syndicate results
    :param name_column: Column name for name results
    :return: Path to the output file
    """
    from .scraper import get_engineer_syndicate_safe

    if output_path is None:
        output_path = file_path

    # Read the Excel file as strings
    df = pd.read_excel(file_path, dtype=str)

    id_col = _find_id_column(df, id_column)
    if id_col is None:
        raise ValueError(f"Could not find National ID column in file. Columns: {', '.join(df.columns)}")

    # Process each national ID
    syndicates = []
    names = []
    for national_id in df[id_col].dropna().astype(str):
        clean_id = _clean_id_value(national_id)
        result = get_engineer_syndicate_safe(clean_id)
        syndicates.append(result.get('syndicate', result.get('error', 'Error')))
        names.append(result.get('name', ''))

    # Add results to dataframe
    df[name_column] = names
    df[syndicate_column] = syndicates

    # Write to Excel
    df.to_excel(output_path, index=False, engine='openpyxl')

    return output_path


def create_sample_excel(output_path: str, sample_ids: List[str] = None):
    """
    Create a sample Excel file with national IDs for testing.
    
    :param output_path: Path where the sample file will be saved
    :param sample_ids: Optional list of sample national IDs
    """
    if sample_ids is None:
        sample_ids = [
            "29501011234567",
            "30012251234568",
            "28803151234569"
        ]
    
    df = pd.DataFrame({"National ID": sample_ids})
    df.to_excel(output_path, index=False, engine='openpyxl')
    
    return output_path