"""Read, validate, and describe uploaded Excel and CSV files."""

from pathlib import Path

import pandas as pd
from streamlit.runtime.uploaded_file_manager import UploadedFile

from config.settings import MAX_FILE_SIZE_MB, FALLBACK_ENCODING
from src.logger_setup import setup_logger

logger = setup_logger(__name__)


class FileReadError(Exception):
    """Raised when a file cannot be read or is malformed."""


class FileSizeError(Exception):
    """Raised when a file exceeds the maximum allowed size."""


def read_uploaded_file(uploaded_file: UploadedFile) -> pd.DataFrame:
    """Read a single uploaded file into a DataFrame.

    Supports .xlsx and .csv files. Attempts UTF-8 first for CSVs,
    then falls back to latin-1.

    Args:
        uploaded_file: A Streamlit UploadedFile object.

    Returns:
        DataFrame containing the file contents.

    Raises:
        FileSizeError: If the file exceeds MAX_FILE_SIZE_MB.
        FileReadError: If the file cannot be parsed.
    """
    _check_file_size(uploaded_file)

    ext = Path(uploaded_file.name).suffix.lower()
    logger.info("Reading file: %s (%.1f KB)", uploaded_file.name,
                 uploaded_file.size / 1024)

    try:
        if ext == ".csv":
            return _read_csv(uploaded_file)
        return _read_excel(uploaded_file)
    except FileReadError:
        raise
    except Exception as exc:
        logger.exception("Failed to read %s", uploaded_file.name)
        raise FileReadError(
            f"Could not read '{uploaded_file.name}': {exc}"
        ) from exc


def _check_file_size(uploaded_file: UploadedFile) -> None:
    """Check that the uploaded file is under the size limit."""
    size_mb = uploaded_file.size / (1024 * 1024)
    if size_mb > MAX_FILE_SIZE_MB:
        raise FileSizeError(
            f"'{uploaded_file.name}' is {size_mb:.1f} MB, "
            f"which exceeds the {MAX_FILE_SIZE_MB} MB limit."
        )


def _read_csv(uploaded_file: UploadedFile) -> pd.DataFrame:
    """Try to read a CSV with UTF-8, fall back to latin-1."""
    try:
        return pd.read_csv(uploaded_file, encoding="utf-8")
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        logger.warning("UTF-8 decode failed for %s, trying latin-1",
                       uploaded_file.name)
        df = pd.read_csv(uploaded_file, encoding=FALLBACK_ENCODING)
        return df


def _read_excel(uploaded_file: UploadedFile) -> pd.DataFrame:
    """Read all sheets from an Excel file and return the first sheet.

    In a later milestone we'll let the user pick a specific sheet.
    """
    excel_data = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
    sheet_name = list(excel_data.keys())[0]
    logger.info("Using sheet '%s' from %s", sheet_name, uploaded_file.name)
    return excel_data[sheet_name]


def describe_dataframe(df: pd.DataFrame, name: str) -> dict:
    """Return a summary dictionary for a DataFrame.

    Args:
        df: The DataFrame to describe.
        name: Display name (usually the filename).

    Returns:
        Dict with keys: name, rows, columns, column_list, missing_cells,
        duplicate_rows, dtypes.
    """
    return {
        "name": name,
        "rows": len(df),
        "columns": len(df.columns),
        "column_list": list(df.columns),
        "missing_cells": int(df.isna().sum().sum()),
        "duplicate_rows": int(df.duplicated().sum()),
        "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
    }
