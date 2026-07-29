"""Export a DataFrame to Excel or CSV bytes.

Each function returns (bytes, suggested_filename). No files are written
to disk — everything stays in memory for Streamlit's download buttons.
"""

import re
from datetime import date
from io import BytesIO

import openpyxl
import pandas as pd

from src.logger_setup import setup_logger

logger = setup_logger(__name__)

EXCEL_SHEET_NAME = "Cleaned Data"
MAX_COL_WIDTH = 40


class ExportError(Exception):
    """Raised when export fails."""


def _safe_filename(base: str, extension: str) -> str:
    """Build a safe filename by removing characters unsafe for file systems.

    Args:
        base: The user-chosen or default base name.
        extension: File extension including the dot (e.g. '.xlsx').

    Returns:
        A sanitised filename with the extension appended.
    """
    safe = re.sub(r'[<>:"/\\|?*]', "_", base.strip())
    safe = re.sub(r"\s+", "_", safe)
    if not safe:
        safe = "exported_data"
    return f"{safe}{extension}"


def export_to_excel(
    df: pd.DataFrame,
    filename_base: str = "",
) -> tuple[bytes, str]:
    """Export a DataFrame to an Excel .xlsx byte stream.

    Args:
        df: DataFrame to export (not modified).
        filename_base: Desired base name (without extension). Auto-generates
                       from current date if empty.

    Returns:
        Tuple of (file_bytes, suggested_filename).

    Raises:
        ExportError: If the DataFrame is empty or export fails.
    """
    if df.empty:
        raise ExportError(
            "Cannot export an empty DataFrame. "
            "Exporting an empty file (headers only) is blocked because "
            "it would not be useful — please verify your cleaning step."
        )

    if not filename_base:
        filename_base = f"cleaned_dataset_{date.today().isoformat()}"

    filename = _safe_filename(filename_base, ".xlsx")

    try:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=EXCEL_SHEET_NAME, index=False)

            from openpyxl.styles import Font

            workbook = writer.book
            worksheet = workbook[EXCEL_SHEET_NAME]

            # Bold headers
            header_font = Font(bold=True)
            for col_idx in range(len(df.columns)):
                cell = worksheet.cell(row=1, column=col_idx + 1)
                cell.font = header_font

            # Freeze top row
            worksheet.freeze_panes = "A2"

            # Auto-adjust column widths (capped)
            for col_idx, col_name in enumerate(df.columns):
                max_len = max(
                    df[col_name].fillna("").astype(str).map(len).max()
                    if len(df) > 0 else len(str(col_name)),
                    len(str(col_name)),
                )
                col_letter = openpyxl.utils.get_column_letter(col_idx + 1)
                worksheet.column_dimensions[col_letter].width = min(max_len + 2, MAX_COL_WIDTH)

        buffer.seek(0)
        result = buffer.getvalue()

        logger.info(
            "Exported Excel: %s (%.1f KB, %d rows, %d cols)",
            filename, len(result) / 1024, len(df), len(df.columns),
        )
        return result, filename

    except ExportError:
        raise
    except Exception as exc:
        logger.exception("Excel export failed")
        raise ExportError(f"Excel export failed: {exc}") from exc


def export_to_csv(
    df: pd.DataFrame,
    filename_base: str = "",
) -> tuple[bytes, str]:
    """Export a DataFrame to a CSV byte stream (UTF-8 with BOM).

    UTF-8 BOM ensures Excel on Windows opens the file with correct encoding.

    Args:
        df: DataFrame to export (not modified).
        filename_base: Desired base name (without extension). Auto-generates
                       from current date if empty.

    Returns:
        Tuple of (file_bytes, suggested_filename).

    Raises:
        ExportError: If the DataFrame is empty or export fails.
    """
    if df.empty:
        raise ExportError(
            "Cannot export an empty DataFrame. "
            "Exporting an empty file (headers only) is blocked because "
            "it would not be useful — please verify your cleaning step."
        )

    if not filename_base:
        filename_base = f"cleaned_dataset_{date.today().isoformat()}"

    filename = _safe_filename(filename_base, ".csv")

    try:
        buffer = BytesIO()
        df.to_csv(buffer, index=False, encoding="utf-8-sig")
        buffer.seek(0)
        result = buffer.getvalue()

        logger.info(
            "Exported CSV: %s (%.1f KB, %d rows, %d cols)",
            filename, len(result) / 1024, len(df), len(df.columns),
        )
        return result, filename

    except ExportError:
        raise
    except Exception as exc:
        logger.exception("CSV export failed")
        raise ExportError(f"CSV export failed: {exc}") from exc
