"""Shared utility functions used across modules."""

from pathlib import Path

from config.settings import ALLOWED_EXTENSIONS


def is_allowed_file(filename: str) -> bool:
    """Check whether a filename has an allowed extension.

    Args:
        filename: Original name of the uploaded file.

    Returns:
        True if the extension is in ALLOWED_EXTENSIONS.
    """
    ext = Path(filename).suffix.lower()
    return ext in ALLOWED_EXTENSIONS


def format_bytes(size_bytes: int) -> str:
    """Convert a byte count to a human-readable string.

    Args:
        size_bytes: File size in bytes.

    Returns:
        e.g. "4.2 MB", "890.0 KB"
    """
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"
