"""Generate charts. Read-only — never modifies the input DataFrame."""

from io import BytesIO

import matplotlib
matplotlib.use("Agg")

import pandas as pd
import matplotlib.pyplot as plt
from src.logger_setup import setup_logger

logger = setup_logger(__name__)

plt.style.use("seaborn-v0_8-whitegrid")


# ── Numeric: histogram ───────────────────────────────────


def plot_histogram(df: pd.DataFrame, column: str, bins: int = 20) -> plt.Figure:
    """Return a histogram figure for a numeric column.

    Args:
        df: Source DataFrame (not modified).
        column: Column name to plot.
        bins: Number of histogram bins.

    Returns:
        Matplotlib Figure.
    """
    data = df[column].dropna()
    if data.empty:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, "No data to plot", ha="center", va="center")
        return fig

    fig, ax = plt.subplots(figsize=(8, 4))
    ax.hist(data, bins=bins, edgecolor="white", color="steelblue")
    ax.set_title(f"Distribution of {column}", fontsize=13)
    ax.set_xlabel(column)
    ax.set_ylabel("Frequency")
    plt.tight_layout()
    return fig


# ── Numeric: box plot ────────────────────────────────────


def plot_boxplot(df: pd.DataFrame, column: str) -> plt.Figure:
    """Return a box plot figure for a numeric column.

    Args:
        df: Source DataFrame (not modified).
        column: Column name to plot.

    Returns:
        Matplotlib Figure.
    """
    data = df[column].dropna()
    if data.empty:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, "No data to plot", ha="center", va="center")
        return fig

    fig, ax = plt.subplots(figsize=(8, 3))
    ax.boxplot(data, orientation="horizontal", patch_artist=True,
               boxprops={"facecolor": "steelblue", "alpha": 0.7})
    ax.set_title(f"Box plot of {column}", fontsize=13)
    ax.set_xlabel(column)
    plt.tight_layout()
    return fig


# ── Categorical: bar chart ───────────────────────────────


def plot_bar_chart(
    df: pd.DataFrame,
    column: str,
    top_n: int = 10,
) -> plt.Figure:
    """Return a horizontal bar chart for a categorical column.

    Args:
        df: Source DataFrame (not modified).
        column: Column name to plot.
        top_n: Show only the top N categories.

    Returns:
        Matplotlib Figure.
    """
    data = df[column].dropna()
    if data.empty:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, "No data to plot", ha="center", va="center")
        return fig

    value_counts = data.value_counts()
    unique_count = len(value_counts)

    if unique_count > 50:
        logger.warning(
            "Column '%s' has %d unique values; showing top %d",
            column, unique_count, top_n,
        )

    top = value_counts.head(top_n).sort_values(ascending=True)

    fig, ax = plt.subplots(figsize=(8, max(3, len(top) * 0.35)))
    ax.barh(top.index.astype(str), top.values, color="steelblue")
    ax.set_title(f"Top {min(top_n, len(top))} values in {column}", fontsize=13)
    ax.set_xlabel("Count")
    plt.tight_layout()
    return fig


# ── Date + numeric: line chart ───────────────────────────


def plot_line_chart(
    df: pd.DataFrame,
    date_column: str,
    value_column: str,
    agg: str = "sum",
) -> plt.Figure:
    """Return a line chart of a numeric column over time.

    Args:
        df: Source DataFrame (not modified).
        date_column: Column with datetime values.
        value_column: Numeric column to aggregate.
        agg: Aggregation function — 'sum' or 'mean'.

    Returns:
        Matplotlib Figure.
    """
    data = df[[date_column, value_column]].dropna()
    if data.empty:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, "No data to plot", ha="center", va="center")
        return fig

    date_col = data[date_column]
    if not pd.api.types.is_datetime64_any_dtype(date_col):
        try:
            date_col = pd.to_datetime(date_col, errors="coerce")
        except (ValueError, TypeError):
            pass
        data = data.copy()
        data[date_column] = date_col
        data = data.dropna(subset=[date_column])

    if data.empty:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, "No valid dates to plot", ha="center", va="center")
        return fig

    data = data.set_index(date_column)
    monthly = data[value_column].resample("ME").agg(agg)

    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(monthly.index, monthly.values, marker="o", linewidth=1.5, color="steelblue")
    ax.set_title(f"{value_column} by month ({agg})", fontsize=13)
    ax.set_ylabel(value_column)
    plt.xticks(rotation=45)
    plt.tight_layout()
    return fig


# ── Helper: convert Figure to PNG bytes ──────────────────


def figure_to_bytes(fig: plt.Figure, dpi: int = 150) -> bytes:
    """Save a matplotlib Figure to PNG bytes.

    Args:
        fig: A matplotlib Figure.
        dpi: Resolution in dots per inch.

    Returns:
        PNG image bytes.
    """
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf.getvalue()
