"""Generate charts. Read-only — never modifies the input DataFrame."""

from io import BytesIO

import matplotlib
matplotlib.use("Agg")

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter, StrMethodFormatter

from src.analyzer import friendly_column_name
from src.insights import is_currency_column
from src.logger_setup import setup_logger

logger = setup_logger(__name__)

PRIMARY_BLUE = "#2563EB"
SECONDARY_BLUE = "#60A5FA"
TEXT_COLOR = "#0F172A"
SECONDARY_TEXT = "#475569"
GRID_COLOR = "#E2E8F0"
SURFACE_COLOR = "#FFFFFF"


def _new_figure(*, height: float = 4.0) -> tuple[plt.Figure, plt.Axes]:
    fig, ax = plt.subplots(figsize=(8.5, height))
    fig.patch.set_facecolor(SURFACE_COLOR)
    ax.set_facecolor(SURFACE_COLOR)
    ax.tick_params(colors=SECONDARY_TEXT, labelsize=9)
    ax.xaxis.label.set_color(SECONDARY_TEXT)
    ax.yaxis.label.set_color(SECONDARY_TEXT)
    ax.title.set_color(TEXT_COLOR)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(GRID_COLOR)
    ax.spines["bottom"].set_color(GRID_COLOR)
    return fig, ax


def _empty_figure(message: str) -> plt.Figure:
    fig, ax = _new_figure(height=3.2)
    ax.text(
        0.5,
        0.5,
        message,
        ha="center",
        va="center",
        color=SECONDARY_TEXT,
        transform=ax.transAxes,
    )
    ax.set_axis_off()
    fig.tight_layout()
    return fig


def _set_business_axis_format(ax: plt.Axes, column: object, *, axis: str) -> None:
    formatter = (
        FuncFormatter(lambda value, _position: f"${value:,.0f}")
        if is_currency_column(column)
        else StrMethodFormatter("{x:,.0f}")
    )
    if axis == "x":
        ax.xaxis.set_major_formatter(formatter)
    else:
        ax.yaxis.set_major_formatter(formatter)


def _format_reference_value(value: float, column: object) -> str:
    if is_currency_column(column):
        return f"${value:,.2f}"
    return f"{value:,.2f}"


def _finish_chart(
    fig: plt.Figure,
    ax: plt.Axes,
    title: str,
    subtitle: str,
) -> plt.Figure:
    ax.set_title(title, loc="left", fontsize=14, fontweight=600, pad=20)
    ax.text(
        0,
        1.02,
        subtitle,
        transform=ax.transAxes,
        color=SECONDARY_TEXT,
        fontsize=9,
        va="bottom",
    )
    fig.tight_layout()
    return fig


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
        return _empty_figure("No numeric values are available.")

    fig, ax = _new_figure(height=3.5)
    ax.hist(data, bins=bins, edgecolor=SURFACE_COLOR, color=PRIMARY_BLUE)
    average = float(data.mean())
    median = float(data.median())
    ax.axvline(
        average,
        color=SECONDARY_BLUE,
        linewidth=2,
        linestyle="--",
        label=f"Average {_format_reference_value(average, column)}",
    )
    ax.axvline(
        median,
        color=TEXT_COLOR,
        linewidth=1.6,
        linestyle=":",
        label=f"Median {_format_reference_value(median, column)}",
    )
    ax.set_xlabel(friendly_column_name(column))
    ax.set_ylabel("Records")
    ax.grid(axis="y", color=GRID_COLOR, linewidth=0.8)
    ax.grid(axis="x", visible=False)
    _set_business_axis_format(ax, column, axis="x")
    ax.legend(frameon=False, fontsize=8, labelcolor=SECONDARY_TEXT)
    return _finish_chart(
        fig,
        ax,
        f"{friendly_column_name(column)} Distribution",
        "Each bar shows how many records fall within a value range.",
    )


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
        return _empty_figure("No numeric values are available.")

    fig, ax = _new_figure(height=2.8)
    ax.boxplot(data, orientation="horizontal", patch_artist=True,
               boxprops={"facecolor": SECONDARY_BLUE, "alpha": 0.75},
               medianprops={"color": TEXT_COLOR, "linewidth": 1.8},
               whiskerprops={"color": SECONDARY_TEXT},
               capprops={"color": SECONDARY_TEXT},
               flierprops={
                   "markerfacecolor": PRIMARY_BLUE,
                   "markeredgecolor": PRIMARY_BLUE,
                   "markersize": 5,
               })
    ax.set_xlabel(friendly_column_name(column))
    ax.set_yticks([])
    ax.grid(axis="x", color=GRID_COLOR, linewidth=0.8)
    ax.grid(axis="y", visible=False)
    _set_business_axis_format(ax, column, axis="x")
    return _finish_chart(
        fig,
        ax,
        f"{friendly_column_name(column)} Range and Review Points",
        "Points beyond the whiskers are statistically unusual, not automatically incorrect.",
    )


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
        return _empty_figure("No category values are available.")

    value_counts = data.value_counts()
    unique_count = len(value_counts)

    if unique_count > 50:
        logger.warning(
            "Column '%s' has %d unique values; showing top %d",
            column, unique_count, top_n,
        )

    top = value_counts.head(top_n).copy()
    remaining_count = int(value_counts.iloc[top_n:].sum())
    if remaining_count:
        top.loc["Other"] = remaining_count
    top = top.sort_values(ascending=True)

    fig, ax = _new_figure(height=min(4.55, max(2.8, len(top) * 0.34)))
    ax.barh(top.index.astype(str), top.values, color=PRIMARY_BLUE)
    ax.set_xlabel("Records")
    ax.grid(axis="x", color=GRID_COLOR, linewidth=0.8)
    ax.grid(axis="y", visible=False)
    ax.xaxis.set_major_formatter(StrMethodFormatter("{x:,.0f}"))
    return _finish_chart(
        fig,
        ax,
        f"Most Frequent {friendly_column_name(column)} Values",
        (
            f"Showing the top {min(top_n, unique_count):,} categories"
            + (" with the remainder grouped as Other." if remaining_count else ".")
        ),
    )


# ── Date + numeric: line chart ───────────────────────────


def plot_line_chart(
    df: pd.DataFrame,
    date_column: str,
    value_column: str,
    agg: str = "sum",
    group_by: str = "Month",
) -> plt.Figure:
    """Return a line chart of a numeric column over time.

    Args:
        df: Source DataFrame (not modified).
        date_column: Column with datetime values.
        value_column: Numeric column to aggregate.
        agg: Aggregation function — 'sum', 'mean', 'median', or 'count'.
        group_by: Calendar grouping — 'Day', 'Week', or 'Month'.

    Returns:
        Matplotlib Figure.
    """
    data = df[[date_column, value_column]].dropna()
    if data.empty:
        return _empty_figure("No complete date and measure pairs are available.")

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
        return _empty_figure("No valid dates are available.")

    frequency = {
        "Day": "D",
        "Week": "W-SUN",
        "Month": "MS",
    }.get(group_by)
    if frequency is None:
        raise ValueError("Group by must be Day, Week, or Month.")
    if agg not in {"sum", "mean", "median", "count"}:
        raise ValueError(
            "Aggregation must be sum, mean, median, or count."
        )

    data = data.set_index(date_column)
    grouped = data[value_column].resample(frequency).agg(agg)

    fig, ax = _new_figure(height=3.5)
    ax.plot(
        grouped.index,
        grouped.values,
        marker="o",
        markersize=4,
        linewidth=2,
        color=PRIMARY_BLUE,
    )
    ax.set_xlabel(friendly_column_name(date_column))
    ax.set_ylabel(
        "Records"
        if agg == "count"
        else friendly_column_name(value_column)
    )
    ax.grid(axis="y", color=GRID_COLOR, linewidth=0.8)
    ax.grid(axis="x", visible=False)
    if agg == "count":
        ax.yaxis.set_major_formatter(StrMethodFormatter("{x:,.0f}"))
    else:
        _set_business_axis_format(ax, value_column, axis="y")
    fig.autofmt_xdate(rotation=30)
    aggregation_label = {
        "sum": "Total",
        "mean": "Average",
        "median": "Median",
        "count": "Record count",
    }[agg]
    measure_label = friendly_column_name(value_column)
    normalized_measure = str(value_column).strip().lower().replace(" ", "_")
    if normalized_measure == "total" and agg == "sum":
        title = f"Sales Total by {group_by}"
    else:
        title = f"{aggregation_label} {measure_label} by {group_by}"
    return _finish_chart(
        fig,
        ax,
        title,
        (
            f"Using {friendly_column_name(date_column)}, grouped by "
            f"{group_by.lower()}; calculation: {aggregation_label.lower()}."
        ),
    )


def plot_date_counts(df: pd.DataFrame, column: str) -> plt.Figure:
    """Return a compact records-over-time chart for a date column."""
    parsed = pd.to_datetime(df[column], errors="coerce").dropna()
    if parsed.empty:
        return _empty_figure("No valid dates are available.")

    span_days = max(0, int((parsed.max() - parsed.min()).days))
    if span_days <= 45:
        frequency, period_label = "D", "day"
    elif span_days <= 730:
        frequency, period_label = "ME", "month"
    else:
        frequency, period_label = "YE", "year"

    counts = (
        pd.Series(1, index=parsed)
        .resample(frequency)
        .sum()
    )
    fig, ax = _new_figure(height=3.5)
    ax.plot(
        counts.index,
        counts.values,
        marker="o",
        markersize=4,
        linewidth=2,
        color=PRIMARY_BLUE,
    )
    ax.fill_between(
        counts.index,
        counts.values,
        color=SECONDARY_BLUE,
        alpha=0.18,
    )
    ax.set_xlabel("Date")
    ax.set_ylabel("Records")
    ax.grid(axis="y", color=GRID_COLOR, linewidth=0.8)
    ax.grid(axis="x", visible=False)
    ax.yaxis.set_major_formatter(StrMethodFormatter("{x:,.0f}"))
    fig.autofmt_xdate(rotation=30)
    return _finish_chart(
        fig,
        ax,
        f"{friendly_column_name(column)} Records Over Time",
        f"Record counts are grouped by {period_label}.",
    )


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
