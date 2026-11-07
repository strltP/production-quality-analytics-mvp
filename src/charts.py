"""Render headless, reusable charts for both reports."""

from datetime import timedelta
from textwrap import shorten

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator, PercentFormatter

TEAL = "#007D79"
NAVY = "#16334A"
MUTED = "#465E71"
GRID = "#DDE6EE"
ORANGE = "#BD6024"
MAX_LINES = 12
MAX_DEFECTS = 10


def figure():
    """Use shared chart styling and a wide layout for Office reports."""
    fig, ax = plt.subplots(figsize=(10, 4.5), layout="constrained")
    fig.set_facecolor("white")
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["bottom", "left"]].set_color(GRID)
    ax.tick_params(colors=MUTED, labelsize=10, length=0)
    ax.grid(axis="y", color=GRID, linewidth=0.7)
    ax.set_axisbelow(True)
    return fig, ax


def label(value):
    """Constrain chart tick labels. Reports retain the full category names."""
    return shorten(str(value)[:40], width=32, placeholder="…") or str(value)[:30] + "…"


def save(fig, output, name):
    path = output / f"{name}.png"
    fig.savefig(path, dpi=155)
    plt.close(fig)
    return path


def plot_daily(result, output):
    """Leave a visible gap when a calendar day has no observed denominator."""
    fig, ax = figure()
    dates, rates = [], []
    for row in result.daily.itertuples(index=False):
        if dates and (row.date - dates[-1]).days > 1:
            dates.append(dates[-1] + timedelta(days=1))
            rates.append(float("nan"))
        dates.append(row.date)
        rates.append(row.defect_rate)
    ax.plot(dates, rates, color=TEAL, linewidth=1.7, marker="o", markersize=3)
    ax.axhline(result.defect_rate, color=ORANGE, linestyle="--", linewidth=1.3,
               label=f"Overall rate {result.defect_rate:.2%}")
    locator = mdates.AutoDateLocator(minticks=3, maxticks=8)
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    ax.yaxis.set_major_formatter(PercentFormatter(1))
    ax.set_ylim(bottom=0)
    ax.set_ylabel("Defect rate", color=NAVY, labelpad=10)
    ax.set_xlabel("Inspection date", color=NAVY, labelpad=10)
    ax.legend(loc="upper right", frameon=True, edgecolor="white", labelcolor=MUTED)
    return save(fig, output, "daily_trend")


def plot_lines(result, output):
    """Rank lines with defined rates and display their exact rounded rates."""
    rows = result.by_line.dropna(subset=["defect_rate"]).head(MAX_LINES)
    fig, ax = figure()
    ax.grid(False, axis="y")
    ax.grid(True, axis="x", color=GRID, linewidth=0.7)
    positions = range(len(rows))
    bars = ax.barh(positions, rows["defect_rate"], color=TEAL, height=0.58)
    ax.bar_label(bars, labels=[f"{value:.2%}" for value in rows["defect_rate"]],
                 padding=5, color=NAVY, fontsize=10)
    ax.set_yticks(positions, [label(value) for value in rows["production_line"]])
    ax.invert_yaxis()
    ax.set_xlim(0, max(float(rows["defect_rate"].max()) * 1.24, 0.01))
    ax.xaxis.set_major_formatter(PercentFormatter(1))
    ax.set_xlabel("Defect rate", color=NAVY, labelpad=10)
    return save(fig, output, "by_line")


def plot_pareto(result, output):
    """Plot sorted unit counts against cumulative share of all classified units."""
    rows = result.pareto.head(MAX_DEFECTS)
    fig, ax = figure()
    positions = range(len(rows))
    bars = ax.bar(positions, rows["defective_units"], color=TEAL, width=0.62)
    ax.bar_label(bars, labels=[f"{value:,}" for value in rows["defective_units"]],
                 color=NAVY, fontsize=9, padding=4)
    ax.set_xticks(positions, [label(value) for value in rows["defect_type"]],
                  rotation=23, ha="right")
    ax.set_ylim(0, float(rows["defective_units"].max()) * 1.2)
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax.set_ylabel("Defective units", color=NAVY, labelpad=10)
    cumulative = ax.twinx()
    cumulative.spines[["top", "left", "bottom"]].set_visible(False)
    cumulative.spines["right"].set_color(GRID)
    cumulative.tick_params(colors=MUTED, labelsize=10, length=0)
    cumulative.plot(positions, rows["cumulative_share"], color=ORANGE, marker="o")
    cumulative.set_ylim(0, 1.08)
    cumulative.set_yticks([0, .25, .50, .75, 1])
    cumulative.yaxis.set_major_formatter(PercentFormatter(1))
    cumulative.set_ylabel("Cumulative share", color=ORANGE, labelpad=10)
    return save(fig, output, "defect_pareto")


def write_charts(result, output):
    """Write only supported visuals and return paths keyed by chart name."""
    output.mkdir()
    charts = {}
    if result.defect_rate is not None:
        charts["daily_trend"] = plot_daily(result, output)
    if result.by_line["defect_rate"].notna().any():
        charts["by_line"] = plot_lines(result, output)
    if not result.pareto.empty:
        charts["defect_pareto"] = plot_pareto(result, output)
    return charts
