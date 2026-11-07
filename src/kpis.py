"""Compute weighted inspection KPIs and optional single-label Pareto data."""

from dataclasses import dataclass

import pandas as pd

COUNT_FIELDS = ["inspected_units", "defective_units"]


@dataclass
class Analysis:
    inspected: int
    defective: int
    defect_rate: float | None
    inspection_yield: float | None
    daily: pd.DataFrame
    by_line: pd.DataFrame
    pareto: pd.DataFrame
    pareto_status: str
    unassigned_line_units: int


def rate(numerator, denominator):
    """Return a fraction or None for an undefined ratio."""
    return numerator / denominator if denominator else None


def summarize(rows, group):
    """Sum quantities first, then compute ratios for each observed group."""
    counts = rows.astype({field: object for field in COUNT_FIELDS})
    grouped = counts.groupby(group, sort=True)[COUNT_FIELDS].sum().reset_index()
    grouped["defect_rate"] = [
        rate(defective, inspected)
        for inspected, defective in grouped[COUNT_FIELDS].itertuples(index=False, name=None)
    ]
    grouped["inspection_yield"] = 1 - pd.to_numeric(grouped["defect_rate"])
    return grouped


def defect_pareto(rows, confirmed):
    """Count defective units with one verified label per unit, if supported."""
    fields = ["defect_type", "defective_units", "share", "cumulative_share"]
    empty = pd.DataFrame(columns=fields)
    if not confirmed:
        return empty, "Skipped. Per-unit allocation to a single defect type was not confirmed."
    failed = rows[rows["defective_units"] > 0]
    if failed.empty:
        return empty, "Skipped. No defective units in the accepted records."
    if failed["defect_type"].isna().any():
        return empty, "Skipped. At least one positive defective count lacks a defect type."
    pareto = (failed.astype({"defective_units": object})
              .groupby("defect_type", as_index=False)["defective_units"].sum()
              .sort_values(["defective_units", "defect_type"], ascending=[False, True])
              .reset_index(drop=True))
    counts = pareto["defective_units"]
    total = sum(int(value) for value in counts)
    pareto["share"] = counts.map(lambda value: value / total)
    pareto["cumulative_share"] = counts.cumsum().map(lambda value: value / total)
    return pareto, "Each defective unit contributes to one verified category."


def analyze(rows: pd.DataFrame, pareto_confirmed: bool = False) -> Analysis:
    """Build one set of metrics for both reports and all charts."""
    inspected = sum(int(value) for value in rows["inspected_units"])
    defective = sum(int(value) for value in rows["defective_units"])
    defect_rate = rate(defective, inspected)
    known_lines = rows["production_line"].notna()
    by_line = summarize(rows[known_lines], "production_line").sort_values(
        ["defect_rate", "production_line"], ascending=[False, True], na_position="last"
    )
    pareto, status = defect_pareto(rows, pareto_confirmed)
    unassigned = sum(int(value) for value in rows.loc[~known_lines, "inspected_units"])
    return Analysis(inspected, defective, defect_rate,
                    None if defect_rate is None else 1 - defect_rate,
                    summarize(rows, "date"), by_line, pareto, status, unassigned)


def percent(value):
    """Use the same percentage display throughout the narrative."""
    return "N/A" if pd.isna(value) else f"{value:.2%}"


def observations(result):
    """Describe the highest observed rates with numerators and denominators."""
    findings = {}
    daily = result.daily.dropna(subset=["defect_rate"])
    if not daily.empty:
        peak = daily.sort_values(["defect_rate", "date"], ascending=[False, True]).iloc[0]
        findings["Daily peak"] = (f"{peak['date']} had the highest observed daily rate at "
                                  f"{percent(peak.defect_rate)} "
                                  f"({peak.defective_units:,} of {peak.inspected_units:,} units).")
    lines = result.by_line.dropna(subset=["defect_rate"])
    if not lines.empty:
        top = lines.iloc[0]
        findings["Highest line rate"] = (f"{top.production_line} recorded {percent(top.defect_rate)} "
                                         f"({top.defective_units:,} of {top.inspected_units:,} units).")
    if not result.pareto.empty:
        top = result.pareto.iloc[0]
        findings["Largest defect category"] = (f"{top.defect_type} accounted for {top.defective_units:,} "
                                               f"defective units ({percent(top.share)}).")
    return findings
