"""Run the local production quality reporting pipeline."""

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from uuid import uuid4

from src.charts import write_charts
from src.ingest import Issue, Source, ingest
from src.kpis import analyze, percent
from src.reports import LOG_HEADERS, RunInfo, log_rows, write_excel, write_powerpoint
from src.validation import validate


def run(input_dir, output_dir, date_order="dmy", pareto=False, label="User-supplied inspection records"):
    """Publish a complete report bundle or a diagnostic-only run directory."""
    intake = ingest(input_dir)
    clean, issues, rejected = validate(intake, date_order)
    if clean.empty:
        issues.append(Issue(Source(""), "ERROR", "NO_VALID_RECORDS",
                            "No usable inspection records. The KPI reports were skipped."))
    output_dir.mkdir(parents=True, exist_ok=True)
    run_name = f"run-{datetime.now():%Y%m%d-%H%M%S}-{uuid4().hex[:8]}"
    destination = output_dir / run_name

    # Publishing the directory in one operation avoids mixed results on retries.
    with TemporaryDirectory(prefix=".quality-", dir=output_dir) as temporary:
        stage = Path(temporary) / "bundle"
        stage.mkdir()
        with (stage / "data_quality_log.csv").open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            writer.writerow(LOG_HEADERS)
            writer.writerows(log_rows(issues))
        if not clean.empty:
            clean.to_csv(stage / "clean_inspections.csv", index=False, encoding="utf-8-sig")
            result = analyze(clean, pareto)
            info = RunInfo(label, intake.files, intake.tables, len(clean), rejected)
            charts = write_charts(result, stage / "charts")
            write_excel(stage / "quality_report.xlsx", result, info, issues, charts)
            write_powerpoint(stage / "quality_report.pptx", result, info, charts)
        stage.rename(destination)

    if clean.empty:
        print(f"No usable records. See {destination / 'data_quality_log.csv'}", file=sys.stderr)
        return 2
    print(f"Accepted {len(clean):,} rows, excluded {rejected:,} rows, logged {len(issues):,} issues.")
    print(f"Inspected {result.inspected:,} | Defective {result.defective:,} | "
          f"Defect rate {percent(result.defect_rate)} | Inspection yield {percent(result.inspection_yield)}")
    print(f"Pareto: {result.pareto_status}")
    print(f"Output: {destination}")
    return 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=Path("data/input"), help="Folder of CSV/XLSX files")
    parser.add_argument("--output", type=Path, default=Path("reports"), help="Parent folder for run bundles")
    parser.add_argument("--date-order", choices=("dmy", "mdy"), default="dmy",
                        help="Interpretation of slash dates. ISO dates are always YYYY-MM-DD.")
    parser.add_argument("--pareto", action="store_true",
                        help="Confirm each defective unit has exactly one category in the input.")
    parser.add_argument("--dataset-label", default="User-supplied inspection records",
                        help="Short source description printed in the reports")
    args = parser.parse_args()
    label = " ".join(args.dataset_label.split())
    if not label or len(label) > 100:
        parser.error("--dataset-label must be 1 to 100 characters")
    try:
        return run(args.input, args.output, args.date_order, args.pareto, label)
    except OSError as error:
        print(f"Unable to access files or publish a report: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
