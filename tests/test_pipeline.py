"""Exercise the normal CLI pipeline and edge cases with small input tables."""

import csv
from datetime import date, datetime

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from main import run
from src.ingest import ingest
from src.kpis import analyze
from src.validation import parse_date, validate


HEADER = ["Inspection Date", "Units Checked", "Defective Units",
          "Production Line", "Defect Type", "Inspection ID", "Batch ID", "Product Name"]


def write_csv(path, rows, header=HEADER):
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


def record(day="2025-01-01", inspected=10, defective=2, line="Line A",
           kind="Scratch", identifier="I1", batch="", product="Model X"):
    return day, inspected, defective, line, kind, identifier, batch, product


def load_test_records(folder, date_order="dmy"):
    return validate(ingest(folder), date_order)


def test_empty_input_publishes_diagnostic(tmp_path):
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    output = tmp_path / "reports"
    assert run(inputs, output) == 2
    (bundle,) = output.iterdir()
    assert [path.name for path in bundle.iterdir()] == ["data_quality_log.csv"]
    log = (bundle / "data_quality_log.csv").read_text(encoding="utf-8-sig")
    assert "NO_INPUT_FILES" in log
    assert "NO_VALID_RECORDS" in log


@pytest.mark.parametrize("header, expected", [
    (["Date", "Units Checked", "Unrelated"], "Missing required columns"),
    (["Date", "Inspection Date", "Units Checked", "Defective Units"], "Ambiguous columns"),
])
def test_unmappable_headers_are_logged(tmp_path, header, expected):
    write_csv(tmp_path / "unrelated.csv", [], header)
    clean, issues, rejected = load_test_records(tmp_path)
    assert clean.empty and rejected == 0
    assert [(issue.source.file, issue.code) for issue in issues] == [
        ("unrelated.csv", "UNMAPPABLE_TABLE")
    ]
    assert expected in issues[0].detail


def test_mixed_files_sheets_provenance_and_duplicates(tmp_path):
    write_csv(tmp_path / "01.csv", [record(identifier="0001")])
    workbook = Workbook()
    first = workbook.active
    first.title = "Inspections"
    first.append(["date", "  UNITS-CHECKED  ", "defective_units",
                  "production_line", "defect_type", "record_id", "batch_id", "product"])
    first.append(record(day=datetime(2025, 1, 1, 12), line="  LINE   A ", identifier="0001"))
    first.append(record(day=datetime(2025, 1, 2), identifier="0002"))
    unrelated = workbook.create_sheet("Inventory")
    unrelated.append(["product", "quantity"])
    unrelated.append(["Model X", 100])
    workbook.save(tmp_path / "02.xlsx")

    clean, issues, rejected = load_test_records(tmp_path)
    assert clean["inspection_id"].tolist() == ["0001", "0002"]
    assert list(clean[["source_file", "source_sheet", "source_row"]].itertuples(index=False, name=None)) == [
        ("01.csv", "", 2), ("02.xlsx", "Inspections", 3)
    ]
    assert rejected == 1
    assert {issue.code for issue in issues} == {"LABEL_NORMALIZED", "DUPLICATE_ROW", "UNMAPPABLE_TABLE"}


def test_conflicting_explicit_ids_exclude_all_versions(tmp_path):
    write_csv(tmp_path / "one.csv", [record(), record(inspected=19), record(identifier="I2")])
    clean, issues, rejected = load_test_records(tmp_path)
    assert clean["inspection_id"].tolist() == ["I2"]
    assert rejected == 2
    assert [issue.source.row for issue in issues if issue.code == "CONFLICTING_KEY"] == [2, 3]


def test_repeated_batch_id_is_not_a_unique_inspection_key(tmp_path):
    write_csv(tmp_path / "one.csv", [
        record(identifier="", batch="B1", product="Model X"),
        record(identifier="", batch="B1", product="Model Y"),
        record(identifier="", batch="B1", product="Model Y"),
        record(identifier="", batch="B1", day="2025-01-02"),
    ])
    clean, issues, rejected = load_test_records(tmp_path)
    assert len(clean) == 3 and rejected == 1
    assert {"DUPLICATE_ROW", "BATCH_ID_REUSED"} <= {issue.code for issue in issues}


@pytest.mark.parametrize("inspected, defective", [
    (-1, 0), (3, 4), (8, -1), (0, 1), (10.2, 1), ("", 1), ("NaN", 1), ("x", 1),
])
def test_invalid_quantities_are_excluded(tmp_path, inspected, defective):
    write_csv(tmp_path / "bad.csv", [record(inspected=inspected, defective=defective)])
    clean, issues, rejected = load_test_records(tmp_path)
    assert clean.empty and rejected == 1
    assert (issues[0].source.row, issues[0].code) == (2, "INVALID_ROW")


def test_dates_follow_explicit_slash_order_and_invalid_dates_are_excluded(tmp_path):
    write_csv(tmp_path / "dates.csv", [
        record(day="2025-02-01"), record(day="02/01/2025", identifier="I2"),
        record(day="31/01/2025", identifier="I3"),
        record(day="02/30/2025", identifier="I4"),
        record(day="2025-02-30", identifier="I5"),
    ])
    dmy, _, dmy_rejected = load_test_records(tmp_path, "dmy")
    mdy, _, mdy_rejected = load_test_records(tmp_path, "mdy")
    assert dmy["date"].tolist() == [date(2025, 2, 1), date(2025, 1, 2), date(2025, 1, 31)]
    assert mdy["date"].tolist() == [date(2025, 2, 1), date(2025, 2, 1)]
    assert (dmy_rejected, mdy_rejected) == (2, 3)
    assert parse_date(datetime(2025, 2, 1, 14), "dmy") == date(2025, 2, 1)


def test_zero_denominator_is_blank_and_does_not_fake_zero_rate(tmp_path):
    write_csv(tmp_path / "zero.csv", [record(inspected=0, defective=0, kind="")])
    clean, issues, _ = load_test_records(tmp_path)
    result = analyze(clean, pareto_confirmed=True)
    assert "ZERO_INSPECTED" in {issue.code for issue in issues}
    assert result.defect_rate is None and result.inspection_yield is None
    assert result.daily["defect_rate"].isna().all()
    assert result.by_line["defect_rate"].isna().all()
    assert result.pareto.empty
    output = tmp_path / "out"
    assert run(tmp_path, output) == 0
    (bundle,) = output.iterdir()
    with (bundle / "quality_report.xlsx").open("rb") as handle:
        book = load_workbook(handle, read_only=True)
        assert book["Executive Summary"]["B6"].value == "N/A"
        assert book["Daily Trend"]["D2"].value is None
        book.close()
    assert list((bundle / "charts").iterdir()) == []


def test_weighted_rates_and_missing_line_reconciliation(tmp_path):
    write_csv(tmp_path / "weighted.csv", [
        record(inspected=100, defective=1),
        record(inspected=10, defective=5, identifier="I2"),
        record(inspected=4, defective=3, identifier="I3", line="  n/a "),
        record(inspected=0, defective=0, identifier="I4", line="Line B"),
    ])
    clean, issues, _ = load_test_records(tmp_path)
    result = analyze(clean)
    assert result.defect_rate == pytest.approx(9 / 114)
    assert result.inspection_yield == pytest.approx(105 / 114)
    assert result.daily["defect_rate"].iloc[0] == pytest.approx(9 / 114)
    assert result.by_line["defect_rate"].iloc[0] == pytest.approx(6 / 110)
    assert pd.isna(result.by_line["defect_rate"].iloc[1])
    assert result.unassigned_line_units == 4
    assert "LINE_MISSING" in {issue.code for issue in issues}


def test_pareto_requires_explicit_contract_and_complete_labels(tmp_path):
    rows = [record(defective=4), record(defective=2, identifier="I2", kind="Dirt"),
            record(defective=0, identifier="I3", kind="Dirt")]
    write_csv(tmp_path / "pareto.csv", rows)
    clean, _, _ = load_test_records(tmp_path)
    assert analyze(clean).pareto.empty
    result = analyze(clean, pareto_confirmed=True)
    assert result.pareto["defective_units"].tolist() == [4, 2]
    assert result.pareto["cumulative_share"].tolist() == pytest.approx([4 / 6, 1])
    clean.loc[0, "defect_type"] = None
    incomplete = analyze(clean, pareto_confirmed=True)
    assert incomplete.pareto.empty and "lacks a defect type" in incomplete.pareto_status


def test_ragged_csv_row_logged_without_shifting_the_next_row(tmp_path):
    write_csv(tmp_path / "ragged.csv", [record(), ["2025-01-01", 10], record(identifier="I2")])
    clean, issues, rejected = load_test_records(tmp_path)
    assert clean["source_row"].tolist() == [2, 4]
    assert rejected == 1
    assert [(issue.code, issue.source.row) for issue in issues] == [("CSV_WIDTH", 3)]


def test_malformed_xlsx_does_not_block_other_files(tmp_path):
    (tmp_path / "broken.xlsx").write_text("Not a workbook", encoding="utf-8")
    write_csv(tmp_path / "valid.csv", [record()])
    clean, issues, rejected = load_test_records(tmp_path)
    assert len(clean) == 1 and rejected == 0
    assert [issue.code for issue in issues] == ["READ_ERROR"]
