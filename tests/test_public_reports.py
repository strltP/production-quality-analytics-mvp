"""Compare the generated workbook and presentation to independent source totals."""

from datetime import datetime
from pathlib import Path
import subprocess
import sys
from zipfile import ZipFile

import pandas as pd
import pytest
from openpyxl import load_workbook
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from prepare_pareto_demo import prepare


ROOT = Path(__file__).resolve().parents[1]


def aggregate_source(source, dimension):
    return source.groupby(dimension)[["inspected", "defective"]].sum().to_dict("index")


def check_breakdown(worksheet, expected):
    output = {
        (key.date().isoformat() if isinstance(key, datetime) else key):
        (inspected, defective, defect_rate, inspection_yield)
        for key, inspected, defective, defect_rate, inspection_yield, *_ in
        worksheet.iter_rows(min_row=2, max_col=5, values_only=True)
        if key is not None
    }
    assert output.keys() == expected.keys()
    for key, (inspected, defective, rate, yield_rate) in output.items():
        assert (inspected, defective) == (expected[key]["inspected"], expected[key]["defective"])
        assert rate == pytest.approx(defective / inspected)
        assert yield_rate == pytest.approx(1 - defective / inspected)


def check_media(bundle, deck, expected):
    charts = {path.read_bytes() for path in (bundle / "charts").glob("*.png")}
    assert len(charts) == expected
    with ZipFile(bundle / "quality_report.xlsx") as workbook_zip:
        assert workbook_zip.testzip() is None
        assert charts == {workbook_zip.read(name) for name in workbook_zip.namelist()
                          if name.startswith("xl/media/")}
    with ZipFile(bundle / "quality_report.pptx") as deck_zip:
        assert deck_zip.testzip() is None
    assert charts == {shape.image.blob for slide in deck.slides for shape in slide.shapes
                      if shape.shape_type == MSO_SHAPE_TYPE.PICTURE}


@pytest.mark.parametrize("case,expected_counts", [
    ("workbook", (42_259, 5_062)), ("pareto", (60_000, 3_334)),
])
def test_public_sources_end_to_end(case, expected_counts, tmp_path):
    output = tmp_path / "reports with spaces"
    if case == "workbook":
        inputs = ROOT / "data/input"
        source = pd.read_excel(inputs / "Quality_Inspection_Dataset.xlsx").rename(columns={
            "Inspection_Date": "date", "Line_Number": "production_line",
            "Units_Checked": "inspected", "Defective_Units": "defective",
        })
        options = []
    else:
        raw = ROOT / "data/raw/quality_inspections.csv"
        source = pd.read_csv(raw)
        source["inspected"] = 1
        source["defective"] = source["status"].eq("Fail").astype(int)
        inputs = tmp_path / "prepared input"
        normalized = inputs / "public_synthetic_inspections.csv"
        assert prepare(raw, normalized) == len(source)
        expected = ROOT / "data/pareto_input/public_synthetic_inspections.csv"
        assert normalized.read_text(encoding="utf-8-sig") == expected.read_text(encoding="utf-8-sig")
        options = ["--pareto"]
    process = subprocess.run(
        [sys.executable, str(ROOT / "main.py"), "--input", str(inputs), "--output", str(output), *options],
        cwd=ROOT, capture_output=True, text=True, check=False,
    )
    assert process.returncode == 0, process.stdout + process.stderr
    (bundle,) = output.iterdir()
    inspected, defective = map(int, source[["inspected", "defective"]].sum())
    assert (inspected, defective) == expected_counts
    clean = pd.read_csv(bundle / "clean_inspections.csv")
    assert len(clean) == len(source)
    assert tuple(clean[["inspected_units", "defective_units"]].sum()) == (inspected, defective)
    workbook = load_workbook(bundle / "quality_report.xlsx", read_only=True, data_only=True)
    summary = {key: value for key, value in workbook["Executive Summary"].values}
    assert summary["Inspected units"] == inspected
    assert summary["Defective units"] == defective
    assert summary["Defect rate"] == pytest.approx(defective / inspected)
    assert summary["Inspection yield"] == pytest.approx(1 - defective / inspected)
    assert summary["Accepted rows"] == len(source)
    assert summary["Excluded rows"] == 0
    check_breakdown(workbook["Daily Trend"], aggregate_source(source, "date"))
    check_breakdown(workbook["By Line"], aggregate_source(source, "production_line"))
    if case == "pareto":
        categories = source.loc[source["defective"].eq(1), "defect_type"].value_counts()
        rows = list(workbook["Defect Pareto"].iter_rows(min_row=2, max_col=4, values_only=True))
        assert [(kind, count) for kind, count, *_ in rows] == list(categories.items())
        assert [share for _, _, share, _ in rows] == pytest.approx((categories / defective).tolist())
        assert [running for _, _, _, running in rows] == pytest.approx(
            (categories.cumsum() / defective).tolist())
    else:
        assert "Defect Pareto" not in workbook.sheetnames
        assert "BATCH_ID_REUSED" in {row[1] for row in workbook["Data Quality Log"].values}
    workbook.close()

    deck = Presentation(bundle / "quality_report.pptx")
    assert len(deck.slides) == (5 if case == "pareto" else 4)
    title_text = {shape.text for shape in deck.slides[0].shapes if shape.has_text_frame}
    assert {f"{inspected:,}", f"{defective:,}", f"{defective / inspected:.2%}",
            f"{1 - defective / inspected:.2%}"} <= title_text
    check_media(bundle, deck, 3 if case == "pareto" else 2)
