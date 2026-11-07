"""Export consistent Excel and PowerPoint views of the same analysis."""

from dataclasses import dataclass
from textwrap import shorten

import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

from src.charts import MAX_DEFECTS, MAX_LINES
from src.kpis import observations, percent

NAVY = "16334A"
TEAL = "007D79"
MUTED = "465E71"
PALE = "EEF5F9"
WHITE = "FFFFFF"
COUNT = "#,##0"
RATE = "0.00%"
BASIS = "Rates assume distinct inspection cohorts. First-pass and rework history are unverified."


@dataclass(frozen=True)
class RunInfo:
    label: str
    files: int
    tables: int
    accepted: int
    rejected: int


def period(result):
    if result.daily.empty:
        return "No accepted dates"
    return f"{result.daily['date'].iloc[0]} to {result.daily['date'].iloc[-1]}"


def add_excel_table(ws, headers, rows, widths, formats=()):
    """Write a static report table with formatted numeric cells and filters."""
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 27
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for row_number, values in enumerate(rows, start=2):
        for column, value in enumerate(values, start=1):
            if pd.isna(value):
                value = None
            if isinstance(value, int) and abs(value) >= 10**15:
                value = str(value)
            if isinstance(value, str):
                value = ILLEGAL_CHARACTERS_RE.sub("�", value)
            cell = ws.cell(row_number, column, value)
            if isinstance(value, str):
                # Keep input labels literal, even if they begin with '='.
                cell.data_type = "s"
            if row_number % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=PALE)
            cell.font = Font(name="Calibri", size=11, color=NAVY)
            cell.alignment = Alignment(vertical="center")
            if formats and formats[column - 1]:
                cell.number_format = formats[column - 1]
        ws.row_dimensions[row_number].height = 18
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"


def log_rows(issues):
    for issue in issues:
        yield (issue.severity, issue.code, issue.source.file, issue.source.sheet,
               issue.source.row, issue.detail)


LOG_HEADERS = ("Severity", "Code", "File", "Sheet", "Row", "Detail")


def add_chart(ws, charts, key, anchor):
    if key in charts:
        image = Image(charts[key])
        image.width, image.height = 700, 315
        ws.add_image(image, anchor)


def add_note(ws, cell, text):
    ws[cell] = text
    ws[cell].font = Font(name="Calibri", size=11, italic=True, color=MUTED)


def write_excel(path, result, info, issues, charts):
    """Save a report workbook with auditable totals and a source-level log."""
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Executive Summary"
    metrics = [
        ("Dataset", info.label),
        ("Inspection period", period(result)),
        ("Inspected units", result.inspected),
        ("Defective units", result.defective),
        ("Defect rate", result.defect_rate if result.defect_rate is not None else "N/A"),
        ("Inspection yield", result.inspection_yield if result.inspection_yield is not None else "N/A"),
        ("Accepted rows", info.accepted),
        ("Excluded rows", info.rejected),
        ("Candidate files", info.files),
        ("Mapped tables", info.tables),
        ("Inspected units without line", result.unassigned_line_units),
        ("Pareto basis", result.pareto_status),
        ("Cohort basis", BASIS),
        *observations(result).items(),
    ]
    add_excel_table(summary, ("Measure", "Result"), metrics, (34, 106))
    summary["B2"].alignment = Alignment(wrap_text=True, vertical="center")
    summary.row_dimensions[2].height = 33
    for row in (4, 5, 8, 9, 10, 11, 12):
        summary[f"B{row}"].number_format = COUNT
    for row in (6, 7):
        summary[f"B{row}"].number_format = RATE
        summary[f"B{row}"].font = Font(name="Calibri", bold=True, size=13, color=TEAL)
    for row in range(13, summary.max_row + 1):
        summary[f"B{row}"].alignment = Alignment(wrap_text=True, vertical="center")
        summary.row_dimensions[row].height = 34

    daily = workbook.create_sheet("Daily Trend")
    line = workbook.create_sheet("By Line")
    for ws, frame, dimension in (
        (daily, result.daily, "Inspection date"), (line, result.by_line, "Production line")
    ):
        add_excel_table(ws, (dimension, "Inspected units", "Defective units",
                             "Defect rate", "Inspection yield"),
                        frame.itertuples(index=False, name=None),
                        (27, 20, 20, 17, 20), ("yyyy-mm-dd" if ws is daily else None, COUNT, COUNT, RATE, RATE))
        add_note(ws, "G1", "Blank rate cells indicate zero inspected units.")
    add_chart(daily, charts, "daily_trend", "G3")
    if result.by_line.empty:
        add_note(line, "G1", "No production line values were available for comparison.")
    elif len(result.by_line.dropna(subset=["defect_rate"])) > MAX_LINES:
        add_note(line, "G1", f"Chart shows the top {MAX_LINES} lines with defined rates.")
    add_chart(line, charts, "by_line", "G3")

    if not result.pareto.empty:
        pareto = workbook.create_sheet("Defect Pareto")
        add_excel_table(pareto, ("Defect type", "Defective units", "Share", "Cumulative share"),
                        result.pareto.itertuples(index=False, name=None),
                        (30, 20, 17, 22), (None, COUNT, RATE, RATE))
        add_note(pareto, "F1", result.pareto_status)
        if len(result.pareto) > MAX_DEFECTS:
            add_note(pareto, "F2", f"Chart shows the top {MAX_DEFECTS}. Shares use the full total.")
        add_chart(pareto, charts, "defect_pareto", "F3")

    log = workbook.create_sheet("Data Quality Log")
    add_excel_table(log, LOG_HEADERS, log_rows(issues), (13, 25, 37, 20, 12, 106))
    if not issues:
        log["A3"] = "No issues logged."
    for row in range(2, log.max_row + 1):
        log[f"F{row}"].alignment = Alignment(wrap_text=True, vertical="center")
        log.row_dimensions[row].height = 34
    workbook.save(path)


def text_box(slide, text, x, y, width, height, size, color=NAVY, bold=False):
    """Place editable text using the same typography throughout the deck."""
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = 0
    frame.margin_top = frame.margin_bottom = 0
    paragraph = frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.name = "Aptos"
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold
    paragraph.font.color.rgb = RGBColor.from_string(color)
    return shape


def add_slide(deck, title, info, *, cover=False):
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor.from_string(WHITE)
    text_box(slide, "PRODUCTION QUALITY", 0.68, 0.39, 7, 0.25, 12, TEAL, True)
    text_box(slide, title, 0.68, 0.78, 12, 0.76, 43 if cover else 34, NAVY, True)
    text_box(slide, info.label, 0.68, 7.04, 11.1, 0.3, 11, MUTED)
    text_box(slide, f"{len(deck.slides):02}", 12.14, 7.04, 0.5, 0.3, 11, MUTED)
    return slide


def chart_slide(deck, title, info, chart, heading, finding, note):
    slide = add_slide(deck, title, info)
    if chart:
        slide.shapes.add_picture(str(chart), Inches(0.55), Inches(1.8), width=Inches(9.0))
        text_box(slide, heading, 9.8, 1.93, 2.8, 0.62, 15, TEAL, True)
        text_box(slide, shorten(finding, width=143, placeholder="…"),
                 9.8, 2.59, 2.8, 2.16, 19, NAVY, True)
        text_box(slide, note, 9.8, 4.83, 2.8, 1.9, 15, MUTED)
    else:
        text_box(slide, note, 0.7, 2.07, 10.6, 1.45, 23, MUTED)
    return slide


def write_powerpoint(path, result, info, charts):
    """Save a four- or five-slide briefing with figures from the workbook."""
    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
    deck.core_properties.title = "Production Quality Report"
    slide = add_slide(deck, "Production quality report", info, cover=True)
    text_box(slide, period(result), 0.7, 1.77, 10, 0.5, 21, MUTED)
    for index, (title, value) in enumerate((
        ("Defect rate", percent(result.defect_rate)),
        ("Inspection yield", percent(result.inspection_yield)),
        ("Inspected units", f"{result.inspected:,}"),
        ("Defective units", f"{result.defective:,}"),
    )):
        x, y = 0.72 + (index % 2) * 5.92, 2.78 + (index // 2) * 1.63
        text_box(slide, title, x, y, 5.5, 0.32, 17, MUTED)
        text_box(slide, value, x, y + 0.38, 5.5, 0.76,
                 33 if len(value) <= 15 else 26, TEAL, True)

    findings = observations(result)
    daily_note = ("Each point uses the units inspected on that date. Gaps indicate no defined rate."
                  if "daily_trend" in charts else "The accepted records have zero inspected units.")
    chart_slide(deck, "Daily defect rate", info, charts.get("daily_trend"),
                "DAILY PEAK", findings.get("Daily peak", ""), daily_note)

    line_note = "Line rates use each line's inspected units. Lines with zero volume have no rate."
    if len(result.by_line.dropna(subset=["defect_rate"])) > MAX_LINES:
        line_note = f"Showing the {MAX_LINES} highest defined rates. See Excel for all lines."
    if "by_line" not in charts:
        line_note = "No production line has a defined rate in the accepted records."
    chart_slide(deck, "Defect rate by line", info, charts.get("by_line"),
                "HIGHEST OBSERVED RATE", findings.get("Highest line rate", ""), line_note)

    if "defect_pareto" in charts:
        pareto_note = "Each defective unit counts once under its assigned defect type."
        if len(result.pareto) > MAX_DEFECTS:
            pareto_note = f"Showing the top {MAX_DEFECTS} categories. Shares use the full total."
        chart_slide(deck, "Defect Pareto", info, charts["defect_pareto"],
                    "LARGEST CATEGORY", findings["Largest defect category"], pareto_note)

    slide = add_slide(deck, "Observations and limits", info)
    readout = [
        ("Accepted data", f"{info.accepted:,} rows kept and {info.rejected:,} rows excluded. "
                          "See the Data Quality Log for details."),
        ("Rate interpretation", BASIS),
        ("Defect categories", result.pareto_status),
    ]
    if result.unassigned_line_units:
        readout.append(("Line coverage", f"{result.unassigned_line_units:,} inspected units have no line "
                                         "and appear in the overall and daily rates."))
    for index, (heading, body) in enumerate(readout):
        y = 1.78 + index * (1.2 if len(readout) == 4 else 1.46)
        text_box(slide, heading, 0.71, y, 11.4, 0.35, 16, TEAL, True)
        text_box(slide, body, 0.71, y + 0.42, 11.6, 0.82, 20)
    for slide in deck.slides:
        slide.notes_slide.notes_text_frame.text = f"Dataset: {info.label}\nSee DATASET_SOURCE.md for sample provenance."
    deck.save(path)
