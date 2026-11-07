"""Read input tables and map known header names to the inspection schema."""

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


ALIASES = {
    "date": ("date", "inspection_date"),
    "inspected_units": ("inspected_units", "units_inspected", "units_checked", "units_check"),
    "defective_units": ("defective_units", "units_defective", "rejected_units", "units_rejected"),
    "production_line": ("production_line", "line_number", "line"),
    "defect_type": ("defect_type",),
    "inspection_id": ("inspection_id", "record_id"),
    "batch_id": ("batch_id",),
    "product": ("product", "product_name", "product_model"),
}
REQUIRED = {"date", "inspected_units", "defective_units"}


@dataclass(frozen=True)
class Source:
    file: str
    sheet: str = ""
    row: int | None = None


@dataclass(frozen=True)
class Issue:
    source: Source
    severity: str
    code: str
    detail: str


@dataclass
class Intake:
    rows: list[tuple[Source, dict]]
    issues: list[Issue]
    files: int
    tables: int
    rejected_rows: int


def normalize_header(value):
    """Treat spaces, underscores, dashes and casing as header formatting."""
    return re.sub(r"[\s_-]+", "_", str(value).strip().casefold()).strip("_")


def csv_rows(path):
    """Return one table with physical starting line numbers and raw strings."""
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle, strict=True)
        header = next(reader, [])
        rows = []
        while True:
            line = reader.line_num + 1
            row = next(reader, None)
            if row is None:
                break
            rows.append((line, row))
    return [("", header, rows)]


def excel_tables(path):
    """Read each XLSX worksheet as a separate candidate table."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in workbook.worksheets:
            values = sheet.values
            header = list(next(values, ()))
            while header and header[-1] is None:
                header.pop()
            yield sheet.title, header, enumerate(values, start=2)
    finally:
        workbook.close()


def column_mapping(header):
    """Reject missing required fields or two headers mapping to one field."""
    normalized = [normalize_header(cell) for cell in header]
    mapping = {}
    for field, aliases in ALIASES.items():
        positions = [i for i, label in enumerate(normalized) if label in aliases]
        if len(positions) > 1:
            raise ValueError(f"Ambiguous columns for {field}")
        if positions:
            mapping[field] = positions[0]
    missing = REQUIRED - mapping.keys()
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(sorted(missing))}")
    return mapping


def ingest(folder: Path) -> Intake:
    """Collect compatible inspection rows and audit skipped files and sheets."""
    result = Intake([], [], 0, 0, 0)
    if not folder.is_dir():
        result.issues.append(Issue(Source(folder.name), "ERROR", "INPUT_FOLDER",
                                   "Input folder does not exist or is not a directory."))
        return result

    for path in sorted(folder.iterdir(), key=lambda item: (item.name.casefold(), item.name)):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in {".csv", ".xlsx"}:
            result.issues.append(Issue(Source(path.name), "INFO", "IGNORED_FILE",
                                       "Supported file extensions are .csv and .xlsx."))
            continue
        result.files += 1
        # Build a local result first. A parse failure discards that entire file.
        file_rows, file_issues = [], []
        file_tables = rejected = 0
        try:
            tables = csv_rows(path) if path.suffix.lower() == ".csv" else excel_tables(path)
            for sheet, header, rows in tables:
                table = Source(path.name, sheet)
                try:
                    mapping = column_mapping(header)
                except ValueError as error:
                    file_issues.append(Issue(table, "ERROR", "UNMAPPABLE_TABLE", str(error)))
                    continue
                file_tables += 1
                count = 0
                for row_number, raw in rows:
                    if not any(value is not None and str(value).strip() for value in raw):
                        continue
                    count += 1
                    source = Source(path.name, sheet, row_number)
                    if not sheet and len(raw) != len(header):
                        rejected += 1
                        file_issues.append(Issue(source, "ERROR", "CSV_WIDTH",
                                                 f"Expected {len(header)} fields, found {len(raw)}."))
                        continue
                    data = {field: raw[i] for field, i in mapping.items()}
                    file_rows.append((source, data))
                if not count:
                    file_issues.append(Issue(table, "WARNING", "NO_DATA_ROWS",
                                             "Mapped headers but no nonblank data rows."))
        except (OSError, UnicodeError, csv.Error, BadZipFile, InvalidFileException,
                ValueError, KeyError, IndexError, SyntaxError) as error:
            result.issues.append(Issue(Source(path.name), "ERROR", "READ_ERROR", str(error)))
            continue
        result.rows.extend(file_rows)
        result.issues.extend(file_issues)
        result.tables += file_tables
        result.rejected_rows += rejected

    if result.files == 0:
        result.issues.append(Issue(Source(folder.name), "ERROR", "NO_INPUT_FILES",
                                   "No CSV or XLSX input files were found."))
    return result
