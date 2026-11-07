"""Validate counts and dates, standardize labels and resolve duplicate records."""

import re
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

import pandas as pd

from src.ingest import Intake, Issue, Source


COLUMNS = [
    "date", "inspected_units", "defective_units", "production_line", "defect_type",
    "inspection_id", "batch_id", "product", "source_file", "source_sheet", "source_row",
]
LABELS = ("production_line", "defect_type", "product")
MISSING_LABELS = {"none", "unknown", "n/a", "na"}
# Limits keep source quantities within Excel's exact 15-digit integer range.
MAX_COUNT = 999_999_999_999_999


def parse_date(value, date_order: str) -> date:
    """Accept native Excel dates, ISO dates, or an explicitly ordered slash date."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    value = str(value).strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        return date.fromisoformat(value)
    if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", value):
        pattern = "%d/%m/%Y" if date_order == "dmy" else "%m/%d/%Y"
        return datetime.strptime(value, pattern).date()
    raise ValueError(f"Unsupported inspection date {value!r}")


def parse_count(value) -> int:
    """Parse nonnegative whole quantities without floating-point rounding."""
    if isinstance(value, bool):
        raise ValueError(f"Invalid unit count {value!r}")
    try:
        number = Decimal(str(value).strip())
    except InvalidOperation as error:
        raise ValueError(f"Invalid unit count {value!r}") from error
    if not number.is_finite() or number < 0 or number > MAX_COUNT:
        raise ValueError(f"Unit count outside 0 to {MAX_COUNT}: {value!r}")
    if number != number.to_integral_value():
        raise ValueError(f"Unit count must be whole: {value!r}")
    return int(number)


def clean_text(value):
    """Collapse internal whitespace and convert empty cells to None."""
    return (" ".join(str(value).split()) or None) if value is not None else None


def clean_row(data, date_order):
    """Parse typed data and reject rows whose numerator exceeds its denominator."""
    clean = {field: clean_text(data.get(field)) for field in COLUMNS[3:8]}
    for field in ("date", "inspected_units", "defective_units"):
        try:
            clean[field] = (parse_date(data.get(field), date_order) if field == "date"
                            else parse_count(data.get(field)))
        except ValueError as error:
            raise ValueError(f"{field}: {error}") from error
    if clean["defective_units"] > clean["inspected_units"]:
        raise ValueError("defective_units exceeds inspected_units")
    return clean


def record_key(clean):
    """Use the inspection ID, the full batch key, or an exact row fingerprint."""
    if clean["inspection_id"]:
        return ("inspection_id", clean["inspection_id"])
    if all(clean[field] for field in ("batch_id", "production_line", "product")):
        return ("batch", clean["batch_id"], clean["date"],
                clean["production_line"], clean["product"])
    return ("row", *(clean[field] for field in COLUMNS[:8]))


def validate(intake: Intake, date_order: str = "dmy"):
    """Return an accepted DataFrame, a full issue log and a rejected-row count."""
    issues = list(intake.issues)
    rejected = intake.rejected_rows
    candidates = defaultdict(list)
    label_spellings = {field: {} for field in LABELS}

    for source, raw in intake.rows:
        try:
            clean = clean_row(raw, date_order)
        except ValueError as error:
            issues.append(Issue(source, "ERROR", "INVALID_ROW", str(error)))
            rejected += 1
            continue
        for field in LABELS:
            value = clean[field]
            if value:
                if field in {"production_line", "defect_type"} and value.casefold() in MISSING_LABELS:
                    clean[field] = None
                    continue
                canonical = label_spellings[field].setdefault(value.casefold(), value)
                if canonical != raw.get(field):
                    issues.append(Issue(source, "INFO", "LABEL_NORMALIZED",
                                        f"{field}: {raw[field]!r} -> {canonical!r}"))
                clean[field] = canonical
        candidates[record_key(clean)].append((source, clean))

    accepted = []
    for entries in candidates.values():
        first_source, first_row = entries[0]
        if any(row != first_row for _, row in entries[1:]):
            for source, _ in entries:
                issues.append(Issue(source, "ERROR", "CONFLICTING_KEY",
                                    "The same inspection key has conflicting values. All versions were excluded."))
                rejected += 1
            continue
        for source, _ in entries[1:]:
            issues.append(Issue(source, "WARNING", "DUPLICATE_ROW",
                                f"Retained {first_source.file} / {first_source.sheet or 'CSV'} / row {first_source.row}."))
            rejected += 1
        accepted.append({**first_row, "source_file": first_source.file,
                         "source_sheet": first_source.sheet, "source_row": first_source.row})
        if first_row["inspected_units"] == 0:
            issues.append(Issue(first_source, "WARNING", "ZERO_INSPECTED",
                                "Zero units contribute no volume. Rates are N/A if a whole group has zero units."))

    if accepted:
        batches = Counter(row["batch_id"] for row in accepted if row["batch_id"])
        repeated = sum(count > 1 for count in batches.values())
        if repeated:
            issues.append(Issue(Source(""), "WARNING", "BATCH_ID_REUSED",
                                f"{repeated:,} batch IDs recur across retained inspection keys."))
        missing_line = sum(not row["production_line"] for row in accepted)
        if missing_line:
            issues.append(Issue(Source(""), "WARNING", "LINE_MISSING",
                                f"{missing_line:,} accepted rows contribute to overall and daily KPIs only."))

        labels_on_pass = sum(row["defect_type"] is not None and row["defective_units"] == 0
                             for row in accepted)
        if labels_on_pass:
            issues.append(Issue(Source(""), "INFO", "TYPE_WITH_ZERO_DEFECTS",
                                f"{labels_on_pass:,} zero-defect rows have a type label. "
                                "They contribute zero units to any category."))

    return pd.DataFrame(accepted, columns=COLUMNS), issues, rejected
