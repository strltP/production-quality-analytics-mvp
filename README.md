# Production Quality Analytics & Automated Reporting Tool

A local Python CLI that validates inspection records and exports a reusable
Excel workbook, a PowerPoint briefing, and an audit trail. The two included
report bundles use separate public datasets. See [DATASET_SOURCE.md](DATASET_SOURCE.md)
for attribution, provenance, and the unit of analysis.

## Run on Windows 11 with VS Code

1. Extract the ZIP. Open the inner `production-quality-analytics-mvp` folder
   in VS Code using **File → Open Folder**.
2. Install Python 3.11 or 3.12 and the Microsoft Python extension for VS Code.
   Open **Terminal → New Terminal** and choose PowerShell.
3. Create the virtual environment and install the pinned dependencies.

```powershell
py -3.11 -m venv .venv
& ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt
```

If Python 3.12 is installed, use `py -3.12` in the first command. If the Python
launcher is unavailable, check `python --version` and use
`python -m venv .venv`. The commands call the virtual environment interpreter
directly, so PowerShell activation is optional. In VS Code, select
`.venv\Scripts\python.exe` with **Python: Select Interpreter**.

### 1. Run the public workbook

```powershell
& ".\.venv\Scripts\python.exe" main.py --input data/input --output reports --dataset-label "Kajal Panigrahi public sample (factory provenance unverified)"
```

Expected result is **42,259 inspected units**, **5,062 defective units**,
**11.98% defect rate** and **88.02% inspection yield**. The supplied workbook
has 350 accepted rows. The program logs repeated batch IDs and defect labels
on zero-defect rows. This source does not establish a per-unit defect-type
allocation, so the Pareto is skipped.

### 2. Run the public synthetic Pareto demonstration

```powershell
& ".\.venv\Scripts\python.exe" main.py --input data/pareto_input --output reports --pareto --dataset-label "Oussama Elayeb public synthetic inspections (simulated)"
```

Expected result is **60,000 simulated inspections**, **3,334 defective units**,
**5.56% defect rate** and **94.44% inspection yield**. The Pareto has six
categories. To recreate the prepared CSV from the included raw file, run
`& ".\.venv\Scripts\python.exe" prepare_pareto_demo.py`.

### 3. Run the tests

```powershell
& ".\.venv\Scripts\python.exe" -m pytest -q
```

The dependency installation typically needs internet access. Once dependencies
are installed, both sample runs and the tests work offline. Read
[TEST_RESULTS.md](TEST_RESULTS.md) for the tested environment and results.

## Inputs

Put exports for one inspection population in a folder and pass its path as
`--input`. Use comma-separated UTF-8 CSV with or without BOM, or XLSX with
headers in the first row of each worksheet. The CLI scans files directly in
the folder and validates each sheet independently. Excel lock files are ignored.
An input table must contain the three required fields below. A mapping error
skips that table and records its filename and worksheet in the audit log.

| Canonical header | Required | Meaning and accepted aliases |
| --- | --- | --- |
| `date` | Yes | Inspection date. Alias `inspection_date`. ISO `YYYY-MM-DD`, a native Excel date, or a slash date. |
| `inspected_units` | Yes | Inspected units for the record. Aliases `units_inspected`, `units_checked`, `units_check`. |
| `defective_units` | Yes | Distinct defective units in the same record. Aliases `units_defective`, `rejected_units`, `units_rejected`. |
| `production_line` | For line analysis | Actual line identifier. Aliases `line_number`, `line`. |
| `defect_type` | For Pareto | Single verified label for all defective units in the row. |
| `inspection_id` | Preferred | Unique inspection record ID across the input folder. Alias `record_id`. |
| `batch_id` | Optional | Batch identifier that can recur across independent inspection cohorts. |
| `product` | Optional | Product identifier. Aliases `product_name`, `product_model`. |

Header matching ignores case and treats spaces, hyphens, and underscores as
equivalent separators. Column order does not matter. Extra fields are left in
the original input. Two headers mapping to the same field make a table ambiguous.
Counts must be finite, whole, nonnegative numbers of at most 15 digits. An
Excel date may contain a time, which the daily rollup drops. Slash dates use
DMY by default. Pass `--date-order mdy` for month-first files. Normalize exports
to ISO if day-first and month-first dates coexist. An ambiguous slash date
follows the selected order. XLSX formula cells need saved cached values.

### Validation and duplicate policy

- Reject unusable required values, negative or fractional counts, and
  `defective_units > inspected_units`. A ragged CSV row is logged and skipped.
  A file-level parse failure skips that file.
- Keep a valid `0/0` record for audit and report its group rate as undefined
  when the group has no volume. A day with no defined rate creates a gap.
- Clean optional labels by trimming, collapsing spaces, and unifying case to
  the first observed spelling. For line and defect labels, `none`, `unknown`,
  `n/a` and `na` represent missing values. Missing line rows stay in the
  overall and daily totals. They are outside the line comparison.
- Deduplicate across files using `inspection_id` first. When absent, the complete
  tuple `(batch_id, date, production_line, product)` is used when available.
  Otherwise use the full mapped and cleaned row as the fingerprint.
  Keep the first exact repeat in deterministic file, sheet, row order.
  If valid rows share an ID or full batch key and disagree on mapped values,
  exclude all versions of that key. Provide a consistent key across exports
  and distinct IDs for legitimate repeats with identical measurements.
- Source filename, sheet and row trace each accepted record and row-level
  issue back to input. CSV row references use the starting physical line.
  Excluded-row counts cover mapped row rejections and deduplication. A whole
  unreadable or unmappable table is tracked as a separate file or table issue.

## KPI definitions

- **Defect rate (%)** = 100 × sum of defective units ÷ sum of inspected units.
- **Inspection yield (%)** = 100 × (sum of inspected units − sum of defective
  units) ÷ sum of inspected units. Both measures use the same accepted cohort.
- **Pareto** sums positive defective-unit counts by verified single category,
  sorts descending and calculates cumulative share against the full category
  total. Supply `--pareto` only when each defective unit in every positive
  row can be assigned to exactly one type. A missing label in any positive
  row suppresses the chart. A label on a zero-defect row contributes zero.

The system aggregates counts before calculating any rate. For example, one
defect in 100 inspections and five in 10 gives 6 ÷ 110 = 5.45%.
The workbook stores rates as numeric fractions formatted as percentages.
Counts longer than 15 digits are written as text in Excel to preserve precision.

## Outputs

The CLI prints a unique `reports/run-...` folder. Each successful run includes:

| File | Contents |
| --- | --- |
| `quality_report.xlsx` | Executive Summary, Daily Trend, By Line, optional Defect Pareto and Data Quality Log. |
| `quality_report.pptx` | Four or five slides with shared chart images and editable text. |
| `charts/*.png` | Daily trend, line comparison and Pareto where supported. |
| `clean_inspections.csv` | Accepted records with standardized fields and row provenance. |
| `data_quality_log.csv` | Issues encountered at file, sheet, row or run level. |

Packaged examples are in `reports/public-workbook/` and `reports/pareto-demo/`.
Each run publishes a new directory after generation. A run with no usable rows
publishes only the issue log and exits with code 2. A file-system failure exits
with code 1. A completed report returns 0 even if some source records or tables
were excluded. Check the log before using the results. If opening the CSVs in
Excel, import text fields as Text to preserve leading zeros and literal labels.

## Code walkthrough

1. `main.py` handles arguments, runs the steps in order and publishes results.
2. `src/ingest.py` reads CSV and XLSX, maps headers and captures source positions.
3. `src/validation.py` parses dates and quantities, normalizes categories and resolves duplicates.
4. `src/kpis.py` aggregates counts and builds the daily, line and Pareto tables.
5. `src/charts.py` draws PNGs from those tables using Matplotlib.
6. `src/reports.py` writes Excel with openpyxl and PowerPoint with python-pptx.
7. `tests/` covers edge cases and compares fresh outputs with both upstream datasets.

To trace a number, open **Executive Summary**, find its detail in **Daily Trend**
or **By Line**, and follow the cleaned CSV source positions to the original file.
To study the calculation in a VS Code notebook, set the working directory to
the project root and use the project interpreter:

```python
from pathlib import Path
from src.ingest import ingest
from src.validation import validate
from src.kpis import analyze

clean, issues, rejected = validate(ingest(Path("data/input")))
result = analyze(clean)
display(result.by_line)
```

A notebook kernel requires Jupyter support in VS Code and `ipykernel` in the
environment. The CLI has no notebook dependency.

## Scope and limitations

- Each accepted row must represent an additive inspection cohort. The raw
  workbook does not provide unit-level IDs to prove cohorts never overlap.
- The program measures inspection yield. It does not have a first-pass or
  rework identifier. It does not infer a cause from a high line or daily rate.
- The Pareto implementation expects one assignable category per defective
  unit. A multi-defect event log would need a separate, documented event grain.
- Different sources with the same compatible schema may still describe
  different populations. Supply one cohort per run and harmonize record IDs
  before combining files. The bundled examples stay separate.
- The reports are static snapshots. To refresh charts after editing input,
  rerun the CLI. Charts show up to 12 lines or 10 categories. Excel tables
  retain all groups.
- The pipeline works in memory and targets small to moderate exports. Source
  formulas, exotic spreadsheet layouts, and Windows Office rendering are
  outside the verified scope. See the test record for platform coverage.
   you need to calculate first-pass yield?
3. How do you distinguish a duplicate inspection from a reused batch ID?
4. Why is the public workbook's Pareto omitted, and what makes the simulated
   source's Pareto defensible?
5. How did you independently verify that the exported Excel and PowerPoint
   match the input data and one another?
