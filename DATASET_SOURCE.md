# Dataset sources and record grain

Inspected on 2026-11-07. Original files are retained byte-for-byte. The two
sources describe separate populations and are run independently.

## Public inspection workbook

**Source:** Kajal Panigrahi, *Quality Inspection Dashboard*.
[Workbook](https://github.com/KAJALPANIGRAHI1997/Quality-Inspection-Dashboard/blob/main/Quality_Inspection_Dataset.xlsx)
and [repository README](https://github.com/KAJALPANIGRAHI1997/Quality-Inspection-Dashboard).

- Original file is `data/input/Quality_Inspection_Dataset.xlsx`.
- SHA-256 is `56651c6f72ec31bd834b5ad29b408780229f35e2abb4540f951aae1ba6fd62a0`.
- The README's license section permits learning and showcase sharing with
  author credit. The repository supplies no standard license file. The copy
  here is attributed for that stated purpose.
- The repository does not establish how or where the records were collected.
  The workbook is a public sample with unverified factory provenance.

Actual file inspection:

| Check | Result |
| --- | --- |
| Sheets and size | `Sheet1` only, 350 data rows, 9 columns |
| Actual column headers | `Batch_ID`, `Product_Name`, `Units_Checked`, `Defective_Units`, `Defect_Type`, `Supplier`, `Inspector_Name`, `Line_Number`, `Inspection_Date` |
| Missing cells and exact duplicates | 0 missing cells in each column and 0 exact duplicate rows |
| Inspection date and line | January 1–31, 2024, across 4 labeled lines |
| Counts | 42,259 checked, 5,062 defective, 0 negative, 0 zero checked, 0 defective counts above checked |
| IDs | 296 distinct batch IDs, 46 IDs recur, 54 rows are additional occurrences after the first per batch ID |
| Batch-key candidates | 2 excess rows for batch/date, 1 for batch/date/line, 0 for batch/date/line/product |
| Defect labels | 4 distinct labels. All 14 zero-defective-unit rows carry a type label. |

**Working grain:** one recorded inspection cohort per row. Use
`(Batch_ID, Inspection_Date, Line_Number, Product_Name)` as the observed
composite key. This tuple is unique in the file. Retaining these rows assumes
the checked cohorts do not overlap. The source has no individual unit ID or
first-pass/rework field to verify that assumption. The README describes the
count header as `Units_Check`. The actual XLSX uses `Units_Checked`.

The defect type is one label per row. There is no documented allocation of
individual defective units to that label and no per-type occurrence count.
This report omits the Pareto rather than assigning the entire row count to
the label. Supplier and inspector columns remain in the source workbook.

## Public synthetic Pareto demonstration

**Source:** Oussama Elayeb, *Manufacturing Quality Control Dashboard*.
[Original CSV](https://github.com/elayeboussama/manufacturing-quality-dashboard/blob/main/data/quality_inspections.csv),
[data generator](https://github.com/elayeboussama/manufacturing-quality-dashboard/blob/main/src/generate_data.py),
and [MIT license](https://github.com/elayeboussama/manufacturing-quality-dashboard/blob/main/LICENSE).

- Original file is `data/raw/quality_inspections.csv`. The upstream MIT notice
  is retained as `data/raw/LICENSE-OUSSAMA-ELAYEB-MIT.txt`.
- SHA-256 is `a5ed3c827ad2cd5ed5019927fab4286abab33021c3c1065b2a3b31fd0bdd6b5c`.
- The upstream seeded generator simulates a single inspection and one pass/fail
  result per ID. It assigns one category and severity to each failed record.
  These are simulated observations.

| Check | Result |
| --- | --- |
| File shape | CSV, 60,000 rows and 8 columns |
| Actual column headers | `inspection_id`, `date`, `production_line`, `shift`, `product_model`, `status`, `defect_type`, `severity` |
| Missing values | 56,666 blanks in each of `defect_type` and `severity`, all on Pass rows. 0 missing values in the other fields. |
| Duplicates | 0 exact duplicate rows and 0 repeated inspection IDs |
| Date and line | January 1–December 31, 2025, across 4 lines |
| Outcomes | 56,666 Pass and 3,334 Fail. All failed inspections have one of 6 types. |

`prepare_pareto_demo.py` maps each ID to `inspected_units = 1` and each Fail
to `defective_units = 1`, otherwise 0. The date, line, product and defect
type come directly from the upstream record. The normalized output is
`data/pareto_input/public_synthetic_inspections.csv`. The Pareto counts
simulated failed units per assigned type. It has no information on multiple
physical defect events per unit.

## Another candidate considered

[UCI SECOM](https://archive.ics.uci.edu/dataset/179/secom), CC BY 4.0, has
timestamped pass/fail labels for semiconductor production entities. The
published [labels file](https://archive.ics.uci.edu/ml/machine-learning-databases/secom/secom_labels.data)
provides status and timestamp. It does not provide identified production lines
or defect-type counts, so it was excluded from the bundled demonstrations.
