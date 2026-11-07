# Verification record

Executed on 2026-11-07 with Python 3.12.14 on Linux x86-64. Runtime package
versions were pandas 2.2.3, openpyxl 3.1.5, Matplotlib 3.10.8, python-pptx 1.0.2
and pytest 8.4.2. The four reporting dependencies were already available.
The missing pytest packages were installed locally from checksum-verified
PyPI wheels after a direct pip request timed out.

## Commands and results

```text
python prepare_pareto_demo.py
python main.py --input data/input --output reports --dataset-label "Kajal Panigrahi public sample (factory provenance unverified)"
python main.py --input data/pareto_input --output reports --pareto --dataset-label "Oussama Elayeb public synthetic inspections (simulated)"
python -m pytest -q
```

Both CLI runs exited with code 0. Bundled example directories are named for
their source. A normal run writes a unique `run-...` directory.

| Check | Public workbook | Public synthetic demo |
| --- | ---: | ---: |
| Accepted inspection records | 350 | 60,000 |
| Excluded rows | 0 | 0 |
| Logged issues | 2 | 0 |
| Inspected units | 42,259 | 60,000 |
| Defective units | 5,062 | 3,334 |
| Defect rate | 11.98% | 5.56% |
| Inspection yield | 88.02% | 94.44% |
| Excel sheets | 4 | 5 |
| PowerPoint slides | 4 | 5 |

**Final pytest result: 22 passed.** Tests covered empty input, missing and
ambiguous headers, CSV/XLSX input in one run, worksheet selection and audit
provenance, exact duplicates, conflicting IDs, reused batch IDs, invalid counts,
zero denominator, date parsing, weighted aggregation, missing line values,
Pareto eligibility, malformed input files, and both public-source CLI runs.

For each fresh CLI run the tests calculate source totals independently and
reopen Excel and PowerPoint. They check all daily and line groups, both KPI
values, the optional Pareto totals and cumulative shares, expected slide counts,
and that the same generated chart bytes appear in the workbook and deck.
The first test run exposed a verification mismatch between Excel date objects
and source date strings. The verifier now normalizes the Excel date keys
before comparing them.

Both bundled decks were rendered with LibreOffice Impress and checked for
off-slide overflow. The workbook summary, daily chart, line chart, audit page
and synthetic Pareto were inspected using LibreOffice Calc print previews.

## Platform coverage

Windows 11, PowerShell, VS Code, and Microsoft Office desktop were outside
the execution environment. The README includes the PowerShell commands.
The file structure and CLI were verified on Linux, including paths with spaces.
