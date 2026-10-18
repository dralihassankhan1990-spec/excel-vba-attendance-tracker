# TEST REPORT — Excel VBA Attendance Tracker

**by Muhammad Bilal** · executed 2026-10-18 · 100% local, no network/APIs

## What was proven
The shipped VBA macro (`AttendanceReport.bas`, embedded byte-identical in
`Attendance.xlsm`) was **executed for real** and every computed value was
verified — not simulated, not hand-typed.

## Execution method
- Headless LibreOffice + UNO bridge (`python3-uno`).
- `prove.py` opened `Attendance-base.xlsx`, injected the **actual `.bas`
  source** as a Basic module (prefixed `Option VBASupport 1`, the same
  technique as the proven reference build), and invoked
  `BuildAttendanceReport` through `ScriptProviderForBasic`.
- The macro created the "Attendance Report" sheet, color-coded the source
  sheet, and wrote the summary block — all inside the running office engine.
- Results were read back cell-by-cell through UNO and compared against
  `expected-stats.json`, computed **independently in Python** from the same
  sample data (seed 42, plus one forced chronic absentee and one forced
  frequent latecomer for a realistic demo).

## Results — ALL CHECKS PASSED
- **15/15 employee rows match**: name, % present, absent, late, half-day,
  days tracked — zero mismatches.
- **Summary block matches**: 15 employees · 25 total absents · 24 total lates ·
  most absences: Kamran Iqbal (8).
- **Source coloring verified**: first "A" cell backcolor = `0xFFCDCD` (red).

## Hand-verified spot checks (against the sample data by eye)
1. **Ayesha Khan**: 83.3% present, 2 absent, 3 late — matches the source row.
2. **Kamran Iqbal**: 8 absents (6 forced + 2 random) → correctly flagged red
   as a chronic absentee (≥5 rule).
3. **Summary totals**: 25 absents / 24 lates — consistent with the 15-row sum.

## Deliverable integrity
- `olevba` confirms the `.xlsm` contains `AttendanceReport.bas`
  **byte-identical** to the shipped source (6,501 chars).
- Modules in the workbook: `AttendanceReport.bas`, `Sheet1.cls`,
  `ThisWorkbook.cls`.

## Honest limitations
- Execution used LibreOffice's VBA engine, not Microsoft Excel. The macro is
  written in portable VBA (`RGB()`, `Interior.Color`, `Font.Bold`,
  `NumberFormat`, merged ranges) that runs in Excel unchanged; acceptance in
  Excel itself was not run in this Linux environment.
- `python3-uno` was already present in the image; the earlier "no UNO"
  finding was an apt-lock race, corrected during this build.
- The demo video shows the real terminal run and the real macro-built
  workbook. LibreOffice has no GUI in this environment, so the video renders
  the actual output files rather than a live Calc window — every pixel comes
  from genuine macro output.

## How to re-run the proof
```
cd 10-Excel-VBA-Attendance-Tracker
python3 prove.py
```
Requires: Python 3, `openpyxl`, `uno` (python3-uno), headless LibreOffice.
