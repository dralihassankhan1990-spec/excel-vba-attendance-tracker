# Excel VBA Attendance Tracker

**by Muhammad Bilal**

## The problem
Tracking monthly attendance for a team means staring at hundreds of daily
codes — P (present), A (absent), L (late), H (half day) — and hand-computing
each person's present percentage, absent count, and late count. It is slow,
mind-numbing, and one slip ruins the payroll summary.

## How it works
`Attendance.xlsm` ships with a VBA macro, **`BuildAttendanceReport`**
(source: `AttendanceReport.bas`), that does the whole month in one click:

1. Reads the **Attendance** sheet — 15 fictional employees × 30 days of codes.
2. Builds a new **"Attendance Report"** sheet: per-employee % present (P + ½·H),
   absent / late / half-day counts, and days tracked.
3. Color-codes the source sheet — absent cells red, late cells amber — so
   patterns jump out instantly.
4. Flags chronic absentees (≥5 absents, red) and frequent latecomers
   (≥3 lates, amber) right in the report.
5. Writes a **summary dashboard** block: employees tracked, total absents,
   total lates, and the worst absentee.

No external APIs, no network — 100% local. All demo data is fictional.

## Run it (Excel)
1. Open `Attendance.xlsm`. If Excel shows a **security warning**, click
   **Enable Content** (the macro is unsigned; that is normal for a demo).
2. Press **Alt+F8**, select `BuildAttendanceReport`, click **Run**.
   (Or add a button: Developer → Insert → Button, and assign the macro.)
3. The **Attendance Report** sheet appears with the full breakdown.

## Run it (LibreOffice)
Macros run in LibreOffice too: open the `.xlsm`, enable macros when asked,
then Tools → Macros → Run Macro → `AttendanceReport.BuildAttendanceReport`.

## Files
| File | What it is |
|---|---|
| `Attendance.xlsm` | The workbook: sample data + embedded VBA macro |
| `AttendanceReport.bas` | The VBA source (also embedded byte-identical in the `.xlsm`) |
| `build_data.py` | Generates the fictional sample data + `expected-stats.json` |
| `build_vba_project.py` | Builds a genuine `vbaProject.bin` and injects it into the `.xlsm` |
| `build_all.py` | One-shot builder: data → workbook → macro injection → verification |
| `prove.py` | Executes the real macro via LibreOffice UNO and verifies every number |
| `Attendance-base.xlsx` | Plain data workbook used as the proof input |
| `proof-output.xlsx` | Workbook produced by the actual macro run (video source) |
| `Attendance-Tracker-demo.mp4` | ~100s demo video: real macro execution, real numbers |
| `TEST-REPORT.md` | Execution evidence and verification results |

## Proof
See `TEST-REPORT.md` — the macro was executed for real (LibreOffice UNO
bridge), and all 15 employee rows plus the summary were verified against
independently computed values. Spot-check: **Kamran Iqbal — 8 absents**
(flagged red), **15 employees, 25 total absents, 24 total lates**.
