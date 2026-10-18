"""Proof: execute the REAL AttendanceReport VBA macro in LibreOffice and verify.

Pipeline (mirrors the proven approach from automation-build/03):
1. Start headless LibreOffice with a UNO bridge.
2. Open Attendance-base.xlsx, inject the .bas source as a Basic module
   (prefixed with "Option VBASupport 1", exactly like the reference build).
3. Invoke BuildAttendanceReport via ScriptProviderForBasic.
4. Read back the "Attendance Report" sheet through UNO and compare every
   number against expected-stats.json (independently computed in Python).
5. Save the macro-built workbook for the video (proof-output.xlsx).

100% local. by Muhammad Bilal.
"""
import json
import logging
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
BASE_XLSX = os.path.join(HERE, "Attendance-base.xlsx")
BAS_FILE = os.path.join(HERE, "AttendanceReport.bas")
EXPECTED_JSON = os.path.join(HERE, "expected-stats.json")
PROOF_XLSX = os.path.join(HERE, "proof-output.xlsx")
PORT = 2002

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("prove")


def pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def start_soffice():
    subprocess.run(["pkill", "-f", "[s]office.bin"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(1)
    proc = subprocess.Popen(
        ["soffice", "--headless", "--nologo", "--norestore",
         f"--accept=socket,host=localhost,port={PORT};urp;StarOffice.ServiceManager",
         "--nofirststartwizard"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    for _ in range(60):
        try:
            ctx = resolver.resolve(
                f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
            log.info("UNO bridge connected.")
            return proc, ctx
        except Exception:
            time.sleep(1)
    raise RuntimeError("Could not connect to LibreOffice UNO bridge")


def main() -> int:
    with open(EXPECTED_JSON) as f:
        expected = json.load(f)
    with open(BAS_FILE, encoding="utf-8") as f:
        vba = f.read()
    vba_lines = [ln for ln in vba.splitlines() if not ln.startswith("Attribute ")]
    module_code = "Option VBASupport 1\n" + "\n".join(vba_lines)

    proc, ctx = start_soffice()
    try:
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(BASE_XLSX), "_blank", 0, ())
        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        std = libs.getByName("Standard")
        if std.hasByName("AttendanceReport"):
            std.removeByName("AttendanceReport")
        std.insertByName("AttendanceReport", module_code)

        sp = smgr.createInstanceWithContext(
            "com.sun.star.script.provider.ScriptProviderForBasic", ctx)
        sp.initialize((doc,))
        script = sp.getScript(
            "vnd.sun.star.script:Standard.AttendanceReport.BuildAttendanceReport"
            "?language=Basic&location=document")
        script.invoke((), (), ())
        log.info("BuildAttendanceReport executed without error.")

        # ---- verify the report sheet ----
        sheets = doc.getSheets()
        assert sheets.hasByName("Attendance Report"), "report sheet missing!"
        assert sheets.hasByName("Attendance"), "source sheet missing!"
        rep = sheets.getByName("Attendance Report")
        src = sheets.getByName("Attendance")

        assert rep.getCellByPosition(0, 0).getString() == "MONTHLY ATTENDANCE REPORT"
        errors = []
        for i, emp in enumerate(expected["per_employee"]):
            r = 4 + i  # 0-based row; data starts at Excel row 5
            name = rep.getCellByPosition(0, r).getString()
            pct = rep.getCellByPosition(1, r).getValue() * 100
            absent = int(rep.getCellByPosition(2, r).getValue())
            late = int(rep.getCellByPosition(3, r).getValue())
            half = int(rep.getCellByPosition(4, r).getValue())
            tracked = int(rep.getCellByPosition(5, r).getValue())
            if name != emp["employee"]:
                errors.append(f"row {i}: name {name!r} != {emp['employee']!r}")
            if abs(pct - emp["pct"]) > 0.06:
                errors.append(f"row {i} ({name}): pct {pct:.1f} != {emp['pct']}")
            for got, want, label in [(absent, emp["absent"], "absent"),
                                     (late, emp["late"], "late"),
                                     (half, emp["half"], "half"),
                                     (tracked, emp["tracked"], "tracked")]:
                if got != want:
                    errors.append(f"row {i} ({name}): {label} {got} != {want}")
        # summary block: s = 20 + 2 = 22 (0-based) -> Excel rows 23..27
        s = 4 + len(expected["per_employee"]) + 2
        got_emp = int(rep.getCellByPosition(1, s + 1).getValue())
        got_abs = int(rep.getCellByPosition(1, s + 2).getValue())
        got_late = int(rep.getCellByPosition(1, s + 3).getValue())
        got_max = rep.getCellByPosition(1, s + 4).getString()
        if got_emp != expected["employees"]:
            errors.append(f"summary employees {got_emp} != {expected['employees']}")
        if got_abs != expected["total_absent"]:
            errors.append(f"summary total_absent {got_abs} != {expected['total_absent']}")
        if got_late != expected["total_late"]:
            errors.append(f"summary total_late {got_late} != {expected['total_late']}")
        want_max = f"{expected['max_absent_emp']} ({expected['max_absent']})"
        if got_max != want_max:
            errors.append(f"summary max {got_max!r} != {want_max!r}")

        # spot-check source-sheet coloring: find first "A" in sample data and
        # confirm the cell got the red fill (0xFFCDCD)
        colored = src.getCellByPosition(1, 1).CellBackColor  # header, skip
        check = None
        for rr in range(1, 16):
            for cc in range(1, 31):
                if src.getCellByPosition(cc, rr).getString() == "A":
                    check = (cc, rr, src.getCellByPosition(cc, rr).CellBackColor)
                    break
            if check:
                break
        log.info("First 'A' cell at col=%d row=%d backcolor=0x%06X",
                 check[0], check[1], check[2])
        if check[2] != 0xFFCDCD:
            errors.append(f"source coloring: expected red 0xFFCDCD, got 0x{check[2]:06X}")

        if errors:
            for e in errors:
                log.error("MISMATCH: %s", e)
            raise SystemExit(f"PROOF FAILED: {len(errors)} mismatches")
        log.info("All %d employee rows + summary match expected stats. "
                 "Source coloring verified red.",
                 len(expected["per_employee"]))

        # hand-verification spot checks (printed for the TEST-REPORT)
        e0 = expected["per_employee"][0]
        log.info("HAND-CHECK e.g. %s: pct=%.1f%% absent=%d late=%d (from seed-42 data)",
                 e0["employee"], e0["pct"], e0["absent"], e0["late"])
        log.info("HAND-CHECK summary: employees=%d total_absent=%d total_late=%d max=%s(%d)",
                 expected["employees"], expected["total_absent"],
                 expected["total_late"], expected["max_absent_emp"],
                 expected["max_absent"])

        if os.path.exists(PROOF_XLSX):
            os.remove(PROOF_XLSX)
        doc.storeAsURL(uno.systemPathToFileUrl(PROOF_XLSX),
                       (pv("FilterName", "Calc MS Excel 2007 XML"), pv("Overwrite", True)))
        log.info("Macro-built workbook saved: %s", PROOF_XLSX)
        doc.close(True)
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
    log.info("PROOF PASSED.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
