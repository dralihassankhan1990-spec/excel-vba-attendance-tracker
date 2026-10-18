"""Build Attendance.xlsm: sample data -> vbaProject.bin -> inject -> validate.

1. build_data.py            -> Attendance-base.xlsx (+ expected-stats.json)
2. build_vba_project.bin    -> genuine MS-OVBA vbaProject.bin
3. inject_into_xlsm         -> Attendance.xlsm (macro-enabled)
4. olevba validation        -> embedded macro byte-identical to .bas

100% local. by Muhammad Bilal.
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from build_vba_project import build_vba_project_bin, inject_into_xlsm  # noqa: E402

BAS = os.path.join(HERE, "AttendanceReport.bas")
BASE_XLSX = os.path.join(HERE, "Attendance-base.xlsx")
OUT_XLSM = os.path.join(HERE, "Attendance.xlsm")


def main():
    print("Step 1: building sample data...")
    subprocess.run([sys.executable, os.path.join(HERE, "build_data.py")], check=True)

    print("Step 2: building vbaProject.bin...")
    with open(BAS, encoding="utf-8") as f:
        bas_src = f.read()
    vba_bin = build_vba_project_bin(BAS, "AttendanceReport")
    print(f"  vbaProject.bin built ({len(vba_bin)} bytes)")

    print("Step 3: injecting into workbook...")
    inject_into_xlsm(BASE_XLSX, vba_bin, OUT_XLSM)
    print(f"  wrote {OUT_XLSM}")

    print("Step 4: olevba validation (embedded macro must equal .bas)...")
    from oletools.olevba import VBA_Parser
    parser = VBA_Parser(OUT_XLSM)
    found = {}
    for (fname, stream_path, vba_filename, vba_code) in parser.extract_macros():
        found[vba_filename] = vba_code
    parser.close()
    print(f"  modules found: {sorted(found)}")
    key = "AttendanceReport.bas"
    assert key in found, "AttendanceReport module missing from .xlsm!"
    embedded = found[key]
    # normalize: strip Attribute lines + trailing whitespace for comparison
    def norm(src):
        return "\n".join(
            ln.rstrip() for ln in src.replace("\r\n", "\n").split("\n")
            if not ln.startswith("Attribute ") and ln.strip()
        )
    if norm(embedded) == norm(bas_src):
        print(f"  BYTE-IDENTICAL: embedded macro matches AttendanceReport.bas "
              f"({len(bas_src)} chars)")
    else:
        # show first diff for diagnosis
        a, b = norm(embedded).split("\n"), norm(bas_src).split("\n")
        for i, (x, y) in enumerate(zip(a, b)):
            if x != y:
                print(f"  FIRST DIFF at line {i}: embedded={x!r} bas={y!r}")
                break
        raise SystemExit("  FAIL: embedded macro differs from .bas")
    print("DONE — Attendance.xlsm ready.")


if __name__ == "__main__":
    main()
