# -*- coding: utf-8 -*-
"""
run_fatigue_workbook.py
Feed the delta-S11 reports written by extract_stress_ranges.py into the fatigue
workbook and run its own macro workflow (import -> fatigue damage -> summary)
for every case, through Excel.

Needs Windows, Excel and pywin32:

    pip install pywin32
    python run_fatigue_workbook.py

The original workbook is not modified: the script works on a copy (OUTPUT).
"""
import json
import os
import shutil
import sys

# ============================ USER INPUT =====================================
WORKBOOK = r"C:\SeabedAnalysis\FatigueDamageCal_ManualWorkflows_INPUT.xlsm"
OUTPUT = r"C:\SeabedAnalysis\FatigueDamageCal_Results.xlsm"
MANIFEST = r"C:\SeabedAnalysis\fatigue_reports\manifest.json"
VISIBLE = False            # True to watch Excel while it runs
# =============================================================================

MACRO = "FatigueDamageManualWorkflow.RunManualFatigueWorkflowFromPath"
CONTROL = "Workflow Controls"
XL_MANUAL, XL_AUTOMATIC = -4135, -4105
FATIGUE_SHEETS = [  # (sheet, first damage column, number of damage columns)
    ("FatigueID_Hydrotest", 5, 4), ("FatigueOD_Hydrotest", 6, 4),
    ("FatigueID_DesignOp", 5, 4), ("FatigueOD_DesignOp", 6, 4),
    ("FatigueID_EarlyUndr", 5, 16), ("FatigueOD_EarlyUndr", 6, 16),
    ("FatigueID_Earlydr", 5, 16), ("FatigueOD_Earlydr", 6, 16),
    ("FatigueID_Middr", 5, 16), ("FatigueOD_Middr", 6, 16),
    ("FatigueID_Latdr", 5, 16), ("FatigueOD_Latdr", 6, 16),
]


def row_max(ws, row, first_col, count):
    best = None
    for c in range(first_col, first_col + count):
        v = ws.Cells(row, c).Value
        if isinstance(v, (int, float)) and (best is None or v > best):
            best = v
    return best


def main():
    import win32com.client as win32

    with open(MANIFEST) as fh:
        cases = json.load(fh)
    if not cases:
        sys.exit("No cases in %s" % MANIFEST)
    for c in cases:
        if not os.path.isfile(c["report"]):
            sys.exit("Report not found: %s" % c["report"])
        if not c.get("phase") and not c["load_case"].replace(" ", "").startswith(("HYDROTEST", "DESIGN")):
            sys.exit("Case %s has no phase. Rerun the extraction with --phase."
                     % c["label"])
    if os.path.abspath(OUTPUT).lower() == os.path.abspath(WORKBOOK).lower():
        sys.exit("OUTPUT must differ from WORKBOOK")
    shutil.copyfile(WORKBOOK, OUTPUT)

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = VISIBLE
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1          # allow the workbook's macros to run
    failed = []
    wb = None
    try:
        wb = xl.Workbooks.Open(OUTPUT)
        xl.Calculation = XL_MANUAL
        xl.ScreenUpdating = False
        ctl = wb.Worksheets(CONTROL)
        macro = "'%s'!%s" % (wb.Name, MACRO)
        for c in cases:
            ctl.Range("I4").Value = "DELTA S11"
            if c.get("phase"):
                ctl.Range("I5").Value = c["phase"]
            ctl.Range("I6").Value = c["load_case"]
            ctl.Range("I7").Value = ""           # one pair per report: no filter
            ctl.Range("I10").Value = ""
            xl.Run(macro, c["report"])
            status = str(ctl.Range("I10").Value or "")
            ok = status.startswith("PROGRAMMATIC TEST PASSED")
            print("%-20s %-18s %-5s %s" % (c["label"], c["phase"], c["load_case"],
                                           "OK" if ok else status))
            if not ok:
                failed.append((c["label"], status))
        xl.Calculation = XL_AUTOMATIC
        xl.ScreenUpdating = True

        print("\nKP rows: %s" % ctl.Range("I12").Value)
        print("Maximum damage per sheet (largest of the angles / load cases):")
        print("  %-22s %-28s %12s %12s" % ("sheet", "stress source", "DBM", "RB"))
        for name, col, count in FATIGUE_SHEETS:
            ws = wb.Worksheets(name)
            source = str(ws.Range("C1").Value or "")
            if "no data" in source.lower():
                continue
            dbm, rb = row_max(ws, 5, col, count), row_max(ws, 8, col, count)
            print("  %-22s %-28s %12s %12s"
                  % (name, source,
                     "-" if dbm is None else "%.3e" % dbm,
                     "-" if rb is None else "%.3e" % rb))
        wb.Save()
        print("\nSaved %s" % OUTPUT)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()

    if failed:
        print("\nCases with a workflow error (message from the workbook macro):")
        for label, status in failed:
            print("  %s: %s" % (label, status))
        sys.exit(1)


if __name__ == "__main__":
    main()
