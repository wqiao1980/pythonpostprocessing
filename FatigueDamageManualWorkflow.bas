Attribute VB_Name = "FatigueDamageManualWorkflow"
Option Explicit

' VBA-only manual fatigue workflow for FatigueDamageCal.xlsx.
' No Workbook_Open or Worksheet_Change automation is used.

Private Const CONTROL_SHEET As String = "Workflow Controls"
Private Const FIRST_DATA_ROW As Long = 15
Private Const LAST_TEMPLATE_ROW As Long = 16756
Private Const SEP As String = "~|~"

Private Const MODE_CELL As String = "I4"
Private Const PHASE_CELL As String = "I5"
Private Const LOAD_CASE_CELL As String = "I6"
Private Const PAIR_FILTER_CELL As String = "I7"
Private Const ALLOWABLE_CELL As String = "I8"
Private Const STATUS_CELL As String = "I10"
Private Const LAST_MODE_CELL As String = "I11"
Private Const KP_COUNT_CELL As String = "I12"
Private Const FIRST_STEP_CELL As String = "I13"
Private Const SECOND_STEP_CELL As String = "I14"
Private Const SELECTED_PAIR_CELL As String = "I15"
Private Const SOURCE_FILE_CELL As String = "I16"

' INPUT tab: one column per fatigue tab; fatigue-tab row r is INPUT row r + offset.
Private Const INPUT_SHEET As String = "INPUT"
Private Const INPUT_ROW_OFFSET As Long = 4
Private Const INPUT_FIRST_COLUMN As Long = 2
Private Const INPUT_DBM_HEADER_ROW As Long = 25
' User S-N curve table on INPUT: starts in this column, same header row.
'   +0 Range ID, +1 KP Start, +2 KP End, +3..+6 inner fibre (name, m, C2, KDF),
'   +7..+10 outer fibre (name, m, C2, KDF), +11 KP rows matched (output)
Private Const USER_CURVE_FIRST_COLUMN As Long = 8
' Result KP ranges table on INPUT (Range ID, KP Start, KP End) and its output tab.
Private Const RESULT_RANGE_FIRST_COLUMN As Long = 21
Private Const RANGE_SUMMARY_SHEET As String = "Fatigue_KPRangeSUM"
' Plots tab: charts at the top, their helper data from this row down.
Private Const PLOTS_SHEET As String = "Plots"
Private Const PLOTS_DATA_ROW As Long = 102
Private Const U2_SHEET As String = "U2"
Private Const U2_HEADER_ROW As Long = 4
Private Const LOAD_CASE_LIST As String = "FCD,HCD,PCD,HYDROTEST 1,HYDROTEST 2,DESIGN"
Private Const PHASE_LIST As String = "EARLY UNDRAINED,EARLY DRAINED,MIDDLE DRAINED,LATE DRAINED,HYDROTEST 1,HYDROTEST 2,DESIGN OPERATION"
Private Const CONTROL_NOTE_1 As String = "Enter cycle counts, SCF, wall thickness, S-N curves, DBM KP ranges and optional user S-N curve KP ranges on the INPUT tab."
Private Const CONTROL_NOTE_2 As String = "If middle/late stress is absent, those tabs use Early Drained (or Early Undrained). Target phase is ignored when the load case is HYDROTEST 1, HYDROTEST 2 or DESIGN."

Private mProgrammaticMode As Boolean
Private mUserCurveCount As Long
Private mUserCurveStart() As Double
Private mUserCurveEnd() As Double
Private mUserCurveRow() As Long
Private mUserCurveHasId() As Boolean
Private mUserCurveHasOd() As Boolean
Private mCurrentStage As String

' ============================================================================
' PUBLIC SETUP AND WORKFLOW BUTTON MACROS
' ============================================================================

Public Sub SetupManualFatigueWorkflow()
    Dim oldCalculation As XlCalculation
    On Error GoTo Failed
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    EnsureLateLifeSheets
    EnsureHydrotestAndDesignSheets
    ConfigureFatigueParameterBlocks
    ConfigureStressRangeDbmArea
    BuildInputSheetCore
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_IDSUMEarlyProfile"), True, True
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_ODSUMEArlyProfile"), False, True
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_IDSUM"), True, False
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_ODSUM"), False, False
    ApplySinglePcdLayout
    BuildWorkflowControlsSheet
    SetWorkflowStatus "Ready. Select an input option, phase, and load case."

CleanExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    MsgBox "Workflow setup failed: " & Err.Description, vbCritical, "Fatigue workflow"
    Resume CleanExit
End Sub

Public Sub RunAllManualFatigueWorkflows()
    Dim oldCalculation As XlCalculation, importMode As String
    On Error GoTo Failed
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Running the manual fatigue workflow..."

    importMode = ImportFatigueReportCore(vbNullString)
    If Len(importMode) = 0 Then GoTo CleanExit
    If importMode = "S11" Then
        SeparateSheet1IntoIdOdCore
        CalculateStressRangesCore
    End If
    CalculateFatigueDamageCore
    SummarizeFatigueDamageCore
    Application.CalculateFullRebuild
    SetWorkflowStatus "Run All completed. Last input mode: " & importMode
    MsgBox "The manual fatigue workflow is complete.", vbInformation, "Fatigue workflow"

CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetWorkflowStatus "ERROR: " & Err.Description
    MsgBox "Run All failed: " & Err.Description, vbCritical, "Fatigue workflow"
    Resume CleanExit
End Sub

Public Sub ImportS11OrDeltaS11Report()
    Dim oldCalculation As XlCalculation, importMode As String
    On Error GoTo Failed
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Importing the S11 report..."
    importMode = ImportFatigueReportCore(vbNullString)
    If Len(importMode) > 0 Then
        SetWorkflowStatus "Button 1 complete. Imported " & importMode & "."
        If importMode = "DELTA S11" Then
            MsgBox "Delta S11 was imported directly into StressRangeID and StressRangeOD." & _
                   vbCrLf & "You can now run button 4 and button 5.", _
                   vbInformation, "Fatigue workflow"
        Else
            MsgBox "S11 was imported into Sheet1. Run button 2 next.", _
                   vbInformation, "Fatigue workflow"
        End If
    End If
CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetWorkflowStatus "ERROR: " & Err.Description
    MsgBox "Import failed: " & Err.Description, vbCritical, "Fatigue workflow"
    Resume CleanExit
End Sub

Public Sub SeparateSheet1IntoIdOd()
    Dim oldCalculation As XlCalculation
    On Error GoTo Failed
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Separating inner and outer fiber S11..."
    SeparateSheet1IntoIdOdCore
    SetWorkflowStatus "Button 2 complete. Sheet1 was separated into ID and OD."
    MsgBox "Sheet1 was separated into the ID and OD tabs.", vbInformation, "Fatigue workflow"
CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetWorkflowStatus "ERROR: " & Err.Description
    MsgBox "ID/OD separation failed: " & Err.Description, vbCritical, "Fatigue workflow"
    Resume CleanExit
End Sub

Public Sub CalculateStressRanges()
    Dim oldCalculation As XlCalculation
    On Error GoTo Failed
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Calculating heat-up minus cool-down stress ranges..."
    CalculateStressRangesCore
    SetWorkflowStatus "Button 3 complete. Stress ranges were written to the selected phase/load case."
    MsgBox "Stress ranges were calculated in StressRangeID and StressRangeOD.", _
           vbInformation, "Fatigue workflow"
CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetWorkflowStatus "ERROR: " & Err.Description
    MsgBox "Stress-range calculation failed: " & Err.Description, vbCritical, "Fatigue workflow"
    Resume CleanExit
End Sub

Public Sub CalculateFatigueDamage()
    Dim oldCalculation As XlCalculation
    On Error GoTo Failed
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Calculating fatigue damage with DBM and RB S-N curves..."
    CalculateFatigueDamageCore
    Application.CalculateFullRebuild
    SetWorkflowStatus "Button 4 complete. Fatigue damage was calculated using DBM/RB KP classification."
    MsgBox "Fatigue damage was calculated along the pipeline KP.", vbInformation, "Fatigue workflow"
CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetWorkflowStatus "ERROR: " & Err.Description
    MsgBox "Fatigue calculation failed: " & Err.Description, vbCritical, "Fatigue workflow"
    Resume CleanExit
End Sub

Public Sub SummarizeFatigueDamageAndUC()
    Dim oldCalculation As XlCalculation
    On Error GoTo Failed
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Summarizing fatigue damage and utilization..."
    SummarizeFatigueDamageCore
    Application.CalculateFullRebuild
    SetWorkflowStatus "Button 5 complete. Summaries and fatigue charts were updated."
    MsgBox "Fatigue damage, UC summaries, and charts were updated.", _
           vbInformation, "Fatigue workflow"
CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetWorkflowStatus "ERROR: " & Err.Description
    MsgBox "Summary calculation failed: " & Err.Description, vbCritical, "Fatigue workflow"
    Resume CleanExit
End Sub

' Button 6: imports a U2 report of extract_stress_ranges.py (U2_steps.rpt or
' CHECK_..._U2.rpt) into the U2 tab and adds the U2 chart to the Plots tab.
Public Sub ImportU2Report()
    Dim picked As Variant, stepCount As Long
    On Error GoTo Failed
    picked = Application.GetOpenFilename("U2 report (*.rpt;*.txt),*.rpt;*.txt,All files (*.*),*.*", , _
                                         "Select the U2 report (U2_steps.rpt)")
    If VarType(picked) = vbBoolean Then Exit Sub
    Application.ScreenUpdating = False
    stepCount = ImportU2Core(CStr(picked))
    SetWorkflowStatus "Button 6 complete. U2 of " & CStr(stepCount) & " step(s) imported; Plots tab updated."
    MsgBox "U2 of " & CStr(stepCount) & " step(s) was imported into the U2 tab." & vbCrLf & _
           "The U2 chart is on the Plots tab.", vbInformation, "Fatigue workflow"
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetWorkflowStatus "ERROR: " & Err.Description
    MsgBox "U2 import failed: " & Err.Description, vbCritical, "Fatigue workflow"
    Resume CleanExit
End Sub

' Same as button 6 without dialogs, for controlled testing or batch calls.
Public Sub ImportU2ReportFromPath(ByVal reportPath As String)
    Dim stepCount As Long
    On Error GoTo Failed
    Application.ScreenUpdating = False
    stepCount = ImportU2Core(reportPath)
    SetWorkflowStatus "PROGRAMMATIC U2 IMPORT PASSED: " & CStr(stepCount) & " step(s)"
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetWorkflowStatus "ERROR: " & Err.Description
    Resume CleanExit
End Sub

' This argument-taking procedure is for controlled testing or batch calls.
' It does not appear in the Alt+F8 no-argument macro list.
Public Sub RunManualFatigueWorkflowFromPath(ByVal reportPath As String)
    Dim importMode As String
    On Error GoTo Failed
    mProgrammaticMode = True
    mCurrentStage = "button 1 report import"
    importMode = ImportFatigueReportCore(reportPath)
    If importMode = "S11" Then
        mCurrentStage = "button 2 ID/OD separation"
        SeparateSheet1IntoIdOdCore
        mCurrentStage = "button 3 stress-range calculation"
        CalculateStressRangesCore
    End If
    mCurrentStage = "button 4 fatigue calculation"
    CalculateFatigueDamageCore
    mCurrentStage = "button 5 summary calculation"
    SummarizeFatigueDamageCore
    mCurrentStage = "Excel recalculation"
    Application.CalculateFullRebuild
    SetWorkflowStatus "PROGRAMMATIC TEST PASSED: " & importMode
CleanExit:
    mProgrammaticMode = False
    Exit Sub
Failed:
    SetWorkflowStatus "PROGRAMMATIC ERROR in " & mCurrentStage & ": " & Err.Description
    Resume CleanExit
End Sub

Public Sub SelfCheckManualFatigueWorkflow()
    Dim requiredSheets As Variant, item As Variant, ws As Worksheet
    On Error GoTo Failed
    requiredSheets = Array("Sheet1", "ID", "OD", "StressRangeID", "StressRangeOD", _
        "FatigueID_EarlyUndr", "FatigueID_Earlydr", "FatigueID_Middr", "FatigueID_Latdr", _
        "FatigueID_Hydrotest", "FatigueID_DesignOp", _
        "FatigueOD_EarlyUndr", "FatigueOD_Earlydr", "FatigueOD_Middr", "FatigueOD_Latdr", _
        "FatigueOD_Hydrotest", "FatigueOD_DesignOp", _
        "Fatigue_IDSUMEarlyProfile", "Fatigue_ODSUMEArlyProfile", _
        "Fatigue_IDSUM", "Fatigue_ODSUM", CONTROL_SHEET, INPUT_SHEET)
    For Each item In requiredSheets
        If Not WorksheetExists(CStr(item)) Then Err.Raise vbObjectError + 900, , _
            "Required worksheet is missing: " & CStr(item)
    Next item
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    If ws.Shapes.Count < 6 Then Err.Raise vbObjectError + 901, , _
        "The Workflow Controls sheet does not contain all six workflow buttons."
    If CStr(ThisWorkbook.Worksheets("StressRangeID").Range("BN1").Value2) <> "DBM ID" Then _
        Err.Raise vbObjectError + 902, , "StressRangeID!BN:BP is not configured for DBM input."
    For Each item In Array("FatigueID_Hydrotest", "FatigueOD_Hydrotest", _
                           "FatigueID_DesignOp", "FatigueOD_DesignOp")
        If Not IsPositiveNumber(ThisWorkbook.Worksheets(CStr(item)).Range("B1").Value2) Then _
            Err.Raise vbObjectError + 903, , CStr(item) & _
                " needs a positive total cycle count on the INPUT tab."
    Next item
    SetWorkflowStatus "SELF-CHECK PASSED - VBA-only manual workflow is ready"
    Exit Sub
Failed:
    SetWorkflowStatus "PROGRAMMATIC ERROR: " & Err.Description
End Sub

' Used when producing a clean reusable .xlsm template from the original workbook.
' The Boolean argument keeps this procedure out of the Alt+F8 no-argument macro list.
Public Sub InitializeManualFatigueTemplateProgrammatic(ByVal confirmed As Boolean)
    If Not confirmed Then Exit Sub
    PrepareTemplateForDelivery
    SetWorkflowStatus "Ready. Select an input option, phase, and load case."
End Sub

' Used by the workbook builder so the delivered template already contains the
' same native charts that button 5 will refresh after a calculation.
Public Sub RefreshFatigueSummaryChartsProgrammatic(ByVal confirmed As Boolean)
    Dim kpCount As Long
    If Not confirmed Then Exit Sub
    kpCount = ExistingKPCount(ThisWorkbook.Worksheets("StressRangeID"))
    If kpCount <= 0 Then Exit Sub
    BuildSummaryCharts ThisWorkbook.Worksheets("Fatigue_IDSUMEarlyProfile"), True, True, kpCount
    BuildSummaryCharts ThisWorkbook.Worksheets("Fatigue_ODSUMEArlyProfile"), False, True, kpCount
    BuildSummaryCharts ThisWorkbook.Worksheets("Fatigue_IDSUM"), True, False, kpCount
    BuildSummaryCharts ThisWorkbook.Worksheets("Fatigue_ODSUM"), False, False, kpCount
End Sub

' Creates or refreshes the INPUT tab and links the fatigue tabs to it.
' Safe to run again: inputs already on the INPUT tab are kept.
Public Sub BuildFatigueInputSheet()
    Dim oldCalculation As XlCalculation
    On Error GoTo Failed
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    BuildInputSheetCore
    ApplySinglePcdLayout
    UpdateControlSheetForInput
    BuildKpRangeSummary ExistingKPCount(ThisWorkbook.Worksheets("StressRangeID"))
    If Not WorksheetExists(PLOTS_SHEET) Then BuildPlotsSheet 0
    Application.Calculate
    SetWorkflowStatus "Ready. Inputs are on the INPUT tab. Load cases: FCD, HCD, PCD, HYDROTEST 1, HYDROTEST 2, DESIGN."
CleanExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetWorkflowStatus "ERROR: " & Err.Description
    MsgBox "INPUT tab setup failed: " & Err.Description, vbCritical, "Fatigue workflow"
    Resume CleanExit
End Sub

' ============================================================================
' INPUT TAB
' ============================================================================

Private Function FatigueSheetNames() As Variant
    FatigueSheetNames = Array("FatigueID_EarlyUndr", "FatigueID_Earlydr", _
        "FatigueID_Middr", "FatigueID_Latdr", "FatigueID_Hydrotest", _
        "FatigueID_DesignOp", "FatigueOD_EarlyUndr", "FatigueOD_Earlydr", _
        "FatigueOD_Middr", "FatigueOD_Latdr", "FatigueOD_Hydrotest", _
        "FatigueOD_DesignOp")
End Function

Private Function InputSheetOrRaise() As Worksheet
    If Not WorksheetExists(INPUT_SHEET) Then Err.Raise vbObjectError + 800, , _
        "The INPUT tab is missing. Run the macro BuildFatigueInputSheet."
    Set InputSheetOrRaise = ThisWorkbook.Worksheets(INPUT_SHEET)
End Function

Private Function IsPositiveOrZeroNumber(ByVal cellValue As Variant) As Boolean
    If IsEmpty(cellValue) Or IsError(cellValue) Then Exit Function
    If Len(CStr(cellValue)) = 0 Then Exit Function
    IsPositiveOrZeroNumber = IsNumeric(cellValue)
End Function

Private Function IsPositiveNumber(ByVal cellValue As Variant) As Boolean
    If IsEmpty(cellValue) Or IsError(cellValue) Then Exit Function
    If Not IsNumeric(cellValue) Then Exit Function
    IsPositiveNumber = CDbl(cellValue) > 0
End Function

' Blank cycle cells are allowed (the damage columns then stay blank).
Private Function CycleCellInvalid(ByVal cellValue As Variant) As Boolean
    If IsError(cellValue) Then CycleCellInvalid = True: Exit Function
    If Len(CStr(cellValue)) = 0 Then Exit Function
    CycleCellInvalid = Not IsNumeric(cellValue)
End Function

Private Sub BuildInputSheetCore()
    Dim ws As Worksheet, src As Worksheet, sheetNames As Variant, labels As Variant
    Dim j As Long, r As Long, col As Long, n As Long, lastRow As Long
    Dim cell As Range, target As Range, linkAddress As String, isNewDbmTable As Boolean
    Dim c0 As Long, userHeaders As Variant

    If WorksheetExists(INPUT_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = INPUT_SHEET
    End If
    sheetNames = FatigueSheetNames()
    labels = Array("FCD cycles (life tabs) / Hydrotest 1 cycles / Design cycles", _
        "HCD cycles (life tabs) / Hydrotest 2 cycles", "PCD cycles", "SCF (SCFID / SCFOD)", "CorrFactor", _
        "Wall Thickness [mm]", "Tref [mm]", "Exponent k", "WTKcorrection", _
        "DBM S-N curve", "DBM m", "DBM C2", "DBM KDF", _
        "RB S-N curve", "RB m", "RB C2", "RB KDF")

    ws.Range("A1").Value2 = "Fatigue inputs"
    ws.Range("A2").Value2 = "Yellow cells are inputs. Each fatigue tab reads its cells B1:B17 from its column here."
    ws.Cells(INPUT_ROW_OFFSET, 1).Value2 = "Parameter"
    For r = 1 To 17
        ws.Cells(r + INPUT_ROW_OFFSET, 1).Value2 = CStr(labels(r - 1))
    Next r

    For j = 0 To UBound(sheetNames)
        Set src = ThisWorkbook.Worksheets(CStr(sheetNames(j)))
        col = INPUT_FIRST_COLUMN + j
        If InStr(1, src.Name, "FatigueID_", vbTextCompare) = 1 Then
            ws.Cells(INPUT_ROW_OFFSET - 1, col).Value2 = "INNER FIBER (ID)"
        Else
            ws.Cells(INPUT_ROW_OFFSET - 1, col).Value2 = "OUTER FIBER (OD)"
        End If
        ws.Cells(INPUT_ROW_OFFSET, col).Value2 = src.Name
        For r = 1 To 17
            Set cell = src.Cells(r, 2)
            Set target = ws.Cells(r + INPUT_ROW_OFFSET, col)
            If Not (cell.HasFormula And InStr(1, cell.Formula, INPUT_SHEET & "!", vbTextCompare) > 0) Then
                target.NumberFormat = "General"
                If cell.HasFormula Then
                    target.Formula = TranslateInputFormula(cell.Formula, src.Name, ColumnLetter(col))
                Else
                    target.Value2 = cell.Value2
                End If
                linkAddress = INPUT_SHEET & "!" & target.address(True, True)
                cell.Formula = "=IF(" & linkAddress & "="""",""""," & linkAddress & ")"
            End If
        Next r
    Next j

    ' DBM KP ranges: moved here from StressRangeID!BN:BP the first time.
    isNewDbmTable = (CStr(ws.Cells(INPUT_DBM_HEADER_ROW, 1).Value2) <> "DBM ID")
    ws.Cells(INPUT_DBM_HEADER_ROW - 2, 1).Value2 = "DBM KP ranges"
    ws.Cells(INPUT_DBM_HEADER_ROW - 1, 1).Value2 = "KP inside a range uses the DBM S-N curve; all other KP use the RB curve. Button 4 copies this table to StressRangeID!BN:BP."
    ws.Cells(INPUT_DBM_HEADER_ROW, 1).Value2 = "DBM ID"
    ws.Cells(INPUT_DBM_HEADER_ROW, 2).Value2 = "KP Start [m]"
    ws.Cells(INPUT_DBM_HEADER_ROW, 3).Value2 = "KP End [m]"
    If isNewDbmTable Then
        Set src = ThisWorkbook.Worksheets("StressRangeID")
        lastRow = Application.Max(src.Cells(src.Rows.Count, 66).End(xlUp).Row, _
                                  src.Cells(src.Rows.Count, 67).End(xlUp).Row, _
                                  src.Cells(src.Rows.Count, 68).End(xlUp).Row)
        For r = 2 To lastRow
            If Not IsEmpty(src.Cells(r, 67).Value2) Or Not IsEmpty(src.Cells(r, 68).Value2) Then
                n = n + 1
                ws.Cells(INPUT_DBM_HEADER_ROW + n, 1).Value2 = src.Cells(r, 66).Value2
                ws.Cells(INPUT_DBM_HEADER_ROW + n, 2).FormulaR1C1 = src.Cells(r, 67).FormulaR1C1
                ws.Cells(INPUT_DBM_HEADER_ROW + n, 3).FormulaR1C1 = src.Cells(r, 68).FormulaR1C1
            End If
        Next r
        src.Range("BN2:BP1000").Interior.ColorIndex = xlNone
        src.Range("BM1").Value2 = "DBM ranges are entered on the INPUT tab"
    End If

    ' Earlier layout had the optional tables three columns further left: move them
    ' (with any data; formulas that point at them follow) to free D:F.
    If CStr(ws.Cells(INPUT_DBM_HEADER_ROW, 5).Value2) = "Range ID" Then
        ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW - 2, 5), ws.Cells(1000, 20)).Cut _
            Destination:=ws.Cells(INPUT_DBM_HEADER_ROW - 2, 8)
        Application.CutCopyMode = False
    End If
    ws.Cells(INPUT_DBM_HEADER_ROW, 4).Value2 = "First stress row"
    ws.Cells(INPUT_DBM_HEADER_ROW, 5).Value2 = "Last stress row"
    ws.Cells(INPUT_DBM_HEADER_ROW, 6).Value2 = "Match status"
    With ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW, 4), ws.Cells(INPUT_DBM_HEADER_ROW, 6))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW + 1, 4), ws.Cells(INPUT_DBM_HEADER_ROW + 40, 6))
        .Interior.Color = RGB(242, 242, 242)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
    End With
    ws.Cells(INPUT_DBM_HEADER_ROW - 1, 1).Value2 = "KP inside a range uses the DBM S-N curve; other KP use the RB curve. Button 4 fills the grey columns: first and last row of StressRangeID / StressRangeOD inside each range."

    ' User S-N curve table (optional), to the right of the DBM table
    c0 = USER_CURVE_FIRST_COLUMN
    ws.Cells(INPUT_DBM_HEADER_ROW - 2, c0).Value2 = "User S-N curve KP ranges (optional)"
    ws.Cells(INPUT_DBM_HEADER_ROW - 1, c0).Value2 = "KP inside a range uses the curve entered here instead of the DBM/RB curve, on every fatigue tab. Leave the inner or outer curve blank to keep DBM/RB for that fibre. Run button 4 after changing KP ranges."
    userHeaders = Array("Range ID", "KP Start [m]", "KP End [m]", _
        "INNER curve name", "INNER m", "INNER C2", "INNER KDF", _
        "OUTER curve name", "OUTER m", "OUTER C2", "OUTER KDF", "KP rows matched")
    For j = 0 To 11
        ws.Cells(INPUT_DBM_HEADER_ROW, c0 + j).Value2 = CStr(userHeaders(j))
    Next j

    ' Result KP ranges table (optional), to the right of the user-curve table
    ws.Cells(INPUT_DBM_HEADER_ROW - 2, RESULT_RANGE_FIRST_COLUMN).Value2 = "Result KP ranges (optional)"
    ws.Cells(INPUT_DBM_HEADER_ROW - 1, RESULT_RANGE_FIRST_COLUMN).Value2 = "Button 5 reports the maximum total damage and UC inside each range on the " & RANGE_SUMMARY_SHEET & " tab."
    ws.Cells(INPUT_DBM_HEADER_ROW, RESULT_RANGE_FIRST_COLUMN).Value2 = "Range ID"
    ws.Cells(INPUT_DBM_HEADER_ROW, RESULT_RANGE_FIRST_COLUMN + 1).Value2 = "KP Start [m]"
    ws.Cells(INPUT_DBM_HEADER_ROW, RESULT_RANGE_FIRST_COLUMN + 2).Value2 = "KP End [m]"

    ' Formatting
    ws.Cells.Font.Name = "Arial"
    ws.Cells(INPUT_DBM_HEADER_ROW - 2, RESULT_RANGE_FIRST_COLUMN).Font.Size = 12
    ws.Cells(INPUT_DBM_HEADER_ROW - 2, RESULT_RANGE_FIRST_COLUMN).Font.Bold = True
    With ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW, RESULT_RANGE_FIRST_COLUMN), _
                  ws.Cells(INPUT_DBM_HEADER_ROW, RESULT_RANGE_FIRST_COLUMN + 2))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW + 1, RESULT_RANGE_FIRST_COLUMN), _
                  ws.Cells(INPUT_DBM_HEADER_ROW + 40, RESULT_RANGE_FIRST_COLUMN + 2))
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
    End With
    ws.Range(ws.Columns(RESULT_RANGE_FIRST_COLUMN), ws.Columns(RESULT_RANGE_FIRST_COLUMN + 2)).ColumnWidth = 24
    ws.Range(ws.Columns(USER_CURVE_FIRST_COLUMN + 6), ws.Columns(USER_CURVE_FIRST_COLUMN + 11)).ColumnWidth = 18
    ws.Cells(INPUT_DBM_HEADER_ROW - 2, c0).Font.Size = 12
    ws.Cells(INPUT_DBM_HEADER_ROW - 2, c0).Font.Bold = True
    With ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW, c0), ws.Cells(INPUT_DBM_HEADER_ROW, c0 + 11))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW + 1, c0), ws.Cells(INPUT_DBM_HEADER_ROW + 40, c0 + 10))
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
    End With
    With ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW + 1, c0 + 11), ws.Cells(INPUT_DBM_HEADER_ROW + 40, c0 + 11))
        .Interior.Color = RGB(242, 242, 242)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
    End With
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Cells(INPUT_DBM_HEADER_ROW - 2, 1).Font.Size = 12
    ws.Cells(INPUT_DBM_HEADER_ROW - 2, 1).Font.Bold = True
    With ws.Range(ws.Cells(INPUT_ROW_OFFSET - 1, 1), _
                  ws.Cells(INPUT_ROW_OFFSET, INPUT_FIRST_COLUMN + UBound(sheetNames)))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(INPUT_ROW_OFFSET + 1, INPUT_FIRST_COLUMN), _
                  ws.Cells(INPUT_ROW_OFFSET + 17, INPUT_FIRST_COLUMN + UBound(sheetNames)))
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlRight
    End With
    ws.Range(ws.Cells(INPUT_ROW_OFFSET + 1, 1), ws.Cells(INPUT_ROW_OFFSET + 17, 1)).Font.Bold = True
    With ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW, 1), ws.Cells(INPUT_DBM_HEADER_ROW, 3))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW + 1, 1), ws.Cells(INPUT_DBM_HEADER_ROW + 40, 3))
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
    End With
    ws.Columns(1).ColumnWidth = 60
    ws.Range(ws.Columns(INPUT_FIRST_COLUMN), _
             ws.Columns(INPUT_FIRST_COLUMN + UBound(sheetNames))).ColumnWidth = 24
End Sub

' Moves a fatigue-tab input formula to the INPUT tab: references to B1:B17 of
' the same tab become the matching INPUT cells, any other reference is
' qualified with the fatigue tab's name.
Private Function TranslateInputFormula(ByVal formulaText As String, _
                                       ByVal sheetName As String, _
                                       ByVal inputColumn As String) As String
    Dim re As Object, matches As Object, i As Long, m As Object
    Dim prefixText As String, refText As String, newRef As String, rowNumber As Long
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(^|[^A-Z0-9_.!$'])(\$?([A-Z]{1,3})\$?([0-9]+))(?![0-9A-Z_(.!])"
    Set matches = re.Execute(formulaText)
    TranslateInputFormula = formulaText
    For i = matches.Count - 1 To 0 Step -1
        Set m = matches(i)
        prefixText = CStr(m.SubMatches(0))
        refText = CStr(m.SubMatches(1))
        rowNumber = CLng(m.SubMatches(3))
        If UCase$(CStr(m.SubMatches(2))) = "B" And rowNumber >= 1 And rowNumber <= 17 Then
            newRef = inputColumn & CStr(rowNumber + INPUT_ROW_OFFSET)
        Else
            newRef = "'" & sheetName & "'!" & refText
        End If
        TranslateInputFormula = Left$(TranslateInputFormula, m.FirstIndex) & prefixText & _
            newRef & Mid$(TranslateInputFormula, m.FirstIndex + m.Length + 1)
    Next i
End Function

' ---- total damage / UC inside user-specified KP ranges -------------------------

Private Function EnsureRangeSummarySheet() As Worksheet
    Dim previousSheet As Object
    If Not WorksheetExists(RANGE_SUMMARY_SHEET) Then
        Set previousSheet = ActiveSheet
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Fatigue_ODSUM")).Name = _
            RANGE_SUMMARY_SHEET
        On Error Resume Next
        previousSheet.Activate
        On Error GoTo 0
    End If
    Set EnsureRangeSummarySheet = ThisWorkbook.Worksheets(RANGE_SUMMARY_SHEET)
End Function

Private Sub BuildKpRangeSummary(ByVal kpCount As Long)
    Dim ws As Worksheet, inputWs As Worksheet, lastRow As Long, r As Long
    Dim outRow As Long, c0 As Long, lastDataRow As Long, kpRef As String
    Dim idCell As String, startCell As String, endCell As String, headers As Variant, j As Long
    Set inputWs = InputSheetOrRaise()
    Set ws = EnsureRangeSummarySheet()
    c0 = RESULT_RANGE_FIRST_COLUMN
    ws.Cells.ClearContents
    ws.Cells.Font.Name = "Arial"
    ws.Range("A1").Value2 = "Total fatigue damage and UC inside the result KP ranges"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value2 = "Largest total damage over the four angles inside each KP range. UC = damage / allowable damage. The ranges are entered on the INPUT tab; press button 5 after adding or removing ranges."
    ws.Range("E3").Value2 = "INNER FIBER - all life"
    ws.Range("H3").Value2 = "OUTER FIBER - all life"
    ws.Range("K3").Value2 = "INNER FIBER - early profile"
    ws.Range("M3").Value2 = "OUTER FIBER - early profile"
    headers = Array("Range ID", "KP Start [m]", "KP End [m]", "KP rows", _
        "Max total damage", "UC", "KP at max [m]", "Max total damage", "UC", "KP at max [m]", _
        "Max total damage", "UC", "Max total damage", "UC")
    For j = 0 To 13
        ws.Cells(4, 1 + j).Value2 = CStr(headers(j))
    Next j
    With ws.Range("A3:N4")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns("A:N").ColumnWidth = 18
    ws.Range("E5:E1000,H5:H1000,K5:K1000,M5:M1000").NumberFormat = "0.000E+00"
    ws.Range("F5:F1000,I5:I1000,L5:L1000,N5:N1000").NumberFormat = "0.0000"
    ws.Range("B5:C1000,G5:G1000,J5:J1000").NumberFormat = "0.00"

    If kpCount <= 0 Then
        ws.Range("A5").Value2 = "No KP grid yet. Import a stress-range report, then run buttons 4 and 5."
        Exit Sub
    End If
    If CStr(inputWs.Cells(INPUT_DBM_HEADER_ROW, c0).Value2) <> "Range ID" Then Exit Sub
    lastRow = Application.Max(inputWs.Cells(inputWs.Rows.Count, c0 + 1).End(xlUp).Row, _
                              inputWs.Cells(inputWs.Rows.Count, c0 + 2).End(xlUp).Row)
    lastDataRow = FIRST_DATA_ROW + kpCount - 1
    ' KP is read from StressRangeID: the summary tabs' own KP column is empty
    ' when the early-life tabs have no data.
    kpRef = "StressRangeID!$C$" & CStr(FIRST_DATA_ROW) & ":$C$" & CStr(lastDataRow)
    outRow = 5
    For r = INPUT_DBM_HEADER_ROW + 1 To lastRow
        If IsPositiveOrZeroNumber(inputWs.Cells(r, c0 + 1).Value2) And _
           IsPositiveOrZeroNumber(inputWs.Cells(r, c0 + 2).Value2) Then
            idCell = INPUT_SHEET & "!" & inputWs.Cells(r, c0).address(True, True)
            startCell = INPUT_SHEET & "!" & inputWs.Cells(r, c0 + 1).address(True, True)
            endCell = INPUT_SHEET & "!" & inputWs.Cells(r, c0 + 2).address(True, True)
            ws.Cells(outRow, 1).Formula = "=IF(" & idCell & "="""",""""," & idCell & ")"
            ws.Cells(outRow, 2).Formula = "=MIN(" & startCell & "," & endCell & ")"
            ws.Cells(outRow, 3).Formula = "=MAX(" & startCell & "," & endCell & ")"
            ws.Cells(outRow, 4).Formula = "=COUNTIFS(" & kpRef & ","">=""&$B" & CStr(outRow) & _
                "," & kpRef & ",""<=""&$C" & CStr(outRow) & ")"
            WriteRangeBlock ws, outRow, 5, "Fatigue_IDSUM", lastDataRow, True
            WriteRangeBlock ws, outRow, 8, "Fatigue_ODSUM", lastDataRow, True
            WriteRangeBlock ws, outRow, 11, "Fatigue_IDSUMEarlyProfile", lastDataRow, False
            WriteRangeBlock ws, outRow, 13, "Fatigue_ODSUMEArlyProfile", lastDataRow, False
            outRow = outRow + 1
        End If
    Next r
    If outRow = 5 Then ws.Range("A5").Value2 = "No result KP ranges are entered on the INPUT tab."
End Sub

' Writes Max total damage, UC and (optionally) KP at max for one range row.
' Total damage is in columns U:X (21 to 24) of the summary tabs, KP in column D.
Private Sub WriteRangeBlock(ByVal ws As Worksheet, ByVal outRow As Long, _
                            ByVal firstColumn As Long, ByVal summarySheetName As String, _
                            ByVal lastDataRow As Long, ByVal withKp As Boolean)
    Dim kpRef As String, colRef As String, maxArgs As String, kpArgs As String
    Dim c As Long, rowText As String, maxCell As String, columnName As String
    rowText = CStr(outRow)
    kpRef = "StressRangeID!$C$" & CStr(FIRST_DATA_ROW) & ":$C$" & CStr(lastDataRow)
    maxCell = ws.Cells(outRow, firstColumn).address(False, False)
    For c = 21 To 24
        columnName = ColumnLetter(c)
        colRef = summarySheetName & "!$" & columnName & "$" & CStr(FIRST_DATA_ROW) & _
                 ":$" & columnName & "$" & CStr(lastDataRow)
        If Len(maxArgs) > 0 Then maxArgs = maxArgs & ",": kpArgs = kpArgs & ","
        maxArgs = maxArgs & "MAXIFS(" & colRef & "," & kpRef & ","">=""&$B" & rowText & _
                  "," & kpRef & ",""<=""&$C" & rowText & ")"
        kpArgs = kpArgs & "IFERROR(AGGREGATE(15,6," & kpRef & "/((" & kpRef & ">=$B" & rowText & _
                 ")*(" & kpRef & "<=$C" & rowText & ")*(" & colRef & "=" & maxCell & ")),1),1E+99)"
    Next c
    ws.Cells(outRow, firstColumn).Formula = "=IF(OR($D" & rowText & "=0,COUNT(" & _
        summarySheetName & "!$U$" & CStr(FIRST_DATA_ROW) & ":$X$" & CStr(lastDataRow) & _
        ")=0),"""",MAX(" & maxArgs & "))"
    ws.Cells(outRow, firstColumn + 1).Formula = "=IF(" & maxCell & "="""",""""," & maxCell & _
        "/'" & CONTROL_SHEET & "'!$I$8)"
    If withKp Then
        On Error Resume Next
        ws.Cells(outRow, firstColumn + 2).Formula2 = "=IF(" & maxCell & "="""","""",MIN(" & kpArgs & "))"
        If Err.Number <> 0 Then
            Err.Clear
            ws.Cells(outRow, firstColumn + 2).FormulaArray = "=IF(" & maxCell & "="""","""",MIN(" & kpArgs & "))"
        End If
        On Error GoTo 0
    End If
End Sub

' ---- Plots tab: stress range, total damage and UC along KP ---------------------

Private Sub BuildPlotsSheet(ByVal kpCount As Long)
    Dim ws As Worksheet, previousSheet As Object, i As Long, r As Long, f As Long
    Dim labels As Variant, phases As Variant, loadCases As Variant
    Dim startColumn As Long, nextColumn As Long, lastRow As Long, excelRow As Long
    Dim kpRange As Range, stressSheet As Worksheet, sourceRange As Range
    Dim output() As Variant, cellRange As String, sheetName As String
    Dim plotCharts(0 To 5) As Chart, leftPos(0 To 1) As Double, topPos As Double
    Dim summaryName As String, fiberText As String
    Dim xMin As Double, xMax As Double, kpFirst As Double, kpLast As Double, unitStep As Double

    If Not WorksheetExists(PLOTS_SHEET) Then
        Set previousSheet = ActiveSheet
        ThisWorkbook.Worksheets.Add(After:=EnsureRangeSummarySheet()).Name = PLOTS_SHEET
        On Error Resume Next
        previousSheet.Activate
        On Error GoTo 0
    End If
    Set ws = ThisWorkbook.Worksheets(PLOTS_SHEET)
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
    ws.Cells.ClearContents
    ws.Cells.Font.Name = "Arial"
    ws.Range("A1").Value2 = "Stress range, total fatigue damage and UC along KP"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    If kpCount <= 0 Then
        ws.Range("A3").Value2 = "No stress-range data yet. Import a stress-range report, then run buttons 4 and 5."
        AddU2Plot ws, ws.Range("A5").Left, ws.Range("A5").Top, False, 0, 0
        Exit Sub
    End If

    ' KP column of the chart data
    lastRow = PLOTS_DATA_ROW + kpCount - 1
    ws.Cells(PLOTS_DATA_ROW - 2, 1).Value2 = "Chart data: largest |delta S11| over the four angles [MPa] for each imported phase / load case"
    ws.Cells(PLOTS_DATA_ROW - 2, 1).Font.Bold = True
    ws.Cells(PLOTS_DATA_ROW - 1, 1).Value2 = "KP [m]"
    ReDim output(1 To kpCount, 1 To 1)
    For r = 1 To kpCount
        output(r, 1) = "=StressRangeID!C" & CStr(FIRST_DATA_ROW + r - 1)
    Next r
    ws.Cells(PLOTS_DATA_ROW, 1).Resize(kpCount, 1).Formula = output
    Set kpRange = ws.Range(ws.Cells(PLOTS_DATA_ROW, 1), ws.Cells(lastRow, 1))

    ' Fixed KP axis limits (the DBM bands are positioned against them)
    With ThisWorkbook.Worksheets("StressRangeID")
        kpFirst = CDbl(.Cells(FIRST_DATA_ROW, 3).Value2)
        kpLast = CDbl(.Cells(FIRST_DATA_ROW + kpCount - 1, 3).Value2)
    End With
    If kpLast > kpFirst Then
        unitStep = 10 ^ Int(Log(kpLast - kpFirst) / Log(10#)) / 10#
        xMin = Int(kpFirst / unitStep) * unitStep
        xMax = -Int(-kpLast / unitStep) * unitStep
    Else
        xMin = kpFirst: xMax = kpFirst + 1
    End If

    leftPos(0) = ws.Range("A3").Left
    leftPos(1) = leftPos(0) + 740
    topPos = ws.Range("A3").Top
    For f = 0 To 1
        If f = 0 Then fiberText = "ID" Else fiberText = "OD"
        Set plotCharts(f) = NewPlotChart(ws, "PL_Stress" & fiberText, leftPos(f), topPos)
        Set plotCharts(2 + f) = NewPlotChart(ws, "PL_Damage" & fiberText, leftPos(f), topPos + 320)
        Set plotCharts(4 + f) = NewPlotChart(ws, "PL_UC" & fiberText, leftPos(f), topPos + 640)
    Next f

    ' Stress range: one line per imported phase / load case and fibre
    labels = Array("Early undrained FCD", "Early undrained HCD", "Early undrained PCD", _
        "Early drained FCD", "Early drained HCD", "Early drained PCD", _
        "Middle drained FCD", "Middle drained HCD", "Middle drained PCD", _
        "Late drained FCD", "Late drained HCD", "Late drained PCD", "Hydrotest 1", "Hydrotest 2", "Design")
    phases = Array("EARLY UNDRAINED", "EARLY UNDRAINED", "EARLY UNDRAINED", _
        "EARLY DRAINED", "EARLY DRAINED", "EARLY DRAINED", _
        "MIDDLE DRAINED", "MIDDLE DRAINED", "MIDDLE DRAINED", _
        "LATE DRAINED", "LATE DRAINED", "LATE DRAINED", "HYDROTEST", "HYDROTEST 2", "DESIGN OPERATION")
    loadCases = Array("FCD", "HCD", "PCD", "FCD", "HCD", "PCD", "FCD", "HCD", "PCD", _
        "FCD", "HCD", "PCD", "FCD", "FCD", "FCD")
    nextColumn = 2
    For f = 0 To 1
        If f = 0 Then sheetName = "StressRangeID" Else sheetName = "StressRangeOD"
        If f = 0 Then fiberText = "inner" Else fiberText = "outer"
        Set stressSheet = ThisWorkbook.Worksheets(sheetName)
        For i = 0 To UBound(labels)
            startColumn = StressStartColumn(CStr(phases(i)), CStr(loadCases(i)))
            Set sourceRange = stressSheet.Range(stressSheet.Cells(FIRST_DATA_ROW, startColumn), _
                stressSheet.Cells(FIRST_DATA_ROW + kpCount - 1, startColumn + 3))
            If Application.WorksheetFunction.Count(sourceRange) > 0 Then
                ws.Cells(PLOTS_DATA_ROW - 1, nextColumn).Value2 = CStr(labels(i)) & " (" & fiberText & ")"
                For r = 1 To kpCount
                    excelRow = FIRST_DATA_ROW + r - 1
                    cellRange = sheetName & "!" & _
                        stressSheet.Cells(excelRow, startColumn).address(False, False) & ":" & _
                        stressSheet.Cells(excelRow, startColumn + 3).address(False, False)
                    output(r, 1) = "=IF(COUNT(" & cellRange & ")=0,NA(),MAX(MAX(" & cellRange & _
                        "),-MIN(" & cellRange & "))/1000000)"
                Next r
                ws.Cells(PLOTS_DATA_ROW, nextColumn).Resize(kpCount, 1).Formula = output
                AddPlotSeries plotCharts(f), CStr(labels(i)), kpRange, _
                    ws.Range(ws.Cells(PLOTS_DATA_ROW, nextColumn), ws.Cells(lastRow, nextColumn)), 0
                nextColumn = nextColumn + 1
            End If
        Next i
    Next f
    ws.Range(ws.Cells(PLOTS_DATA_ROW - 1, 1), ws.Cells(PLOTS_DATA_ROW - 1, nextColumn)).Font.Bold = True
    ws.Range(ws.Cells(PLOTS_DATA_ROW, 2), ws.Cells(lastRow, nextColumn)).NumberFormat = "0.00"

    ' Total damage (columns U:X) and UC (columns Y:AB) of the all-life summary tabs
    For f = 0 To 1
        If f = 0 Then summaryName = "Fatigue_IDSUM" Else summaryName = "Fatigue_ODSUM"
        With ThisWorkbook.Worksheets(summaryName)
            For i = 0 To 3
                AddPlotSeries plotCharts(2 + f), CStr(Array("-90", "0", "90", "180")(i)) & ChrW$(176), kpRange, _
                    .Range(.Cells(FIRST_DATA_ROW, 21 + i), .Cells(FIRST_DATA_ROW + kpCount - 1, 21 + i)), i + 1
                AddPlotSeries plotCharts(4 + f), CStr(Array("-90", "0", "90", "180")(i)) & ChrW$(176), kpRange, _
                    .Range(.Cells(FIRST_DATA_ROW, 25 + i), .Cells(FIRST_DATA_ROW + kpCount - 1, 25 + i)), i + 1
            Next i
        End With
    Next f

    FinishPlotChart plotCharts(0), "Stress range along KP - inner fibre", "Max |delta S11| over angles [MPa]", "0", xMin, xMax
    FinishPlotChart plotCharts(1), "Stress range along KP - outer fibre", "Max |delta S11| over angles [MPa]", "0", xMin, xMax
    FinishPlotChart plotCharts(2), "Total fatigue damage along KP - inner fibre (all life)", "Total fatigue damage", "0.0E+00", xMin, xMax
    FinishPlotChart plotCharts(3), "Total fatigue damage along KP - outer fibre (all life)", "Total fatigue damage", "0.0E+00", xMin, xMax
    FinishPlotChart plotCharts(4), "Fatigue UC along KP - inner fibre (all life)", "UC = damage / allowable", "0.0E+00", xMin, xMax
    FinishPlotChart plotCharts(5), "Fatigue UC along KP - outer fibre (all life)", "UC = damage / allowable", "0.0E+00", xMin, xMax
    AddU2Plot ws, leftPos(0), topPos + 960, True, xMin, xMax
End Sub

Private Function ImportU2Core(ByVal reportPath As String) As Long
    Dim reportLines As Variant, parts As Variant, headerIndex As Long, i As Long, c As Long
    Dim ws As Worksheet, previousSheet As Object, output() As Variant
    Dim rowCount As Long, columnCount As Long, numberValue As Double
    If Len(Dir$(reportPath)) = 0 Then Err.Raise vbObjectError + 700, , "File not found: " & reportPath
    reportLines = ReadReportLines(reportPath)
    headerIndex = -1
    For i = LBound(reportLines) To UBound(reportLines)
        If UCase$(Left$(Trim$(CStr(reportLines(i))), 2)) = "KP" And InStr(CStr(reportLines(i)), vbTab) > 0 Then
            headerIndex = i
            Exit For
        End If
    Next i
    If headerIndex < 0 Then Err.Raise vbObjectError + 701, , _
        "No header line starting with 'KP' was found. Use U2_steps.rpt or CHECK_..._U2.rpt of extract_stress_ranges.py."
    parts = Split(CStr(reportLines(headerIndex)), vbTab)
    columnCount = UBound(parts) + 1
    If columnCount < 2 Then Err.Raise vbObjectError + 702, , "The U2 report has no step columns."
    ReDim output(1 To UBound(reportLines) - headerIndex + 1, 1 To columnCount)
    For i = headerIndex + 1 To UBound(reportLines)
        parts = Split(CStr(reportLines(i)), vbTab)
        If UBound(parts) >= 1 Then
            If TryNumber(parts(0), numberValue) Then
                rowCount = rowCount + 1
                output(rowCount, 1) = numberValue
                For c = 1 To Application.Min(UBound(parts), columnCount - 1)
                    If TryNumber(parts(c), numberValue) Then
                        output(rowCount, c + 1) = numberValue
                    Else
                        output(rowCount, c + 1) = CVErr(xlErrNA)
                    End If
                Next c
            End If
        End If
    Next i
    If rowCount = 0 Then Err.Raise vbObjectError + 703, , "The U2 report has no data rows."

    If Not WorksheetExists(U2_SHEET) Then
        Set previousSheet = ActiveSheet
        ThisWorkbook.Worksheets.Add(After:=EnsureRangeSummarySheet()).Name = U2_SHEET
        On Error Resume Next
        previousSheet.Activate
        On Error GoTo 0
    End If
    Set ws = ThisWorkbook.Worksheets(U2_SHEET)
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial"
    ws.Range("A1").Value2 = "Local U2 along KP (imported with button 6)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value2 = "Source file: " & reportPath
    parts = Split(CStr(reportLines(headerIndex)), vbTab)
    For c = 0 To columnCount - 1
        ws.Cells(U2_HEADER_ROW, c + 1).Value2 = Trim$(CStr(parts(c)))
    Next c
    ws.Cells(U2_HEADER_ROW, 1).Resize(1, columnCount).Font.Bold = True
    ws.Cells(U2_HEADER_ROW + 1, 1).Resize(rowCount, columnCount).Value = output
    ws.Columns(1).Resize(, columnCount).ColumnWidth = 16
    BuildPlotsSheet ExistingKPCount(ThisWorkbook.Worksheets("StressRangeID"))
    ImportU2Core = columnCount - 1
End Function

' U2 chart of the Plots tab: one line per step of the U2 tab, with the DBM bands.
' useLimits = True keeps the KP axis of the other charts.
Private Sub AddU2Plot(ByVal ws As Worksheet, ByVal chartLeft As Double, ByVal chartTop As Double, _
                      ByVal useLimits As Boolean, ByVal xMin As Double, ByVal xMax As Double)
    Dim u2Ws As Worksheet, lastRow As Long, lastColumn As Long, c As Long
    Dim plotChart As Chart, kpFirst As Double, kpLast As Double, unitStep As Double
    If Not WorksheetExists(U2_SHEET) Then Exit Sub
    Set u2Ws = ThisWorkbook.Worksheets(U2_SHEET)
    lastRow = u2Ws.Cells(u2Ws.Rows.Count, 1).End(xlUp).Row
    lastColumn = u2Ws.Cells(U2_HEADER_ROW, u2Ws.Columns.Count).End(xlToLeft).Column
    If lastRow <= U2_HEADER_ROW + 1 Or lastColumn < 2 Then Exit Sub
    If Not useLimits Then
        kpFirst = CDbl(u2Ws.Cells(U2_HEADER_ROW + 1, 1).Value2)
        kpLast = CDbl(u2Ws.Cells(lastRow, 1).Value2)
        If kpLast > kpFirst Then
            unitStep = 10 ^ Int(Log(kpLast - kpFirst) / Log(10#)) / 10#
            xMin = Int(kpFirst / unitStep) * unitStep
            xMax = -Int(-kpLast / unitStep) * unitStep
        Else
            xMin = kpFirst: xMax = kpFirst + 1
        End If
    End If
    Set plotChart = NewPlotChart(ws, "PL_U2", chartLeft, chartTop)
    For c = 2 To lastColumn
        AddPlotSeries plotChart, CStr(u2Ws.Cells(U2_HEADER_ROW, c).Value2), _
            u2Ws.Range(u2Ws.Cells(U2_HEADER_ROW + 1, 1), u2Ws.Cells(lastRow, 1)), _
            u2Ws.Range(u2Ws.Cells(U2_HEADER_ROW + 1, c), u2Ws.Cells(lastRow, c)), 0
    Next c
    FinishPlotChart plotChart, "Local U2 (lateral displacement) along KP", "U2 [model length unit]", "General", xMin, xMax
    On Error Resume Next
    plotChart.Axes(xlValue).MinimumScaleIsAuto = True
    plotChart.Axes(xlCategory).TickLabelPosition = xlLow
    On Error GoTo 0
End Sub

Private Function NewPlotChart(ByVal ws As Worksheet, ByVal chartName As String, _
                              ByVal chartLeft As Double, ByVal chartTop As Double) As Chart
    Dim plotObject As ChartObject
    Set plotObject = ws.ChartObjects.Add(chartLeft, chartTop, 720, 300)
    plotObject.Name = chartName
    plotObject.Chart.ChartType = xlXYScatterLinesNoMarkers
    Do While plotObject.Chart.SeriesCollection.Count > 0
        plotObject.Chart.SeriesCollection(1).Delete
    Loop
    Set NewPlotChart = plotObject.Chart
End Function

' colorIndex 1 to 4 uses the angle colours of the summary charts; 0 keeps Excel's own.
Private Sub AddPlotSeries(ByVal plotChart As Chart, ByVal seriesName As String, _
                          ByVal xRange As Range, ByVal yRange As Range, ByVal colorIndex As Long)
    Dim plotSeries As Series, seriesColors As Variant
    Set plotSeries = plotChart.SeriesCollection.NewSeries
    plotSeries.Name = seriesName
    plotSeries.XValues = xRange
    plotSeries.values = yRange
    plotSeries.MarkerStyle = xlMarkerStyleNone
    plotSeries.Format.Line.Weight = 1.5
    If colorIndex >= 1 And colorIndex <= 4 Then
        seriesColors = Array(RGB(31, 78, 121), RGB(237, 125, 49), RGB(112, 173, 71), RGB(112, 48, 160))
        plotSeries.Format.Line.ForeColor.RGB = CLng(seriesColors(colorIndex - 1))
    End If
End Sub

Private Sub FinishPlotChart(ByVal plotChart As Chart, ByVal titleText As String, _
                            ByVal yAxisText As String, ByVal yFormat As String, _
                            ByVal xMin As Double, ByVal xMax As Double)
    On Error Resume Next
    With plotChart
        .HasTitle = True
        .ChartTitle.Text = titleText
        .ChartTitle.Font.Name = "Arial"
        .ChartTitle.Font.Size = 12
        .HasLegend = True
        .Legend.position = xlLegendPositionTop
        .DisplayBlanksAs = xlNotPlotted
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Size = 9
        .ChartArea.Format.Line.ForeColor.RGB = RGB(166, 166, 166)
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "KP (m)"
        .Axes(xlCategory).TickLabels.NumberFormat = "0"
        .Axes(xlCategory).MinimumScale = xMin
        .Axes(xlCategory).MaximumScale = xMax
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = yAxisText
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).TickLabels.NumberFormat = yFormat
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
        ' Fixed plot area, so the DBM bands stay aligned with the KP axis
        .PlotArea.InsideLeft = 62
        .PlotArea.InsideTop = 50
        .PlotArea.InsideWidth = 636
        .PlotArea.InsideHeight = 196
    End With
    On Error GoTo 0
    AddDbmBands plotChart, xMin, xMax
End Sub

' Shades the DBM KP ranges of the INPUT tab on a chart: one semi-transparent
' band per range, from KP Start to KP End over the full height of the plot.
Private Sub AddDbmBands(ByVal plotChart As Chart, ByVal xMin As Double, ByVal xMax As Double)
    Dim inputWs As Worksheet, lastRow As Long, r As Long
    Dim startKP As Double, endKP As Double, tempValue As Double
    Dim band As Shape, note As Shape, leftPos As Double, widthPos As Double, bandCount As Long
    If xMax <= xMin Or Not WorksheetExists(INPUT_SHEET) Then Exit Sub
    If plotChart.SeriesCollection.Count = 0 Then Exit Sub
    Set inputWs = ThisWorkbook.Worksheets(INPUT_SHEET)
    lastRow = Application.Max(inputWs.Cells(inputWs.Rows.Count, 2).End(xlUp).Row, _
                              inputWs.Cells(inputWs.Rows.Count, 3).End(xlUp).Row)
    On Error Resume Next
    For r = INPUT_DBM_HEADER_ROW + 1 To lastRow
        If IsPositiveOrZeroNumber(inputWs.Cells(r, 2).Value2) And _
           IsPositiveOrZeroNumber(inputWs.Cells(r, 3).Value2) Then
            startKP = CDbl(inputWs.Cells(r, 2).Value2)
            endKP = CDbl(inputWs.Cells(r, 3).Value2)
            If startKP > endKP Then
                tempValue = startKP: startKP = endKP: endKP = tempValue
            End If
            If startKP < xMin Then startKP = xMin
            If endKP > xMax Then endKP = xMax
            If endKP > startKP Then
                leftPos = plotChart.PlotArea.InsideLeft + _
                    (startKP - xMin) / (xMax - xMin) * plotChart.PlotArea.InsideWidth
                widthPos = (endKP - startKP) / (xMax - xMin) * plotChart.PlotArea.InsideWidth
                If widthPos < 1.5 Then widthPos = 1.5
                Set band = plotChart.Shapes.AddShape(msoShapeRectangle, leftPos, _
                    plotChart.PlotArea.InsideTop, widthPos, plotChart.PlotArea.InsideHeight)
                band.Name = "DBM_band_" & CStr(r)
                band.Fill.ForeColor.RGB = RGB(255, 192, 0)
                band.Fill.Transparency = 0.6
                band.Line.Visible = msoFalse
                bandCount = bandCount + 1
            End If
        End If
    Next r
    If bandCount > 0 Then
        Set note = plotChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 560, 4, 156, 16)
        note.Name = "DBM_band_note"
        note.TextFrame2.TextRange.Text = "Shaded bands = DBM KP ranges"
        note.TextFrame2.TextRange.Font.Size = 8
        note.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(127, 96, 0)
        note.Line.Visible = msoFalse
        note.Fill.Visible = msoFalse
    End If
    On Error GoTo 0
End Sub

' ---- user-specified S-N curves for KP ranges ---------------------------------

Private Sub LoadUserCurveRanges()
    Dim ws As Worksheet, lastRow As Long, r As Long, i As Long, j As Long, c0 As Long
    Dim startValue As Double, endValue As Double, tempValue As Double
    mUserCurveCount = 0
    Set ws = InputSheetOrRaise()
    c0 = USER_CURVE_FIRST_COLUMN
    If CStr(ws.Cells(INPUT_DBM_HEADER_ROW, c0).Value2) <> "Range ID" Then Exit Sub
    lastRow = Application.Max(ws.Cells(ws.Rows.Count, c0 + 1).End(xlUp).Row, _
                              ws.Cells(ws.Rows.Count, c0 + 2).End(xlUp).Row)
    For r = INPUT_DBM_HEADER_ROW + 1 To lastRow
        If IsPositiveOrZeroNumber(ws.Cells(r, c0 + 1).Value2) And _
           IsPositiveOrZeroNumber(ws.Cells(r, c0 + 2).Value2) Then
            mUserCurveCount = mUserCurveCount + 1
            ReDim Preserve mUserCurveStart(1 To mUserCurveCount)
            ReDim Preserve mUserCurveEnd(1 To mUserCurveCount)
            ReDim Preserve mUserCurveRow(1 To mUserCurveCount)
            ReDim Preserve mUserCurveHasId(1 To mUserCurveCount)
            ReDim Preserve mUserCurveHasOd(1 To mUserCurveCount)
            startValue = CDbl(ws.Cells(r, c0 + 1).Value2)
            endValue = CDbl(ws.Cells(r, c0 + 2).Value2)
            If startValue > endValue Then
                tempValue = startValue: startValue = endValue: endValue = tempValue
            End If
            mUserCurveStart(mUserCurveCount) = startValue
            mUserCurveEnd(mUserCurveCount) = endValue
            mUserCurveRow(mUserCurveCount) = r
            mUserCurveHasId(mUserCurveCount) = UserCurveDefined(ws, r, c0 + 4, "INNER")
            mUserCurveHasOd(mUserCurveCount) = UserCurveDefined(ws, r, c0 + 8, "OUTER")
            If Not mUserCurveHasId(mUserCurveCount) And Not mUserCurveHasOd(mUserCurveCount) Then _
                Err.Raise vbObjectError + 810, , "INPUT row " & CStr(r) & _
                    ": enter m, C2 and KDF for the INNER and/or OUTER curve of this KP range."
        End If
    Next r
    For i = 1 To mUserCurveCount - 1
        For j = i + 1 To mUserCurveCount
            If mUserCurveStart(i) <= mUserCurveEnd(j) And mUserCurveStart(j) <= mUserCurveEnd(i) Then _
                Err.Raise vbObjectError + 811, , "User S-N curve KP ranges overlap: INPUT rows " & _
                    CStr(mUserCurveRow(i)) & " and " & CStr(mUserCurveRow(j)) & "."
        Next j
    Next i
End Sub

Private Function UserCurveDefined(ByVal ws As Worksheet, ByVal r As Long, _
                                  ByVal firstColumn As Long, ByVal fiberText As String) As Boolean
    Dim c As Long, filled As Long, cellValue As Variant
    For c = firstColumn To firstColumn + 2
        cellValue = ws.Cells(r, c).Value2
        If IsError(cellValue) Then
            filled = filled + 1
        ElseIf Len(CStr(cellValue)) > 0 Then
            filled = filled + 1
        End If
    Next c
    If filled = 0 Then Exit Function
    For c = firstColumn To firstColumn + 2
        If Not IsPositiveNumber(ws.Cells(r, c).Value2) Then Err.Raise vbObjectError + 812, , _
            "INPUT row " & CStr(r) & ": the " & fiberText & " S-N curve needs positive m, C2 and KDF."
    Next c
    UserCurveDefined = True
End Function

' Replaces the DBM/RB curve cell references with the user curve when the KP
' lies in a user range that defines a curve for this fibre.
Private Sub ApplyUserCurve(ByVal kp As Double, ByVal innerFiber As Boolean, _
                           ByRef curveM As String, ByRef curveC As String, _
                           ByRef curveKdf As String)
    Dim i As Long, firstColumn As Long, ws As Worksheet
    For i = 1 To mUserCurveCount
        If kp >= mUserCurveStart(i) And kp <= mUserCurveEnd(i) Then
            If innerFiber Then
                If Not mUserCurveHasId(i) Then Exit Sub
                firstColumn = USER_CURVE_FIRST_COLUMN + 4
            Else
                If Not mUserCurveHasOd(i) Then Exit Sub
                firstColumn = USER_CURVE_FIRST_COLUMN + 8
            End If
            Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
            curveM = INPUT_SHEET & "!" & ws.Cells(mUserCurveRow(i), firstColumn).address(True, True)
            curveC = INPUT_SHEET & "!" & ws.Cells(mUserCurveRow(i), firstColumn + 1).address(True, True)
            curveKdf = INPUT_SHEET & "!" & ws.Cells(mUserCurveRow(i), firstColumn + 2).address(True, True)
            Exit Sub
        End If
    Next i
End Sub

Private Sub WriteUserCurveMatches(ByVal kpCount As Long)
    Dim ws As Worksheet, stressSheet As Worksheet, kpValues As Variant
    Dim i As Long, r As Long, hits As Long, kp As Double
    If Not WorksheetExists(INPUT_SHEET) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
    If CStr(ws.Cells(INPUT_DBM_HEADER_ROW, USER_CURVE_FIRST_COLUMN).Value2) <> "Range ID" Then Exit Sub
    ws.Range(ws.Cells(INPUT_DBM_HEADER_ROW + 1, USER_CURVE_FIRST_COLUMN + 11), _
             ws.Cells(INPUT_DBM_HEADER_ROW + 1000, USER_CURVE_FIRST_COLUMN + 11)).ClearContents
    If mUserCurveCount = 0 Then Exit Sub
    Set stressSheet = ThisWorkbook.Worksheets("StressRangeID")
    kpValues = stressSheet.Range(stressSheet.Cells(FIRST_DATA_ROW, 3), _
                                 stressSheet.Cells(FIRST_DATA_ROW + kpCount, 3)).Value2
    For i = 1 To mUserCurveCount
        hits = 0
        For r = 1 To kpCount
            kp = CDbl(kpValues(r, 1))
            If kp >= mUserCurveStart(i) And kp <= mUserCurveEnd(i) Then hits = hits + 1
        Next r
        ws.Cells(mUserCurveRow(i), USER_CURVE_FIRST_COLUMN + 11).Value2 = hits
    Next i
End Sub

' Moves a four-column stress-range block (rows 3 to the last data row) from an
' old position to the new one if the old one holds data, then empties the old one.
Private Sub MoveStressBlock(ByVal ws As Worksheet, ByVal oldColumn As Long, ByVal newColumn As Long)
    Dim oldData As Range, labelText As String
    Set oldData = ws.Range(ws.Cells(FIRST_DATA_ROW, oldColumn), ws.Cells(LAST_TEMPLATE_ROW, oldColumn + 3))
    On Error Resume Next
    ws.Range(ws.Cells(3, oldColumn), ws.Cells(3, oldColumn + 3)).UnMerge
    On Error GoTo 0
    If Application.WorksheetFunction.Count(oldData) > 0 Then
        labelText = CStr(ws.Cells(3, oldColumn).Value2)
        ws.Range(ws.Cells(FIRST_DATA_ROW, newColumn), _
                 ws.Cells(LAST_TEMPLATE_ROW, newColumn + 3)).Value2 = oldData.Value2
        On Error Resume Next
        ws.Range(ws.Cells(3, newColumn), ws.Cells(3, newColumn + 3)).UnMerge
        On Error GoTo 0
        ws.Range(ws.Cells(3, newColumn), ws.Cells(3, newColumn + 3)).Merge
        ws.Cells(3, newColumn).Value2 = labelText
        ws.Range(ws.Cells(4, newColumn), ws.Cells(5, newColumn + 3)).Value2 = _
            ws.Range(ws.Cells(4, oldColumn), ws.Cells(5, oldColumn + 3)).Value2
    End If
    ws.Range(ws.Cells(1, oldColumn), ws.Cells(LAST_TEMPLATE_ROW, oldColumn + 3)).ClearContents
End Sub

Private Sub ApplySinglePcdLayout()
    Dim sheetName As Variant, ws As Worksheet, firstColumn As Long, c As Long
    Dim startColumn As Variant
    For Each sheetName In Array("FatigueID_EarlyUndr", "FatigueID_Earlydr", _
        "FatigueID_Middr", "FatigueID_Latdr", "FatigueOD_EarlyUndr", _
        "FatigueOD_Earlydr", "FatigueOD_Middr", "FatigueOD_Latdr")
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        If InStr(1, ws.Name, "FatigueID_", vbTextCompare) = 1 Then firstColumn = 5 Else firstColumn = 6
        For c = 0 To 3
            ws.Cells(1, firstColumn + 8 + c).Value2 = "PCD"
        Next c
        ws.Range(ws.Cells(2, firstColumn + 8), ws.Cells(2, firstColumn + 11)).ClearContents
        ws.Range(ws.Cells(1, firstColumn + 12), _
                 ws.Cells(LAST_TEMPLATE_ROW, firstColumn + 15)).ClearContents
        ws.Range(ws.Columns(firstColumn + 12), ws.Columns(firstColumn + 15)).Hidden = True
    Next sheetName
    For Each sheetName In Array("StressRangeID", "StressRangeOD")
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        For Each startColumn In Array(24, 40, 56, 103)
            On Error Resume Next
            ws.Range(ws.Cells(3, CLng(startColumn)), ws.Cells(3, CLng(startColumn) + 3)).UnMerge
            On Error GoTo 0
            ws.Range(ws.Cells(1, CLng(startColumn)), _
                     ws.Cells(LAST_TEMPLATE_ROW, CLng(startColumn) + 3)).ClearContents
            ws.Range(ws.Columns(CLng(startColumn)), ws.Columns(CLng(startColumn) + 3)).Hidden = True
        Next startColumn
    Next sheetName
    For Each sheetName In Array("StressRangeID", "StressRangeOD")
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        MoveStressBlock ws, 111, 4
        MoveStressBlock ws, 115, 12
    Next sheetName
    For Each sheetName In Array("Fatigue_IDSUMEarlyProfile", "Fatigue_ODSUMEArlyProfile", _
        "Fatigue_IDSUM", "Fatigue_ODSUM")
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        ws.Range("Q1:T" & LAST_TEMPLATE_ROW).ClearContents
        ws.Range("Q:T").EntireColumn.Hidden = True
    Next sheetName
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_IDSUMEarlyProfile"), True, True
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_ODSUMEArlyProfile"), False, True
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_IDSUM"), True, False
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_ODSUM"), False, False
    If WorksheetExists(CONTROL_SHEET) Then
        If UCase$(Trim$(CStr(ControlSheet.Range(LOAD_CASE_CELL).Value2))) Like "PCD#" Then _
            ControlSheet.Range(LOAD_CASE_CELL).Value2 = "PCD"
    End If
End Sub

Private Sub UpdateControlSheetForInput()
    If Not WorksheetExists(CONTROL_SHEET) Then Exit Sub
    With ControlSheet
        On Error Resume Next
        .Range(LOAD_CASE_CELL).Validation.Delete
        On Error GoTo 0
        .Range(LOAD_CASE_CELL).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, _
            LOAD_CASE_LIST
        On Error Resume Next
        .Range(PHASE_CELL).Validation.Delete
        On Error GoTo 0
        .Range(PHASE_CELL).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, _
            PHASE_LIST
        .Range("H19").Value2 = CONTROL_NOTE_1
        .Range("H21").Value2 = CONTROL_NOTE_2
    End With
    On Error Resume Next
    ControlSheet.Shapes("FD_U2").Delete
    On Error GoTo 0
    AddWorkflowShape ControlSheet, "FD_U2", "6. Import U2 Report (optional plot)", _
        "ImportU2Report", "B22", RGB(89, 89, 89)
End Sub

' Reads the phase and load case from Workflow Controls. HYDROTEST or DESIGN
' as load case selects that block whatever the target phase is.
Private Sub ReadPhaseAndLoadCase(ByRef phaseName As String, ByRef loadCase As String)
    loadCase = NormalizeLoadCase(CStr(ControlSheet.Range(LOAD_CASE_CELL).Value2))
    If loadCase = "HYDROTEST" Or loadCase = "HYDROTEST 2" Or loadCase = "DESIGN OPERATION" Then
        phaseName = loadCase
        loadCase = "FCD"
    Else
        phaseName = NormalizePhase(CStr(ControlSheet.Range(PHASE_CELL).Value2))
    End If
End Sub

' ============================================================================
' SETUP HELPERS
' ============================================================================

Private Sub EnsureLateLifeSheets()
    Dim ws As Worksheet
    If Not WorksheetExists("FatigueID_Latdr") Then
        ThisWorkbook.Worksheets("FatigueID_Middr").Copy _
            After:=ThisWorkbook.Worksheets("FatigueID_Middr")
        Set ws = ActiveSheet
        ws.Name = "FatigueID_Latdr"
        ws.Range("B1:B3").ClearContents
    End If
    If Not WorksheetExists("FatigueOD_Latdr") Then
        ThisWorkbook.Worksheets("FatigueOD_Middr").Copy _
            After:=ThisWorkbook.Worksheets("FatigueOD_Middr")
        Set ws = ActiveSheet
        ws.Name = "FatigueOD_Latdr"
        ws.Range("B1:B3").ClearContents
    End If
End Sub

Private Sub EnsureHydrotestAndDesignSheets()
    Dim ws As Worksheet
    If Not WorksheetExists("FatigueID_Hydrotest") Then
        ThisWorkbook.Worksheets("FatigueID_EarlyUndr").Copy _
            After:=ThisWorkbook.Worksheets("FatigueID_Latdr")
        Set ws = ActiveSheet
        ws.Name = "FatigueID_Hydrotest"
    End If
    If Not WorksheetExists("FatigueID_DesignOp") Then
        ThisWorkbook.Worksheets("FatigueID_EarlyUndr").Copy _
            After:=ThisWorkbook.Worksheets("FatigueID_Hydrotest")
        Set ws = ActiveSheet
        ws.Name = "FatigueID_DesignOp"
    End If
    If Not WorksheetExists("FatigueOD_Hydrotest") Then
        ThisWorkbook.Worksheets("FatigueOD_EarlyUndr").Copy _
            After:=ThisWorkbook.Worksheets("FatigueOD_Latdr")
        Set ws = ActiveSheet
        ws.Name = "FatigueOD_Hydrotest"
    End If
    If Not WorksheetExists("FatigueOD_DesignOp") Then
        ThisWorkbook.Worksheets("FatigueOD_EarlyUndr").Copy _
            After:=ThisWorkbook.Worksheets("FatigueOD_Hydrotest")
        Set ws = ActiveSheet
        ws.Name = "FatigueOD_DesignOp"
    End If

    ConfigureSupplementalFatigueSheet ThisWorkbook.Worksheets("FatigueID_Hydrotest"), _
        "HYDROTEST", 2#, True
    ConfigureSupplementalFatigueSheet ThisWorkbook.Worksheets("FatigueID_DesignOp"), _
        "DESIGN OPERATION", 1#, True
    ConfigureSupplementalFatigueSheet ThisWorkbook.Worksheets("FatigueOD_Hydrotest"), _
        "HYDROTEST", 2#, False
    ConfigureSupplementalFatigueSheet ThisWorkbook.Worksheets("FatigueOD_DesignOp"), _
        "DESIGN OPERATION", 1#, False
End Sub

Private Sub ConfigureSupplementalFatigueSheet(ByVal ws As Worksheet, _
                                               ByVal phaseName As String, _
                                               ByVal cycleCount As Double, _
                                               ByVal innerFiber As Boolean)
    Dim kpColumn As Long, firstDamageColumn As Long, c As Long, b As Long, blockCount As Long
    Dim angleList As Variant, fiberText As String, displayPhaseName As String
    If innerFiber Then
        kpColumn = 4
        fiberText = "INNER"
    Else
        kpColumn = 5
        fiberText = "OUTER"
    End If
    firstDamageColumn = kpColumn + 1
    angleList = Array(-90#, 0#, 90#, 180#)
    If phaseName = "DESIGN OPERATION" Then _
        displayPhaseName = "DESIGN OP." Else displayPhaseName = phaseName

    On Error Resume Next
    ws.Range("C1:D1").UnMerge
    On Error GoTo 0
    ws.Range("D1:U14").ClearContents
    ws.Range(ws.Cells(FIRST_DATA_ROW, kpColumn), _
             ws.Cells(LAST_TEMPLATE_ROW, firstDamageColumn + 15)).ClearContents
    ws.Range("A2:A3").ClearContents
    If phaseName = "HYDROTEST" Then
        ws.Range("A1").Value2 = "Hydrotest 1 cycles="
        ws.Range("A2").Value2 = "Hydrotest 2 cycles="
    Else
        ws.Range("A1").Value2 = "Total cycles="
    End If
    If Not ws.Range("B1").HasFormula And IsEmpty(ws.Range("B1").Value2) Then _
        ws.Range("B1").Value2 = cycleCount
    ws.Range("C1:D1").Merge
    ws.Range("C1").Value2 = "Source: no data"
    ws.Range("C1").WrapText = True
    ws.Range("D5").Value2 = "DBM Max"
    ws.Range("D8").Value2 = "RB Max"
    ws.Cells(14, kpColumn).Value2 = "KP [m]"
    If phaseName = "HYDROTEST" Then blockCount = 2 Else blockCount = 1
    For b = 0 To blockCount - 1
        If phaseName = "HYDROTEST" Then displayPhaseName = "HYDROTEST " & CStr(b + 1)
        For c = 0 To 3
            ws.Cells(1, firstDamageColumn + 4 * b + c).Value2 = displayPhaseName
            ws.Cells(3, firstDamageColumn + 4 * b + c).Value2 = "Radius = " & fiberText & " FIBER"
            ws.Cells(4, firstDamageColumn + 4 * b + c).Value2 = "Angle = " & CStr(angleList(c))
        Next c
        ws.Cells(2, firstDamageColumn + 4 * b).Formula = "=$B$" & CStr(b + 1) & "&"" cycles"""
    Next b
    ws.Range(ws.Cells(1, firstDamageColumn), _
             ws.Cells(1, firstDamageColumn + 4 * blockCount - 1)).EntireColumn.ColumnWidth = 20
    ws.Range(ws.Cells(FIRST_DATA_ROW, firstDamageColumn), _
             ws.Cells(LAST_TEMPLATE_ROW, firstDamageColumn + 4 * blockCount - 1)).NumberFormat = "0.000E+00"
End Sub

Private Sub ConfigureFatigueParameterBlocks()
    Dim sheetName As Variant, ws As Worksheet
    For Each sheetName In Array("FatigueID_EarlyUndr", "FatigueID_Earlydr", _
        "FatigueID_Middr", "FatigueID_Latdr", "FatigueID_Hydrotest", _
        "FatigueID_DesignOp", "FatigueOD_EarlyUndr", "FatigueOD_Earlydr", _
        "FatigueOD_Middr", "FatigueOD_Latdr", "FatigueOD_Hydrotest", _
        "FatigueOD_DesignOp")
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        ws.Range("A10").Value2 = "S-N curve"
        ws.Range("C10").Value2 = "DBM"
        ws.Range("A11").Value2 = "m="
        ws.Range("A12").Value2 = "C2="
        ws.Range("A13").Value2 = "KDFID="
        ws.Range("A14").Value2 = "S-N curve"
        ws.Range("C14").Value2 = "RB"
        ws.Range("A15").Value2 = "m="
        ws.Range("A16").Value2 = "C2="
        ws.Range("A17").Value2 = "KDFID="
        ws.Range("A10:C13").Interior.Color = RGB(221, 235, 247)
        ws.Range("A14:C17").Interior.Color = RGB(226, 239, 218)
        ws.Range("C10").Font.Bold = True
        ws.Range("C14").Font.Bold = True
    Next sheetName
End Sub

Private Sub ConfigureStressRangeDbmArea()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("StressRangeID")
    ws.Range("BN1").Value2 = "DBM ID"
    ws.Range("BO1").Value2 = "KP Start"
    ws.Range("BP1").Value2 = "KP End"
    ws.Range("BQ1").Value2 = "Matched start row"
    ws.Range("BR1").Value2 = "Matched end row"
    ws.Range("BS1").Value2 = "Matched start KP"
    ws.Range("BT1").Value2 = "Matched end KP"
    ws.Range("BU1").Value2 = "Match status"
    ws.Range("BW1").Value2 = "Segment type"
    ws.Range("BX1").Value2 = "Segment ID"
    ws.Range("BY1").Value2 = "Row start"
    ws.Range("BZ1").Value2 = "Row end"
    ws.Range("CA1").Value2 = "KP start"
    ws.Range("CB1").Value2 = "KP end"
    ws.Range("BN1:BU1").Font.Bold = True
    ws.Range("BN1:BU1").Font.Color = vbWhite
    ws.Range("BN1:BU1").Interior.Color = RGB(31, 78, 121)
    ws.Range("BW1:CB1").Font.Bold = True
    ws.Range("BW1:CB1").Font.Color = vbWhite
    ws.Range("BW1:CB1").Interior.Color = RGB(68, 114, 196)
    ws.Range("BN2:BP1000").Interior.ColorIndex = xlNone
    ws.Range("BM1").Value2 = "DBM ranges are entered on the INPUT tab"
    ws.Columns("BN:CB").ColumnWidth = 16

    Set ws = ThisWorkbook.Worksheets("StressRangeOD")
    ws.Range("BN1:CO1000").ClearContents
    ws.Range("BN1").Value2 = "DBM KP ranges are entered on the INPUT tab; matched rows are in StressRangeID!BN:CB"
End Sub

Private Sub BuildWorkflowControlsSheet()
    Dim ws As Worksheet, i As Long
    If WorksheetExists(CONTROL_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONTROL_SHEET
    End If
    ws.Cells.Clear
    For i = ws.Shapes.Count To 1 Step -1: ws.Shapes(i).Delete: Next i
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Range("B1").Value2 = "Manual Fatigue Workflow Controls"
    ws.Range("B1").Font.Name = "Arial"
    ws.Range("B1").Font.Size = 16
    ws.Range("B1").Font.Bold = True
    ws.Range("B2").Value2 = "Nothing runs automatically. Use buttons 1 to 5 in sequence. Delta S11 skips buttons 2 and 3."
    ws.Range("B2").Font.Name = "Arial"
    ws.Range("B2").Font.Size = 10

    ws.Range("H3:J3").Merge
    ws.Range("H3").Value2 = "Workflow settings"
    ws.Range("H3:J3").Font.Bold = True
    ws.Range("H3:J3").Font.Color = vbWhite
    ws.Range("H3:J3").Interior.Color = RGB(31, 78, 121)
    ws.Range("H4").Value2 = "Input option"
    ws.Range("H5").Value2 = "Target phase"
    ws.Range("H6").Value2 = "Load case"
    ws.Range("H7").Value2 = "Pair-name filter"
    ws.Range("H8").Value2 = "Allowable damage"
    ws.Range("H10").Value2 = "Status"
    ws.Range("H11").Value2 = "Last imported mode"
    ws.Range("H12").Value2 = "KP count"
    ws.Range("H13").Value2 = "Heat-up step"
    ws.Range("H14").Value2 = "Cool-down step"
    ws.Range("H15").Value2 = "Selected pair"
    ws.Range("H16").Value2 = "Source file"
    ws.Range(MODE_CELL).Value2 = "AUTO"
    ws.Range(PHASE_CELL).Value2 = "EARLY UNDRAINED"
    ws.Range(LOAD_CASE_CELL).Value2 = "FCD"
    ws.Range(PAIR_FILTER_CELL).Value2 = vbNullString
    ws.Range(ALLOWABLE_CELL).Value2 = 0.128
    ws.Range(LAST_MODE_CELL & ":" & SOURCE_FILE_CELL).ClearContents
    ws.Range("I10:M10").Merge
    ws.Range("I16:M16").Merge
    ws.Range("I10:M10").WrapText = True
    ws.Range("I16:M16").WrapText = True
    ws.Range("H4:H16").Font.Bold = True
    ws.Range("I4:I8").Interior.Color = RGB(255, 242, 204)
    ws.Range("H4:M16").Borders.LineStyle = xlContinuous
    ws.Range("H4:M16").Borders.Color = RGB(217, 217, 217)

    ws.Range("H18:M18").Merge
    ws.Range("H18").Value2 = "Before button 4"
    ws.Range("H18:M18").Font.Bold = True
    ws.Range("H18:M18").Font.Color = vbWhite
    ws.Range("H18:M18").Interior.Color = RGB(68, 114, 196)
    ws.Range("H19:M20").Merge
    ws.Range("H19").Value2 = CONTROL_NOTE_1
    ws.Range("H21:M22").Merge
    ws.Range("H21").Value2 = CONTROL_NOTE_2
    ws.Range("H19:M22").WrapText = True
    ws.Range("H19:M22").VerticalAlignment = xlCenter
    ws.Range("H18:M22").Borders.LineStyle = xlContinuous
    ws.Range("H18:M22").Borders.Color = RGB(217, 217, 217)

    On Error Resume Next
    ws.Range(MODE_CELL).Validation.Delete
    ws.Range(PHASE_CELL).Validation.Delete
    ws.Range(LOAD_CASE_CELL).Validation.Delete
    On Error GoTo 0
    ws.Range(MODE_CELL).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, _
        "AUTO,S11,DELTA S11"
    ws.Range(PHASE_CELL).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, _
        PHASE_LIST
    ws.Range(LOAD_CASE_CELL).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, _
        LOAD_CASE_LIST

    AddWorkflowShape ws, "FD_RunAll", "Run All Applicable Workflows", _
        "RunAllManualFatigueWorkflows", "B4", RGB(31, 78, 121)
    AddWorkflowShape ws, "FD_Import", "1. Import S11 or Delta S11 Report", _
        "ImportS11OrDeltaS11Report", "B7", RGB(91, 155, 213)
    AddWorkflowShape ws, "FD_Separate", "2. Separate Sheet1 into ID and OD", _
        "SeparateSheet1IntoIdOd", "B10", RGB(112, 173, 71)
    AddWorkflowShape ws, "FD_Stress", "3. Calculate Stress Ranges", _
        "CalculateStressRanges", "B13", RGB(237, 125, 49)
    AddWorkflowShape ws, "FD_Fatigue", "4. Calculate Fatigue Damage", _
        "CalculateFatigueDamage", "B16", RGB(112, 48, 160)
    AddWorkflowShape ws, "FD_Summary", "5. Summarize Damage, UC, and Charts", _
        "SummarizeFatigueDamageAndUC", "B19", RGB(0, 112, 192)
    AddWorkflowShape ws, "FD_U2", "6. Import U2 Report (optional plot)", _
        "ImportU2Report", "B22", RGB(89, 89, 89)

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 2
    ws.Columns("H").ColumnWidth = 20
    ws.Columns("I:M").ColumnWidth = 18
    ws.Rows("1:22").RowHeight = 21
    ws.Rows("10:10").RowHeight = 32
    ws.Rows("16:16").RowHeight = 32
    ws.Rows("19:22").RowHeight = 24
    ws.Range("B1:M22").Font.Name = "Arial"
End Sub

Private Sub AddWorkflowShape(ByVal ws As Worksheet, ByVal shapeName As String, _
                             ByVal captionText As String, ByVal macroName As String, _
                             ByVal anchorCell As String, ByVal fillColor As Long)
    Dim Shape As Shape, anchor As Range
    Set anchor = ws.Range(anchorCell)
    Set Shape = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, 280, 43.2)
    Shape.Name = shapeName
    Shape.OnAction = macroName
    Shape.Fill.ForeColor.RGB = fillColor
    Shape.Line.ForeColor.RGB = RGB(255, 255, 255)
    Shape.TextFrame2.TextRange.Text = captionText
    Shape.TextFrame2.TextRange.Font.Name = "Arial"
    Shape.TextFrame2.TextRange.Font.Size = 13
    Shape.TextFrame2.TextRange.Font.Bold = msoTrue
    Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Shape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Shape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
End Sub

' ============================================================================
' BUTTON 1: IMPORT S11 OR DIRECT DELTA S11
' ============================================================================

Private Function ImportFatigueReportCore(ByVal suppliedPath As String) As String
    Dim reportPath As Variant, reportLines As Variant, headerIndex As Long
    Dim headerText As String, requestedMode As String, detectedMode As String
    Dim preamble As Collection

    If Len(suppliedPath) = 0 Then
        reportPath = Application.GetOpenFilename( _
            "Report files (*.rpt;*.txt),*.rpt;*.txt,All files (*.*),*.*", _
            , "Select an S11 or delta-S11 report")
        If VarType(reportPath) = vbBoolean Then Exit Function
    Else
        reportPath = suppliedPath
    End If
    If Len(Dir$(CStr(reportPath))) = 0 Then Err.Raise vbObjectError + 100, , _
        "Report file not found: " & CStr(reportPath)

    reportLines = ReadReportLines(CStr(reportPath))
    Set preamble = New Collection
    headerIndex = FindReportHeaderIndex(reportLines, preamble)
    If headerIndex < 0 Then Err.Raise vbObjectError + 101, , _
        "A Pipe Distance or Pipeline Distance header was not found."
    headerText = CStr(reportLines(headerIndex))
    If InStr(1, headerText, "HEATUP LAST S11", vbTextCompare) > 0 And _
       InStr(1, headerText, "COOLDOWN LAST S11", vbTextCompare) > 0 Then
        detectedMode = "S11"
    ElseIf InStr(1, headerText, "DELTA_S11", vbTextCompare) > 0 And _
           InStr(1, headerText, "MAX_DELTA_S11", vbTextCompare) = 0 And _
           InStr(1, headerText, "MIN_DELTA_S11", vbTextCompare) = 0 Then
        detectedMode = "DELTA S11"
    Else
        Err.Raise vbObjectError + 102, , _
            "The report is not a supported last-frame S11 or direct delta-S11 report."
    End If

    requestedMode = NormalizeInputMode(CStr(ControlSheet.Range(MODE_CELL).Value2))
    If requestedMode <> "AUTO" And requestedMode <> detectedMode Then _
        Err.Raise vbObjectError + 103, , "The selected input option is " & requestedMode & _
            ", but the report was detected as " & detectedMode & "."

    If Len(Trim$(CStr(ControlSheet.Range(LAST_MODE_CELL).Value2))) = 0 Then _
        ClearLegacyWorkflowData

    If detectedMode = "S11" Then
        ImportS11ToSheet1 reportLines, headerIndex, preamble
    Else
        ImportDeltaDirect reportLines, headerIndex
    End If
    ControlSheet.Range(LAST_MODE_CELL).Value2 = detectedMode
    ControlSheet.Range(SOURCE_FILE_CELL).Value2 = CStr(reportPath)
    ImportFatigueReportCore = detectedMode
End Function

Private Sub ImportS11ToSheet1(ByVal reportLines As Variant, ByVal headerIndex As Long, _
                              ByVal preamble As Collection)
    Dim ws As Worksheet, headers As Variant, fields As Variant, output() As Variant
    Dim lineIndex As Long, dataCount As Long, rowIndex As Long, columnIndex As Long
    Dim columnCount As Long, kp As Double, valueNumber As Double, item As Variant
    Dim firstStep As String, secondStep As String, previousCount As Long, clearRow As Long

    headers = Split(CStr(reportLines(headerIndex)), vbTab)
    columnCount = UBound(headers) + 1
    For lineIndex = headerIndex + 1 To UBound(reportLines)
        If Len(Trim$(CStr(reportLines(lineIndex)))) > 0 Then
            fields = Split(CStr(reportLines(lineIndex)), vbTab)
            If UBound(fields) >= 0 And TryNumber(fields(0), kp) Then dataCount = dataCount + 1
        End If
    Next lineIndex
    If dataCount = 0 Then Err.Raise vbObjectError + 104, , "The S11 report contains no numeric KP rows."
    ReDim output(1 To dataCount + 1, 1 To columnCount)
    For columnIndex = 0 To UBound(headers)
        output(1, columnIndex + 1) = CStr(headers(columnIndex))
    Next columnIndex
    rowIndex = 1
    For lineIndex = headerIndex + 1 To UBound(reportLines)
        If Len(Trim$(CStr(reportLines(lineIndex)))) > 0 Then
            fields = Split(CStr(reportLines(lineIndex)), vbTab)
            If UBound(fields) >= 0 And TryNumber(fields(0), kp) Then
                rowIndex = rowIndex + 1
                For columnIndex = 0 To UBound(headers)
                    If columnIndex <= UBound(fields) Then
                        If TryNumber(fields(columnIndex), valueNumber) Then
                            output(rowIndex, columnIndex + 1) = valueNumber
                        Else
                            output(rowIndex, columnIndex + 1) = CStr(fields(columnIndex))
                        End If
                    End If
                Next columnIndex
            End If
        End If
    Next lineIndex

    Set ws = ThisWorkbook.Worksheets("Sheet1")
    If IsNumeric(ControlSheet.Range(KP_COUNT_CELL).Value2) Then _
        previousCount = CLng(ControlSheet.Range(KP_COUNT_CELL).Value2)
    clearRow = Application.Max(previousCount + 1, dataCount + 1, 2)
    ws.Range("A1:V" & CStr(clearRow)).ClearContents
    ws.Range("A1").Resize(dataCount + 1, columnCount).Value2 = output
    FormatImportedTable ws, dataCount + 1, columnCount

    For Each item In preamble
        If InStr(1, CStr(item), "First step:", vbTextCompare) = 1 Then _
            firstStep = StepNameFromPreamble(CStr(item))
        If InStr(1, CStr(item), "Second step:", vbTextCompare) = 1 Then _
            secondStep = StepNameFromPreamble(CStr(item))
    Next item
    If Len(firstStep) = 0 Then firstStep = "HEAT-UP LAST"
    If Len(secondStep) = 0 Then secondStep = "COOL-DOWN LAST"
    ControlSheet.Range(FIRST_STEP_CELL).Value2 = firstStep
    ControlSheet.Range(SECOND_STEP_CELL).Value2 = secondStep
    ControlSheet.Range(SELECTED_PAIR_CELL).Value2 = firstStep & " - " & secondStep
    ControlSheet.Range(KP_COUNT_CELL).Value2 = dataCount
End Sub

Private Sub ImportDeltaDirect(ByVal reportLines As Variant, ByVal headerIndex As Long)
    Dim headers As Variant, fields As Variant, pairs As Object, selectedPair As String
    Dim metaFiber() As String, metaAngle() As Double, metaUse() As Boolean
    Dim idValues As Object, odValues As Object, kpValues As Object
    Dim lineIndex As Long, i As Long, kp As Double, valueS11 As Double
    Dim pairName As String, fiberName As String, angleValue As Double, key As String
    Dim existing As Double, phaseName As String, loadCase As String

    mCurrentStage = "button 1 delta header parsing"
    headers = Split(CStr(reportLines(headerIndex)), vbTab)
    Set pairs = CreateObject("Scripting.Dictionary")
    pairs.CompareMode = vbTextCompare
    For i = 1 To UBound(headers)
        pairName = BracketText(CStr(headers(i)))
        If Len(pairName) > 0 Then pairs(pairName) = True
    Next i
    If pairs.Count = 0 Then Err.Raise vbObjectError + 105, , _
        "No step-pair names were found in the delta-S11 header."
    selectedPair = SelectPairName(pairs, Trim$(CStr(ControlSheet.Range(PAIR_FILTER_CELL).Value2)))

    ReDim metaFiber(0 To UBound(headers))
    ReDim metaAngle(0 To UBound(headers))
    ReDim metaUse(0 To UBound(headers))
    For i = 1 To UBound(headers)
        If StrComp(BracketText(CStr(headers(i))), selectedPair, vbTextCompare) = 0 Then
            fiberName = FiberFromHeader(CStr(headers(i)))
            angleValue = AngleFromHeader(CStr(headers(i)))
            If (fiberName = "INNER" Or fiberName = "OUTER") And IsTargetAngle(angleValue) Then
                metaFiber(i) = fiberName
                metaAngle(i) = angleValue
                metaUse(i) = True
            End If
        End If
    Next i

    mCurrentStage = "button 1 delta data parsing"
    Set idValues = CreateObject("Scripting.Dictionary")
    Set odValues = CreateObject("Scripting.Dictionary")
    Set kpValues = CreateObject("Scripting.Dictionary")
    For lineIndex = headerIndex + 1 To UBound(reportLines)
        mCurrentStage = "button 1 delta data line " & CStr(lineIndex + 1)
        If Len(Trim$(CStr(reportLines(lineIndex)))) > 0 Then
            fields = Split(CStr(reportLines(lineIndex)), vbTab)
            If UBound(fields) >= 0 And TryNumber(fields(0), kp) Then
                kpValues(CanonicalNumber(kp)) = kp
                For i = 1 To UBound(headers)
                    mCurrentStage = "button 1 delta data line " & CStr(lineIndex + 1) & _
                        ", column " & CStr(i + 1)
                    If metaUse(i) And i <= UBound(fields) Then
                        If TryNumber(fields(i), valueS11) Then
                            key = CanonicalNumber(metaAngle(i)) & SEP & CanonicalNumber(kp)
                            If metaFiber(i) = "INNER" Then
                                AddLargestAbsolute idValues, key, valueS11
                            Else
                                AddLargestAbsolute odValues, key, valueS11
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    Next lineIndex
    If idValues.Count = 0 Or odValues.Count = 0 Then Err.Raise vbObjectError + 106, , _
        "The selected pair does not contain both INNER and OUTER four-angle delta S11 data."

    mCurrentStage = "button 1 KP-grid preparation"
    EnsureStressKPGrid kpValues
    ReadPhaseAndLoadCase phaseName, loadCase
    mCurrentStage = "button 1 direct delta write"
    WriteStressBlock ThisWorkbook.Worksheets("StressRangeID"), idValues, kpValues, _
        phaseName, loadCase, "INNER", selectedPair
    WriteStressBlock ThisWorkbook.Worksheets("StressRangeOD"), odValues, kpValues, _
        phaseName, loadCase, "OUTER", selectedPair
    ControlSheet.Range(SELECTED_PAIR_CELL).Value2 = selectedPair
    ControlSheet.Range(FIRST_STEP_CELL & ":" & SECOND_STEP_CELL).ClearContents
    ControlSheet.Range(KP_COUNT_CELL).Value2 = kpValues.Count
End Sub

' ============================================================================
' BUTTON 2: SEPARATE SHEET1 INTO ID AND OD
' ============================================================================

Private Sub SeparateSheet1IntoIdOdCore()
    Dim sourceSheet As Worksheet, idSheet As Worksheet, odSheet As Worksheet
    Dim lastRow As Long, lastColumn As Long, idColumns As Collection, odColumns As Collection
    Dim columnIndex As Long, headerText As String
    Set sourceSheet = ThisWorkbook.Worksheets("Sheet1")
    If IsNumeric(ControlSheet.Range(KP_COUNT_CELL).Value2) Then _
        lastRow = CLng(ControlSheet.Range(KP_COUNT_CELL).Value2) + 1
    If lastRow < 2 Then lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row
    lastColumn = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastColumn < 4 Then Err.Raise vbObjectError + 200, , _
        "Sheet1 does not contain an imported S11 table. Run button 1 first."

    Set idColumns = New Collection
    Set odColumns = New Collection
    For columnIndex = 1 To Application.Min(3, lastColumn)
        idColumns.Add columnIndex
        odColumns.Add columnIndex
    Next columnIndex
    For columnIndex = 4 To lastColumn
        headerText = UCase$(CStr(sourceSheet.Cells(1, columnIndex).Value2))
        If InStr(1, headerText, "INNER", vbTextCompare) > 0 Then idColumns.Add columnIndex
        If InStr(1, headerText, "OUTER", vbTextCompare) > 0 Then odColumns.Add columnIndex
    Next columnIndex
    If idColumns.Count <= 3 Or odColumns.Count <= 3 Then Err.Raise vbObjectError + 201, , _
        "Sheet1 does not contain both INNER and OUTER S11 columns."

    Set idSheet = ThisWorkbook.Worksheets("ID")
    Set odSheet = ThisWorkbook.Worksheets("OD")
    CopySelectedColumns sourceSheet, idSheet, idColumns, lastRow
    CopySelectedColumns sourceSheet, odSheet, odColumns, lastRow
End Sub

Private Sub CopySelectedColumns(ByVal sourceSheet As Worksheet, ByVal targetSheet As Worksheet, _
                                ByVal sourceColumns As Collection, ByVal lastRow As Long)
    Dim output() As Variant, r As Long, c As Long, sourceColumn As Long
    Dim clearLastRow As Long, clearLastColumn As Long
    ReDim output(1 To lastRow, 1 To sourceColumns.Count)
    For c = 1 To sourceColumns.Count
        sourceColumn = CLng(sourceColumns(c))
        For r = 1 To lastRow
            output(r, c) = sourceSheet.Cells(r, sourceColumn).Value2
        Next r
    Next c
    clearLastRow = LastUsedRow(targetSheet)
    clearLastColumn = LastUsedColumn(targetSheet)
    If clearLastRow > 0 And clearLastColumn > 0 Then _
        targetSheet.Range(targetSheet.Cells(1, 1), _
                          targetSheet.Cells(clearLastRow, clearLastColumn)).ClearContents
    targetSheet.Range("A1").Resize(lastRow, sourceColumns.Count).Value2 = output
    FormatImportedTable targetSheet, lastRow, sourceColumns.Count
End Sub

' ============================================================================
' BUTTON 3: CALCULATE STRESS RANGES
' ============================================================================

Private Sub CalculateStressRangesCore()
    Dim idValues As Object, odValues As Object, idKPs As Object, odKPs As Object
    Dim phaseName As String, loadCase As String, pairName As String
    Set idValues = CreateObject("Scripting.Dictionary")
    Set odValues = CreateObject("Scripting.Dictionary")
    Set idKPs = CreateObject("Scripting.Dictionary")
    Set odKPs = CreateObject("Scripting.Dictionary")
    ReadDeltasFromSeparatedSheet ThisWorkbook.Worksheets("ID"), idValues, idKPs
    ReadDeltasFromSeparatedSheet ThisWorkbook.Worksheets("OD"), odValues, odKPs
    If idKPs.Count = 0 Or odKPs.Count = 0 Then Err.Raise vbObjectError + 300, , _
        "ID or OD does not contain numeric KP rows. Run button 2 first."
    VerifySameKPGrid idKPs, odKPs
    EnsureStressKPGrid idKPs

    ReadPhaseAndLoadCase phaseName, loadCase
    pairName = Trim$(CStr(ControlSheet.Range(SELECTED_PAIR_CELL).Value2))
    If Len(pairName) = 0 Then pairName = "HEAT-UP LAST - COOL-DOWN LAST"
    WriteStressBlock ThisWorkbook.Worksheets("StressRangeID"), idValues, idKPs, _
        phaseName, loadCase, "INNER", pairName
    WriteStressBlock ThisWorkbook.Worksheets("StressRangeOD"), odValues, odKPs, _
        phaseName, loadCase, "OUTER", pairName
    ControlSheet.Range(KP_COUNT_CELL).Value2 = idKPs.Count
End Sub

Private Sub ReadDeltasFromSeparatedSheet(ByVal ws As Worksheet, ByVal values As Object, _
                                         ByVal kpValues As Object)
    Dim heatColumns As Object, coolColumns As Object, lastRow As Long, lastColumn As Long
    Dim columnIndex As Long, rowIndex As Long, headerText As String, angleValue As Double
    Dim keyAngle As String, kp As Double, heatValue As Double, coolValue As Double
    Dim valueKey As String, deltaValue As Double, item As Variant
    Set heatColumns = CreateObject("Scripting.Dictionary")
    Set coolColumns = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For columnIndex = 2 To lastColumn
        headerText = CStr(ws.Cells(1, columnIndex).Value2)
        If InStr(1, headerText, "S11", vbTextCompare) > 0 Then
            angleValue = AngleFromHeader(headerText)
            If IsTargetAngle(angleValue) Then
                keyAngle = CanonicalNumber(angleValue)
                If InStr(1, headerText, "HEATUP LAST S11", vbTextCompare) > 0 Then _
                    heatColumns(keyAngle) = columnIndex
                If InStr(1, headerText, "COOLDOWN LAST S11", vbTextCompare) > 0 Then _
                    coolColumns(keyAngle) = columnIndex
            End If
        End If
    Next columnIndex
    For Each item In Array(-90#, 0#, 90#, 180#)
        keyAngle = CanonicalNumber(CDbl(item))
        If Not heatColumns.Exists(keyAngle) Or Not coolColumns.Exists(keyAngle) Then _
            Err.Raise vbObjectError + 301, , ws.Name & _
                " is missing heat-up or cool-down S11 at angle " & CStr(item) & "."
    Next item
    For rowIndex = 2 To lastRow
        If TryNumber(ws.Cells(rowIndex, 1).Value2, kp) Then
            kpValues(CanonicalNumber(kp)) = kp
            For Each item In heatColumns.keys
                If coolColumns.Exists(item) Then
                    If TryNumber(ws.Cells(rowIndex, CLng(heatColumns(item))).Value2, heatValue) And _
                       TryNumber(ws.Cells(rowIndex, CLng(coolColumns(item))).Value2, coolValue) Then
                        deltaValue = heatValue - coolValue
                        valueKey = CStr(item) & SEP & CanonicalNumber(kp)
                        AddLargestAbsolute values, valueKey, deltaValue
                    End If
                End If
            Next item
        End If
    Next rowIndex
End Sub

Private Sub VerifySameKPGrid(ByVal firstGrid As Object, ByVal secondGrid As Object)
    Dim key As Variant
    If firstGrid.Count <> secondGrid.Count Then Err.Raise vbObjectError + 302, , _
        "The ID and OD KP counts are different."
    For Each key In firstGrid.keys
        If Not secondGrid.Exists(CStr(key)) Then Err.Raise vbObjectError + 303, , _
            "The ID and OD KP values do not match at KP " & CStr(firstGrid(key)) & "."
    Next key
End Sub

Private Sub EnsureStressKPGrid(ByVal kpValues As Object)
    Dim numbers() As Double, existingCount As Long, i As Long, mismatch As Boolean
    Dim ws As Worksheet, output() As Variant, answer As VbMsgBoxResult
    If kpValues.Count = 0 Then Err.Raise vbObjectError + 304, , "No numeric KP values were found."
    If FIRST_DATA_ROW + kpValues.Count - 1 > LAST_TEMPLATE_ROW Then _
        Err.Raise vbObjectError + 305, , "The report has more KP rows than the workbook template supports."
    numbers = SortedKPs(kpValues)
    Set ws = ThisWorkbook.Worksheets("StressRangeID")
    existingCount = ExistingKPCount(ws)
    If existingCount > 0 Then
        If existingCount <> kpValues.Count Then
            mismatch = True
        Else
            For i = 1 To existingCount
                If Abs(CDbl(ws.Cells(FIRST_DATA_ROW + i - 1, 3).Value2) - numbers(i)) > 0.001 Then
                    mismatch = True
                    Exit For
                End If
            Next i
        End If
    End If
    If mismatch Then
        If mProgrammaticMode Then
            ResetStressGrid
        Else
            answer = MsgBox("The imported KP grid differs from the current stress-range grid." & _
                vbCrLf & "Replace the KP grid and clear all existing stress-range results?", _
                vbYesNo + vbQuestion, "Replace KP grid")
            If answer <> vbYes Then Err.Raise vbObjectError + 306, , "The KP-grid replacement was cancelled."
            ResetStressGrid
        End If
        existingCount = 0
    End If
    If existingCount = 0 Then
        ReDim output(1 To kpValues.Count, 1 To 2)
        For i = 1 To kpValues.Count
            output(i, 1) = i
            output(i, 2) = numbers(i)
        Next i
        For Each ws In ArrayWorksheet("StressRangeID", "StressRangeOD")
            ws.Range("B" & FIRST_DATA_ROW).Resize(kpValues.Count, 2).Value2 = output
        Next ws
    End If
End Sub

Private Sub ResetStressGrid()
    Dim sheetName As Variant, ws As Worksheet
    For Each sheetName In Array("StressRangeID", "StressRangeOD")
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        ws.Range("B" & FIRST_DATA_ROW & ":BK" & LAST_TEMPLATE_ROW).ClearContents
        ws.Range("CQ" & FIRST_DATA_ROW & ":DF" & LAST_TEMPLATE_ROW).ClearContents
        ws.Range("DG" & FIRST_DATA_ROW & ":DN" & LAST_TEMPLATE_ROW).ClearContents
    Next sheetName
End Sub

Private Sub WriteStressBlock(ByVal ws As Worksheet, ByVal values As Object, _
                             ByVal kpValues As Object, ByVal phaseName As String, _
                             ByVal loadCase As String, ByVal fiberName As String, _
                             ByVal pairName As String)
    Dim numbers() As Double, output() As Variant, angleList As Variant
    Dim startColumn As Long, r As Long, c As Long, key As String
    numbers = SortedKPs(kpValues)
    startColumn = StressStartColumn(phaseName, loadCase)
    angleList = Array(-90#, 0#, 90#, 180#)
    ws.Range(ws.Cells(FIRST_DATA_ROW, startColumn), _
             ws.Cells(LAST_TEMPLATE_ROW, startColumn + 3)).ClearContents
    ReDim output(1 To UBound(numbers), 1 To 4)
    For r = 1 To UBound(numbers)
        For c = 0 To 3
            key = CanonicalNumber(CDbl(angleList(c))) & SEP & CanonicalNumber(numbers(r))
            If values.Exists(key) Then output(r, c + 1) = CDbl(values(key))
        Next c
    Next r
    On Error Resume Next
    ws.Range(ws.Cells(3, startColumn), ws.Cells(3, startColumn + 3)).UnMerge
    On Error GoTo 0
    ws.Range(ws.Cells(3, startColumn), ws.Cells(3, startColumn + 3)).Merge
    If NormalizePhase(phaseName) = "HYDROTEST" Or _
       NormalizePhase(phaseName) = "HYDROTEST 2" Or _
       NormalizePhase(phaseName) = "DESIGN OPERATION" Then
        If NormalizePhase(phaseName) = "HYDROTEST" Then
            ws.Cells(3, startColumn).Value2 = "HYDROTEST 1 / " & pairName
        Else
            ws.Cells(3, startColumn).Value2 = NormalizePhase(phaseName) & " / " & pairName
        End If
    Else
        ws.Cells(3, startColumn).Value2 = phaseName & " / " & loadCase & " / " & pairName
    End If
    For c = 0 To 3
        ws.Cells(4, startColumn + c).Value2 = "Radius = " & fiberName & " FIBER"
        ws.Cells(5, startColumn + c).Value2 = "Angle = " & CStr(angleList(c))
    Next c
    ws.Cells(FIRST_DATA_ROW, startColumn).Resize(UBound(numbers), 4).Value2 = output
    ws.Cells(FIRST_DATA_ROW, startColumn).Resize(UBound(numbers), 4).NumberFormat = "0.000000E+00"
    For c = 0 To 3
        If ws.Columns(startColumn + c).ColumnWidth < 14 Then ws.Columns(startColumn + c).ColumnWidth = 14
    Next c
End Sub

' ============================================================================
' BUTTON 4: DBM/RB CLASSIFICATION AND FATIGUE DAMAGE
' ============================================================================

Private Sub CalculateFatigueDamageCore()
    Dim kpCount As Long
    EnsureLateLifeSheets
    EnsureHydrotestAndDesignSheets
    ConfigureFatigueParameterBlocks
    kpCount = ExistingKPCount(ThisWorkbook.Worksheets("StressRangeID"))
    If kpCount <= 0 Then Err.Raise vbObjectError + 400, , _
        "No KP grid is available in StressRangeID. Import delta S11 or run buttons 1 to 3 first."
    BuildDbmAndRbSegments kpCount
    LoadUserCurveRanges
    WriteUserCurveMatches kpCount
    BuildFatigueSheet "FatigueID_EarlyUndr", "StressRangeID", "EARLY UNDRAINED", kpCount, True
    BuildFatigueSheet "FatigueID_Earlydr", "StressRangeID", "EARLY DRAINED", kpCount, True
    BuildFatigueSheet "FatigueID_Middr", "StressRangeID", "MIDDLE DRAINED", kpCount, True
    BuildFatigueSheet "FatigueID_Latdr", "StressRangeID", "LATE DRAINED", kpCount, True
    BuildFatigueSheet "FatigueOD_EarlyUndr", "StressRangeOD", "EARLY UNDRAINED", kpCount, False
    BuildFatigueSheet "FatigueOD_Earlydr", "StressRangeOD", "EARLY DRAINED", kpCount, False
    BuildFatigueSheet "FatigueOD_Middr", "StressRangeOD", "MIDDLE DRAINED", kpCount, False
    BuildFatigueSheet "FatigueOD_Latdr", "StressRangeOD", "LATE DRAINED", kpCount, False
    BuildSupplementalFatigueSheet "FatigueID_Hydrotest", "StressRangeID", _
        "HYDROTEST", kpCount, True, 2#
    BuildSupplementalFatigueSheet "FatigueID_DesignOp", "StressRangeID", _
        "DESIGN OPERATION", kpCount, True, 1#
    BuildSupplementalFatigueSheet "FatigueOD_Hydrotest", "StressRangeOD", _
        "HYDROTEST", kpCount, False, 2#
    BuildSupplementalFatigueSheet "FatigueOD_DesignOp", "StressRangeOD", _
        "DESIGN OPERATION", kpCount, False, 1#
End Sub

Private Sub BuildDbmAndRbSegments(ByVal kpCount As Long)
    Dim ws As Worksheet, lastInputRow As Long, inputRow As Long, rangeCount As Long
    Dim ids() As String, starts() As Double, ends() As Double
    Dim startValue As Double, endValue As Double, tempValue As Double, idValue As String
    Dim i As Long, j As Long, startIndex As Long, endIndex As Long, statusText As String
    Dim segmentStart As Long, segmentType As String, currentType As String
    Dim segmentId As String, currentId As String, segmentRow As Long, kp As Double
    Dim inputWs As Worksheet, inputRows() As Long, rowSwap As Long
    Dim kpValues As Variant, firstRow As Long, lastRow As Long, r As Long
    Set ws = ThisWorkbook.Worksheets("StressRangeID")
    Set inputWs = InputSheetOrRaise()
    lastInputRow = Application.Max(inputWs.Cells(inputWs.Rows.Count, 1).End(xlUp).Row, _
                                   inputWs.Cells(inputWs.Rows.Count, 2).End(xlUp).Row, _
                                   inputWs.Cells(inputWs.Rows.Count, 3).End(xlUp).Row)
    For inputRow = INPUT_DBM_HEADER_ROW + 1 To lastInputRow
        If IsPositiveOrZeroNumber(inputWs.Cells(inputRow, 2).Value2) And _
           IsPositiveOrZeroNumber(inputWs.Cells(inputRow, 3).Value2) Then
            rangeCount = rangeCount + 1
            ReDim Preserve ids(1 To rangeCount)
            ReDim Preserve starts(1 To rangeCount)
            ReDim Preserve ends(1 To rangeCount)
            idValue = Trim$(CStr(inputWs.Cells(inputRow, 1).Value2))
            If Len(idValue) = 0 Then idValue = "DBM-" & CStr(rangeCount)
            startValue = CDbl(inputWs.Cells(inputRow, 2).Value2)
            endValue = CDbl(inputWs.Cells(inputRow, 3).Value2)
            If startValue > endValue Then
                tempValue = startValue: startValue = endValue: endValue = tempValue
            End If
            ids(rangeCount) = idValue
            starts(rangeCount) = startValue
            ends(rangeCount) = endValue
            ReDim Preserve inputRows(1 To rangeCount)
            inputRows(rangeCount) = inputRow
        End If
    Next inputRow
    For i = 1 To rangeCount - 1
        For j = i + 1 To rangeCount
            If starts(j) < starts(i) Then
                tempValue = starts(i): starts(i) = starts(j): starts(j) = tempValue
                tempValue = ends(i): ends(i) = ends(j): ends(j) = tempValue
                idValue = ids(i): ids(i) = ids(j): ids(j) = idValue
                rowSwap = inputRows(i): inputRows(i) = inputRows(j): inputRows(j) = rowSwap
            End If
        Next j
    Next i
    For i = 2 To rangeCount
        If starts(i) <= ends(i - 1) Then Err.Raise vbObjectError + 401, , _
            "DBM KP ranges overlap: " & ids(i - 1) & " and " & ids(i) & "."
    Next i

    ws.Range("BN2:CO1000").ClearContents
    For i = 1 To rangeCount
        inputRow = i + 1
        ws.Cells(inputRow, 66).Value2 = ids(i)
        ws.Cells(inputRow, 67).Value2 = starts(i)
        ws.Cells(inputRow, 68).Value2 = ends(i)
        If ends(i) < CDbl(ws.Cells(FIRST_DATA_ROW, 3).Value2) Or _
           starts(i) > CDbl(ws.Cells(FIRST_DATA_ROW + kpCount - 1, 3).Value2) Then
            statusText = "OUTSIDE KP GRID"
        Else
            startIndex = FindFloorKPIndex(starts(i), kpCount)
            endIndex = FindFloorKPIndex(ends(i), kpCount)
            ws.Cells(inputRow, 69).Value2 = FIRST_DATA_ROW + startIndex - 1
            ws.Cells(inputRow, 70).Value2 = FIRST_DATA_ROW + endIndex - 1
            ws.Cells(inputRow, 71).Value2 = ws.Cells(FIRST_DATA_ROW + startIndex - 1, 3).Value2
            ws.Cells(inputRow, 72).Value2 = ws.Cells(FIRST_DATA_ROW + endIndex - 1, 3).Value2
            statusText = "MATCHED"
        End If
        ws.Cells(inputRow, 73).Value2 = statusText
    Next i

    ' INPUT tab: first and last stress-range row whose KP lies inside each DBM range
    inputWs.Range(inputWs.Cells(INPUT_DBM_HEADER_ROW + 1, 4), inputWs.Cells(1000, 6)).ClearContents
    kpValues = ws.Range(ws.Cells(FIRST_DATA_ROW, 3), ws.Cells(FIRST_DATA_ROW + kpCount, 3)).Value2
    For i = 1 To rangeCount
        firstRow = 0: lastRow = 0
        For r = 1 To kpCount
            If CDbl(kpValues(r, 1)) >= starts(i) And CDbl(kpValues(r, 1)) <= ends(i) Then
                If firstRow = 0 Then firstRow = FIRST_DATA_ROW + r - 1
                lastRow = FIRST_DATA_ROW + r - 1
            End If
        Next r
        If firstRow > 0 Then
            inputWs.Cells(inputRows(i), 4).Value2 = firstRow
            inputWs.Cells(inputRows(i), 5).Value2 = lastRow
            inputWs.Cells(inputRows(i), 6).Value2 = "MATCHED (" & CStr(lastRow - firstRow + 1) & " rows)"
        Else
            inputWs.Cells(inputRows(i), 6).Value2 = "NO KP IN RANGE"
        End If
    Next i

    segmentRow = 2
    segmentStart = 1
    currentId = DbmIdForKP(CDbl(ws.Cells(FIRST_DATA_ROW, 3).Value2), ids, starts, ends, rangeCount)
    If Len(currentId) > 0 Then currentType = "DBM" Else currentType = "RB"
    For i = 2 To kpCount + 1
        If i <= kpCount Then
            kp = CDbl(ws.Cells(FIRST_DATA_ROW + i - 1, 3).Value2)
            segmentId = DbmIdForKP(kp, ids, starts, ends, rangeCount)
            If Len(segmentId) > 0 Then segmentType = "DBM" Else segmentType = "RB"
        Else
            segmentType = "END"
            segmentId = vbNullString
        End If
        If segmentType <> currentType Or segmentId <> currentId Then
            ws.Cells(segmentRow, 75).Value2 = currentType
            If currentType = "DBM" Then ws.Cells(segmentRow, 76).Value2 = currentId Else _
                ws.Cells(segmentRow, 76).Value2 = "RB-" & CStr(segmentRow - 1)
            ws.Cells(segmentRow, 77).Value2 = FIRST_DATA_ROW + segmentStart - 1
            ws.Cells(segmentRow, 78).Value2 = FIRST_DATA_ROW + i - 2
            ws.Cells(segmentRow, 79).Value2 = ws.Cells(FIRST_DATA_ROW + segmentStart - 1, 3).Value2
            ws.Cells(segmentRow, 80).Value2 = ws.Cells(FIRST_DATA_ROW + i - 2, 3).Value2
            segmentRow = segmentRow + 1
            segmentStart = i
            currentType = segmentType
            currentId = segmentId
        End If
    Next i
End Sub

Private Sub BuildFatigueSheet(ByVal fatigueSheetName As String, _
                              ByVal stressSheetName As String, _
                              ByVal phaseName As String, ByVal kpCount As Long, _
                              ByVal innerFiber As Boolean)
    Dim ws As Worksheet, stressSheet As Worksheet, output() As Variant
    Dim kpColumn As Long, firstDamageColumn As Long, rowIndex As Long
    Dim loadCaseIndex As Long, angleIndex As Long, sourceColumn As Long
    Dim loadCase As String, cycleAddress As String, sourceAddress As String
    Dim formulaText As String, curveM As String, curveC As String, curveKdf As String
    Dim kp As Double, isDbm As Boolean, hasPhaseData As Boolean
    Dim sourcePhaseName As String
    Set ws = ThisWorkbook.Worksheets(fatigueSheetName)
    Set stressSheet = ThisWorkbook.Worksheets(stressSheetName)
    If innerFiber Then kpColumn = 4 Else kpColumn = 5
    firstDamageColumn = kpColumn + 1
    sourcePhaseName = phaseName
    hasPhaseData = PhaseHasStressData(stressSheet, sourcePhaseName, kpCount)
    If Not hasPhaseData And (NormalizePhase(phaseName) = "MIDDLE DRAINED" Or _
                             NormalizePhase(phaseName) = "LATE DRAINED") Then
        If PhaseHasStressData(stressSheet, "EARLY DRAINED", kpCount) Then
            sourcePhaseName = "EARLY DRAINED"
            hasPhaseData = True
        ElseIf PhaseHasStressData(stressSheet, "EARLY UNDRAINED", kpCount) Then
            sourcePhaseName = "EARLY UNDRAINED"
            hasPhaseData = True
        End If
    End If
    ws.Range(ws.Cells(FIRST_DATA_ROW, kpColumn), _
             ws.Cells(LAST_TEMPLATE_ROW, firstDamageColumn + 15)).ClearContents
    ws.Range("C1").Value2 = "Stress source: " & sourcePhaseName
    If Not hasPhaseData Then
        ws.Range("C1").Value2 = "Stress source: no data"
        Exit Sub
    End If
    ValidateFatigueInputs ws, sourcePhaseName

    ReDim output(1 To kpCount, 1 To 13)
    For rowIndex = 1 To kpCount
        output(rowIndex, 1) = "='" & stressSheetName & "'!C" & _
            CStr(FIRST_DATA_ROW + rowIndex - 1)
        kp = CDbl(stressSheet.Cells(FIRST_DATA_ROW + rowIndex - 1, 3).Value2)
        isDbm = IsDbmKP(kp)
        If isDbm Then
            curveM = "$B$11": curveC = "$B$12": curveKdf = "$B$13"
        Else
            curveM = "$B$15": curveC = "$B$16": curveKdf = "$B$17"
        End If
        ApplyUserCurve kp, innerFiber, curveM, curveC, curveKdf
        For loadCaseIndex = 0 To 2
            loadCase = Array("FCD", "HCD", "PCD")(loadCaseIndex)
            If loadCase = "FCD" Then
                cycleAddress = "$B$1"
            ElseIf loadCase = "HCD" Then
                cycleAddress = "$B$2"
            Else
                cycleAddress = "$B$3"
            End If
            sourceColumn = StressStartColumn(sourcePhaseName, loadCase)
            For angleIndex = 0 To 3
                sourceAddress = "'" & stressSheetName & "'!" & _
                    stressSheet.Cells(FIRST_DATA_ROW + rowIndex - 1, _
                    sourceColumn + angleIndex).address(False, False)
                formulaText = "=IF(OR(" & sourceAddress & "=""""," & _
                    "NOT(ISNUMBER(" & cycleAddress & "))),""""," & cycleAddress & _
                    "/(" & curveC & "*((ABS(" & sourceAddress & ")/1000000)*" & _
                    "$B$4*$B$5*$B$9^$B$8)^(-" & curveM & ")*(1/" & curveKdf & ")))"
                output(rowIndex, 2 + loadCaseIndex * 4 + angleIndex) = formulaText
            Next angleIndex
        Next loadCaseIndex
    Next rowIndex
    ws.Cells(14, kpColumn).Value2 = "KP [m]"
    ws.Cells(FIRST_DATA_ROW, kpColumn).Resize(kpCount, 13).Formula = output
    ws.Cells(FIRST_DATA_ROW, firstDamageColumn).Resize(kpCount, 12).NumberFormat = "0.000E+00"
    ws.Range(ws.Cells(5, firstDamageColumn + 12), ws.Cells(8, firstDamageColumn + 15)).ClearContents
    UpdateFatigueMaximumRows ws, firstDamageColumn, 12
End Sub

' Hydrotest tab: two blocks (HYDROTEST with cycles in B1, HYDROTEST 2 with cycles
' in B2), four angle columns each. Design tab: one block with cycles in B1.
Private Sub BuildSupplementalFatigueSheet(ByVal fatigueSheetName As String, _
                                          ByVal stressSheetName As String, _
                                          ByVal phaseName As String, _
                                          ByVal kpCount As Long, _
                                          ByVal innerFiber As Boolean, _
                                          ByVal cycleCount As Double)
    Dim ws As Worksheet, stressSheet As Worksheet, output() As Variant
    Dim kpColumn As Long, firstDamageColumn As Long, sourceColumn As Long
    Dim rowIndex As Long, angleIndex As Long, kp As Double, isDbm As Boolean
    Dim sourceAddress As String, formulaText As String, cycleCell As String
    Dim curveM As String, curveC As String, curveKdf As String
    Dim blockPhases As Variant, b As Long, blockCount As Long
    Dim hasData() As Boolean, anyData As Boolean, sourceText As String
    Set ws = ThisWorkbook.Worksheets(fatigueSheetName)
    Set stressSheet = ThisWorkbook.Worksheets(stressSheetName)
    If innerFiber Then kpColumn = 4 Else kpColumn = 5
    firstDamageColumn = kpColumn + 1
    If NormalizePhase(phaseName) = "HYDROTEST" Then
        blockPhases = Array("HYDROTEST", "HYDROTEST 2")
    Else
        blockPhases = Array(phaseName)
    End If
    blockCount = UBound(blockPhases) + 1

    ws.Range(ws.Cells(FIRST_DATA_ROW, kpColumn), _
             ws.Cells(LAST_TEMPLATE_ROW, firstDamageColumn + 15)).ClearContents
    ws.Range(ws.Cells(5, firstDamageColumn), _
             ws.Cells(8, firstDamageColumn + 15)).ClearContents
    ReDim hasData(0 To blockCount - 1)
    For b = 0 To blockCount - 1
        hasData(b) = PhaseHasStressData(stressSheet, CStr(blockPhases(b)), kpCount)
        If hasData(b) Then
            anyData = True
            If Len(sourceText) > 0 Then sourceText = sourceText & ", "
            sourceText = sourceText & CStr(blockPhases(b))
            If Not IsPositiveNumber(ws.Cells(b + 1, 2).Value2) Then Err.Raise vbObjectError + 409, , _
                ws.Name & ": enter a positive cycle count for " & CStr(blockPhases(b)) & _
                " on the INPUT tab (row " & CStr(b + 1 + INPUT_ROW_OFFSET) & ")."
        End If
    Next b
    If Not anyData Then
        ws.Range("C1").Value2 = "Source: no data"
        Exit Sub
    End If
    ws.Range("C1").Value2 = "Source: " & sourceText
    ValidateSupplementalFatigueInputs ws

    ReDim output(1 To kpCount, 1 To 1 + 4 * blockCount)
    For rowIndex = 1 To kpCount
        output(rowIndex, 1) = "='" & stressSheetName & "'!C" & _
            CStr(FIRST_DATA_ROW + rowIndex - 1)
        kp = CDbl(stressSheet.Cells(FIRST_DATA_ROW + rowIndex - 1, 3).Value2)
        isDbm = IsDbmKP(kp)
        If isDbm Then
            curveM = "$B$11": curveC = "$B$12": curveKdf = "$B$13"
        Else
            curveM = "$B$15": curveC = "$B$16": curveKdf = "$B$17"
        End If
        ApplyUserCurve kp, innerFiber, curveM, curveC, curveKdf
        For b = 0 To blockCount - 1
            If hasData(b) Then
                sourceColumn = StressStartColumn(CStr(blockPhases(b)), "FCD")
                cycleCell = "$B$" & CStr(b + 1)
                For angleIndex = 0 To 3
                    sourceAddress = "'" & stressSheetName & "'!" & _
                        stressSheet.Cells(FIRST_DATA_ROW + rowIndex - 1, _
                        sourceColumn + angleIndex).address(False, False)
                    formulaText = "=IF(OR(" & sourceAddress & "=""""," & _
                        "NOT(ISNUMBER(" & cycleCell & "))),""""," & cycleCell & "/(" & curveC & _
                        "*((ABS(" & sourceAddress & ")/1000000)*$B$4*$B$5*$B$9^$B$8)^(-" & _
                        curveM & ")*(1/" & curveKdf & ")))"
                    output(rowIndex, 2 + 4 * b + angleIndex) = formulaText
                Next angleIndex
            End If
        Next b
    Next rowIndex
    ws.Cells(14, kpColumn).Value2 = "KP [m]"
    ws.Cells(FIRST_DATA_ROW, kpColumn).Resize(kpCount, 1 + 4 * blockCount).Formula = output
    ws.Cells(FIRST_DATA_ROW, firstDamageColumn).Resize(kpCount, 4 * blockCount).NumberFormat = "0.000E+00"
    UpdateFatigueMaximumRows ws, firstDamageColumn, 4 * blockCount
End Sub

Private Sub ValidateSupplementalFatigueInputs(ByVal ws As Worksheet)
    Dim address As Variant
    For Each address In Array("B4", "B5", "B8", "B9", _
                              "B11", "B12", "B13", "B15", "B16", "B17")
        If Not IsNumeric(ws.Range(CStr(address)).Value2) Then Err.Raise vbObjectError + 408, , _
            ws.Name & " requires a numeric value in " & CStr(address) & "."
    Next address
    If CDbl(ws.Range("B12").Value2) <= 0 Or CDbl(ws.Range("B16").Value2) <= 0 Then _
        Err.Raise vbObjectError + 410, , ws.Name & " requires positive C2 values."
    If CDbl(ws.Range("B13").Value2) <= 0 Or CDbl(ws.Range("B17").Value2) <= 0 Then _
        Err.Raise vbObjectError + 411, , ws.Name & " requires positive KDFID values."
End Sub

Private Sub UpdateFatigueMaximumRows(ByVal ws As Worksheet, _
                                     ByVal firstDamageColumn As Long, _
                                     Optional ByVal damageColumnCount As Long = 16)
    Dim stressSheet As Worksheet, lastSegmentRow As Long, damageColumn As Long
    Dim formulaArgs As String
    Set stressSheet = ThisWorkbook.Worksheets("StressRangeID")
    lastSegmentRow = stressSheet.Cells(stressSheet.Rows.Count, 75).End(xlUp).Row
    ws.Range(ws.Cells(5, firstDamageColumn), _
             ws.Cells(5, firstDamageColumn + damageColumnCount - 1)).ClearContents
    ws.Range(ws.Cells(8, firstDamageColumn), _
             ws.Cells(8, firstDamageColumn + damageColumnCount - 1)).ClearContents
    For damageColumn = firstDamageColumn To firstDamageColumn + damageColumnCount - 1
        formulaArgs = SegmentRangeArguments(stressSheet, ws, lastSegmentRow, "DBM", damageColumn)
        If Len(formulaArgs) > 0 Then ws.Cells(5, damageColumn).Formula = "=MAX(" & formulaArgs & ")"
        formulaArgs = SegmentRangeArguments(stressSheet, ws, lastSegmentRow, "RB", damageColumn)
        If Len(formulaArgs) > 0 Then ws.Cells(8, damageColumn).Formula = "=MAX(" & formulaArgs & ")"
    Next damageColumn
End Sub

Private Sub ValidateFatigueInputs(ByVal ws As Worksheet, ByVal phaseName As String)
    Dim address As Variant
    For Each address In Array("B4", "B5", "B8", "B9", "B11", "B12", "B13", "B15", "B16", "B17")
        If Not IsNumeric(ws.Range(CStr(address)).Value2) Then Err.Raise vbObjectError + 402, , _
            ws.Name & " requires a numeric value in " & CStr(address) & "."
    Next address
    If CDbl(ws.Range("B12").Value2) <= 0 Or CDbl(ws.Range("B16").Value2) <= 0 Then _
        Err.Raise vbObjectError + 403, , ws.Name & " requires positive C2 values."
    If CDbl(ws.Range("B13").Value2) <= 0 Or CDbl(ws.Range("B17").Value2) <= 0 Then _
        Err.Raise vbObjectError + 404, , ws.Name & " requires positive KDFID values."
    If CycleCellInvalid(ws.Range("B1").Value2) And PhaseLoadCaseHasData(phaseName, "FCD") Then _
        Err.Raise vbObjectError + 405, , ws.Name & " requires the FCD cycle count in B1."
    If CycleCellInvalid(ws.Range("B2").Value2) And PhaseLoadCaseHasData(phaseName, "HCD") Then _
        Err.Raise vbObjectError + 406, , ws.Name & " requires the HCD cycle count in B2."
    If CycleCellInvalid(ws.Range("B3").Value2) And _
       PhaseLoadCaseHasData(phaseName, "PCD") Then _
        Err.Raise vbObjectError + 407, , ws.Name & " requires the PCD cycle count in B3."
End Sub

' ============================================================================
' BUTTON 5: EARLY-PROFILE AND ALL-LIFE SUMMARIES
' ============================================================================

Private Sub SummarizeFatigueDamageCore()
    Dim kpCount As Long, allowable As Double
    kpCount = ExistingKPCount(ThisWorkbook.Worksheets("StressRangeID"))
    If kpCount <= 0 Then Err.Raise vbObjectError + 500, , "No KP grid is available for the summary."
    If Not IsNumeric(ControlSheet.Range(ALLOWABLE_CELL).Value2) Then _
        Err.Raise vbObjectError + 501, , "Allowable damage in Workflow Controls!" & ALLOWABLE_CELL & " must be numeric."
    allowable = CDbl(ControlSheet.Range(ALLOWABLE_CELL).Value2)
    If allowable <= 0 Then Err.Raise vbObjectError + 502, , "Allowable damage must be greater than zero."

    BuildSummarySheet "Fatigue_IDSUMEarlyProfile", True, True, kpCount
    BuildSummarySheet "Fatigue_ODSUMEArlyProfile", False, True, kpCount
    BuildSummarySheet "Fatigue_IDSUM", True, False, kpCount
    BuildSummarySheet "Fatigue_ODSUM", False, False, kpCount
    BuildKpRangeSummary kpCount
    BuildPlotsSheet kpCount
End Sub

Private Sub BuildSummarySheet(ByVal summarySheetName As String, ByVal innerFiber As Boolean, _
                              ByVal earlyOnly As Boolean, ByVal kpCount As Long)
    Dim ws As Worksheet, output() As Variant, sourceSheets As Variant
    Dim sourceKPColumn As Long, sourceDamageColumn As Long, rowIndex As Long, columnIndex As Long
    Dim excelRow As Long, sourceList As String, sourceCount As String, sheetName As Variant
    Dim totalColumn As Long, formulaText As String, outputColumnCount As Long
    Dim hydroSheetName As String, designSheetName As String
    Dim hydroAddress As String, designAddress As String
    Set ws = ThisWorkbook.Worksheets(summarySheetName)
    If innerFiber Then
        sourceKPColumn = 4
        sourceDamageColumn = 5
        hydroSheetName = "FatigueID_Hydrotest"
        designSheetName = "FatigueID_DesignOp"
        If earlyOnly Then
            sourceSheets = Array("FatigueID_EarlyUndr", "FatigueID_Earlydr")
        Else
            sourceSheets = Array("FatigueID_EarlyUndr", "FatigueID_Earlydr", _
                                 "FatigueID_Middr", "FatigueID_Latdr")
        End If
    Else
        sourceKPColumn = 5
        sourceDamageColumn = 6
        hydroSheetName = "FatigueOD_Hydrotest"
        designSheetName = "FatigueOD_DesignOp"
        If earlyOnly Then
            sourceSheets = Array("FatigueOD_EarlyUndr", "FatigueOD_Earlydr")
        Else
            sourceSheets = Array("FatigueOD_EarlyUndr", "FatigueOD_Earlydr", _
                                 "FatigueOD_Middr", "FatigueOD_Latdr")
        End If
    End If
    ws.Range("D" & FIRST_DATA_ROW & ":AJ" & LAST_TEMPLATE_ROW).ClearContents
    If earlyOnly Then outputColumnCount = 25 Else outputColumnCount = 33
    ReDim output(1 To kpCount, 1 To outputColumnCount)
    For rowIndex = 1 To kpCount
        excelRow = FIRST_DATA_ROW + rowIndex - 1
        output(rowIndex, 1) = "=StressRangeID!C" & CStr(excelRow)
        For columnIndex = 0 To 11
            sourceList = vbNullString
            sourceCount = vbNullString
            For Each sheetName In sourceSheets
                If Len(sourceList) > 0 Then sourceList = sourceList & ","
                If Len(sourceCount) > 0 Then sourceCount = sourceCount & ","
                sourceList = sourceList & "'" & CStr(sheetName) & "'!" & _
                    ThisWorkbook.Worksheets(CStr(sheetName)).Cells(excelRow, _
                    sourceDamageColumn + columnIndex).address(False, False)
                sourceCount = sourceCount & "'" & CStr(sheetName) & "'!" & _
                    ThisWorkbook.Worksheets(CStr(sheetName)).Cells(excelRow, _
                    sourceDamageColumn + columnIndex).address(False, False)
            Next sheetName
            output(rowIndex, 2 + columnIndex) = "=IF(COUNT(" & sourceCount & ")=0,"""",SUM(" & sourceList & "))"
        Next columnIndex
        For totalColumn = 0 To 3
            If earlyOnly Then
                formulaText = "=IF(COUNT(" & ws.Cells(excelRow, 5 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 9 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 13 + totalColumn).address(False, False) & ")=0,"""",SUM(" & _
                    ws.Cells(excelRow, 5 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 9 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 13 + totalColumn).address(False, False) & "))"
            Else
                hydroAddress = "'" & hydroSheetName & "'!" & _
                    ThisWorkbook.Worksheets(hydroSheetName).Cells(excelRow, _
                    sourceDamageColumn + totalColumn).address(False, False) & ",'" & _
                    hydroSheetName & "'!" & _
                    ThisWorkbook.Worksheets(hydroSheetName).Cells(excelRow, _
                    sourceDamageColumn + 4 + totalColumn).address(False, False)
                designAddress = "'" & designSheetName & "'!" & _
                    ThisWorkbook.Worksheets(designSheetName).Cells(excelRow, _
                    sourceDamageColumn + totalColumn).address(False, False)
                output(rowIndex, 26 + totalColumn) = "=IF(COUNT(" & hydroAddress & _
                    ")=0,"""",SUM(" & hydroAddress & "))"
                output(rowIndex, 30 + totalColumn) = "=IF(COUNT(" & designAddress & _
                    ")=0,""""," & designAddress & ")"
                formulaText = "=IF(COUNT(" & ws.Cells(excelRow, 5 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 9 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 13 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 29 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 33 + totalColumn).address(False, False) & ")=0,"""",SUM(" & _
                    ws.Cells(excelRow, 5 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 9 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 13 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 29 + totalColumn).address(False, False) & "," & _
                    ws.Cells(excelRow, 33 + totalColumn).address(False, False) & "))"
            End If
            output(rowIndex, 18 + totalColumn) = formulaText
            output(rowIndex, 22 + totalColumn) = "=IF(" & _
                ws.Cells(excelRow, 21 + totalColumn).address(False, False) & _
                "="""",""""," & ws.Cells(excelRow, 21 + totalColumn).address(False, False) & _
                "/'" & CONTROL_SHEET & "'!$I$8)"
        Next totalColumn
    Next rowIndex
    WriteSummaryHeaders ws, innerFiber, earlyOnly
    ws.Range("D" & FIRST_DATA_ROW).Resize(kpCount, outputColumnCount).Formula = output
    If earlyOnly Then
        ws.Range("E" & FIRST_DATA_ROW & ":AB" & (FIRST_DATA_ROW + kpCount - 1)).NumberFormat = "0.000E+00"
    Else
        ws.Range("E" & FIRST_DATA_ROW & ":AJ" & (FIRST_DATA_ROW + kpCount - 1)).NumberFormat = "0.000E+00"
    End If
    UpdateProfileMaximumRows ws, kpCount, earlyOnly
    BuildSummaryCharts ws, innerFiber, earlyOnly, kpCount
End Sub

Private Sub BuildSummaryCharts(ByVal ws As Worksheet, ByVal innerFiber As Boolean, _
                               ByVal earlyOnly As Boolean, ByVal kpCount As Long)
    Dim fiberText As String, lifeText As String, lastDataRow As Long
    Dim chartLeft As Double, chartTop As Double, chartWidth As Double, chartHeight As Double
    If kpCount <= 0 Then Exit Sub
    If innerFiber Then fiberText = "ID" Else fiberText = "OD"
    If earlyOnly Then lifeText = "Early-Profile " Else lifeText = vbNullString
    RemoveLegacySummaryCharts ws
    lastDataRow = FIRST_DATA_ROW + kpCount - 1
    If earlyOnly Then
        chartLeft = ws.Range("AD2").Left
    Else
        chartLeft = ws.Range("AL2").Left
    End If
    chartTop = ws.Range("AD2").Top
    chartWidth = 720
    chartHeight = 300

    CreateOrUpdateFatigueChart ws, "FD_TotalDamage", _
        lifeText & "Total Fatigue Damage (" & fiberText & ")", _
        "Fatigue damage", 21, lastDataRow, chartLeft, chartTop, chartWidth, chartHeight
    CreateOrUpdateFatigueChart ws, "FD_TotalUC", _
        lifeText & "Fatigue Utilization (UC) (" & fiberText & ")", _
        "Fatigue UC", 25, lastDataRow, chartLeft, chartTop + chartHeight + 18, _
        chartWidth, chartHeight
End Sub

Private Sub RemoveLegacySummaryCharts(ByVal ws As Worksheet)
    Dim chartIndex As Long, existingName As String
    For chartIndex = ws.ChartObjects.Count To 1 Step -1
        existingName = ws.ChartObjects(chartIndex).Name
        If StrComp(existingName, "FD_TotalDamage", vbTextCompare) <> 0 And _
           StrComp(existingName, "FD_TotalUC", vbTextCompare) <> 0 Then
            ws.ChartObjects(chartIndex).Delete
        End If
    Next chartIndex
End Sub

Private Sub CreateOrUpdateFatigueChart(ByVal ws As Worksheet, ByVal chartName As String, _
                                       ByVal titleText As String, ByVal yAxisText As String, _
                                       ByVal firstValueColumn As Long, ByVal lastDataRow As Long, _
                                       ByVal chartLeft As Double, ByVal chartTop As Double, _
                                       ByVal chartWidth As Double, ByVal chartHeight As Double)
    Dim ChartObject As ChartObject, chartSeries As Series
    Dim seriesIndex As Long, seriesLabels As Variant, seriesColors As Variant
    On Error Resume Next
    Set ChartObject = ws.ChartObjects(chartName)
    On Error GoTo 0
    If ChartObject Is Nothing Then
        Set ChartObject = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
        ChartObject.Name = chartName
    Else
        ChartObject.Left = chartLeft
        ChartObject.Top = chartTop
        ChartObject.Width = chartWidth
        ChartObject.Height = chartHeight
    End If

    seriesLabels = Array("-90" & ChrW$(176), "0" & ChrW$(176), _
                         "90" & ChrW$(176), "180" & ChrW$(176))
    seriesColors = Array(RGB(31, 78, 121), RGB(237, 125, 49), _
                         RGB(112, 173, 71), RGB(112, 48, 160))
    With ChartObject.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For seriesIndex = 0 To 3
            Set chartSeries = .SeriesCollection.NewSeries
            chartSeries.Name = CStr(seriesLabels(seriesIndex))
            chartSeries.XValues = ws.Range(ws.Cells(FIRST_DATA_ROW, 4), _
                                           ws.Cells(lastDataRow, 4))
            chartSeries.values = ws.Range(ws.Cells(FIRST_DATA_ROW, _
                                          firstValueColumn + seriesIndex), _
                                          ws.Cells(lastDataRow, firstValueColumn + seriesIndex))
            chartSeries.MarkerStyle = xlMarkerStyleNone
            chartSeries.Format.Line.ForeColor.RGB = CLng(seriesColors(seriesIndex))
            chartSeries.Format.Line.Weight = 1.75
        Next seriesIndex
        .HasTitle = True
        .ChartTitle.Text = titleText
        .HasLegend = True
        .Legend.position = xlLegendPositionTop
        .DisplayBlanksAs = xlNotPlotted
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Size = 9
        .ChartTitle.Font.Name = "Arial"
        .ChartTitle.Font.Size = 12
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "KP (m)"
        .Axes(xlCategory).TickLabels.NumberFormat = "0"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = yAxisText
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).TickLabels.NumberFormat = "0.0E+00"
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
        .ChartArea.Format.Line.ForeColor.RGB = RGB(166, 166, 166)
    End With
End Sub

Private Sub WriteSummaryHeaders(ByVal ws As Worksheet, ByVal innerFiber As Boolean, _
                                ByVal earlyOnly As Boolean)
    Dim c As Long, angleList As Variant, fiberText As String
    angleList = Array(-90#, 0#, 90#, 180#)
    If innerFiber Then fiberText = "Inner FIBER" Else fiberText = "Outer FIBER"
    ws.Range("E1:AJ4").ClearContents
    If Not earlyOnly Then
        ws.Range("U1:X14").Copy
        ws.Range("AC1:AF14").PasteSpecial xlPasteFormats
        ws.Range("Y1:AB14").Copy
        ws.Range("AG1:AJ14").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        For c = 0 To 3
            ws.Columns(29 + c).ColumnWidth = 14
            ws.Columns(33 + c).ColumnWidth = 14
        Next c
    End If
    For c = 0 To 3
        ws.Cells(1, 5 + c).Value2 = "FCD"
        ws.Cells(1, 9 + c).Value2 = "HCD"
        ws.Cells(1, 13 + c).Value2 = "PCD"
        ws.Cells(1, 21 + c).Value2 = "TOTAL"
        ws.Cells(1, 25 + c).Value2 = "UC"
        If Not earlyOnly Then
            ws.Cells(1, 29 + c).Value2 = "HYDROTEST"
            ws.Cells(1, 33 + c).Value2 = "DESIGN OP."
        End If
        ws.Cells(3, 5 + c).Value2 = "Radius = " & fiberText
        ws.Cells(3, 9 + c).Value2 = "Radius = " & fiberText
        ws.Cells(3, 13 + c).Value2 = "Radius = " & fiberText
        ws.Cells(3, 21 + c).Value2 = "Radius = " & fiberText
        ws.Cells(3, 25 + c).Value2 = "Radius = " & fiberText
        If Not earlyOnly Then
            ws.Cells(3, 29 + c).Value2 = "Radius = " & fiberText
            ws.Cells(3, 33 + c).Value2 = "Radius = " & fiberText
        End If
        ws.Cells(4, 5 + c).Value2 = "Angle = " & CStr(angleList(c))
        ws.Cells(4, 9 + c).Value2 = "Angle = " & CStr(angleList(c))
        ws.Cells(4, 13 + c).Value2 = "Angle = " & CStr(angleList(c))
        ws.Cells(4, 21 + c).Value2 = "Angle = " & CStr(angleList(c))
        ws.Cells(4, 25 + c).Value2 = "Angle = " & CStr(angleList(c))
        If Not earlyOnly Then
            ws.Cells(4, 29 + c).Value2 = "Angle = " & CStr(angleList(c))
            ws.Cells(4, 33 + c).Value2 = "Angle = " & CStr(angleList(c))
        End If
    Next c
    ws.Range("E2").Value2 = "SUM"
    ws.Range("I2").Value2 = "SUM"
    ws.Range("M2").Value2 = "SUM"
    ws.Range("U2").Value2 = "SUM"
    ws.Range("Y2").Value2 = "TOTAL / ALLOWABLE"
    If Not earlyOnly Then
        If innerFiber Then
            ws.Range("AC2").Formula = "=SUM(FatigueID_Hydrotest!$B$1:$B$2)&"" CYCLE(S)"""
            ws.Range("AG2").Formula = "=FatigueID_DesignOp!$B$1&"" CYCLE(S)"""
        Else
            ws.Range("AC2").Formula = "=SUM(FatigueOD_Hydrotest!$B$1:$B$2)&"" CYCLE(S)"""
            ws.Range("AG2").Formula = "=FatigueOD_DesignOp!$B$1&"" CYCLE(S)"""
        End If
    End If
    ws.Range("D14").Value2 = "KP [m]"
End Sub

Private Sub UpdateProfileMaximumRows(ByVal ws As Worksheet, ByVal kpCount As Long, _
                                     ByVal earlyOnly As Boolean)
    Dim stressSheet As Worksheet, lastSegmentRow As Long, summaryColumn As Long
    Dim segmentRow As Long, segmentType As String, formulaArgs As String
    Dim startRow As Long, endRow As Long
    Set stressSheet = ThisWorkbook.Worksheets("StressRangeID")
    lastSegmentRow = stressSheet.Cells(stressSheet.Rows.Count, 75).End(xlUp).Row
    If earlyOnly Then
        ws.Range("D5:AB9").ClearContents
    Else
        ws.Range("D5:AJ9").ClearContents
    End If
    ws.Range("D5").Value2 = "DBM Max"
    ws.Range("D6").Value2 = "DBM UC"
    ws.Range("D8").Value2 = "RB Max"
    ws.Range("D9").Value2 = "RB UC"
    For summaryColumn = 5 To 24
        If summaryColumn >= 17 And summaryColumn <= 20 Then GoTo NextSummaryColumn
        formulaArgs = SegmentRangeArguments(stressSheet, ws, lastSegmentRow, _
            "DBM", summaryColumn)
        If Len(formulaArgs) > 0 Then ws.Cells(5, summaryColumn).Formula = "=MAX(" & formulaArgs & ")"
        formulaArgs = SegmentRangeArguments(stressSheet, ws, lastSegmentRow, _
            "RB", summaryColumn)
        If Len(formulaArgs) > 0 Then ws.Cells(8, summaryColumn).Formula = "=MAX(" & formulaArgs & ")"
NextSummaryColumn:
    Next summaryColumn
    If Not earlyOnly Then
        For summaryColumn = 29 To 36
            formulaArgs = SegmentRangeArguments(stressSheet, ws, lastSegmentRow, _
                "DBM", summaryColumn)
            If Len(formulaArgs) > 0 Then ws.Cells(5, summaryColumn).Formula = "=MAX(" & formulaArgs & ")"
            formulaArgs = SegmentRangeArguments(stressSheet, ws, lastSegmentRow, _
                "RB", summaryColumn)
            If Len(formulaArgs) > 0 Then ws.Cells(8, summaryColumn).Formula = "=MAX(" & formulaArgs & ")"
        Next summaryColumn
    End If
    For summaryColumn = 21 To 24
        ws.Cells(6, summaryColumn).Formula = "=IF(" & ws.Cells(5, summaryColumn).address(False, False) & _
            "="""",""""," & ws.Cells(5, summaryColumn).address(False, False) & _
            "/'" & CONTROL_SHEET & "'!$I$8)"
        ws.Cells(9, summaryColumn).Formula = "=IF(" & ws.Cells(8, summaryColumn).address(False, False) & _
            "="""",""""," & ws.Cells(8, summaryColumn).address(False, False) & _
            "/'" & CONTROL_SHEET & "'!$I$8)"
    Next summaryColumn
End Sub

Private Function SegmentRangeArguments(ByVal segmentSheet As Worksheet, _
                                       ByVal summarySheet As Worksheet, _
                                       ByVal lastSegmentRow As Long, _
                                       ByVal requestedType As String, _
                                       ByVal summaryColumn As Long) As String
    Dim r As Long, startRow As Long, endRow As Long, argumentText As String
    For r = 2 To lastSegmentRow
        If StrComp(Trim$(CStr(segmentSheet.Cells(r, 75).Value2)), requestedType, vbTextCompare) = 0 Then
            If IsNumeric(segmentSheet.Cells(r, 77).Value2) And _
               IsNumeric(segmentSheet.Cells(r, 78).Value2) Then
                startRow = CLng(segmentSheet.Cells(r, 77).Value2)
                endRow = CLng(segmentSheet.Cells(r, 78).Value2)
                If Len(argumentText) > 0 Then argumentText = argumentText & ","
                argumentText = argumentText & summarySheet.Range( _
                    summarySheet.Cells(startRow, summaryColumn), _
                    summarySheet.Cells(endRow, summaryColumn)).address(False, False)
            End If
        End If
    Next r
    SegmentRangeArguments = argumentText
End Function

' ============================================================================
' SHARED LOOKUPS, PARSING, AND FORMATTING
' ============================================================================

Private Function ControlSheet() As Worksheet
    Set ControlSheet = ThisWorkbook.Worksheets(CONTROL_SHEET)
End Function

Private Sub SetWorkflowStatus(ByVal statusText As String)
    If WorksheetExists(CONTROL_SHEET) Then _
        ThisWorkbook.Worksheets(CONTROL_SHEET).Range(STATUS_CELL).Value2 = statusText
End Sub

Private Sub ClearLegacyWorkflowData()
    Dim ws As Worksheet, sheetName As Variant, firstDamageColumn As Long
    mCurrentStage = "button 1 first-use cleanup: Sheet1"
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    ws.Range("A1:V2").ClearContents

    For Each sheetName In Array("ID", "OD")
        mCurrentStage = "button 1 first-use cleanup: " & CStr(sheetName)
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        If LastUsedRow(ws) > 0 And LastUsedColumn(ws) > 0 Then _
            ws.Range(ws.Cells(1, 1), ws.Cells(LastUsedRow(ws), LastUsedColumn(ws))).ClearContents
    Next sheetName

    For Each sheetName In Array("StressRangeID", "StressRangeOD")
        mCurrentStage = "button 1 first-use cleanup: " & CStr(sheetName)
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        ws.Range("D14:DN14").ClearContents
        ws.Range("B14").Value2 = "No"
        ws.Range("C14").Value2 = "KP [m]"
        ws.Range("B" & FIRST_DATA_ROW & ":BK" & LAST_TEMPLATE_ROW).ClearContents
        ws.Range("CQ" & FIRST_DATA_ROW & ":DF" & LAST_TEMPLATE_ROW).ClearContents
        ws.Range("DG" & FIRST_DATA_ROW & ":DN" & LAST_TEMPLATE_ROW).ClearContents
    Next sheetName
    Set ws = ThisWorkbook.Worksheets("StressRangeID")
    mCurrentStage = "button 1 first-use cleanup: StressRangeID helper area"
    ws.Range("BQ2:CO1000").ClearContents

    For Each sheetName In Array("FatigueID_EarlyUndr", "FatigueID_Earlydr", _
        "FatigueID_Middr", "FatigueID_Latdr", "FatigueID_Hydrotest", _
        "FatigueID_DesignOp", "FatigueOD_EarlyUndr", "FatigueOD_Earlydr", _
        "FatigueOD_Middr", "FatigueOD_Latdr", "FatigueOD_Hydrotest", _
        "FatigueOD_DesignOp")
        mCurrentStage = "button 1 first-use cleanup: " & CStr(sheetName)
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        If InStr(1, CStr(sheetName), "FatigueID_", vbTextCompare) = 1 Then _
            firstDamageColumn = 5 Else firstDamageColumn = 6
        ws.Range(ws.Cells(5, firstDamageColumn), _
                 ws.Cells(LAST_TEMPLATE_ROW, firstDamageColumn + 15)).ClearContents
        If firstDamageColumn = 5 Then
            ws.Range("D" & FIRST_DATA_ROW & ":D" & LAST_TEMPLATE_ROW).ClearContents
        Else
            ws.Range("E" & FIRST_DATA_ROW & ":E" & LAST_TEMPLATE_ROW).ClearContents
        End If
    Next sheetName

    For Each sheetName In Array("Fatigue_IDSUMEarlyProfile", "Fatigue_ODSUMEArlyProfile", _
        "Fatigue_IDSUM", "Fatigue_ODSUM")
        mCurrentStage = "button 1 first-use cleanup: " & CStr(sheetName)
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        ws.Range("E5:AJ9").ClearContents
        ws.Range("D" & FIRST_DATA_ROW & ":AJ" & LAST_TEMPLATE_ROW).ClearContents
    Next sheetName

    If WorksheetExists(CONTROL_SHEET) Then
        mCurrentStage = "button 1 first-use cleanup: control metadata"
        ClearControlImportMetadata
        ControlSheet.Range(KP_COUNT_CELL).Value2 = 0
    End If
End Sub

Private Sub PrepareTemplateForDelivery()
    Dim ws As Worksheet, sheetName As Variant, firstDamageColumn As Long
    ClearLegacyWorkflowData
    ThisWorkbook.Worksheets("StressRangeID").Range("BQ2:CO1000").ClearContents
    For Each sheetName In Array("FatigueID_EarlyUndr", "FatigueID_Earlydr", _
        "FatigueID_Middr", "FatigueID_Latdr", "FatigueID_Hydrotest", _
        "FatigueID_DesignOp", "FatigueOD_EarlyUndr", "FatigueOD_Earlydr", _
        "FatigueOD_Middr", "FatigueOD_Latdr", "FatigueOD_Hydrotest", _
        "FatigueOD_DesignOp")
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        If InStr(1, CStr(sheetName), "FatigueID_", vbTextCompare) = 1 Then _
            firstDamageColumn = 5 Else firstDamageColumn = 6
        ws.Range(ws.Cells(5, firstDamageColumn), ws.Cells(14, firstDamageColumn + 15)).ClearContents
    Next sheetName
    For Each sheetName In Array("Fatigue_IDSUMEarlyProfile", "Fatigue_ODSUMEArlyProfile", _
        "Fatigue_IDSUM", "Fatigue_ODSUM")
        ThisWorkbook.Worksheets(CStr(sheetName)).Range("E5:AJ9").ClearContents
    Next sheetName
    ClearControlImportMetadata
    ControlSheet.Range(KP_COUNT_CELL).Value2 = 0
End Sub

Private Sub ClearControlImportMetadata()
    With ControlSheet
        .Range("I11:I15").ClearContents
        .Range(SOURCE_FILE_CELL).MergeArea.ClearContents
    End With
End Sub

Private Function ReadReportLines(ByVal reportPath As String) As Variant
    Dim fileNumber As Integer, fileText As String, byteCount As Long
    fileNumber = FreeFile
    Open reportPath For Binary Access Read As #fileNumber
    byteCount = LOF(fileNumber)
    If byteCount > 0 Then
        fileText = Space$(byteCount)
        Get #fileNumber, , fileText
    End If
    Close #fileNumber
    fileText = Replace(fileText, vbCrLf, vbLf)
    fileText = Replace(fileText, vbCr, vbLf)
    ReadReportLines = Split(fileText, vbLf)
End Function

Private Function FindReportHeaderIndex(ByVal reportLines As Variant, _
                                       ByVal preamble As Collection) As Long
    Dim lineIndex As Long, lineText As String
    FindReportHeaderIndex = -1
    For lineIndex = LBound(reportLines) To UBound(reportLines)
        lineText = CStr(reportLines(lineIndex))
        If InStr(1, lineText, "Pipeline Distance", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Pipe Distance", vbTextCompare) > 0 Then
            FindReportHeaderIndex = lineIndex
            Exit Function
        End If
        preamble.Add lineText
    Next lineIndex
End Function

Private Function SelectPairName(ByVal pairs As Object, ByVal pairFilter As String) As String
    Dim key As Variant, matches As Collection, promptText As String, entered As String
    Set matches = New Collection
    For Each key In pairs.keys
        If Len(pairFilter) = 0 Or InStr(1, CStr(key), pairFilter, vbTextCompare) > 0 Then _
            matches.Add CStr(key)
    Next key
    If matches.Count = 1 Then SelectPairName = matches(1): Exit Function
    If matches.Count = 0 Then Err.Raise vbObjectError + 700, , _
        "No step pair matches the pair-name filter: " & pairFilter
    If mProgrammaticMode Then Err.Raise vbObjectError + 701, , _
        "The report contains multiple step pairs. Enter a unique pair-name filter in Workflow Controls!" & PAIR_FILTER_CELL & "."
    promptText = "Enter a unique part of one step-pair name:" & vbCrLf
    For Each key In pairs.keys: promptText = promptText & vbCrLf & CStr(key): Next key
    entered = InputBox(promptText, "Select step pair")
    If Len(Trim$(entered)) = 0 Then Err.Raise vbObjectError + 702, , "No step pair was selected."
    SelectPairName = SelectPairName(pairs, entered)
End Function

Private Function StressStartColumn(ByVal phaseName As String, ByVal loadCase As String) As Long
    Select Case NormalizePhase(phaseName)
        Case "EARLY UNDRAINED"
            Select Case NormalizeLoadCase(loadCase)
                Case "PCD": StressStartColumn = 16
                Case "HCD": StressStartColumn = 20
                Case "FCD": StressStartColumn = 28
            End Select
        Case "EARLY DRAINED"
            Select Case NormalizeLoadCase(loadCase)
                Case "PCD": StressStartColumn = 32
                Case "HCD": StressStartColumn = 36
                Case "FCD": StressStartColumn = 44
            End Select
        Case "MIDDLE DRAINED"
            Select Case NormalizeLoadCase(loadCase)
                Case "PCD": StressStartColumn = 48
                Case "HCD": StressStartColumn = 52
                Case "FCD": StressStartColumn = 60
            End Select
        Case "LATE DRAINED"
            Select Case NormalizeLoadCase(loadCase)
                Case "PCD": StressStartColumn = 95
                Case "HCD": StressStartColumn = 99
                Case "FCD": StressStartColumn = 107
            End Select
        Case "HYDROTEST"
            StressStartColumn = 4
        Case "HYDROTEST 2"
            StressStartColumn = 8
        Case "DESIGN OPERATION"
            StressStartColumn = 12
    End Select
    If StressStartColumn = 0 Then Err.Raise vbObjectError + 703, , _
        "Unsupported phase/load-case combination: " & phaseName & " / " & loadCase
End Function

Private Function NormalizeInputMode(ByVal textValue As String) As String
    textValue = UCase$(Trim$(Replace(textValue, "_", " ")))
    Select Case textValue
        Case "", "AUTO": NormalizeInputMode = "AUTO"
        Case "S11", "RAW S11", "LAST FRAME S11": NormalizeInputMode = "S11"
        Case "DELTA", "DELTA S11", "DELTA-S11": NormalizeInputMode = "DELTA S11"
        Case Else: Err.Raise vbObjectError + 704, , "Input option must be AUTO, S11, or DELTA S11."
    End Select
End Function

Private Function NormalizePhase(ByVal textValue As String) As String
    textValue = UCase$(Trim$(Replace(Replace(textValue, "_", " "), "-", " ")))
    Select Case textValue
        Case "EARLY", "EARLY UNDRAINED", "EARLYUNDR": NormalizePhase = "EARLY UNDRAINED"
        Case "EARLY DRAINED", "EARLYDR": NormalizePhase = "EARLY DRAINED"
        Case "MIDDLE", "MIDDLE DRAINED", "MIDDLEDR": NormalizePhase = "MIDDLE DRAINED"
        Case "LATE", "LATE DRAINED", "LATEDR": NormalizePhase = "LATE DRAINED"
        Case "HYDRO", "HYDROTEST", "HYDRO TEST", "HYDROTEST 1", "HYDROTEST1": _
            NormalizePhase = "HYDROTEST"
        Case "HYDROTEST 2", "HYDROTEST2", "HYDRO 2": NormalizePhase = "HYDROTEST 2"
        Case "DESIGN", "DESIGN OP", "DESIGN OPERATION", "DESIGN STEP": _
            NormalizePhase = "DESIGN OPERATION"
        Case Else: Err.Raise vbObjectError + 705, , "Unsupported phase: " & textValue
    End Select
End Function

Private Function NormalizeLoadCase(ByVal textValue As String) As String
    textValue = UCase$(Replace(Trim$(textValue), " ", vbNullString))
    Select Case textValue
        Case "FCD", "HCD": NormalizeLoadCase = textValue
        Case "PCD", "PCD1": NormalizeLoadCase = "PCD"
        Case "PCD2": Err.Raise vbObjectError + 707, , _
            "PCD2 is no longer used. Select load case PCD."
        Case "HYDROTEST", "HYDRO", "HYDROTEST1": NormalizeLoadCase = "HYDROTEST"
        Case "HYDROTEST2", "HYDRO2": NormalizeLoadCase = "HYDROTEST 2"
        Case "DESIGN", "DESIGNOP", "DESIGNOPERATION", "DESIGNSTEP": _
            NormalizeLoadCase = "DESIGN OPERATION"
        Case Else: Err.Raise vbObjectError + 706, , _
            "Load case must be FCD, HCD, PCD, HYDROTEST 1, HYDROTEST 2, or DESIGN."
    End Select
End Function

Private Function FiberFromHeader(ByVal headerText As String) As String
    If InStr(1, headerText, "INNER", vbTextCompare) > 0 Then
        FiberFromHeader = "INNER"
    ElseIf InStr(1, headerText, "OUTER", vbTextCompare) > 0 Then
        FiberFromHeader = "OUTER"
    ElseIf InStr(1, headerText, "MIDDLE", vbTextCompare) > 0 Then
        FiberFromHeader = "MIDDLE"
    End If
End Function

Private Function AngleFromHeader(ByVal headerText As String) As Double
    Dim position As Long, tailText As String, tokens As Variant, token As Variant
    position = InStr(1, headerText, "ANGLE", vbTextCompare)
    If position = 0 Then AngleFromHeader = 1E+99: Exit Function
    tailText = Mid$(headerText, position + 5)
    tailText = Replace(tailText, "=", " ")
    tailText = Replace(tailText, "[", " ")
    tailText = Replace(tailText, "]", " ")
    tokens = Split(WorksheetFunction.Trim(tailText), " ")
    For Each token In tokens
        If IsNumeric(token) Then AngleFromHeader = CDbl(token): Exit Function
    Next token
    AngleFromHeader = 1E+99
End Function

Private Function IsTargetAngle(ByVal angleValue As Double) As Boolean
    IsTargetAngle = Abs(angleValue + 90#) < 0.001 Or Abs(angleValue) < 0.001 Or _
                    Abs(angleValue - 90#) < 0.001 Or Abs(Abs(angleValue) - 180#) < 0.001
End Function

Private Function BracketText(ByVal textValue As String) As String
    Dim openPosition As Long, closePosition As Long
    openPosition = InStrRev(textValue, "[")
    closePosition = InStrRev(textValue, "]")
    If openPosition > 0 And closePosition > openPosition Then _
        BracketText = Mid$(textValue, openPosition + 1, closePosition - openPosition - 1)
End Function

Private Function StepNameFromPreamble(ByVal textValue As String) As String
    Dim equalPosition As Long, parenthesisPosition As Long
    equalPosition = InStr(1, textValue, "=", vbTextCompare)
    If equalPosition > 0 Then
        StepNameFromPreamble = Trim$(Mid$(textValue, equalPosition + 1))
    Else
        StepNameFromPreamble = Trim$(Mid$(textValue, InStr(1, textValue, ":") + 1))
    End If
    parenthesisPosition = InStr(1, StepNameFromPreamble, "(")
    If parenthesisPosition > 1 Then StepNameFromPreamble = Trim$(Left$(StepNameFromPreamble, parenthesisPosition - 1))
End Function

Private Function TryNumber(ByVal textValue As Variant, ByRef numberValue As Double) As Boolean
    Dim cleaned As String
    On Error GoTo NotNumber
    cleaned = Trim$(CStr(textValue))
    If Len(cleaned) = 0 Or UCase$(cleaned) = "NAN" Or UCase$(cleaned) = "INF" Then Exit Function
    cleaned = Replace(cleaned, ",", vbNullString)
    If Not IsNumeric(cleaned) Then Exit Function
    numberValue = CDbl(cleaned)
    TryNumber = True
NotNumber:
End Function

Private Function CanonicalNumber(ByVal numberValue As Double) As String
    CanonicalNumber = Format$(numberValue, "0.000000000000E+00")
End Function

Private Sub AddLargestAbsolute(ByVal values As Object, ByVal key As String, ByVal newValue As Double)
    If values.Exists(key) Then
        If Abs(newValue) > Abs(CDbl(values(key))) Then values(key) = newValue
    Else
        values.Add key, newValue
    End If
End Sub

Private Function SortedKPs(ByVal kpValues As Object) As Double()
    Dim result() As Double, keys As Variant, i As Long
    ReDim result(1 To kpValues.Count)
    keys = kpValues.keys
    For i = 0 To UBound(keys): result(i + 1) = CDbl(kpValues(keys(i))): Next i
    QuickSortDoubles result, LBound(result), UBound(result)
    SortedKPs = result
End Function

Private Sub QuickSortDoubles(ByRef values() As Double, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long, pivot As Double, temporary As Double
    low = first: high = last: pivot = values((first + last) \ 2)
    Do While low <= high
        Do While values(low) < pivot: low = low + 1: Loop
        Do While values(high) > pivot: high = high - 1: Loop
        If low <= high Then
            temporary = values(low): values(low) = values(high): values(high) = temporary
            low = low + 1: high = high - 1
        End If
    Loop
    If first < high Then QuickSortDoubles values, first, high
    If low < last Then QuickSortDoubles values, low, last
End Sub

Private Function ExistingKPCount(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If lastRow >= FIRST_DATA_ROW And IsNumeric(ws.Cells(FIRST_DATA_ROW, 3).Value2) Then _
        ExistingKPCount = lastRow - FIRST_DATA_ROW + 1
End Function

Private Function FindFloorKPIndex(ByVal targetKP As Double, ByVal kpCount As Long) As Long
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Worksheets("StressRangeID")
    FindFloorKPIndex = 1
    For i = 1 To kpCount
        If CDbl(ws.Cells(FIRST_DATA_ROW + i - 1, 3).Value2) <= targetKP Then
            FindFloorKPIndex = i
        Else
            Exit For
        End If
    Next i
End Function

Private Function DbmIdForKP(ByVal kp As Double, ByRef ids() As String, _
                            ByRef starts() As Double, ByRef ends() As Double, _
                            ByVal rangeCount As Long) As String
    Dim i As Long
    For i = 1 To rangeCount
        If kp >= starts(i) And kp <= ends(i) Then DbmIdForKP = ids(i): Exit Function
    Next i
End Function

Private Function IsDbmKP(ByVal kp As Double) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long, startKP As Double, endKP As Double
    Set ws = ThisWorkbook.Worksheets("StressRangeID")
    lastRow = Application.Max(ws.Cells(ws.Rows.Count, 67).End(xlUp).Row, _
                              ws.Cells(ws.Rows.Count, 68).End(xlUp).Row)
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, 67).Value2) And IsNumeric(ws.Cells(r, 68).Value2) Then
            startKP = CDbl(ws.Cells(r, 67).Value2)
            endKP = CDbl(ws.Cells(r, 68).Value2)
            If startKP > endKP Then SwapDoubles startKP, endKP
            If kp >= startKP And kp <= endKP Then IsDbmKP = True: Exit Function
        End If
    Next r
End Function

Private Function PhaseHasStressData(ByVal ws As Worksheet, ByVal phaseName As String, _
                                    ByVal kpCount As Long) As Boolean
    Dim loadCase As Variant, startColumn As Long, normalizedPhase As String
    normalizedPhase = NormalizePhase(phaseName)
    If normalizedPhase = "HYDROTEST" Or normalizedPhase = "HYDROTEST 2" Or _
       normalizedPhase = "DESIGN OPERATION" Then
        startColumn = StressStartColumn(normalizedPhase, "FCD")
        PhaseHasStressData = Application.WorksheetFunction.Count(ws.Range( _
            ws.Cells(FIRST_DATA_ROW, startColumn), _
            ws.Cells(FIRST_DATA_ROW + kpCount - 1, startColumn + 3))) > 0
        Exit Function
    End If
    For Each loadCase In Array("FCD", "HCD", "PCD")
        startColumn = StressStartColumn(phaseName, CStr(loadCase))
        If Application.WorksheetFunction.Count(ws.Range( _
            ws.Cells(FIRST_DATA_ROW, startColumn), _
            ws.Cells(FIRST_DATA_ROW + kpCount - 1, startColumn + 3))) > 0 Then
            PhaseHasStressData = True
            Exit Function
        End If
    Next loadCase
End Function

Private Function PhaseLoadCaseHasData(ByVal phaseName As String, ByVal loadCase As String) As Boolean
    Dim ws As Worksheet, kpCount As Long, startColumn As Long
    Set ws = ThisWorkbook.Worksheets("StressRangeID")
    kpCount = ExistingKPCount(ws)
    If kpCount <= 0 Then Exit Function
    startColumn = StressStartColumn(phaseName, loadCase)
    PhaseLoadCaseHasData = Application.WorksheetFunction.Count(ws.Range( _
        ws.Cells(FIRST_DATA_ROW, startColumn), _
        ws.Cells(FIRST_DATA_ROW + kpCount - 1, startColumn + 3))) > 0
End Function

Private Sub FormatImportedTable(ByVal ws As Worksheet, ByVal rowCount As Long, _
                                ByVal columnCount As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, columnCount))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(2, 1), ws.Cells(rowCount, columnCount)).NumberFormat = "0.000000E+00"
    ws.Columns("A:" & ColumnLetter(Application.Min(columnCount, 50))).AutoFit
End Sub

Private Function ColumnLetter(ByVal columnNumber As Long) As String
    ColumnLetter = Split(Cells(1, columnNumber).address(True, False), "$")(0)
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If Not foundCell Is Nothing Then LastUsedRow = foundCell.Row
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If Not foundCell Is Nothing Then LastUsedColumn = foundCell.Column
End Function

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function ArrayWorksheet(ByVal firstName As String, ByVal secondName As String) As Collection
    Dim result As New Collection
    result.Add ThisWorkbook.Worksheets(firstName)
    result.Add ThisWorkbook.Worksheets(secondName)
    Set ArrayWorksheet = result
End Function

Private Sub SwapDoubles(ByRef firstValue As Double, ByRef secondValue As Double)
    Dim temporary As Double
    temporary = firstValue: firstValue = secondValue: secondValue = temporary
End Sub
