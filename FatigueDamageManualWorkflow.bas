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

Private mProgrammaticMode As Boolean
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
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_IDSUMEarlyProfile"), True, True
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_ODSUMEArlyProfile"), False, True
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_IDSUM"), True, False
    WriteSummaryHeaders ThisWorkbook.Worksheets("Fatigue_ODSUM"), False, False
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
        "Fatigue_IDSUM", "Fatigue_ODSUM", CONTROL_SHEET)
    For Each item In requiredSheets
        If Not WorksheetExists(CStr(item)) Then Err.Raise vbObjectError + 900, , _
            "Required worksheet is missing: " & CStr(item)
    Next item
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    If ws.Shapes.Count < 6 Then Err.Raise vbObjectError + 901, , _
        "The Workflow Controls sheet does not contain all six workflow buttons."
    If CStr(ThisWorkbook.Worksheets("StressRangeID").Range("BN1").Value2) <> "DBM ID" Then _
        Err.Raise vbObjectError + 902, , "StressRangeID!BN:BP is not configured for DBM input."
    If CDbl(ThisWorkbook.Worksheets("FatigueID_Hydrotest").Range("B1").Value2) <> 2# Or _
       CDbl(ThisWorkbook.Worksheets("FatigueOD_Hydrotest").Range("B1").Value2) <> 2# Then _
        Err.Raise vbObjectError + 903, , "Hydrotest total cycles must be fixed at 2."
    If CDbl(ThisWorkbook.Worksheets("FatigueID_DesignOp").Range("B1").Value2) <> 1# Or _
       CDbl(ThisWorkbook.Worksheets("FatigueOD_DesignOp").Range("B1").Value2) <> 1# Then _
        Err.Raise vbObjectError + 904, , "Design Operation total cycles must be fixed at 1."
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
    Dim kpColumn As Long, firstDamageColumn As Long, c As Long
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
    ws.Range("A1").Value2 = "Total cycles="
    ws.Range("B1").Value2 = cycleCount
    ws.Range("A2:B3").ClearContents
    ws.Range("C1:D1").Merge
    ws.Range("C1").Value2 = "Source: no data"
    ws.Range("C1").WrapText = True
    ws.Range("D5").Value2 = "DBM Max"
    ws.Range("D8").Value2 = "RB Max"
    ws.Cells(14, kpColumn).Value2 = "KP [m]"
    For c = 0 To 3
        ws.Cells(1, firstDamageColumn + c).Value2 = displayPhaseName
        ws.Cells(3, firstDamageColumn + c).Value2 = "Radius = " & fiberText & " FIBER"
        ws.Cells(4, firstDamageColumn + c).Value2 = "Angle = " & CStr(angleList(c))
    Next c
    ws.Cells(2, firstDamageColumn).Value2 = CStr(cycleCount) & " total cycles"
    ws.Range(ws.Cells(1, firstDamageColumn), _
             ws.Cells(1, firstDamageColumn + 3)).EntireColumn.ColumnWidth = 20
    ws.Range(ws.Cells(FIRST_DATA_ROW, firstDamageColumn), _
             ws.Cells(LAST_TEMPLATE_ROW, firstDamageColumn + 3)).NumberFormat = "0.000E+00"
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
    ws.Range("BN11:BP1000").ClearContents
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
    ws.Range("BN2:BP1000").Interior.Color = RGB(255, 242, 204)
    ws.Columns("BN:CB").ColumnWidth = 16

    Set ws = ThisWorkbook.Worksheets("StressRangeOD")
    ws.Range("BN1:CO1000").ClearContents
    ws.Range("BN1").Value2 = "DBM inputs and matched rows are maintained in StressRangeID!BN:CB"
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
    ws.Range("H19").Value2 = "Enter life-cycle counts in B1:B3 on the life tabs and DBM/RB S-N inputs in B11:B13 and B15:B17. Hydrotest and Design Operation use fixed totals of 2 and 1 cycles."
    ws.Range("H21:M22").Merge
    ws.Range("H21").Value2 = "If middle/late stress is absent, those tabs use Early Drained (or Early Undrained). Load case is ignored for Hydrotest and Design Operation."
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
        "EARLY UNDRAINED,EARLY DRAINED,MIDDLE DRAINED,LATE DRAINED,HYDROTEST,DESIGN OPERATION"
    ws.Range(LOAD_CASE_CELL).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, _
        "FCD,HCD,PCD1,PCD2"

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
    Dim shape As Shape, anchor As Range
    Set anchor = ws.Range(anchorCell)
    Set shape = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, 280, 43.2)
    shape.Name = shapeName
    shape.OnAction = macroName
    shape.Fill.ForeColor.RGB = fillColor
    shape.Line.ForeColor.RGB = RGB(255, 255, 255)
    shape.TextFrame2.TextRange.Text = captionText
    shape.TextFrame2.TextRange.Font.Name = "Arial"
    shape.TextFrame2.TextRange.Font.Size = 13
    shape.TextFrame2.TextRange.Font.Bold = msoTrue
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
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
    phaseName = NormalizePhase(CStr(ControlSheet.Range(PHASE_CELL).Value2))
    loadCase = NormalizeLoadCase(CStr(ControlSheet.Range(LOAD_CASE_CELL).Value2))
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

    phaseName = NormalizePhase(CStr(ControlSheet.Range(PHASE_CELL).Value2))
    loadCase = NormalizeLoadCase(CStr(ControlSheet.Range(LOAD_CASE_CELL).Value2))
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
            For Each item In heatColumns.Keys
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
    For Each key In firstGrid.Keys
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
       NormalizePhase(phaseName) = "DESIGN OPERATION" Then
        ws.Cells(3, startColumn).Value2 = phaseName & " / TOTAL CYCLES / " & pairName
    Else
        ws.Cells(3, startColumn).Value2 = phaseName & " / " & loadCase & " / " & pairName
    End If
    For c = 0 To 3
        ws.Cells(4, startColumn + c).Value2 = "Radius = " & fiberName & " FIBER"
        ws.Cells(5, startColumn + c).Value2 = "Angle = " & CStr(angleList(c))
    Next c
    ws.Cells(FIRST_DATA_ROW, startColumn).Resize(UBound(numbers), 4).Value2 = output
    ws.Cells(FIRST_DATA_ROW, startColumn).Resize(UBound(numbers), 4).NumberFormat = "0.000000E+00"
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
    Set ws = ThisWorkbook.Worksheets("StressRangeID")
    lastInputRow = Application.Max(ws.Cells(ws.Rows.Count, 66).End(xlUp).Row, _
                                   ws.Cells(ws.Rows.Count, 67).End(xlUp).Row, _
                                   ws.Cells(ws.Rows.Count, 68).End(xlUp).Row)
    If lastInputRow < 2 Then lastInputRow = 2
    For inputRow = 2 To lastInputRow
        If IsNumeric(ws.Cells(inputRow, 67).Value2) And IsNumeric(ws.Cells(inputRow, 68).Value2) Then
            rangeCount = rangeCount + 1
            ReDim Preserve ids(1 To rangeCount)
            ReDim Preserve starts(1 To rangeCount)
            ReDim Preserve ends(1 To rangeCount)
            idValue = Trim$(CStr(ws.Cells(inputRow, 66).Value2))
            If Len(idValue) = 0 Then idValue = "DBM-" & CStr(inputRow - 1)
            startValue = CDbl(ws.Cells(inputRow, 67).Value2)
            endValue = CDbl(ws.Cells(inputRow, 68).Value2)
            If startValue > endValue Then
                tempValue = startValue: startValue = endValue: endValue = tempValue
            End If
            ids(rangeCount) = idValue
            starts(rangeCount) = startValue
            ends(rangeCount) = endValue
        End If
    Next inputRow
    For i = 1 To rangeCount - 1
        For j = i + 1 To rangeCount
            If starts(j) < starts(i) Then
                tempValue = starts(i): starts(i) = starts(j): starts(j) = tempValue
                tempValue = ends(i): ends(i) = ends(j): ends(j) = tempValue
                idValue = ids(i): ids(i) = ids(j): ids(j) = idValue
            End If
        Next j
    Next i
    For i = 2 To rangeCount
        If starts(i) <= ends(i - 1) Then Err.Raise vbObjectError + 401, , _
            "DBM KP ranges overlap: " & ids(i - 1) & " and " & ids(i) & "."
    Next i

    ws.Range("BQ2:CO1000").ClearContents
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

    ReDim output(1 To kpCount, 1 To 17)
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
        For loadCaseIndex = 0 To 3
            loadCase = Array("FCD", "HCD", "PCD1", "PCD2")(loadCaseIndex)
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
                    sourceColumn + angleIndex).Address(False, False)
                formulaText = "=IF(OR(" & sourceAddress & "=""""," & _
                    "NOT(ISNUMBER(" & cycleAddress & "))),""""," & cycleAddress & _
                    "/(" & curveC & "*((ABS(" & sourceAddress & ")/1000000)*" & _
                    "$B$4*$B$5*$B$9^$B$8)^(-" & curveM & ")*(1/" & curveKdf & ")))"
                output(rowIndex, 2 + loadCaseIndex * 4 + angleIndex) = formulaText
            Next angleIndex
        Next loadCaseIndex
    Next rowIndex
    ws.Cells(14, kpColumn).Value2 = "KP [m]"
    ws.Cells(FIRST_DATA_ROW, kpColumn).Resize(kpCount, 17).Formula = output
    ws.Cells(FIRST_DATA_ROW, firstDamageColumn).Resize(kpCount, 16).NumberFormat = "0.000E+00"
    UpdateFatigueMaximumRows ws, firstDamageColumn
End Sub

Private Sub BuildSupplementalFatigueSheet(ByVal fatigueSheetName As String, _
                                          ByVal stressSheetName As String, _
                                          ByVal phaseName As String, _
                                          ByVal kpCount As Long, _
                                          ByVal innerFiber As Boolean, _
                                          ByVal cycleCount As Double)
    Dim ws As Worksheet, stressSheet As Worksheet, output() As Variant
    Dim kpColumn As Long, firstDamageColumn As Long, sourceColumn As Long
    Dim rowIndex As Long, angleIndex As Long, kp As Double, isDbm As Boolean
    Dim sourceAddress As String, formulaText As String
    Dim curveM As String, curveC As String, curveKdf As String
    Set ws = ThisWorkbook.Worksheets(fatigueSheetName)
    Set stressSheet = ThisWorkbook.Worksheets(stressSheetName)
    If innerFiber Then kpColumn = 4 Else kpColumn = 5
    firstDamageColumn = kpColumn + 1
    sourceColumn = StressStartColumn(phaseName, "FCD")

    ws.Range("B1").Value2 = cycleCount
    ws.Range(ws.Cells(FIRST_DATA_ROW, kpColumn), _
             ws.Cells(LAST_TEMPLATE_ROW, firstDamageColumn + 15)).ClearContents
    ws.Range(ws.Cells(5, firstDamageColumn), _
             ws.Cells(8, firstDamageColumn + 15)).ClearContents
    ws.Range("C1").Value2 = "Source: " & phaseName
    If Not PhaseHasStressData(stressSheet, phaseName, kpCount) Then
        ws.Range("C1").Value2 = "Source: no data"
        Exit Sub
    End If
    ValidateSupplementalFatigueInputs ws

    ReDim output(1 To kpCount, 1 To 5)
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
        For angleIndex = 0 To 3
            sourceAddress = "'" & stressSheetName & "'!" & _
                stressSheet.Cells(FIRST_DATA_ROW + rowIndex - 1, _
                sourceColumn + angleIndex).Address(False, False)
            formulaText = "=IF(OR(" & sourceAddress & "=""""," & _
                "NOT(ISNUMBER($B$1))),"""",$B$1/(" & curveC & _
                "*((ABS(" & sourceAddress & ")/1000000)*$B$4*$B$5*$B$9^$B$8)^(-" & _
                curveM & ")*(1/" & curveKdf & ")))"
            output(rowIndex, 2 + angleIndex) = formulaText
        Next angleIndex
    Next rowIndex
    ws.Cells(14, kpColumn).Value2 = "KP [m]"
    ws.Cells(FIRST_DATA_ROW, kpColumn).Resize(kpCount, 5).Formula = output
    ws.Cells(FIRST_DATA_ROW, firstDamageColumn).Resize(kpCount, 4).NumberFormat = "0.000E+00"
    UpdateFatigueMaximumRows ws, firstDamageColumn, 4
End Sub

Private Sub ValidateSupplementalFatigueInputs(ByVal ws As Worksheet)
    Dim address As Variant
    For Each address In Array("B1", "B4", "B5", "B8", "B9", _
                              "B11", "B12", "B13", "B15", "B16", "B17")
        If Not IsNumeric(ws.Range(CStr(address)).Value2) Then Err.Raise vbObjectError + 408, , _
            ws.Name & " requires a numeric value in " & CStr(address) & "."
    Next address
    If CDbl(ws.Range("B1").Value2) <= 0 Then _
        Err.Raise vbObjectError + 409, , ws.Name & " requires a positive total cycle count in B1."
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
    If Not IsNumeric(ws.Range("B1").Value2) And PhaseLoadCaseHasData(phaseName, "FCD") Then _
        Err.Raise vbObjectError + 405, , ws.Name & " requires the FCD cycle count in B1."
    If Not IsNumeric(ws.Range("B2").Value2) And PhaseLoadCaseHasData(phaseName, "HCD") Then _
        Err.Raise vbObjectError + 406, , ws.Name & " requires the HCD cycle count in B2."
    If Not IsNumeric(ws.Range("B3").Value2) And _
       (PhaseLoadCaseHasData(phaseName, "PCD1") Or PhaseLoadCaseHasData(phaseName, "PCD2")) Then _
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
        output(rowIndex, 1) = "='" & CStr(sourceSheets(0)) & "'!" & _
            ThisWorkbook.Worksheets(CStr(sourceSheets(0))).Cells(excelRow, sourceKPColumn).Address(False, False)
        For columnIndex = 0 To 15
            sourceList = vbNullString
            sourceCount = vbNullString
            For Each sheetName In sourceSheets
                If Len(sourceList) > 0 Then sourceList = sourceList & ","
                If Len(sourceCount) > 0 Then sourceCount = sourceCount & ","
                sourceList = sourceList & "'" & CStr(sheetName) & "'!" & _
                    ThisWorkbook.Worksheets(CStr(sheetName)).Cells(excelRow, _
                    sourceDamageColumn + columnIndex).Address(False, False)
                sourceCount = sourceCount & "'" & CStr(sheetName) & "'!" & _
                    ThisWorkbook.Worksheets(CStr(sheetName)).Cells(excelRow, _
                    sourceDamageColumn + columnIndex).Address(False, False)
            Next sheetName
            output(rowIndex, 2 + columnIndex) = "=IF(COUNT(" & sourceCount & ")=0,"""",SUM(" & sourceList & "))"
        Next columnIndex
        For totalColumn = 0 To 3
            If earlyOnly Then
                formulaText = "=IF(COUNT(" & ws.Cells(excelRow, 5 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 9 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 13 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 17 + totalColumn).Address(False, False) & ")=0,"""",SUM(" & _
                    ws.Cells(excelRow, 5 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 9 + totalColumn).Address(False, False) & ",MIN(" & _
                    ws.Cells(excelRow, 13 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 17 + totalColumn).Address(False, False) & ")))"
            Else
                hydroAddress = "'" & hydroSheetName & "'!" & _
                    ThisWorkbook.Worksheets(hydroSheetName).Cells(excelRow, _
                    sourceDamageColumn + totalColumn).Address(False, False)
                designAddress = "'" & designSheetName & "'!" & _
                    ThisWorkbook.Worksheets(designSheetName).Cells(excelRow, _
                    sourceDamageColumn + totalColumn).Address(False, False)
                output(rowIndex, 26 + totalColumn) = "=IF(COUNT(" & hydroAddress & _
                    ")=0,""""," & hydroAddress & ")"
                output(rowIndex, 30 + totalColumn) = "=IF(COUNT(" & designAddress & _
                    ")=0,""""," & designAddress & ")"
                formulaText = "=IF(COUNT(" & ws.Cells(excelRow, 5 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 9 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 13 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 17 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 29 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 33 + totalColumn).Address(False, False) & ")=0,"""",SUM(" & _
                    ws.Cells(excelRow, 5 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 9 + totalColumn).Address(False, False) & ",MIN(" & _
                    ws.Cells(excelRow, 13 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 17 + totalColumn).Address(False, False) & ")," & _
                    ws.Cells(excelRow, 29 + totalColumn).Address(False, False) & "," & _
                    ws.Cells(excelRow, 33 + totalColumn).Address(False, False) & "))"
            End If
            output(rowIndex, 18 + totalColumn) = formulaText
            output(rowIndex, 22 + totalColumn) = "=IF(" & _
                ws.Cells(excelRow, 21 + totalColumn).Address(False, False) & _
                "="""",""""," & ws.Cells(excelRow, 21 + totalColumn).Address(False, False) & _
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
    Dim chartObject As ChartObject, chartSeries As Series
    Dim seriesIndex As Long, seriesLabels As Variant, seriesColors As Variant
    On Error Resume Next
    Set chartObject = ws.ChartObjects(chartName)
    On Error GoTo 0
    If chartObject Is Nothing Then
        Set chartObject = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
        chartObject.Name = chartName
    Else
        chartObject.Left = chartLeft
        chartObject.Top = chartTop
        chartObject.Width = chartWidth
        chartObject.Height = chartHeight
    End If

    seriesLabels = Array("-90" & ChrW$(176), "0" & ChrW$(176), _
                         "90" & ChrW$(176), "180" & ChrW$(176))
    seriesColors = Array(RGB(31, 78, 121), RGB(237, 125, 49), _
                         RGB(112, 173, 71), RGB(112, 48, 160))
    With chartObject.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For seriesIndex = 0 To 3
            Set chartSeries = .SeriesCollection.NewSeries
            chartSeries.Name = CStr(seriesLabels(seriesIndex))
            chartSeries.XValues = ws.Range(ws.Cells(FIRST_DATA_ROW, 4), _
                                           ws.Cells(lastDataRow, 4))
            chartSeries.Values = ws.Range(ws.Cells(FIRST_DATA_ROW, _
                                          firstValueColumn + seriesIndex), _
                                          ws.Cells(lastDataRow, firstValueColumn + seriesIndex))
            chartSeries.MarkerStyle = xlMarkerStyleNone
            chartSeries.Format.Line.ForeColor.RGB = CLng(seriesColors(seriesIndex))
            chartSeries.Format.Line.Weight = 1.75
        Next seriesIndex
        .HasTitle = True
        .ChartTitle.Text = titleText
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
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
        ws.Cells(1, 13 + c).Value2 = "PCD1"
        ws.Cells(1, 17 + c).Value2 = "PCD2"
        ws.Cells(1, 21 + c).Value2 = "TOTAL"
        ws.Cells(1, 25 + c).Value2 = "UC"
        If Not earlyOnly Then
            ws.Cells(1, 29 + c).Value2 = "HYDROTEST"
            ws.Cells(1, 33 + c).Value2 = "DESIGN OP."
        End If
        ws.Cells(3, 5 + c).Value2 = "Radius = " & fiberText
        ws.Cells(3, 9 + c).Value2 = "Radius = " & fiberText
        ws.Cells(3, 13 + c).Value2 = "Radius = " & fiberText
        ws.Cells(3, 17 + c).Value2 = "Radius = " & fiberText
        ws.Cells(3, 21 + c).Value2 = "Radius = " & fiberText
        ws.Cells(3, 25 + c).Value2 = "Radius = " & fiberText
        If Not earlyOnly Then
            ws.Cells(3, 29 + c).Value2 = "Radius = " & fiberText
            ws.Cells(3, 33 + c).Value2 = "Radius = " & fiberText
        End If
        ws.Cells(4, 5 + c).Value2 = "Angle = " & CStr(angleList(c))
        ws.Cells(4, 9 + c).Value2 = "Angle = " & CStr(angleList(c))
        ws.Cells(4, 13 + c).Value2 = "Angle = " & CStr(angleList(c))
        ws.Cells(4, 17 + c).Value2 = "Angle = " & CStr(angleList(c))
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
    ws.Range("Q2").Value2 = "SUM"
    ws.Range("U2").Value2 = "SUM"
    ws.Range("Y2").Value2 = "TOTAL / ALLOWABLE"
    If Not earlyOnly Then
        ws.Range("AC2").Value2 = "2 CYCLES"
        ws.Range("AG2").Value2 = "1 CYCLE"
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
        formulaArgs = SegmentRangeArguments(stressSheet, ws, lastSegmentRow, _
            "DBM", summaryColumn)
        If Len(formulaArgs) > 0 Then ws.Cells(5, summaryColumn).Formula = "=MAX(" & formulaArgs & ")"
        formulaArgs = SegmentRangeArguments(stressSheet, ws, lastSegmentRow, _
            "RB", summaryColumn)
        If Len(formulaArgs) > 0 Then ws.Cells(8, summaryColumn).Formula = "=MAX(" & formulaArgs & ")"
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
        ws.Cells(6, summaryColumn).Formula = "=IF(" & ws.Cells(5, summaryColumn).Address(False, False) & _
            "="""",""""," & ws.Cells(5, summaryColumn).Address(False, False) & _
            "/'" & CONTROL_SHEET & "'!$I$8)"
        ws.Cells(9, summaryColumn).Formula = "=IF(" & ws.Cells(8, summaryColumn).Address(False, False) & _
            "="""",""""," & ws.Cells(8, summaryColumn).Address(False, False) & _
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
                    summarySheet.Cells(endRow, summaryColumn)).Address(False, False)
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
    For Each key In pairs.Keys
        If Len(pairFilter) = 0 Or InStr(1, CStr(key), pairFilter, vbTextCompare) > 0 Then _
            matches.Add CStr(key)
    Next key
    If matches.Count = 1 Then SelectPairName = matches(1): Exit Function
    If matches.Count = 0 Then Err.Raise vbObjectError + 700, , _
        "No step pair matches the pair-name filter: " & pairFilter
    If mProgrammaticMode Then Err.Raise vbObjectError + 701, , _
        "The report contains multiple step pairs. Enter a unique pair-name filter in Workflow Controls!" & PAIR_FILTER_CELL & "."
    promptText = "Enter a unique part of one step-pair name:" & vbCrLf
    For Each key In pairs.Keys: promptText = promptText & vbCrLf & CStr(key): Next key
    entered = InputBox(promptText, "Select step pair")
    If Len(Trim$(entered)) = 0 Then Err.Raise vbObjectError + 702, , "No step pair was selected."
    SelectPairName = SelectPairName(pairs, entered)
End Function

Private Function StressStartColumn(ByVal phaseName As String, ByVal loadCase As String) As Long
    Select Case NormalizePhase(phaseName)
        Case "EARLY UNDRAINED"
            Select Case NormalizeLoadCase(loadCase)
                Case "PCD1": StressStartColumn = 16
                Case "HCD": StressStartColumn = 20
                Case "PCD2": StressStartColumn = 24
                Case "FCD": StressStartColumn = 28
            End Select
        Case "EARLY DRAINED"
            Select Case NormalizeLoadCase(loadCase)
                Case "PCD1": StressStartColumn = 32
                Case "HCD": StressStartColumn = 36
                Case "PCD2": StressStartColumn = 40
                Case "FCD": StressStartColumn = 44
            End Select
        Case "MIDDLE DRAINED"
            Select Case NormalizeLoadCase(loadCase)
                Case "PCD1": StressStartColumn = 48
                Case "HCD": StressStartColumn = 52
                Case "PCD2": StressStartColumn = 56
                Case "FCD": StressStartColumn = 60
            End Select
        Case "LATE DRAINED"
            Select Case NormalizeLoadCase(loadCase)
                Case "PCD1": StressStartColumn = 95
                Case "HCD": StressStartColumn = 99
                Case "PCD2": StressStartColumn = 103
                Case "FCD": StressStartColumn = 107
            End Select
        Case "HYDROTEST"
            StressStartColumn = 111
        Case "DESIGN OPERATION"
            StressStartColumn = 115
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
        Case "HYDRO", "HYDROTEST", "HYDRO TEST": NormalizePhase = "HYDROTEST"
        Case "DESIGN", "DESIGN OP", "DESIGN OPERATION", "DESIGN STEP": _
            NormalizePhase = "DESIGN OPERATION"
        Case Else: Err.Raise vbObjectError + 705, , "Unsupported phase: " & textValue
    End Select
End Function

Private Function NormalizeLoadCase(ByVal textValue As String) As String
    textValue = UCase$(Replace(Trim$(textValue), " ", vbNullString))
    Select Case textValue
        Case "FCD", "HCD", "PCD1", "PCD2": NormalizeLoadCase = textValue
        Case Else: Err.Raise vbObjectError + 706, , "Load case must be FCD, HCD, PCD1, or PCD2."
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
    keys = kpValues.Keys
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
    If normalizedPhase = "HYDROTEST" Or normalizedPhase = "DESIGN OPERATION" Then
        startColumn = StressStartColumn(normalizedPhase, "FCD")
        PhaseHasStressData = Application.WorksheetFunction.Count(ws.Range( _
            ws.Cells(FIRST_DATA_ROW, startColumn), _
            ws.Cells(FIRST_DATA_ROW + kpCount - 1, startColumn + 3))) > 0
        Exit Function
    End If
    For Each loadCase In Array("FCD", "HCD", "PCD1", "PCD2")
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
    ColumnLetter = Split(Cells(1, columnNumber).Address(True, False), "$")(0)
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
