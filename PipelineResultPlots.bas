Attribute VB_Name = "PipelineResultPlots"
Option Explicit

' Plots along pipeline KP for the reports of extract_tension_depth.py and
' extract_mech_strain.py, with the DBM and curve-section KP ranges shaded.

Private Const CONTROL_SHEET As String = "Controls"
Private Const INPUT_SHEET As String = "INPUT"
Private Const TENSION_SHEET As String = "EffTension"
Private Const STRAIN_SHEET As String = "MechStrain"
Private Const U2_SHEET As String = "LocalU2"
Private Const STATUS_CELL As String = "I4"
Private Const INPUT_HEADER_ROW As Long = 5
Private Const DATA_HEADER_ROW As Long = 4
Private Const DATA_FIRST_COLUMN As Long = 17
Private Const ZONE_WIDTH As Long = 80
Private Const KIND_COUNT As Long = 8

' ---- report kinds -------------------------------------------------------------

Private Function KindKeyword(ByVal kind As Long) As String
    KindKeyword = CStr(Array("", "EFFECTIVE TENSION", "Z COORDINATE", "MECHANICAL STRAIN", "LE11", "THE11", _
        "LOCAL U2", "COPEN", "CPRESS")(kind))
End Function

Private Function KindSheet(ByVal kind As Long) As String
    KindSheet = CStr(Array("", TENSION_SHEET, TENSION_SHEET, STRAIN_SHEET, STRAIN_SHEET, STRAIN_SHEET, _
        U2_SHEET, "COPEN", "CPRESS")(kind))
End Function

' Position of the kind's data zone on its tab (1 = first zone).
Private Function KindSlot(ByVal kind As Long) As Long
    KindSlot = CLng(Array(0, 1, 2, 1, 2, 3, 1, 1, 1)(kind))
End Function

Private Function KindTitle(ByVal kind As Long) As String
    KindTitle = CStr(Array("", "Effective tension along KP", "Z coordinate (water depth) along KP", _
        "Mechanical strain (LE11 - THE11) along KP, max / min", _
        "LE11 along KP, max / min", _
        "THE11 along KP, max / min", _
        "Local U2 (lateral displacement) along KP", _
        "Contact opening COPEN along KP", _
        "Contact pressure CPRESS along KP")(kind))
End Function

Private Function KindAxisTitle(ByVal kind As Long) As String
    KindAxisTitle = CStr(Array("", "Effective tension [model force unit]", "Z coordinate [model length unit]", _
        "Mechanical strain [%]", "LE11 [%]", "THE11 [%]", "U2 [model length unit]", _
        "COPEN [model length unit]", "CPRESS [model pressure unit]")(kind))
End Function

Private Function KindFormat(ByVal kind As Long) As String
    KindFormat = CStr(Array("", "#,##0", "#,##0", "0.00%", "0.00%", "0.00%", "General", "General", "General")(kind))
End Function

Private Function ZoneColumn(ByVal kind As Long) As Long
    ZoneColumn = DATA_FIRST_COLUMN + (KindSlot(kind) - 1) * ZONE_WIDTH
End Function

' Last column of the kind's data zone. Tabs with one kind (LocalU2, COPEN,
' CPRESS) use every column to the right, so several element sets fit.
Private Function ZoneEndColumn(ByVal ws As Worksheet, ByVal kind As Long) As Long
    If kind >= 6 Then
        ZoneEndColumn = Application.Max(ws.Cells(DATA_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column, _
                                        ZoneColumn(kind) + ZONE_WIDTH - 1)
    Else
        ZoneEndColumn = ZoneColumn(kind) + ZONE_WIDTH - 1
    End If
End Function

Private Function ResultSheets() As Variant
    ResultSheets = Array(TENSION_SHEET, STRAIN_SHEET, U2_SHEET, "COPEN", "CPRESS")
End Function

Private Sub RedrawAll()
    Dim sheetName As Variant
    For Each sheetName In ResultSheets()
        RedrawSheet CStr(sheetName)
    Next sheetName
End Sub

' ---- public macros ------------------------------------------------------------

' Builds the Controls, INPUT, EffTension and MechStrain tabs. Run once in a new workbook.
Public Sub SetupPipelinePlotsWorkbook()
    Application.ScreenUpdating = False
    BuildSheets
    ThisWorkbook.Worksheets(CONTROL_SHEET).Activate
    Application.ScreenUpdating = True
End Sub

' Setup and save as a macro-enabled workbook (used once to create the file).
Public Sub SetupPipelinePlotsWorkbookAs(ByVal savePath As String)
    SetupPipelinePlotsWorkbook
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=savePath, FileFormat:=52
    Application.DisplayAlerts = True
End Sub

' Button 1: pick one or more reports (EFFTENSION, ZCOORD, MECHSTRAIN, LE11, THE11, LOCALU2, COPEN_<set>, CPRESS_<set>).
Public Sub ImportPipelineReports()
    Dim picked As Variant, i As Long, pathList As String, fileCount As Long
    On Error GoTo Failed
    picked = Application.GetOpenFilename("Reports (*.rpt;*.txt),*.rpt;*.txt,All files (*.*),*.*", , _
        "Select the reports (hold Ctrl to select several)", , True)
    If VarType(picked) = vbBoolean Then Exit Sub
    For i = LBound(picked) To UBound(picked)
        pathList = pathList & "|" & CStr(picked(i))
    Next i
    Application.ScreenUpdating = False
    fileCount = ImportCore(pathList)
    SetStatus "Imported " & CStr(fileCount) & " report(s). Charts updated."
    MsgBox CStr(fileCount) & " report(s) imported. Each result has its own tab with the charts.", _
           vbInformation, "Pipeline plots"
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetStatus "ERROR: " & Err.Description
    MsgBox "Import failed: " & Err.Description, vbCritical, "Pipeline plots"
    Resume CleanExit
End Sub

' Same as button 1 without dialogs; pathList = report paths separated by "|".
Public Sub ImportPipelineReportsFromPaths(ByVal pathList As String)
    Dim fileCount As Long
    On Error GoTo Failed
    Application.ScreenUpdating = False
    fileCount = ImportCore(pathList)
    SetStatus "PROGRAMMATIC IMPORT PASSED: " & CStr(fileCount) & " report(s)"
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetStatus "ERROR: " & Err.Description
    Resume CleanExit
End Sub

' Button 2: redraw all charts (after changing the DBM or curve KP ranges).
Public Sub RedrawPipelineCharts()
    On Error GoTo Failed
    Application.ScreenUpdating = False
    RedrawAll
    SetStatus "Charts redrawn."
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetStatus "ERROR: " & Err.Description
    MsgBox "Redraw failed: " & Err.Description, vbCritical, "Pipeline plots"
    Resume CleanExit
End Sub

' Button 3: copy the DBM and curve-section KP ranges from the INPUT tab of the fatigue workbook.
Public Sub CopyKpRangesFromFatigueWorkbook()
    Dim picked As Variant
    picked = Application.GetOpenFilename("Excel workbooks (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , _
        "Select the fatigue workbook (FatigueDamageCal_ManualWorkflows_INPUT.xlsm)")
    If VarType(picked) = vbBoolean Then Exit Sub
    CopyKpRangesFromPath CStr(picked)
    If Left$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET).Range(STATUS_CELL).Value2), 5) <> "ERROR" Then
        MsgBox "KP ranges copied to the INPUT tab. Charts redrawn.", vbInformation, "Pipeline plots"
    Else
        MsgBox CStr(ThisWorkbook.Worksheets(CONTROL_SHEET).Range(STATUS_CELL).Value2), vbCritical, "Pipeline plots"
    End If
End Sub

Public Sub CopyKpRangesFromPath(ByVal workbookPath As String)
    Dim src As Workbook, srcWs As Worksheet, ws As Worksheet, r As Long, n As Long, c As Long
    Dim wasOpen As Boolean, oldEvents As Boolean, oldSecurity As Long
    On Error GoTo Failed
    Application.ScreenUpdating = False
    oldEvents = Application.EnableEvents
    Application.EnableEvents = False
    For Each src In Application.Workbooks
        If LCase$(src.FullName) = LCase$(workbookPath) Then
            wasOpen = True
            Exit For
        End If
    Next src
    If Not wasOpen Then
        oldSecurity = Application.AutomationSecurity
        Application.AutomationSecurity = 3
        Set src = Application.Workbooks.Open(Filename:=workbookPath, ReadOnly:=True, UpdateLinks:=0)
        Application.AutomationSecurity = oldSecurity
    End If
    Set srcWs = src.Worksheets("INPUT")
    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
    ws.Range(ws.Cells(INPUT_HEADER_ROW + 1, 1), ws.Cells(INPUT_HEADER_ROW + 200, 3)).ClearContents
    ws.Range(ws.Cells(INPUT_HEADER_ROW + 1, 5), ws.Cells(INPUT_HEADER_ROW + 200, 7)).ClearContents
    n = 0
    For r = 26 To 225
        If IsNumeric(srcWs.Cells(r, 2).Value2) And Not IsEmpty(srcWs.Cells(r, 2).Value2) Then
            n = n + 1
            For c = 1 To 3
                ws.Cells(INPUT_HEADER_ROW + n, c).Value2 = srcWs.Cells(r, c).Value2
            Next c
        End If
    Next r
    n = 0
    If CStr(srcWs.Cells(25, 25).Value2) = "Curve ID" Then
        For r = 26 To 225
            If IsNumeric(srcWs.Cells(r, 26).Value2) And Not IsEmpty(srcWs.Cells(r, 26).Value2) Then
                n = n + 1
                For c = 1 To 3
                    ws.Cells(INPUT_HEADER_ROW + n, 4 + c).Value2 = srcWs.Cells(r, 24 + c).Value2
                Next c
            End If
        Next r
    End If
    If Not wasOpen Then src.Close SaveChanges:=False
    RedrawAll
    SetStatus "KP ranges copied from " & workbookPath
CleanExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    SetStatus "ERROR: " & Err.Description
    Resume CleanExit
End Sub

' ---- sheets -------------------------------------------------------------------

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    If Not SheetExists(sheetName) Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = sheetName
    End If
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Private Sub SetStatus(ByVal statusText As String)
    If SheetExists(CONTROL_SHEET) Then ThisWorkbook.Worksheets(CONTROL_SHEET).Range(STATUS_CELL).Value2 = statusText
End Sub

Private Sub BuildSheets()
    Dim ws As Worksheet, i As Long, kind As Long, isNew As Boolean
    ' Controls
    If Not SheetExists(CONTROL_SHEET) Then
        If ThisWorkbook.Worksheets.Count = 1 And Application.WorksheetFunction.CountA(ThisWorkbook.Worksheets(1).Cells) = 0 Then
            ThisWorkbook.Worksheets(1).Name = CONTROL_SHEET
        End If
    End If
    Set ws = EnsureSheet(CONTROL_SHEET)
    ws.Cells.Clear
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.Font.Name = "Arial"
    ws.Range("B1").Value2 = "Pipeline Result Plots along KP"
    ws.Range("B1").Font.Size = 16
    ws.Range("B1").Font.Bold = True
    ws.Range("B2").Value2 = "Effective tension, Z coordinate, strain, local U2, COPEN and CPRESS from the Abaqus extraction scripts."
    ws.Range("H3").Value2 = "Status"
    ws.Range("H3:M3").Font.Bold = True
    ws.Range("H3:M3").Font.Color = vbWhite
    ws.Range("H3:M3").Interior.Color = RGB(31, 78, 121)
    ws.Range("I4:M4").Merge
    ws.Range("I4").WrapText = True
    ws.Range("I4").Value2 = "Ready."
    ws.Range("H4").Value2 = "Last action"
    ws.Range("H4").Font.Bold = True
    ws.Range("H6").Value2 = "How to use"
    ws.Range("H6:M6").Font.Bold = True
    ws.Range("H6:M6").Font.Color = vbWhite
    ws.Range("H6:M6").Interior.Color = RGB(68, 114, 196)
    ws.Range("H7").Value2 = "1. Enter the DBM and curve-section KP ranges on the INPUT tab (or copy them from the fatigue workbook with button 3)."
    ws.Range("H8").Value2 = "2. Button 1: select the .rpt files (hold Ctrl for several): EFFTENSION, ZCOORD, MECHSTRAIN, LE11, THE11, LOCALU2, COPEN_<set>, CPRESS_<set>."
    ws.Range("H9").Value2 = "3. Charts: " & TENSION_SHEET & " (tension, Z), " & STRAIN_SHEET & " (strains), " & U2_SHEET & ", COPEN and CPRESS tabs."
    ws.Range("H10").Value2 = "4. Button 2 redraws the charts after the KP ranges change. A new import replaces the same kind of report (select all COPEN / CPRESS sets together)."
    AddButton ws, "PP_Import", "1. Import Reports", "ImportPipelineReports", "B4", RGB(31, 78, 121)
    AddButton ws, "PP_Redraw", "2. Redraw Charts", "RedrawPipelineCharts", "B7", RGB(112, 173, 71)
    AddButton ws, "PP_Copy", "3. Copy KP Ranges from Fatigue Workbook", "CopyKpRangesFromFatigueWorkbook", "B10", RGB(89, 89, 89)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 2
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I:M").ColumnWidth = 22
    ws.Rows("1:12").RowHeight = 21
    ws.Rows("4:4").RowHeight = 32

    ' INPUT
    isNew = Not SheetExists(INPUT_SHEET)
    Set ws = EnsureSheet(INPUT_SHEET)
    ws.Cells.Font.Name = "Arial"
    ws.Range("A1").Value2 = "KP ranges shown on the charts"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value2 = "Yellow cells are inputs. Press button 2 on the Controls tab after changing them."
    ws.Cells(INPUT_HEADER_ROW - 1, 1).Value2 = "DBM KP ranges (amber bands)"
    ws.Cells(INPUT_HEADER_ROW - 1, 5).Value2 = "Curve section KP ranges (blue bands)"
    ws.Cells(INPUT_HEADER_ROW - 1, 1).Font.Bold = True
    ws.Cells(INPUT_HEADER_ROW - 1, 5).Font.Bold = True
    For i = 0 To 1
        ws.Cells(INPUT_HEADER_ROW, 1 + 4 * i).Value2 = CStr(Array("DBM ID", "Curve ID")(i))
        ws.Cells(INPUT_HEADER_ROW, 2 + 4 * i).Value2 = "KP Start [m]"
        ws.Cells(INPUT_HEADER_ROW, 3 + 4 * i).Value2 = "KP End [m]"
        With ws.Range(ws.Cells(INPUT_HEADER_ROW, 1 + 4 * i), ws.Cells(INPUT_HEADER_ROW, 3 + 4 * i))
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlCenter
        End With
        With ws.Range(ws.Cells(INPUT_HEADER_ROW + 1, 1 + 4 * i), ws.Cells(INPUT_HEADER_ROW + 60, 3 + 4 * i))
            .Interior.Color = RGB(255, 242, 204)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(217, 217, 217)
        End With
    Next i
    ws.Columns("A:C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E:G").ColumnWidth = 18

    ' result tabs
    For i = 0 To UBound(ResultSheets())
        Set ws = EnsureSheet(CStr(ResultSheets()(i)))
        ws.Cells.Font.Name = "Arial"
        ws.Range("A1").Value2 = CStr(Array("Effective tension and Z coordinate along KP", "Strain along KP", _
            "Local U2 along KP", "Contact opening COPEN along KP", "Contact pressure CPRESS along KP")(i))
        ws.Range("A1").Font.Size = 14
        ws.Range("A1").Font.Bold = True
        If ws.ChartObjects.Count = 0 Then ws.Range("A3").Value2 = "No data yet. Use button 1 on the Controls tab."
    Next i
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal captionText As String, _
                      ByVal macroName As String, ByVal anchorCell As String, ByVal fillColor As Long)
    Dim buttonShape As Shape, anchor As Range
    Set anchor = ws.Range(anchorCell)
    Set buttonShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, 250, 43.2)
    buttonShape.Name = shapeName
    buttonShape.OnAction = macroName
    buttonShape.Fill.ForeColor.RGB = fillColor
    buttonShape.Line.ForeColor.RGB = RGB(255, 255, 255)
    buttonShape.TextFrame2.TextRange.Text = captionText
    buttonShape.TextFrame2.TextRange.Font.Name = "Arial"
    buttonShape.TextFrame2.TextRange.Font.Size = 12
    buttonShape.TextFrame2.TextRange.Font.Bold = msoTrue
    buttonShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    buttonShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    buttonShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
End Sub

' ---- import -------------------------------------------------------------------

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

Private Function TryNumber(ByVal textValue As Variant, ByRef numberValue As Double) As Boolean
    Dim cleaned As String
    On Error GoTo NotNumber
    cleaned = Trim$(CStr(textValue))
    If Len(cleaned) = 0 Or UCase$(cleaned) = "NAN" Or UCase$(cleaned) = "INF" Then Exit Function
    If Not IsNumeric(cleaned) Then Exit Function
    numberValue = CDbl(cleaned)
    TryNumber = True
NotNumber:
End Function

Private Function ImportCore(ByVal pathList As String) As Long
    Dim paths As Variant, i As Long, cleared(1 To KIND_COUNT) As Boolean
    If Not SheetExists(INPUT_SHEET) Or Not SheetExists("CPRESS") Then BuildSheets
    paths = Split(pathList, "|")
    For i = LBound(paths) To UBound(paths)
        If Len(Trim$(CStr(paths(i)))) > 0 Then
            ImportOneReport Trim$(CStr(paths(i))), cleared
            ImportCore = ImportCore + 1
        End If
    Next i
    RedrawAll
End Function

Private Sub ImportOneReport(ByVal reportPath As String, ByRef cleared() As Boolean)
    Dim reportLines As Variant, parts As Variant, headerIndex As Long, i As Long, c As Long
    Dim ws As Worksheet, output() As Variant, lineText As String, setName As String
    Dim rowCount As Long, columnCount As Long, numberValue As Double
    Dim kind As Long, k As Long, firstColumn As Long, zoneStart As Long, zoneEnd As Long
    If Len(Dir$(reportPath)) = 0 Then Err.Raise vbObjectError + 800, , "File not found: " & reportPath
    reportLines = ReadReportLines(reportPath)
    headerIndex = -1
    For i = LBound(reportLines) To UBound(reportLines)
        lineText = Trim$(CStr(reportLines(i)))
        If UCase$(Left$(lineText, 2)) = "KP" And InStr(lineText, vbTab) > 0 Then
            headerIndex = i
            Exit For
        End If
        If kind = 0 Then
            For k = 1 To KIND_COUNT
                If UCase$(Left$(lineText, Len(KindKeyword(k)))) = KindKeyword(k) Then kind = k
            Next k
        End If
        If UCase$(Left$(lineText, 4)) = "SET:" Then setName = Trim$(Mid$(lineText, 5))
    Next i
    If headerIndex < 0 Or kind = 0 Then Err.Raise vbObjectError + 801, , _
        "Not a report of the extraction scripts: " & reportPath
    parts = Split(CStr(reportLines(headerIndex)), vbTab)
    columnCount = UBound(parts) + 1
    If columnCount < 2 Then Err.Raise vbObjectError + 802, , "The report has no step columns: " & reportPath
    ReDim output(1 To UBound(reportLines) - headerIndex + 1, 1 To columnCount)
    For i = headerIndex + 1 To UBound(reportLines)
        parts = Split(CStr(reportLines(i)), vbTab)
        If UBound(parts) >= 1 Then
            If TryNumber(parts(0), numberValue) Then
                rowCount = rowCount + 1
                output(rowCount, 1) = numberValue
                For c = 1 To Application.Min(UBound(parts), columnCount - 1)
                    If TryNumber(parts(c), numberValue) Then output(rowCount, c + 1) = numberValue
                Next c
            End If
        End If
    Next i
    If rowCount = 0 Then Err.Raise vbObjectError + 804, , "The report has no data rows: " & reportPath

    Set ws = ThisWorkbook.Worksheets(KindSheet(kind))
    zoneStart = ZoneColumn(kind)
    zoneEnd = ZoneEndColumn(ws, kind)
    If Not cleared(kind) Then
        ws.Range(ws.Cells(1, zoneStart), ws.Cells(ws.Rows.Count, zoneEnd)).Clear
        cleared(kind) = True
    End If
    ' next free block in the zone
    firstColumn = zoneStart
    For c = zoneEnd To zoneStart Step -1
        If Len(CStr(ws.Cells(DATA_HEADER_ROW, c).Value2)) > 0 Then
            firstColumn = c + 2
            Exit For
        End If
    Next c
    If kind < 6 And firstColumn + columnCount - 1 > zoneEnd Then Err.Raise vbObjectError + 803, , _
        "Too many data columns for the " & KindSheet(kind) & " tab (limit " & CStr(ZONE_WIDTH - 2) & _
        "); extract fewer steps: " & reportPath
    ws.Cells(1, firstColumn).Value2 = setName
    ws.Cells(2, firstColumn).Value2 = KindTitle(kind) & " (chart data)"
    ws.Range(ws.Cells(1, firstColumn), ws.Cells(2, firstColumn)).Font.Bold = True
    ws.Cells(3, firstColumn).Value2 = "Source file: " & reportPath
    parts = Split(CStr(reportLines(headerIndex)), vbTab)
    For c = 0 To columnCount - 1
        ws.Cells(DATA_HEADER_ROW, firstColumn + c).Value2 = Trim$(CStr(parts(c)))
    Next c
    ws.Cells(DATA_HEADER_ROW, firstColumn).Resize(1, columnCount).Font.Bold = True
    ws.Cells(DATA_HEADER_ROW + 1, firstColumn).Resize(rowCount, columnCount).Value = output
    ws.Range(ws.Cells(1, firstColumn), ws.Cells(ws.Rows.Count, firstColumn + columnCount - 1)).Font.Name = "Arial"
    ws.Range(ws.Columns(firstColumn), ws.Columns(firstColumn + columnCount - 1)).ColumnWidth = 16
End Sub

' ---- charts -------------------------------------------------------------------

Private Function IsKpHeader(ByVal ws As Worksheet, ByVal c As Long) As Boolean
    IsKpHeader = (Left$(UCase$(Trim$(CStr(ws.Cells(DATA_HEADER_ROW, c).Value2))), 2) = "KP")
End Function

' First and last KP over the blocks of a kind's zone. False when the zone is empty.
Private Function ZoneKpRange(ByVal ws As Worksheet, ByVal kind As Long, _
                             ByRef kpFirst As Double, ByRef kpLast As Double) As Boolean
    Dim c As Long, lastRow As Long, v As Double
    For c = ZoneColumn(kind) To ZoneEndColumn(ws, kind)
        If IsKpHeader(ws, c) Then
            lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
            If lastRow > DATA_HEADER_ROW Then
                v = CDbl(ws.Cells(DATA_HEADER_ROW + 1, c).Value2)
                If Not ZoneKpRange Or v < kpFirst Then kpFirst = v
                v = CDbl(ws.Cells(lastRow, c).Value2)
                If Not ZoneKpRange Or v > kpLast Then kpLast = v
                ZoneKpRange = True
            End If
        End If
    Next c
End Function

Private Sub RedrawSheet(ByVal sheetName As String)
    Dim ws As Worksheet, kind As Long, i As Long
    Dim kpFirst As Double, kpLast As Double, haveKp As Boolean, a As Double, b As Double
    Dim xMin As Double, xMax As Double, unitStep As Double, chartTop As Double
    If Not SheetExists(sheetName) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
    For kind = 1 To KIND_COUNT
        If KindSheet(kind) = sheetName Then
            If ZoneKpRange(ws, kind, a, b) Then
                If Not haveKp Or a < kpFirst Then kpFirst = a
                If Not haveKp Or b > kpLast Then kpLast = b
                haveKp = True
            End If
        End If
    Next kind
    If Not haveKp Then
        ws.Range("A3").Value2 = "No data yet. Use button 1 on the Controls tab."
        Exit Sub
    End If
    ws.Range("A3").ClearContents
    If kpLast > kpFirst Then
        unitStep = 10 ^ Int(Log(kpLast - kpFirst) / Log(10#)) / 10#
        xMin = Int(kpFirst / unitStep) * unitStep
        xMax = -Int(-kpLast / unitStep) * unitStep
    Else
        xMin = kpFirst: xMax = kpFirst + 1
    End If
    chartTop = ws.Range("A3").Top
    For kind = 1 To KIND_COUNT
        If KindSheet(kind) = sheetName Then
            If ZoneKpRange(ws, kind, a, b) Then
                AddZoneChart ws, kind, chartTop, xMin, xMax
                chartTop = chartTop + 320
            End If
        End If
    Next kind
End Sub

Private Sub AddZoneChart(ByVal ws As Worksheet, ByVal kind As Long, ByVal chartTop As Double, _
                         ByVal xMin As Double, ByVal xMax As Double)
    Dim plotObject As ChartObject, plotChart As Chart, plotSeries As Series
    Dim zoneStart As Long, zoneEnd As Long, lastRow As Long, c As Long, headerText As String
    Dim blockColumn As Long, setLabel As String, maxRow As Long
    Dim yLow As Double, yHigh As Double, yUnit As Double
    zoneStart = ZoneColumn(kind)
    zoneEnd = ZoneEndColumn(ws, kind)
    Set plotObject = ws.ChartObjects.Add(ws.Range("A3").Left, chartTop, 720, 300)
    plotObject.Name = "PP_Chart" & CStr(kind)
    Set plotChart = plotObject.Chart
    plotChart.ChartType = xlXYScatterLinesNoMarkers
    Do While plotChart.SeriesCollection.Count > 0
        plotChart.SeriesCollection(1).Delete
    Loop
    For c = zoneStart To zoneEnd
        headerText = Trim$(CStr(ws.Cells(DATA_HEADER_ROW, c).Value2))
        If IsKpHeader(ws, c) Then
            blockColumn = c
            setLabel = Trim$(CStr(ws.Cells(1, c).Value2))
            lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
            If lastRow > maxRow Then maxRow = lastRow
        ElseIf Len(headerText) = 0 Then
            blockColumn = 0
        ElseIf blockColumn > 0 And lastRow > DATA_HEADER_ROW Then
            Set plotSeries = plotChart.SeriesCollection.NewSeries
            If Len(setLabel) > 0 Then
                plotSeries.Name = setLabel & ": " & headerText
            Else
                plotSeries.Name = headerText
            End If
            plotSeries.XValues = ws.Range(ws.Cells(DATA_HEADER_ROW + 1, blockColumn), ws.Cells(lastRow, blockColumn))
            plotSeries.Values = ws.Range(ws.Cells(DATA_HEADER_ROW + 1, c), ws.Cells(lastRow, c))
            plotSeries.MarkerStyle = xlMarkerStyleNone
            plotSeries.Format.Line.Weight = 1.5
        End If
    Next c
    On Error Resume Next
    With plotChart
        .HasTitle = True
        .ChartTitle.Text = KindTitle(kind)
        .ChartTitle.Font.Name = "Arial"
        .ChartTitle.Font.Size = 12
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        .DisplayBlanksAs = xlNotPlotted
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Size = 9
        .ChartArea.Format.Line.ForeColor.RGB = RGB(166, 166, 166)
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "KP (m)"
        .Axes(xlCategory).TickLabels.NumberFormat = "0"
        .Axes(xlCategory).MinimumScale = xMin
        .Axes(xlCategory).MaximumScale = xMax
        .Axes(xlCategory).TickLabelPosition = xlLow
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = KindAxisTitle(kind)
        .Axes(xlValue).TickLabels.NumberFormat = KindFormat(kind)
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
        ' Fixed plot area, so the bands stay aligned with the KP axis
        .PlotArea.InsideLeft = 70
        .PlotArea.InsideTop = 50
        .PlotArea.InsideWidth = 628
        .PlotArea.InsideHeight = 196
    End With
    ' Z coordinate: fit the axis to the data instead of starting at zero
    If kind = 2 And maxRow > DATA_HEADER_ROW Then
        yLow = Application.WorksheetFunction.Min(ws.Range(ws.Cells(DATA_HEADER_ROW + 1, zoneStart + 1), _
            ws.Cells(maxRow, zoneEnd)))
        yHigh = Application.WorksheetFunction.Max(ws.Range(ws.Cells(DATA_HEADER_ROW + 1, zoneStart + 1), _
            ws.Cells(maxRow, zoneEnd)))
        If yHigh > yLow Then
            yUnit = 10 ^ Int(Log(yHigh - yLow) / Log(10#))
            plotChart.Axes(xlValue).MinimumScale = Int(yLow / yUnit) * yUnit
            plotChart.Axes(xlValue).MaximumScale = -Int(-yHigh / yUnit) * yUnit
        End If
    End If
    On Error GoTo 0
    AddBands plotChart, xMin, xMax
End Sub

' DBM KP ranges (amber) and curve-section KP ranges (blue) of the INPUT tab.
Private Sub AddBands(ByVal plotChart As Chart, ByVal xMin As Double, ByVal xMax As Double)
    Dim note As Shape, dbmCount As Long, curveCount As Long, noteText As String
    If xMax <= xMin Or Not SheetExists(INPUT_SHEET) Then Exit Sub
    If plotChart.SeriesCollection.Count = 0 Then Exit Sub
    dbmCount = AddBandsFromTable(plotChart, 2, RGB(255, 192, 0), 0.6, "DBM_band_", xMin, xMax)
    curveCount = AddBandsFromTable(plotChart, 6, RGB(68, 114, 196), 0.75, "Curve_band_", xMin, xMax)
    If dbmCount > 0 Then noteText = "Amber bands = DBM KP ranges"
    If curveCount > 0 Then
        If Len(noteText) > 0 Then noteText = noteText & vbLf
        noteText = noteText & "Blue bands = curve sections"
    End If
    If Len(noteText) = 0 Then Exit Sub
    On Error Resume Next
    Set note = plotChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 560, 2, 156, 28)
    note.Name = "Band_note"
    note.TextFrame2.TextRange.Text = noteText
    note.TextFrame2.TextRange.Font.Size = 8
    note.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(89, 89, 89)
    note.Line.Visible = msoFalse
    note.Fill.Visible = msoFalse
    On Error GoTo 0
End Sub

Private Function IsKpNumber(ByVal cellValue As Variant) As Boolean
    If IsEmpty(cellValue) Or IsError(cellValue) Then Exit Function
    If Len(CStr(cellValue)) = 0 Then Exit Function
    IsKpNumber = IsNumeric(cellValue)
End Function

Private Function AddBandsFromTable(ByVal plotChart As Chart, ByVal startColumn As Long, _
        ByVal bandColor As Long, ByVal bandTransparency As Double, ByVal namePrefix As String, _
        ByVal xMin As Double, ByVal xMax As Double) As Long
    Dim inputWs As Worksheet, lastRow As Long, r As Long
    Dim startKP As Double, endKP As Double, tempValue As Double
    Dim band As Shape, leftPos As Double, widthPos As Double
    Set inputWs = ThisWorkbook.Worksheets(INPUT_SHEET)
    lastRow = Application.Max(inputWs.Cells(inputWs.Rows.Count, startColumn).End(xlUp).Row, _
                              inputWs.Cells(inputWs.Rows.Count, startColumn + 1).End(xlUp).Row)
    On Error Resume Next
    For r = INPUT_HEADER_ROW + 1 To lastRow
        If IsKpNumber(inputWs.Cells(r, startColumn).Value2) And _
           IsKpNumber(inputWs.Cells(r, startColumn + 1).Value2) Then
            startKP = CDbl(inputWs.Cells(r, startColumn).Value2)
            endKP = CDbl(inputWs.Cells(r, startColumn + 1).Value2)
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
                band.Name = namePrefix & CStr(r)
                band.Fill.ForeColor.RGB = bandColor
                band.Fill.Transparency = bandTransparency
                band.Line.Visible = msoFalse
                AddBandsFromTable = AddBandsFromTable + 1
            End If
        End If
    Next r
    On Error GoTo 0
End Function
