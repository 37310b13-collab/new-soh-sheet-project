Attribute VB_Name = "RefreshData_ProductionPlan"
Option Explicit

' ============================================================================
' RefreshData_ProductionPlan module
'
'   RefreshWeeklyBatches : When you select "Powder & Slurry & Pgm Plan"
'                          (revised monthly), updates Production_Plan (production
'                          plan batch counts). The file is a single sheet;
'                          column B = the intermediate/finished product
'                          (Catalyst)/Solution name, and the week-start-date
'                          header row (the row with dates lined up from
'                          column C onward) is auto-detected, with every row
'                          after that being name + weekly quantities.
'                          The per-batch usage rate (M_BOM) is not in this
'                          file - that's handled on the RefreshBOM (Raw
'                          Material - Look Up) side.
'
'   [About automatic row-type detection] Whether a row is an intermediate
'   (Slurry/Powder), a finished product (Catalyst), or a Solution is
'   determined mechanically from the name's prefix and the "Solution Name
'   List" (the T_SolutionNames table on the Control_Panel sheet). A name
'   starting with TSP-/TPP-/TSZ-/TVS-/VSP- is an intermediate; a name listed
'   in T_SolutionNames (aliases like 20P, SH, etc.) is a Solution (the alias
'   is automatically converted to its canonical name SOL-xxx); anything
'   matching neither is treated as a finished product (Catalyst), by
'   elimination. Even as rows are added/removed in the source file or new
'   Solutions are introduced, this determination never keys off a row
'   number anywhere, so no maintenance is ever needed on the macro side as
'   the file's row count changes (a new Solution only needs one row added
'   to T_SolutionNames).
'
' For the overall design rationale (performance, DataBodyRange,
' DisplayAlerts, etc.), see the comment at the top of the
' RefreshData_Utilities module.
' ============================================================================

Sub RefreshWeeklyBatches()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , _
        "Please select the latest version of Powder & Slurry & Pgm Plan")
    If srcPath = False Then Exit Sub

    Dim srcWb As Workbook
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' For a file with "read-only recommended" set, leaving DisplayAlerts=True
    ' causes a confirmation dialog to appear during Workbooks.Open, which
    ' destabilizes processing while waiting for a response (this was the
    ' cause of a bug where srcWb ultimately failed to be obtained correctly)
    ' - so suppress it.
    Application.DisplayAlerts = False

    Set srcWb = Workbooks.Open(CStr(srcPath), ReadOnly:=True, UpdateLinks:=False)

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim ppGrid As ListObject: Set ppGrid = thisWb.Sheets("Production_Plan").ListObjects("Production_Plan")
    Dim calWeeks As ListObject: Set calWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks")

    ' WeekStart (date serial value) -> column number within Production_Plan (within the data range, 1=Intermediate,2=Week1,...)
    Dim weekColByDate As Object: Set weekColByDate = CreateObject("Scripting.Dictionary")
    Dim calN As Long: calN = calWeeks.ListRows.Count
    If calN > 0 Then
        Dim calWeekStartData As Variant
        calWeekStartData = calWeeks.ListColumns("WeekStart").DataBodyRange.Value
        Dim i As Long
        For i = 1 To calN
            weekColByDate(CLng(CDate(calWeekStartData(i, 1)))) = i + 1
        Next i
    End If

    Dim ppIdx As Object: Set ppIdx = BuildNameIndex(ppGrid, "Intermediate")

    ' Load the Solution name list (alias -> canonical name) once.
    ' CompareMode=vbTextCompare absorbs case differences (20p/20P etc.).
    Dim solutionAlias As Object: Set solutionAlias = CreateObject("Scripting.Dictionary")
    solutionAlias.CompareMode = vbTextCompare
    Dim solTbl As ListObject: Set solTbl = thisWb.Sheets("Control_Panel").ListObjects("T_SolutionNames")
    Dim solN As Long: solN = solTbl.ListRows.Count
    If solN > 0 Then
        Dim solData As Variant: solData = solTbl.ListColumns(1).DataBodyRange.Resize(solN, 2).Value
        Dim si As Long
        For si = 1 To solN
            Dim aliasKey As String: aliasKey = Trim(CStr(solData(si, 1)))
            If Len(aliasKey) > 0 Then solutionAlias(aliasKey) = Trim(CStr(solData(si, 2)))
        Next si
    End If

    ' Find the target sheet (a single-sheet format. As a precaution, even
    ' if there are multiple sheets, auto-detect the one whose structure
    ' (the week-start-date header row) matches. Only one sheet per file is
    ' the target).
    Dim foundSheet As Worksheet
    Dim sh As Worksheet
    For Each sh In srcWb.Worksheets
        If IsWeeklyPlanSheet(sh) Then
            Set foundSheet = sh
            Exit For
        End If
    Next sh
    If foundSheet Is Nothing Then
        Err.Raise vbObjectError + 2, , "Could not find weekly data for ""Powder & Slurry & Pgm Plan"". Please check the file format."
    End If

    Const MAX_SCAN_ROWS As Long = 500
    Const MAX_SCAN_COLS As Long = 200
    Dim usedRows As Long, usedCols As Long
    usedRows = foundSheet.UsedRange.Rows.Count
    usedCols = foundSheet.UsedRange.Columns.Count
    If usedRows > MAX_SCAN_ROWS Then usedRows = MAX_SCAN_ROWS
    If usedCols > MAX_SCAN_COLS Then usedCols = MAX_SCAN_COLS

    ' Read the whole sheet into an array just once, and from then on only
    ' reference the in-memory array (data) (calling .Cells(r,c).Value on
    ' every loop iteration piles up COM round-trips and becomes extremely
    ' slow).
    Dim data As Variant
    data = foundSheet.Range(foundSheet.Cells(1, 1), foundSheet.Cells(usedRows, usedCols)).Value

    Dim hdrRow As Long: hdrRow = FindDateHeaderRow(data, usedRows, usedCols)
    If hdrRow = 0 Then
        Err.Raise vbObjectError + 3, , "Could not find the week-start-date header row. Please check the file format."
    End If

    Dim dateCols As Object: Set dateCols = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 3 To usedCols
        If IsDate(data(hdrRow, c)) Then dateCols(c) = CLng(CDate(data(hdrRow, c)))
    Next c

    Dim updatedCells As Long, newInterRows As Long
    updatedCells = 0
    newInterRows = 0

    ' [Important] Production_Plan rows come in two kinds: (1) intermediates/finished
    ' products/Solutions that appear directly in this file (normal rows;
    ' weekly batch counts are written as values), and (2) "pass-through
    ' intermediates" that are derived from other intermediates via M_BOM
    ' (weekly batch counts are a SUMPRODUCT formula; these never appear as
    ' rows in this file). So an approach that reads the whole table into an
    ' array and writes it back as a block (tried in an earlier version)
    ' would overwrite the pass-through intermediates' formula cells with
    ' "the calculated value at that moment," destroying the formulas - an
    ' actual accident. Writes must be limited to only the rows that
    ' actually appear in this file.
    '
    ' Also, calling ListRows.Add in a loop, one new intermediate row at a
    ' time, triggers a dependency recheck on every single call for a
    ' heavyweight table like Production_Plan that's referenced by a large number of
    ' formulas, and when there are many new rows (including a case where a
    ' name-matching gap causes rows that should have matched to all be
    ' judged "new"), this makes Excel stop responding (the same bug as
    ' RefreshBOM's M_BOM previously had). So new rows are first all counted,
    ' then the table is expanded all at once with a single Resize, and only
    ' the intermediate-name column is filled with a single array write.
    Dim canonNames As Object: Set canonNames = CreateObject("Scripting.Dictionary")
    Dim newNames As Object: Set newNames = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = hdrRow + 1 To usedRows
        Dim rawName As String: rawName = Trim(CStr(data(r, 2)))
        If Len(rawName) = 0 Then GoTo NextRowNames

        Dim canonName As String
        If solutionAlias.Exists(rawName) Then
            canonName = solutionAlias(rawName)
        Else
            canonName = rawName
        End If
        canonNames(r) = canonName

        If Not ppIdx.Exists(canonName) And Not newNames.Exists(canonName) Then
            newNames(canonName) = True
        End If
NextRowNames:
    Next r

    If newNames.Count > 0 Then
        Dim oldRowCount As Long: oldRowCount = ppGrid.ListRows.Count
        ppGrid.Resize ppGrid.Range.Resize(ppGrid.Range.Rows.Count + newNames.Count, ppGrid.Range.Columns.Count)
        Dim newNameArr() As Variant
        ReDim newNameArr(1 To newNames.Count, 1 To 1)
        Dim ni As Long: ni = 0
        Dim nameKey As Variant
        For Each nameKey In newNames.Keys
            ni = ni + 1
            newNameArr(ni, 1) = nameKey
            ppIdx(CStr(nameKey)) = oldRowCount + ni
            newInterRows = newInterRows + 1
        Next nameKey
        ppGrid.ListColumns(1).DataBodyRange.Resize(newNames.Count, 1).Offset(oldRowCount, 0).Value = newNameArr
    End If

    ' Only for rows that actually appear in this file, read a whole row at
    ' a time (not cell by cell) -> rewrite just the matching week(s) in
    ' memory -> write the whole row back (pass-through intermediate rows
    ' are never touched at all).
    Dim touchedRows As Object: Set touchedRows = CreateObject("Scripting.Dictionary")
    For r = hdrRow + 1 To usedRows
        If Not canonNames.Exists(r) Then GoTo NextRow
        Dim ppRowIndex As Long: ppRowIndex = ppIdx(canonNames(r))
        If Not touchedRows.Exists(ppRowIndex) Then
            touchedRows(ppRowIndex) = ppGrid.ListRows(ppRowIndex).Range.Value
        End If
        Dim rowArr As Variant: rowArr = touchedRows(ppRowIndex)

        Dim keyVariant As Variant
        For Each keyVariant In dateCols.Keys
            If weekColByDate.Exists(dateCols(keyVariant)) Then
                Dim colIdx As Long: colIdx = weekColByDate(dateCols(keyVariant))
                Dim v As Double: v = 0
                If IsNumeric(data(r, keyVariant)) Then v = data(r, keyVariant)
                rowArr(1, colIdx) = v
                updatedCells = updatedCells + 1
            End If
        Next keyVariant
        touchedRows(ppRowIndex) = rowArr
NextRow:
    Next r

    Dim rowKey As Variant
    For Each rowKey In touchedRows.Keys
        ppGrid.ListRows(CLng(rowKey)).Range.Value = touchedRows(rowKey)
    Next rowKey

    ' Guard the cleanup step itself against failing with "object variable
    ' not set," even in the case where srcWb has already become Nothing
    ' (some automated process on the source file's side can close the
    ' workbook right after it's opened).
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Production_Plan has been updated." & vbCrLf & _
           "Cells updated: " & updatedCells & vbCrLf & _
           "New intermediate/finished product (Cat)/Solution codes added: " & newInterRows & vbCrLf & vbCrLf & _
           "(Note) The per-batch usage rate (M_BOM) is not handled by this macro. If there is" & vbCrLf & _
           "an entirely new combination, please add it via RefreshBOM.", _
           vbInformation
    Exit Sub

ErrHandler:
    ' [Important] On Error Resume Next automatically clears the Err object
    ' (a VBA quirk), so the error number/description must be saved into
    ' variables before any cleanup code runs. Skipping this means the
    ' MsgBox below always shows "(blank)" and the real cause of the error
    ' is never known.
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "An error occurred during the refresh: (" & errNum & ") " & errMsg, vbCritical
End Sub

' Judges whether a sheet looks like a "weekly plan sheet" by whether, near
' the top of the sheet (within the first 10 rows), there is a row with 5 or
' more dates lined up from column C onward (only a small range is scanned,
' so this has no performance impact even with around 40 sheets).
Private Function IsWeeklyPlanSheet(sh As Worksheet) As Boolean
    Dim r As Long, c As Long
    Dim maxC As Long: maxC = Application.WorksheetFunction.Min(20, sh.UsedRange.Columns.Count)
    Dim maxR As Long: maxR = Application.WorksheetFunction.Min(10, sh.UsedRange.Rows.Count)
    For r = 1 To maxR
        Dim dateCount As Long: dateCount = 0
        For c = 3 To maxC
            If IsDate(sh.Cells(r, c).Value) Then dateCount = dateCount + 1
        Next c
        If dateCount >= 5 Then
            IsWeeklyPlanSheet = True
            Exit Function
        End If
    Next r
    IsWeeklyPlanSheet = False
End Function

' Finds the week-start-date header row (a row with 5 or more dates lined up
' from column C onward) from the already-arrayed sheet data (data).
Private Function FindDateHeaderRow(data As Variant, usedRows As Long, usedCols As Long) As Long
    Dim r As Long, c As Long
    Dim maxR As Long: maxR = Application.WorksheetFunction.Min(10, usedRows)
    Dim maxC As Long: maxC = Application.WorksheetFunction.Min(20, usedCols)
    For r = 1 To maxR
        Dim dateCount As Long: dateCount = 0
        For c = 3 To maxC
            If IsDate(data(r, c)) Then dateCount = dateCount + 1
        Next c
        If dateCount >= 5 Then
            FindDateHeaderRow = r
            Exit Function
        End If
    Next r
    FindDateHeaderRow = 0
End Function
