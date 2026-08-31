Attribute VB_Name = "AddDescriptionColumns"
Option Explicit

' ============================================================================
' AddDescriptionColumns module - ONE-TIME MIGRATION
'
' [Why] WeeklyConsumption/Incoming/Stock/TheoreticalStock/CSAstock/TTAFstock
' only ever showed the raw material CODE (Part Name) in column A - with no
' material name visible, it wasn't possible to tell at a glance what a code
' actually refers to. This macro inserts a "Description" column right after
' Part Name on those 6 sheets, showing each row's material name via a
' lookup into M_RawMaterials. Production_Plan intentionally does NOT get a
' Description column (it lists intermediates/finished-product codes, not
' raw materials - M_RawMaterials[Description] wouldn't match anything
' there).
'
' [Run order - important] Run this AFTER RenameProductionAndStockSheets.bas
' (this macro looks for the sheets under their NEW names -
' WeeklyConsumption/Incoming/Stock/TheoreticalStock/CSAstock/TTAFstock - and
' does nothing if a sheet isn't found under that name), and BEFORE
' re-importing the updated RefreshData_*.bas modules (they already assume
' this Description column exists - e.g. AddMaterial writes week-data
' formulas starting one column further right than before. Importing them
' first would make AddMaterial/etc. write into the wrong columns until this
' migration has actually run).
'
' [What it does] For each of the 6 sheets, inserts one column at position B
' (Excel's native column insert, so every other formula/conditional-
' formatting rule elsewhere in the workbook that references a cell in or
' after column B on that sheet is automatically re-pointed one column to
' the right - the same kind of automatic reference update this project has
' already relied on for sheet renames and Table renames), sets its header
' to "Description", and fills every data row with a lookup formula against
' M_RawMaterials[Description]. For the 4 Table-based sheets
' (WeeklyConsumption/Incoming/Stock/TheoreticalStock), the column is added
' via ListColumns.Add so the Table's own metadata stays correct. CSAstock/
' TTAFstock are plain bordered grids (not Tables), so a direct worksheet
' column insert is used instead - this also naturally shifts the small
' "enter week to jump to" panel on row 1 of those two sheets from
' C1/D1/E1/F1 to D1/E1/F1/G1, matching what build_soh.py now generates
' fresh (see RefreshData_Display.bas's JumpToSelectedWeek and
' docs/SOH_System_Guide.md section 5.6 for the updated Worksheet_Change
' wiring you'll need if you already set that up under the old C1/F1
' addresses).
'
' [Caution] Back up first and test on a copy - not the live production
' file directly. After running, spot-check: (1) a few Stock/
' TheoreticalStock/Dashboard/Material_Detail values are unchanged, (2) the
' Description column actually shows the right material name, (3) on
' CSAstock/TTAFstock, the thick-border highlight on the selected week
' still lands on the correct week column after entering a week number in
' D1, and the frozen columns still stop right after the Description
' column. Delete this module once confirmed - it's one-time-use.
'
' Safe to run more than once: the column-insert step is skipped once
' column B's header already says "Description" (inserting it twice would
' create a second, wrong Description column) - but the row-formula and
' freeze-pane steps always run regardless, since they're cheap and
' idempotent. This matters because those steps happen after the header is
' set, so a run that errored out partway through (e.g. only some rows got
' the formula) would otherwise leave the header saying "Description" while
' silently understating how much of the sheet is actually done, and a
' later run would skip it entirely without ever finishing the job.
' ============================================================================

Private resultLog As String

Sub AddDescriptionColumns()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    resultLog = ""

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call AddDescriptionToTableSheet(thisWb, "WeeklyConsumption")
    Call AddDescriptionToTableSheet(thisWb, "Incoming")
    Call AddDescriptionToTableSheet(thisWb, "Stock")
    Call AddDescriptionToTableSheet(thisWb, "TheoreticalStock")
    Call AddDescriptionToGridSheet(thisWb, "CSAstock")
    Call AddDescriptionToGridSheet(thisWb, "TTAFstock")

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Description column migration complete." & vbCrLf & vbCrLf & resultLog & vbCrLf & _
           "Please spot-check a few values (numbers should be unchanged - only a Description " & _
           "column was inserted), and on CSAstock/TTAFstock check that entering a week number in " & _
           "D1 still highlights/scrolls to the correct week, before saving. Re-import the updated " & _
           "RefreshData_*.bas modules only after confirming this looks correct.", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred during the migration: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "Progress so far:" & vbCrLf & resultLog & vbCrLf & _
           "Do NOT save - close the file without saving, restore from your backup, and report this " & _
           "exact error.", vbCritical
End Sub

' Inserts a Description column (position 2) into a Table-based sheet
' (WeeklyConsumption/Incoming/Stock/TheoreticalStock). Skips if the sheet
' isn't found, its Table isn't found, or column B's header is already
' "Description".
Private Sub AddDescriptionToTableSheet(wb As Workbook, sheetName As String)
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then
        resultLog = resultLog & "- " & sheetName & ": sheet not found (skipped)" & vbCrLf
        Exit Sub
    End If
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = sh.ListObjects(sheetName)
    On Error GoTo 0
    If tbl Is Nothing And sh.ListObjects.Count = 1 Then
        ' The Table isn't named to match the sheet - e.g. a leftover from
        ' RenameProductionAndStockSheets not finding it under its expected
        ' old name (see FixMismatchedTableNames.bas). Since this sheet only
        ' has one Table, it must be the right one regardless of its name.
        Set tbl = sh.ListObjects(1)
        resultLog = resultLog & "  (note: this sheet's Table was named """ & tbl.Name & _
            """, not """ & sheetName & """ - proceeding anyway since it's the only Table here. " & _
            "Consider running FixMismatchedTableNames.bas too.)" & vbCrLf
    End If
    If tbl Is Nothing Then
        resultLog = resultLog & "- " & sheetName & ": Table not found (skipped - please check by hand)" & vbCrLf
        Exit Sub
    End If
    Dim alreadyHadColumn As Boolean
    alreadyHadColumn = (Trim(CStr(sh.Cells(1, 2).Value)) = "Description")
    If Not alreadyHadColumn Then
        Dim newCol As ListColumn: Set newCol = tbl.ListColumns.Add(Position:=2)
        newCol.Name = "Description"
        sh.Columns(2).ColumnWidth = 32
    End If

    ' Always (re)populate every row's formula and reset the freeze pane,
    ' even if the column already existed - see the module header comment
    ' for why this must not be skipped just because the column was already
    ' there (a previous run could have been interrupted after inserting
    ' the column but before finishing these steps).
    Dim n As Long: n = tbl.ListRows.Count
    Dim r As Long, rr As Long
    For r = 1 To n
        rr = tbl.ListRows(r).Range.Row
        tbl.ListRows(r).Range.Cells(1, 2).Value = _
            "=IFERROR(INDEX(M_RawMaterials[Description],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),"""")"
    Next r

    ' Re-set the freeze pane to split right after the Description column
    ' (matches build_soh.py's "C2" for freshly generated files). A column
    ' insert normally shifts an existing split point automatically, but
    ' this is set explicitly here rather than relied upon, to be safe.
    sh.Activate
    ActiveWindow.FreezePanes = False
    sh.Range("C2").Select
    ActiveWindow.FreezePanes = True

    If alreadyHadColumn Then
        resultLog = resultLog & "- " & sheetName & ": column already existed - re-verified/refilled all " & n & " rows" & vbCrLf
    Else
        resultLog = resultLog & "- " & sheetName & ": Description column added (" & n & " rows)" & vbCrLf
    End If
End Sub

' Inserts a Description column (position 2) into a plain-grid sheet
' (CSAstock/TTAFstock - not a Table). Skips if the sheet isn't found or
' column B's header (row SS_TABLE_ROW) is already "Description".
Private Sub AddDescriptionToGridSheet(wb As Workbook, sheetName As String)
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then
        resultLog = resultLog & "- " & sheetName & ": sheet not found (skipped)" & vbCrLf
        Exit Sub
    End If
    Dim alreadyHadColumn As Boolean
    alreadyHadColumn = (Trim(CStr(sh.Cells(SS_TABLE_ROW, 2).Value)) = "Description")
    If Not alreadyHadColumn Then
        sh.Columns(2).Insert Shift:=xlToRight
        sh.Cells(SS_TABLE_ROW, 2).Value = "Description"
        sh.Columns(2).ColumnWidth = 32
    End If

    ' Always (re)populate every row's formula and reset the freeze pane,
    ' even if the column already existed - same reasoning as
    ' AddDescriptionToTableSheet.
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, n As Long: n = 0
    For r = SS_TABLE_ROW + 1 To lastRow
        If Len(Trim(CStr(sh.Cells(r, 1).Value))) > 0 Then
            sh.Cells(r, 2).Value = _
                "=IFERROR(INDEX(M_RawMaterials[Description],MATCH($A" & r & ",M_RawMaterials[Part Name],0)),"""")"
            n = n + 1
        End If
    Next r

    ' Re-set the freeze pane to split right after the Description column
    ' (matches build_soh.py's "C{data start row}" for freshly generated
    ' files). See the comment in AddDescriptionToTableSheet for why this is
    ' set explicitly rather than relied upon.
    sh.Activate
    ActiveWindow.FreezePanes = False
    sh.Range("C" & (SS_TABLE_ROW + 1)).Select
    ActiveWindow.FreezePanes = True

    If alreadyHadColumn Then
        resultLog = resultLog & "- " & sheetName & ": column already existed - re-verified/refilled all " & n & " rows" & vbCrLf
    Else
        resultLog = resultLog & "- " & sheetName & ": Description column added (" & n & " rows)" & vbCrLf
    End If
End Sub
