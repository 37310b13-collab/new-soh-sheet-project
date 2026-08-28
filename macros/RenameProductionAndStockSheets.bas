Attribute VB_Name = "RenameProductionAndStockSheets"
Option Explicit

' ============================================================================
' RenameProductionAndStockSheets module - ONE-TIME MIGRATION
'
' [Why] Several sheet names carried internal jargon (PP = Production Plan;
' the "Grid_" prefix; the "T_" prefix on the actual-stock grids) that made
' the sheet list harder to read at a glance. This macro renames:
'   PP_Grid                -> Production_Plan
'   Grid_Requirement        -> WeeklyConsumption
'   Grid_Incoming           -> Incoming
'   Grid_Stock              -> Stock
'   Grid_TheoreticalStock   -> TheoreticalStock
'   T_SelfStock             -> CSAstock   (CSA = our own company, as
'                                          opposed to TTAF the supplier/
'                                          warehouse)
'   T_SelfStock_Log         -> CSAstock_Log
'   T_TTAFStock             -> TTAFstock
'   T_TTAFStock_Log         -> TTAFstock_Log
'
' [Important - Tables, not just sheets] Production_Plan/WeeklyConsumption/
' Incoming/Stock/TheoreticalStock/CSAstock_Log/TTAFstock_Log are Excel
' Tables whose Table name (ListObject.Name) has always been set equal to
' the sheet name. Renaming a WORKSHEET does NOT automatically rename its
' Table - the two are independent objects, and every formula elsewhere in
' the workbook that uses a structured reference (e.g.
' "Production_Plan[#Data]", "CSAstock_Log[Part Name]") follows the TABLE
' name, not the sheet name. So for these 7 sheets, this macro renames the
' Table first (which is what makes every structured-reference formula in
' the workbook follow automatically - the same automatic update Excel
' already does for 'SheetName'!A1-style references on sheet rename), then
' renames the sheet to match. CSAstock/TTAFstock (the visible actual-stock
' grids) are NOT Tables (plain bordered grids, built by hand rather than
' via Excel's Table feature), so only the sheet itself is renamed for
' those two - any formula referencing them (e.g. 'CSAstock'!B6) is a
' sheet-qualified cell reference and updates automatically on sheet
' rename, the same as the T_SelfStock->T_CSAstocks precedent already used
' elsewhere in this project's history.
'
' Each item tries a short list of possible current names in order (in case
' an earlier partial migration already renamed some of them) and uses
' whichever is actually found; an item whose target name already exists is
' treated as already migrated and skipped.
'
' [Caution] Back up first and test on a copy - not the live production
' file directly. VBA code itself is NOT rewritten by this rename (only
' formulas/Table structured-references are). [Run order] After this macro,
' run AddDescriptionColumns.bas next (it looks for the sheets under their
' NEW names) - only THEN re-import the updated RefreshData_*.bas modules
' (VBAProject -> remove the old modules -> File -> Import File for each
' updated .bas); they already assume both this rename AND the Description
' column migration are done (e.g. AddMaterial writes into the columns that
' exist only after AddDescriptionColumns has run). Importing them too early
' would make AddMaterial/RefreshBOM/etc. write into the wrong sheets/
' columns or fail to find them. After running this macro, spot-check a few
' Stock/TheoreticalStock/Dashboard/PO_Draft_* values (the numbers should be
' completely unchanged - only names changed) before saving. Delete this
' module once confirmed - it's one-time-use.
'
' Safe to run more than once: anything already renamed is simply skipped.
' ============================================================================

Private renameLog As String

Sub RenameProductionAndStockSheets()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    renameLog = ""

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call RenameSheetAndTable(thisWb, Array("PP_Grid"), "Production_Plan", True)
    Call RenameSheetAndTable(thisWb, Array("Grid_Requirement"), "WeeklyConsumption", True)
    Call RenameSheetAndTable(thisWb, Array("Grid_Incoming"), "Incoming", True)
    Call RenameSheetAndTable(thisWb, Array("Grid_Stock"), "Stock", True)
    Call RenameSheetAndTable(thisWb, Array("Grid_TheoreticalStock"), "TheoreticalStock", True)
    Call RenameSheetAndTable(thisWb, Array("T_SelfStock_Log", "T_CSAstocks_Log"), "CSAstock_Log", True)
    Call RenameSheetAndTable(thisWb, Array("T_TTAFStock_Log"), "TTAFstock_Log", True)
    Call RenameSheetAndTable(thisWb, Array("T_SelfStock", "T_CSAstocks"), "CSAstock", False)
    Call RenameSheetAndTable(thisWb, Array("T_TTAFStock"), "TTAFstock", False)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Sheet/Table rename complete." & vbCrLf & vbCrLf & renameLog & vbCrLf & _
           "Please spot-check a few Stock/TheoreticalStock/Dashboard/PO_Draft_* values (numbers " & _
           "should be unchanged), then save. Next, run AddDescriptionColumns, and only after that " & _
           "import the updated VBA modules.", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred during the rename: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "Progress so far:" & vbCrLf & renameLog & vbCrLf & _
           "If a sheet/table with the new name already exists (name collision), that is the most " & _
           "likely cause. Do NOT save - close the file without saving, restore from your backup, " & _
           "and report this exact error.", vbCritical
End Sub

' Renames whichever of candidateOldNames is actually found (as a worksheet)
' to newName - renaming its Table first (if hasTable), so every structured-
' reference formula follows, then the sheet itself. If newName already
' exists as a sheet, or none of candidateOldNames is found, does nothing
' (already migrated / nothing to do).
Private Sub RenameSheetAndTable(wb As Workbook, candidateOldNames As Variant, newName As String, hasTable As Boolean)
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets(newName)
    On Error GoTo 0
    If Not sh Is Nothing Then
        renameLog = renameLog & "- " & newName & ": already migrated (skipped)" & vbCrLf
        Exit Sub
    End If

    Dim oldName As Variant
    For Each oldName In candidateOldNames
        Set sh = Nothing
        On Error Resume Next
        Set sh = wb.Sheets(CStr(oldName))
        On Error GoTo 0
        If Not sh Is Nothing Then Exit For
    Next oldName

    If sh Is Nothing Then
        renameLog = renameLog & "- " & newName & ": no matching sheet found (skipped)" & vbCrLf
        Exit Sub
    End If

    Dim actualOldName As String: actualOldName = sh.Name

    If hasTable Then
        Dim lo As ListObject
        On Error Resume Next
        Set lo = sh.ListObjects(actualOldName)
        On Error GoTo 0
        If lo Is Nothing And sh.ListObjects.Count = 1 Then
            Set lo = sh.ListObjects(1)
        End If
        If Not lo Is Nothing Then
            lo.Name = newName
        Else
            renameLog = renameLog & "  (WARNING: expected exactly one Table on """ & actualOldName & _
                """ but found " & sh.ListObjects.Count & " - the Table was NOT renamed, only the sheet " & _
                "was. Please rename the Table by hand in Excel: select any cell in it, go to the " & _
                "Table Design tab, and set Table Name to """ & newName & """.)" & vbCrLf
        End If
    End If

    sh.Name = newName
    renameLog = renameLog & "- " & actualOldName & " -> " & newName & vbCrLf
End Sub
