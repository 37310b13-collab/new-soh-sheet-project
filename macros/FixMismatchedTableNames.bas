Attribute VB_Name = "FixMismatchedTableNames"
Option Explicit

' ============================================================================
' FixMismatchedTableNames module - ONE-TIME DIAGNOSTIC/REPAIR
'
' [Why] RenameProductionAndStockSheets renames both a sheet and its Table
' (ListObject) to the same new name, table first, sheet second. If the
' table-rename step ever failed to find the table under its expected old
' name (or found more than one Table on the sheet), the SHEET still got
' renamed - so a later run's "already migrated" check (which only looks at
' the sheet name) would skip it, silently leaving the Table under its old
' name forever. Since most formulas reference these sheets by plain
' 'SheetName'!CellRef (not by Table name), this mismatch causes no visible
' formula errors anywhere - it only surfaces as runtime error 9
' ("Subscript out of range") in VBA code that looks up the Table by name,
' e.g. FixTheoreticalStockMonthlyReset's or AddMaterial's
' Sheets("TheoreticalStock").ListObjects("TheoreticalStock").
'
' [What it does] For each of the 5 Table-based grid sheets and the 2 hidden
' Log tables, checks whether that sheet has a Table already named to match
' (e.g. sheet "TheoreticalStock" should have a Table named
' "TheoreticalStock"). If not, and the sheet has exactly one Table, renames
' that Table to match (renaming a Table automatically updates every
' structured-reference formula elsewhere in the workbook that used its old
' name, the same as any other Table rename in this project). If a sheet
' has zero or more than one Table with no match, it's reported as a
' warning instead of guessed at.
'
' [Caution] Back up first and test on a copy - not the live production
' file directly, though this only touches Table name metadata (no cell
' values/formulas are written). Delete this module once confirmed - it's
' one-time-use.
'
' Safe to run more than once: anything already matching is simply skipped.
' ============================================================================

Sub FixMismatchedTableNames()
    On Error GoTo ErrHandler
    Dim wb As Workbook: Set wb = ThisWorkbook

    Dim names As Variant
    names = Array("Production_Plan", "WeeklyConsumption", "Incoming", "Stock", _
                  "TheoreticalStock", "CSAstock_Log", "TTAFstock_Log")

    Dim report As String: report = ""
    Dim nm As Variant
    For Each nm In names
        Dim sh As Worksheet
        On Error Resume Next
        Set sh = Nothing
        Set sh = wb.Sheets(CStr(nm))
        On Error GoTo ErrHandler

        If sh Is Nothing Then
            report = report & "- " & nm & ": sheet not found (skipped)" & vbCrLf
        Else
            Dim lo As ListObject
            On Error Resume Next
            Set lo = Nothing
            Set lo = sh.ListObjects(CStr(nm))
            On Error GoTo ErrHandler

            If Not lo Is Nothing Then
                report = report & "- " & nm & ": table name already correct" & vbCrLf
            ElseIf sh.ListObjects.Count = 1 Then
                Dim oldTblName As String: oldTblName = sh.ListObjects(1).Name
                sh.ListObjects(1).Name = CStr(nm)
                report = report & "- " & nm & ": table was named """ & oldTblName & """ - renamed to match" & vbCrLf
            Else
                report = report & "- " & nm & ": WARNING - found " & sh.ListObjects.Count & _
                    " table(s) on this sheet, none named """ & nm & """ - please check and rename by hand " & _
                    "(select a cell in the table, Table Design tab, Table Name box)." & vbCrLf
            End If
        End If
    Next nm

    MsgBox "Table name check/fix complete." & vbCrLf & vbCrLf & report & vbCrLf & _
           "If any rows say ""renamed to match"", please save the file, then try " & _
           "FixTheoreticalStockMonthlyReset (or whichever macro errored) again.", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    MsgBox "An error occurred while checking: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "Progress so far:" & vbCrLf & report & vbCrLf & _
           "Please report this exact error.", vbCritical
End Sub
