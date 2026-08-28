Attribute VB_Name = "RenameSelfStockToCSAstocks"
Option Explicit

' ============================================================================
' RenameSelfStockToCSAstocks module - ONE-TIME MIGRATION
'
' [Why] The visible "T_SelfStock" grid sheet and its hidden raw log
' "T_SelfStock_Log" are being renamed to "T_CSAstocks" / "T_CSAstocks_Log"
' to match the company's own naming convention (CSA = the company itself,
' as distinct from TTAF the supplier/warehouse). Renaming an Excel sheet
' automatically rewrites every formula elsewhere in the workbook that
' references it by name (e.g. 'T_SelfStock'!B6 becomes 'T_CSAstocks'!B6
' on its own) - so this migration is just the two sheet renames below.
' Nothing else needs to change: no formula in Grid_Stock, Dashboard,
' Material_Detail, or anywhere else needs to be touched by hand.
'
' [Caution] Back up first and test on a copy - not the live production
' file directly. After running, spot-check a few Grid_Stock/Dashboard
' values (the numbers should be completely unchanged - only the sheet
' names changed) before saving, then re-import the updated
' RefreshData_MaterialMgmt.bas/RefreshData_StockActuals.bas/
' RefreshData_Display.bas/RefreshData_Utilities.bas modules (which now
' expect "T_CSAstocks"/"T_CSAstocks_Log" as the sheet names). Delete this
' module once confirmed - it's one-time-use.
'
' Safe to run more than once: a sheet that's already been renamed (or was
' never found under the old name) is simply skipped.
' ============================================================================

Sub RenameSelfStockToCSAstocks()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook

    Dim gridRenamed As String: gridRenamed = "already T_CSAstocks / not found"
    Dim logRenamed As String: logRenamed = "already T_CSAstocks_Log / not found"

    Dim oldGrid As Worksheet
    On Error Resume Next
    Set oldGrid = thisWb.Sheets("T_SelfStock")
    On Error GoTo ErrHandler
    If Not oldGrid Is Nothing Then
        oldGrid.Name = "T_CSAstocks"
        gridRenamed = "renamed this run"
    End If

    Dim oldLog As Worksheet
    On Error Resume Next
    Set oldLog = thisWb.Sheets("T_SelfStock_Log")
    On Error GoTo ErrHandler
    If Not oldLog Is Nothing Then
        oldLog.Name = "T_CSAstocks_Log"
        logRenamed = "renamed this run"
    End If

    MsgBox "Sheet rename complete." & vbCrLf & vbCrLf & _
           "T_SelfStock -> T_CSAstocks: " & gridRenamed & vbCrLf & _
           "T_SelfStock_Log -> T_CSAstocks_Log: " & logRenamed & vbCrLf & vbCrLf & _
           "Please spot-check a few Grid_Stock/Dashboard values (numbers should be " & _
           "unchanged), then save the file BEFORE importing the updated VBA modules.", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    MsgBox "An error occurred during the rename: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "If a sheet named ""T_CSAstocks"" or ""T_CSAstocks_Log"" already exists " & _
           "(name collision), that is the most likely cause. Do NOT save - close the " & _
           "file without saving, restore from your backup, and report this exact error.", vbCritical
End Sub
