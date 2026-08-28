Attribute VB_Name = "RemovePODraftStockColumns"
Option Explicit

' ============================================================================
' RemovePODraftStockColumns module - ONE-TIME MIGRATION
'
' [Why] PO_Draft_* sheets are exported (PO_Export.bas) and sent to the
' supplier (TTAF). Their SafetyStock/CurrentStock columns (F/G) exposed our
' own stock levels to the supplier, which we don't want to disclose.
' Excluding them from the print area (the previous approach) wasn't
' enough: ExportPODraft duplicates the whole sheet, formulas and all, so
' the columns were still present - just not printed - in every file
' actually sent out. This macro deletes columns F:G outright from all 4
' PO_Draft_* sheets, matching the new build_soh.py-generated layout (week
' 1 now starts at column F instead of H; the Order Date/Issue Month/Firm
' Month/Revision/Base Week letterhead value cells shift from column P to
' column N, and their labels shift from column N to column L).
'
' Deleting entire columns (rather than trying to rewrite formulas one by
' one) lets Excel's native column-delete automatically re-point every
' other formula on every sheet (including the row 25 week-label formulas,
' which reference the base-week cell directly as $P$13 rather than by the
' "BaseWeek" name), every named range (BaseWeek/PORevision), and every
' merged cell - nothing else needs to be touched by hand. This is the same
' kind of automatic reference update Excel already performs on sheet
' rename (see RenameSelfStockToCSAstocks.bas) or Table row insert/delete
' (see the comment at the top of RefreshData_MaterialMgmt.bas) - here it's
' just column-wise instead of sheet- or row-wise.
'
' [Caution] Back up first and test on a copy - not the live production
' file directly. After running, spot-check a few PO_Draft_* rows (order
' quantities should be completely unchanged - only the SafetyStock/
' CurrentStock columns are gone, and the Order Date/Issue Month/Firm
' Month/Revision/Base Week letterhead cells moved 2 columns to the left)
' before saving, then re-import the updated RefreshData_PODraft.bas /
' RefreshData_MaterialMgmt.bas modules (which now assume the new column
' layout - importing them BEFORE running this macro would make
' AddMaterial/SyncPODraftCategories write new rows in the wrong columns
' until this migration has run). Delete this module once confirmed - it's
' one-time-use.
'
' Safe to run more than once: a sheet whose column F, row 25 header is no
' longer exactly "SafetyStock" is treated as already migrated and skipped.
' ============================================================================

Sub RemovePODraftStockColumns()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook

    Const HDR_ROW As Long = 25  ' row where the "SafetyStock"/"CurrentStock" column headers live

    Dim sheetNames As Variant
    sheetNames = Array("PO_Draft_Chemical", "PO_Draft_Hazardous", _
        "PO_Draft_Substrate_JPN_CHN", "PO_Draft_Substrate_Poland")

    Dim resultMsg As String: resultMsg = ""
    Dim changedCount As Long: changedCount = 0
    Dim si As Long
    For si = LBound(sheetNames) To UBound(sheetNames)
        Dim shName As String: shName = CStr(sheetNames(si))
        Dim sh As Worksheet
        On Error Resume Next
        Set sh = Nothing
        Set sh = thisWb.Sheets(shName)
        On Error GoTo ErrHandler

        If sh Is Nothing Then
            resultMsg = resultMsg & "- " & shName & ": sheet not found (skipped)" & vbCrLf
        ElseIf Trim(CStr(sh.Cells(HDR_ROW, 6).Value)) <> "SafetyStock" Then
            resultMsg = resultMsg & "- " & shName & ": already migrated (skipped)" & vbCrLf
        Else
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            sh.Columns("F:G").Delete
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            resultMsg = resultMsg & "- " & shName & ": columns F:G (SafetyStock/CurrentStock) deleted" & vbCrLf
            changedCount = changedCount + 1
        End If
    Next si

    MsgBox "PO_Draft SafetyStock/CurrentStock column removal complete." & vbCrLf & vbCrLf & _
           resultMsg & vbCrLf & _
           "Sheets changed this run: " & changedCount & vbCrLf & vbCrLf & _
           "Please spot-check a few PO_Draft_* rows (order quantities should be unchanged, and the " & _
           "Order Date/Issue Month/Firm Month/Revision/Base Week letterhead cells should now sit 2 " & _
           "columns to the left of before) before saving. Import the updated RefreshData_PODraft.bas / " & _
           "RefreshData_MaterialMgmt.bas modules only AFTER confirming this migration looks correct.", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred during the migration: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "Some sheets may have already been changed. Do NOT save - close the file without " & _
           "saving, restore from your backup, and report this exact error.", vbCritical
End Sub
