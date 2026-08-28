Attribute VB_Name = "RemoveStockCountFeature"
Option Explicit

' ============================================================================
' RemoveStockCountFeature module - ONE-TIME MIGRATION
'
' [Why] T_StockCount was a manual physical-count override sheet, given top
' priority in Grid_Stock's formula (manual count > self+TTAF actuals sum >
' roll-forward). The weekly "Raw materials daily check" (which already
' feeds T_CSAstocks via RefreshSelfStock) turned out to already BE the
' physical stock count, making T_StockCount's role fully redundant - and
' on the live file it was confirmed to have never actually been used (the
' sample/placeholder row was still sitting there untouched). This macro:
'   1. Rewrites every Grid_Stock cell from the old 3-tier priority formula
'      to the new 2-tier form (self+TTAF sum > roll-forward), matching
'      exactly what the current build_soh.py/AddMaterial now generate.
'   2. Deletes the T_StockCount sheet.
' Nothing else is touched - Grid_TheoreticalStock never referenced
' T_StockCount in the first place, and no other sheet reads from it.
'
' [Caution] Back up first and test on a copy - not the live production
' file directly. No dependency on any other module. After running, read
' the confirmation MsgBox carefully and spot-check that a few Grid_Stock/
' Dashboard values are unchanged (the numbers should be identical - only
' the formula got simpler) before saving. Delete this module once
' confirmed - it's one-time-use.
'
' Safe to run more than once: an already-simplified formula and an
' already-removed sheet are both simply skipped (counted, not re-applied).
' ============================================================================

Sub RemoveStockCountFeature()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim fixedCells As Long: fixedCells = 0
    Dim alreadySimplified As Long: alreadySimplified = 0
    Dim unrecognized As Long: unrecognized = 0
    Call SimplifyGridStockFormulas(thisWb, fixedCells, alreadySimplified, unrecognized)

    Dim sheetDeleted As Boolean: sheetDeleted = False
    Dim scSheet As Worksheet
    On Error Resume Next
    Set scSheet = thisWb.Sheets("T_StockCount")
    On Error GoTo ErrHandler
    If Not scSheet Is Nothing Then
        Application.DisplayAlerts = False
        scSheet.Delete
        Application.DisplayAlerts = True
        sheetDeleted = True
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "T_StockCount removal complete." & vbCrLf & vbCrLf & _
          "Grid_Stock cells simplified this run: " & fixedCells & vbCrLf & _
          "Grid_Stock cells already simplified (skipped): " & alreadySimplified & vbCrLf & _
          "T_StockCount sheet: " & IIf(sheetDeleted, "deleted", "not found (already removed)")
    If unrecognized > 0 Then
        msg = msg & vbCrLf & vbCrLf & "WARNING: " & unrecognized & " cell(s) referenced T_StockCount but " & _
              "in a form that didn't match the expected pattern, and were left untouched. Please check " & _
              "Grid_Stock by hand for any remaining T_StockCount references before relying on this file."
    End If
    MsgBox msg, vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "An error occurred during migration: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "Some changes may have already been applied. Do NOT save - close the file without " & _
           "saving, restore from your backup, and report this exact error message.", vbCritical
End Sub

' Rewrites every data cell of Grid_Stock from the old 3-tier formula
' (manual count via T_StockCount > self+TTAF sum > roll-forward) to the new
' 2-tier form (self+TTAF sum > roll-forward). Works by stripping the exact,
' known "IF(COUNTIFS(T_StockCount...)>0,SUMIFS(T_StockCount...)," prefix
' and its matching outer closing parenthesis - this is not a general
' formula rewrite, so it can't misfire on a formula that doesn't exactly
' match the pattern this project's own code generates. A cell whose
' formula doesn't match is left completely untouched and counted in
' unrecognizedCount instead of being guessed at.
Private Sub SimplifyGridStockFormulas(wb As Workbook, ByRef fixedCount As Long, ByRef skippedCount As Long, ByRef unrecognizedCount As Long)
    fixedCount = 0
    skippedCount = 0
    unrecognizedCount = 0
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets("Grid_Stock")
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = sh.ListObjects("Grid_Stock")
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Dim nRows As Long: nRows = tbl.DataBodyRange.Rows.Count
    Dim nCols As Long: nCols = tbl.DataBodyRange.Columns.Count  ' column 1=Part Name, column 2=Week1, ...
    Dim firstRow As Long: firstRow = tbl.DataBodyRange.Row

    Dim r As Long, c As Long
    For r = 1 To nRows
        Dim actualRow As Long: actualRow = firstRow + r - 1
        For c = 2 To nCols
            Dim w As Long: w = c - 1
            Dim cell As Range: Set cell = sh.Cells(actualRow, c)
            Dim f As String: f = cell.Formula
            If Len(f) = 0 Or Left(f, 1) <> "=" Then GoTo NextCell

            Dim prefix As String
            prefix = "=IF(COUNTIFS(T_StockCount[Part Name],$A" & actualRow & ",T_StockCount[WeekIndex]," & w & ")>0," & _
                     "SUMIFS(T_StockCount[CountedQty],T_StockCount[Part Name],$A" & actualRow & ",T_StockCount[WeekIndex]," & w & "),"

            If Left(f, Len(prefix)) = prefix Then
                Dim innerExpr As String: innerExpr = Mid(f, Len(prefix) + 1)
                If Right(innerExpr, 1) = ")" Then
                    innerExpr = Left(innerExpr, Len(innerExpr) - 1)
                    cell.Formula = "=" & innerExpr
                    fixedCount = fixedCount + 1
                Else
                    unrecognizedCount = unrecognizedCount + 1
                End If
            ElseIf InStr(f, "T_StockCount") = 0 Then
                skippedCount = skippedCount + 1  ' already simplified, or a material added after this migration
            Else
                unrecognizedCount = unrecognizedCount + 1  ' mentions T_StockCount but doesn't match the expected exact pattern
            End If
NextCell:
        Next c
    Next r
End Sub
