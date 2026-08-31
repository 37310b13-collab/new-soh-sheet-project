Attribute VB_Name = "RepairStockAfterTStockCountDeletion"
Option Explicit

' ============================================================================
' RepairStockAfterTStockCountDeletion module - ONE-TIME EMERGENCY REPAIR
'
' [Why] The live file's Stock sheet formulas still had the OLD 3-tier form
' referencing T_StockCount (=IF(COUNTIFS(T_StockCount[...])>0,SUMIFS(...),
' <normal 2-tier formula>)) - meaning RemoveStockCountFeature.bas's formula
' simplification had never actually taken effect on this file, even though
' the sheet itself was believed to already be decommissioned. Deleting the
' T_StockCount sheet directly (as advised, based on that mistaken belief)
' turned every T_StockCount[...] structured reference into #REF!, and since
' those #REF! errors sit inside the OUTER IF's condition/branches, the
' error propagates through the whole formula - every cell in Stock now
' shows #REF! instead of a number.
'
' [What it does] Rather than trying to parse or repair the broken #REF!
' text (fragile, and this file's actual prior formula turned out not to
' exactly match what RemoveStockCountFeature.bas assumed), this macro
' REBUILDS every data cell in Stock from scratch, using the exact same
' formula construction as AddMaterial (RefreshData_MaterialMgmt.bas) and
' build_soh.py: "self+TTAF actuals (if both present) else previous week +
' incoming - consumption". For each Stock row, the matching CSAstock/
' TTAFstock row is found by matching Part Name (column A) - not assumed
' to be at a fixed offset - so this is robust even if row order ever
' differs between sheets.
'
' [Caution] Back up first and test on a copy - not the live production
' file directly. This OVERWRITES every formula in Stock's data area
' (values are formulas, not stored numbers, so nothing is lost - the
' recalculated numbers should come out identical to what Stock showed
' before T_StockCount was deleted, since T_StockCount's manual-count
' priority tier was already supposed to be gone). After running, spot-
' check a few Stock/Dashboard values before saving. If any material's
' row can't be matched to a CSAstock/TTAFstock row, that row is skipped
' and reported - please check it by hand rather than guessing.
'
' Safe to run more than once - it recomputes the same correct formula
' every time regardless of the current (even if already-correct) content.
' ============================================================================

Sub RepairStockAfterTStockCountDeletion()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook

    Dim stTbl As ListObject: Set stTbl = thisWb.Sheets("Stock").ListObjects("Stock")
    Dim ws_st As Worksheet: Set ws_st = thisWb.Sheets("Stock")
    Dim ws_css As Worksheet: Set ws_css = thisWb.Sheets("CSAstock")
    Dim ws_ttf As Worksheet: Set ws_ttf = thisWb.Sheets("TTAFstock")

    Dim dataRange As Range: Set dataRange = stTbl.DataBodyRange
    If dataRange Is Nothing Then
        MsgBox "Stock has no data rows.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Build Part Name -> row index for CSAstock/TTAFstock once (both should
    ' have the exact same set of materials in the exact same order, but
    ' matched independently by name rather than assumed).
    Dim cssIdx As Object: Set cssIdx = BuildNameIndexByColumn(ws_css, 1, 6)
    Dim ttfIdx As Object: Set ttfIdx = BuildNameIndexByColumn(ws_ttf, 1, 6)

    Dim nWeeks As Long: nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count

    Dim fixedRows As Long: fixedRows = 0
    Dim skippedRows As Long: skippedRows = 0
    Dim r As Long, w As Long
    For r = 1 To stTbl.ListRows.Count
        Dim grow As Long: grow = stTbl.ListRows(r).Range.Row
        Dim rmCode As String: rmCode = Trim(CStr(ws_st.Cells(grow, 1).Value))

        Dim ssRow As Variant
        If cssIdx.Exists(rmCode) And ttfIdx.Exists(rmCode) Then
            ssRow = cssIdx(rmCode)
            If ssRow <> ttfIdx(rmCode) Then
                ' CSAstock/TTAFstock disagree on this material's row - use CSAstock's,
                ' but flag it since AppendDashboardRow etc. expect them to match.
                ws_st.Cells(grow, 2).AddComment "WARNING: CSAstock row (" & ssRow & _
                    ") and TTAFstock row (" & ttfIdx(rmCode) & ") differ for this material - please check by hand."
            End If
        Else
            skippedRows = skippedRows + 1
            GoTo NextMaterial
        End If

        For w = 1 To nWeeks
            Dim col As Long: col = 2 + w
            Dim cl As String: cl = ColLetter(col)
            Dim hasSelf As String: hasSelf = "('CSAstock'!" & cl & ssRow & "<>"""")"
            Dim selfVal As String: selfVal = "'CSAstock'!" & cl & ssRow
            Dim hasTTAF As String: hasTTAF = "('TTAFstock'!" & cl & ssRow & "<>"""")"
            Dim ttafVal As String: ttafVal = "'TTAFstock'!" & cl & ssRow

            Dim priorExpr As String
            If w = 1 Then
                priorExpr = "IFERROR(INDEX(T_OpeningStock[OpeningQty],MATCH($A" & grow & ",T_OpeningStock[Part Name],0)),0)"
            Else
                priorExpr = ColLetter(col - 1) & grow
            End If
            Dim normalExpr As String
            normalExpr = priorExpr & "+'Incoming'!" & cl & grow & "-'WeeklyConsumption'!" & cl & grow

            ws_st.Cells(grow, col).Formula = _
                "=IF((" & hasSelf & ")*(" & hasTTAF & ")>0," & selfVal & "+" & ttafVal & "," & normalExpr & ")"
        Next w
        fixedRows = fixedRows + 1
NextMaterial:
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Stock formula repair complete." & vbCrLf & vbCrLf & _
           "Rows rebuilt: " & fixedRows & vbCrLf & _
           "Rows skipped (no matching CSAstock/TTAFstock row found by Part Name): " & skippedRows & vbCrLf & vbCrLf & _
           IIf(skippedRows > 0, "Please check the skipped rows by hand - their Stock formulas were left untouched.", _
               "Please spot-check a few Stock/Dashboard values, then save.") & vbCrLf & vbCrLf & _
           "This is a one-time repair - delete this module once you've confirmed Stock looks correct.", _
           vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred during the repair: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "Do NOT save - close the file without saving, restore from your backup, and report this " & _
           "exact error.", vbCritical
End Sub

' Builds a Part Name (column nameCol) -> row number index for a plain grid
' sheet (not a Table), scanning from startRow down to the last non-blank
' row in nameCol.
Private Function BuildNameIndexByColumn(sh As Worksheet, nameCol As Long, startRow As Long) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbTextCompare
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, nameCol).End(xlUp).Row
    Dim r As Long
    For r = startRow To lastRow
        Dim k As String: k = Trim(CStr(sh.Cells(r, nameCol).Value))
        If Len(k) > 0 And Not idx.Exists(k) Then idx(k) = r
    Next r
    Set BuildNameIndexByColumn = idx
End Function
