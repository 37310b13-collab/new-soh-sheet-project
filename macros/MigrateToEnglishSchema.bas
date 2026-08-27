Attribute VB_Name = "MigrateToEnglishSchema"
Option Explicit

' ============================================================================
' MigrateToEnglishSchema module - ONE-TIME MIGRATION
'
' Purpose: bring an existing, Japanese-labeled production workbook in line
' with the new all-English schema used by the current build_soh.py and the
' current RefreshData_*.bas / PO_Export.bas modules, WITHOUT losing any
' live data (T_Shipments, T_OpeningStock, T_StockCount, actuals logs,
' Material_Detail Order/PO_No entries, SafetyStock values, etc.).
'
' [How to use]
'   1. Back up the workbook first (copy the file). Test on a COPY before
'      touching the real production file.
'   2. Import this module (Alt+F11 -> File -> Import File). It depends on
'      the Public Const values in RefreshData_Utilities, so that module
'      must also be imported (any version - the constant values themselves
'      have not changed).
'   3. Run MigrateToEnglishSchema (Alt+F8) BEFORE importing/using any of
'      the other new RefreshData_*.bas modules or PO_Export.bas. Read the
'      summary MsgBox carefully and spot-check a few materials on
'      Dashboard/Material_Detail by eye.
'   4. Save the workbook.
'   5. Only then import the remaining new VBA modules and resume normal use.
'   6. Once you've confirmed everything works, delete this module - it is
'      one-time-use and should not be run again after the migration is
'      confirmed (running it again is harmless/no-op for already-migrated
'      parts, but it no longer serves any purpose).
'
' [What this migrates] (only what the new VBA code actually depends on
' matching - see each Sub below for exact scope):
'   1. Table column renames: M_RawMaterials / T_OpeningStock / T_StockCount /
'      T_Shipments. Renaming a Table column in Excel automatically updates
'      every structured-reference formula that uses it
'      (TableName[OldName] -> TableName[NewName]) throughout the whole
'      workbook - no formula rewriting needed, and nothing here touches
'      any data values.
'   2. Renames the "操作パネル" sheet to "Control_Panel" (if present under
'      the old name).
'   3. Rewrites Material_Detail's per-material row labels (each
'      intermediate's "Usage" row, Total Usage, TTAF Stock, Self Stock,
'      Order, PO_No, Total Stock) and the sheet's single "Item" header
'      cell, using ROW POSITION within each block rather than matching old
'      Japanese text - this can't silently do nothing even if the exact
'      old wording differs slightly from what's assumed here (see the
'      "No. of batches" caveat in MigrateMaterialDetailLabels).
'   4. Rewrites Dashboard's per-material "Row Type" column (Theoretical
'      Stock / Actual Stock) the same way (row position, not text
'      matching), and updates any conditional-formatting rule whose
'      formula compares against the old Japanese literal "実在庫".
'   5. Replaces the "[済]" completion marker with "[DONE]" wherever it
'      appears in Material_Detail's PO_No rows.
'
' [What this does NOT touch] (cosmetic text only - the new VBA never
' searches for or compares against any of these, so leaving them
' Japanese for now does not break anything): README sheet body text,
' Control_Panel's descriptive text/instructions and button-description
' column, PO_Draft_* letterhead placeholder text, T_StockCount/
' T_SelfStock_Log/T_TTAFStock_Log sample/dummy rows, the MOQ input-field
' comment text, PO_Draft_* header-row labels. These can be updated later
' with ordinary find & replace at your convenience.
'
' Safe to run more than once: every step is written to detect "already
' migrated" and skip (renaming a column that no longer has the old name
' is simply skipped; row-label rewrites always just overwrite with the
' correct current value regardless of what was there before).
' ============================================================================

Sub MigrateToEnglishSchema()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim logMsg As String: logMsg = ""

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ---- 1. Table column renames ----
    Dim colRenamed As Long: colRenamed = 0
    colRenamed = colRenamed + RenameTableColumn(thisWb, "M_RawMaterials", "基準在庫下限_要入力", "SafetyStockMin")
    colRenamed = colRenamed + RenameTableColumn(thisWb, "M_RawMaterials", "基準在庫上限_要入力", "SafetyStockMax")
    colRenamed = colRenamed + RenameTableColumn(thisWb, "M_RawMaterials", "固定週次消費量_要入力", "FixedWeeklyConsumption")
    colRenamed = colRenamed + RenameTableColumn(thisWb, "M_RawMaterials", "LeadTime_Weeks_要入力", "LeadTimeWeeks")
    colRenamed = colRenamed + RenameTableColumn(thisWb, "T_OpeningStock", "Opening_Qty_要入力", "OpeningQty")
    colRenamed = colRenamed + RenameTableColumn(thisWb, "T_StockCount", "Date_棚卸実施日", "CountDate")
    colRenamed = colRenamed + RenameTableColumn(thisWb, "T_Shipments", "Order_Date_発注日", "Order_Date")
    logMsg = logMsg & "Table columns renamed this run: " & colRenamed & vbCrLf

    ' ---- 2. Sheet rename ----
    Dim panelRenamed As String: panelRenamed = "already Control_Panel / not found"
    Dim oldPanel As Worksheet
    On Error Resume Next
    Set oldPanel = thisWb.Sheets("操作パネル")
    On Error GoTo ErrHandler
    If Not oldPanel Is Nothing Then
        oldPanel.Name = "Control_Panel"
        panelRenamed = "renamed this run"
    End If
    logMsg = logMsg & "Control_Panel sheet: " & panelRenamed & vbCrLf

    ' ---- 3 & 5. Material_Detail row labels + [DONE] marker ----
    Dim mdBlocks As Long: mdBlocks = 0
    Dim doneMarkers As Long: doneMarkers = 0
    Call MigrateMaterialDetailLabels(thisWb, mdBlocks, doneMarkers)
    logMsg = logMsg & "Material_Detail blocks relabeled: " & mdBlocks & vbCrLf
    logMsg = logMsg & "[済] markers converted to [DONE]: " & doneMarkers & vbCrLf

    ' ---- 4. Dashboard row types + conditional formatting ----
    ' Isolated from the outer error handler: this step touches conditional-
    ' formatting objects whose exact types on your live sheet we can't know
    ' in advance, so a failure here must not discard the (more important)
    ' work already done above. If it fails outright, dashPairs/cfFixed
    ' report 0 rather than aborting the whole migration.
    Dim dashPairs As Long: dashPairs = 0
    Dim cfFixed As Long: cfFixed = 0
    Dim dashErrMsg As String: dashErrMsg = ""
    On Error Resume Next
    Call MigrateDashboardRowTypes(thisWb, dashPairs, cfFixed)
    If Err.Number <> 0 Then
        dashErrMsg = "(" & Err.Number & ") " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrHandler
    logMsg = logMsg & "Dashboard material row-pairs relabeled: " & dashPairs & vbCrLf
    logMsg = logMsg & "Conditional-format rules updated (実在庫->Actual Stock): " & cfFixed & vbCrLf
    If Len(dashErrMsg) > 0 Then
        logMsg = logMsg & "(Note) Dashboard step hit an error and may be incomplete: " & dashErrMsg & vbCrLf
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If mdBlocks = 0 Or dashPairs = 0 Then
        MsgBox "WARNING: Material_Detail blocks relabeled = " & mdBlocks & _
               ", Dashboard row-pairs relabeled = " & dashPairs & "." & vbCrLf & _
               "If your sheets actually contain materials, a count of 0 here means the block-" & vbCrLf & _
               "detection logic did not recognize the sheet's current layout - please check " & vbCrLf & _
               "Material_Detail/Dashboard by hand before saving, and report back what you see.", vbExclamation
    End If

    MsgBox "Migration to English schema complete." & vbCrLf & vbCrLf & logMsg & vbCrLf & _
           "Please spot-check a few materials on Dashboard and Material_Detail, then save " & _
           "the file BEFORE importing the other new VBA modules.", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred during migration: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "Some changes may have already been applied. Do NOT save - close the file without " & _
           "saving, restore from your backup, and report this exact error message.", vbCritical
End Sub

' Renames a ListColumn from oldName to newName if oldName is currently
' present; does nothing (returns 0) if oldName isn't found (already
' renamed, or the live file's exact wording differs from what's assumed
' here). Returns 1 if a rename was actually performed.
Private Function RenameTableColumn(wb As Workbook, sheetName As String, oldName As String, newName As String) As Long
    RenameTableColumn = 0
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then Exit Function

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = sh.ListObjects(sheetName)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function

    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(oldName)
    On Error GoTo 0
    If col Is Nothing Then Exit Function  ' already renamed, or a name mismatch - skipped, not an error

    col.Name = newName
    RenameTableColumn = 1
End Function

' Rewrites Material_Detail's per-material row labels using ROW POSITION
' within each block (never by matching old Japanese text), so a slightly-
' wrong assumption about the exact old wording can't cause data to be
' silently left half-migrated. Block layout per material (unchanged by
' this session's translation - only the label text differs): header row
' (column A = RM_Code) -> zero or more (batch-count row / usage row) pairs,
' each batch-count row identified by column C = "No. of batches" -> Total
' Usage row -> TTAF Stock row -> Self Stock row -> Order row -> PO_No row
' -> Total Stock row.
' Caveat: "No. of batches" itself was already English before this
' session's translation project and is assumed unchanged; if your live
' file actually has a different label there, the pair-detection loop
' below won't find any pairs (mdBlocks will still be counted correctly,
' but the intermediate rows' "Usage (kg)" relabel will be skipped for
' that block) - the 6 fixed rows after the pairs are still located
' correctly either way, since they're found relative to wherever the
' pair-scan stops.
Private Sub MigrateMaterialDetailLabels(wb As Workbook, ByRef blockCount As Long, ByRef doneMarkerCount As Long)
    blockCount = 0
    doneMarkerCount = 0
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets("Material_Detail")
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub

    ' The single sheet-wide "Item" column header
    sh.Cells(MD_HEADER_ROW, 2).Value = "Item"

    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row
    Dim lastWeekCol As Long: lastWeekCol = MD_WEEK_START_COL + 200 ' generous upper bound; actual data ends well before this
    Dim r As Long: r = MD_HEADER_ROW + 1
    Do While r <= lastRow
        Dim rmCode As String: rmCode = Trim(CStr(sh.Cells(r, 1).Value))
        If Len(rmCode) = 0 Then
            r = r + 1
        Else
            Dim headerRow As Long: headerRow = r
            Dim rr As Long: rr = headerRow + 1
            Do While rr <= lastRow And Trim(CStr(sh.Cells(rr, 3).Value)) = "No. of batches"
                sh.Cells(rr + 1, 2).Value = "Usage (kg)"
                rr = rr + 2
            Loop
            If rr + 5 <= lastRow Then
                sh.Cells(rr, 2).Value = "Total Usage (kg)/week"
                sh.Cells(rr + 1, 2).Value = "TTAF Stock (Actual, kg)"
                sh.Cells(rr + 2, 2).Value = "Self Stock (Actual, kg)"
                sh.Cells(rr + 3, 2).Value = "Order (Planned, kg)"
                sh.Cells(rr + 4, 2).Value = "PO_No"
                sh.Cells(rr + 5, 2).Value = "Total Stock (End of Week, kg)"
                blockCount = blockCount + 1

                ' [済] -> [DONE] on the PO_No row, across all week columns
                Dim poRow As Long: poRow = rr + 4
                Dim actualLastCol As Long: actualLastCol = sh.Cells(poRow, sh.Columns.Count).End(xlToLeft).Column
                Dim c As Long
                For c = MD_WEEK_START_COL To actualLastCol
                    Dim cellVal As String: cellVal = CStr(sh.Cells(poRow, c).Value)
                    If InStr(cellVal, "[済]") > 0 Then
                        sh.Cells(poRow, c).Value = Replace(cellVal, "[済]", "[DONE]")
                        doneMarkerCount = doneMarkerCount + 1
                    End If
                Next c

                r = rr + 6
            Else
                r = rr + 1 ' malformed/truncated block - don't loop forever, move on
            End If
        End If
    Loop
End Sub

' Rewrites Dashboard's per-material "Row Type" column (column 9) using row
' position (consecutive pairs of 2 rows starting at DASH_DATA_START_ROW -
' the first row of a pair has a Part Name in column 1, the pair is
' Theoretical Stock then Actual Stock), and updates any conditional-
' formatting rule on the sheet whose formula references the old Japanese
' literal "実在庫" so it keeps matching after the relabel.
Private Sub MigrateDashboardRowTypes(wb As Workbook, ByRef pairCount As Long, ByRef cfFixedCount As Long)
    pairCount = 0
    cfFixedCount = 0
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets("Dashboard")
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub

    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim r As Long: r = DASH_DATA_START_ROW
    Do While r + 1 <= lastRow
        If Len(Trim(CStr(sh.Cells(r, 1).Value))) > 0 Then
            sh.Cells(r, 9).Value = "Theoretical Stock"
            sh.Cells(r + 1, 9).Value = "Actual Stock"
            pairCount = pairCount + 1
        End If
        r = r + 2
    Loop

    ' Conditional formatting on a real-world sheet can include rule types
    ' other than a plain formula (data bars, color scales, icon sets,
    ' top/bottom, above/below average, etc.), and those object types do
    ' not support .Formula1 - touching it raises runtime error 450 ("Wrong
    ' number of arguments or invalid property assignment"), not a normal
    ' "property not found" error, so a bare On Error Resume Next around
    ' just the read is not enough to guarantee safety on every rule. Each
    ' step below is individually guarded, and .Type (safe to read on every
    ' rule type) is checked before ever touching .Formula1.
    Dim fcs As FormatConditions
    On Error Resume Next
    Set fcs = sh.Cells.FormatConditions
    On Error GoTo 0
    If fcs Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To fcs.Count
        Dim fcType As Long: fcType = -1
        On Error Resume Next
        Err.Clear
        fcType = fcs.Item(i).Type
        On Error GoTo 0
        If fcType = xlExpression Then
            Dim f1 As String: f1 = ""
            On Error Resume Next
            Err.Clear
            f1 = fcs.Item(i).Formula1
            On Error GoTo 0
            If InStr(f1, "実在庫") > 0 Then
                On Error Resume Next
                Err.Clear
                fcs.Item(i).Formula1 = Replace(f1, "実在庫", "Actual Stock")
                If Err.Number = 0 Then cfFixedCount = cfFixedCount + 1
                On Error GoTo 0
            End If
        End If
    Next i
End Sub
