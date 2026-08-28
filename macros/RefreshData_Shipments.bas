Attribute VB_Name = "RefreshData_Shipments"
Option Explicit

' ============================================================================
' RefreshData_Shipments module
'
'   RefreshShipments : When you select the "CSA Report", imports its
'                      "Shipping Schedule" sheet in full and updates
'                      T_Shipments (order-to-arrival actuals/plans).
'                      Imports everything at once rather than per PO No
'                      (since order-to-arrival takes 4-6 months and several
'                      POs are always in progress in parallel).
'
' Columns: D=CSA Product Code (material code), G=CSA Order firm month
' (order month), I=CSA PO No., K=Vessel, L=Container, N=Confirmed Order
' Qty, O=Original ETD, Q=2 week transit to TTAF (Latest ETA [column P] +
' 14 days - this column, not column P itself, is what's used for the
' actual arrival forecast), S=Received At TTAF, T=Status.
' Unlike TTAF PART NUMBER, CSA Product Code is already spelled almost
' identically to our own RM_Code, so instead of matching via
' TTAF_Code/Description (ResolveTTAFPart on the RefreshData_StockActuals
' side), this matches RM_Codes directly against each other
' (case/leading-trailing-space insensitive).
'
' [Important] About T_Shipments' unique key. It used to manage rows
' uniquely by "material name + PO number" alone, but in real CSA Reports,
' the same material and PO number frequently arrive as a split shipment
' across multiple rows (it's common for a single PO to be split across 5
' rows). With a key of just material name + PO number, these rows were
' treated as the same row, so a later-read row would overwrite an earlier
' one and quantities that had actually arrived were silently lost (55
' such cases were found in this project's real data). So rows are now
' distinguished by a composite key that also includes the container
' number (column L) and Original ETD (column O). Even then, in the rare
' case where multiple rows share the exact same combination (e.g. several
' batches consolidated into the same container), they are finally
' distinguished by a sequence number based on their order of appearance in
' the file (see DateKeyStr/BuildShipmentRowIndex).
' The migration that added the Vessel/Container/Original_ETD columns
' (columns 10-12) needed for this composite key to the T_Shipments of an
' existing, already-in-use workbook has already been completed (the
' one-time AddShipmentSplitColumns/CleanupOrphanedPreSplitShipmentRows
' macros have been removed from this module since they were already run).
'
' Since T_Shipments is a heavyweight table referenced extensively by
' Incoming (a material x week SUMIFS), for the same reason as
' RefreshBOM (M_BOM)/RefreshWeeklyBatches (Production_Plan), new rows are first
' all counted and then added with a single Resize, and updates to existing
' rows are also read/written a whole row at a time (calling ListRows.Add
' one row at a time risks making Excel stop responding).
'
' [About order management (Material_Detail integration)] After importing
' T_Shipments, SyncMaterialDetailOrders is called to automatically update
' Material_Detail's Order row (Planned, kg) and PO_No row (right below the
' Order row) to match the CSA Report's latest Status column.
'   - Status="Unconfirmed" with ETA not yet set (TBC): a provisional week
'     is calculated from Order_Month + M_RawMaterials[LeadTimeWeeks], and
'     the Order/PO_No cells are moved to that week (strictly a provisional
'     forecast).
'   - Status="Unconfirmed"/"In-transit" with a known ETA: moved to the week
'     of T_Shipments[Effective_Week] (which already reflects column Q's
'     date). Follows along every time the ETA is updated.
'   - Status="TTAF Stock": fixed at the last known week, the PO_No cell
'     has "[DONE]" appended, and it's excluded from Incoming's
'     calculation target (the number itself is kept as history).
'   - When the same PO number's shipment is split across multiple rows
'     (split shipment), the Order/PO_No cells are also automatically split
'     across that many weeks.
'   - Whenever a cell is moved/split/finalized, the change is recorded in
'     a cell comment.
' Materials with no block on Material_Detail (e.g. packaging materials not
' used in the BOM) and shipments whose PO_No doesn't appear anywhere on
' Material_Detail are excluded (Incoming's fallback that reads
' T_Shipments directly still covers these).
'
' For the overall design rationale (performance, DataBodyRange,
' DisplayAlerts, etc.), see the comment at the top of the
' RefreshData_Utilities module.
' ============================================================================

Sub RefreshShipments()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , _
        "Please select the CSA Report (Shipping Schedule import) file")
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

    ' A pivot table within the same CSA Report might be sent without being
    ' refreshed first (same reasoning as RefreshTTAFStock), so always
    ' refresh it before reading any data.
    srcWb.RefreshAll
    DoEvents

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim shipTbl As ListObject: Set shipTbl = thisWb.Sheets("T_Shipments").ListObjects("T_Shipments")
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    ' If the migration for the composite-key columns (Vessel/Container/
    ' Original_ETD) hasn't happened yet, continuing as-is would shift
    ' column positions and abort with a runtime error (to avoid a
    ' confusing error, show a clear message here and stop safely instead).
    If shipTbl.ListColumns.Count < 12 Then
        srcWb.Close SaveChanges:=False
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        MsgBox "T_Shipments' column layout doesn't match what's expected (the Vessel/Container/" & vbCrLf & _
               "Original_ETD columns were not found). This may be a workbook that needs to be" & vbCrLf & _
               "regenerated with build_soh.py, or the column layout may be broken. Please check the sheet's structure.", vbExclamation
        Exit Sub
    End If

    Dim sh As Worksheet: Set sh = srcWb.Sheets("Shipping Schedule")

    Dim rmCodeIdx As Object: Set rmCodeIdx = CreateObject("Scripting.Dictionary")
    Call BuildRMCodeIndex(rmTbl, rmCodeIdx)
    Dim knownAliasIdx As Object: Set knownAliasIdx = BuildKnownAliasIndex()

    Dim shipIdx As Object: Set shipIdx = BuildShipmentRowIndex(shipTbl)

    ' Reading the sheet cell by cell is slow, so read a generously sized
    ' range as a single array read, then scan it.
    Const MAX_ROWS As Long = 3000
    Dim data As Variant
    data = sh.Range(sh.Cells(2, 1), sh.Cells(MAX_ROWS, 20)).Value

    ' [Important] T_Shipments is a table referenced extensively by
    ' Incoming (a material x week SUMIFS, 90 materials x 104 weeks).
    ' This is the same structure (heavyweight table x adding many rows one
    ' at a time) that previously caused Excel to stop responding when new
    ' rows were added one at a time via ListRows.Add to M_BOM
    ' (RefreshBOM)/Production_Plan (RefreshWeeklyBatches), so here too, new rows
    ' are first counted and then added all at once with a single Resize.
    ' Updates to existing rows are also read/written a whole row at a
    ' time, rather than writing to individual cells one at a time per row.
    Dim r As Long, added As Long, updated As Long, unresolved As String
    added = 0: updated = 0: unresolved = ""
    Dim updateVals As Object: Set updateVals = CreateObject("Scripting.Dictionary")  ' row number -> Array(qty,eta,receivedDate,status,orderMonth,vessel,container,origEtd)
    Dim newRecords As Object: Set newRecords = CreateObject("Scripting.Dictionary")  ' composite key -> Array(partName,poNo,qty,eta,receivedDate,status,orderMonth,vessel,container,origEtd)
    ' The number of times each composite-key base (material+PO number+container+OriginalETD) has appeared so far in this file
    Dim seqCounter As Object: Set seqCounter = CreateObject("Scripting.Dictionary")
    seqCounter.CompareMode = vbTextCompare

    For r = 1 To (MAX_ROWS - 2 + 1)
        Dim rmCodeRaw As String: rmCodeRaw = Trim(CStr(data(r, 4)))
        Dim poNo As String: poNo = Trim(CStr(data(r, 9)))
        If Len(rmCodeRaw) = 0 Or Len(poNo) = 0 Then GoTo NextRow

        Dim kRow As String: kRow = NormalizeText(rmCodeRaw)
        Dim matchedPart As String: matchedPart = ""
        If rmCodeIdx.Exists(kRow) Then
            matchedPart = rmCodeIdx(kRow)
        Else
            ' Try reading it as the "0" (zero)/"O" (letter O) spelling
            ' variant (same reasoning as RefreshData_StockActuals'
            ' ResolveTTAFPart - a spelling variant repeatedly found in the
            ' TTAF-side source data. E.g. the CSA Report's CSA Product Code
            ' is "0JN" while M_RawMaterials' official spelling is "OJN").
            Dim kZeroToO As String: kZeroToO = Replace(kRow, "0", "O")
            If kZeroToO <> kRow And rmCodeIdx.Exists(kZeroToO) Then
                matchedPart = rmCodeIdx(kZeroToO)
            Else
                Dim kOToZero As String: kOToZero = Replace(kRow, "O", "0")
                If kOToZero <> kRow And rmCodeIdx.Exists(kOToZero) Then matchedPart = rmCodeIdx(kOToZero)
            End If
            ' If still not found, try a known alias that symbol-variant
            ' handling alone can't absorb (see BuildKnownAliasIndex).
            If Len(matchedPart) = 0 And knownAliasIdx.Exists(kRow) Then
                Dim aliasKey As String: aliasKey = NormalizeText(CStr(knownAliasIdx(kRow)))
                If rmCodeIdx.Exists(aliasKey) Then matchedPart = rmCodeIdx(aliasKey)
            End If
        End If

        If Len(matchedPart) = 0 Then
            If InStr(unresolved, rmCodeRaw) = 0 Then unresolved = unresolved & rmCodeRaw & "; "
            GoTo NextRow
        End If

        Dim qty As Variant: qty = data(r, 14)
        If Not IsNumeric(qty) Then qty = 0

        Dim eta As Variant: eta = data(r, 17)  ' column Q (2 week transit to TTAF). Column P itself is not used
        If Not IsDate(eta) Then eta = Empty

        Dim receivedDate As Variant: receivedDate = data(r, 19)
        If Not IsDate(receivedDate) Then receivedDate = Empty

        Dim statusVal As String: statusVal = Trim(CStr(data(r, 20)))

        Dim orderMonth As Variant: orderMonth = data(r, 7)
        If Not IsDate(orderMonth) Then orderMonth = Empty

        Dim vessel As String: vessel = Trim(CStr(data(r, 11)))
        Dim container As String: container = Trim(CStr(data(r, 12)))
        Dim origEtd As Variant: origEtd = data(r, 15)
        If Not IsDate(origEtd) Then origEtd = Empty

        Dim baseKey As String
        baseKey = matchedPart & "|" & poNo & "|" & container & "|" & DateKeyStr(origEtd)
        Dim seq As Long
        If seqCounter.Exists(baseKey) Then
            seq = seqCounter(baseKey) + 1
        Else
            seq = 1
        End If
        seqCounter(baseKey) = seq
        Dim key As String: key = baseKey & "|" & seq

        If shipIdx.Exists(key) Then
            updateVals(shipIdx(key)) = Array(CDbl(qty), eta, receivedDate, statusVal, orderMonth, vessel, container, origEtd)
            updated = updated + 1
        Else
            If Not newRecords.Exists(key) Then added = added + 1
            newRecords(key) = Array(matchedPart, poNo, CDbl(qty), eta, receivedDate, statusVal, orderMonth, vessel, container, origEtd)
        End If
NextRow:
    Next r

    ' ---- Apply updates to existing rows in bulk (read -> rewrite -> write back, once per row) ----
    ' [Important - a long-standing bug found and fixed during this pass]
    ' Since column 8 (Effective_Week) is a formula cell, writing the whole
    ' row back in one shot via .Range.Value = array would write the
    ' "already-calculated value" as of when .Value was read, destroying
    ' the formula itself and turning it into a fixed value (previously,
    ' this bug meant that once a row had been updated even once,
    ' Effective_Week stayed frozen at its value at that moment forever,
    ' and even when Latest_ETA/Received_Date changed afterward, the week
    ' never followed along). To avoid touching column 8's formula at all,
    ' the write-back is split into two parts around it: the first half
    ' (columns 1-7) and the second half (columns 9-12).
    Dim rowKey As Variant
    For Each rowKey In updateVals.Keys
        Dim rowN As Long: rowN = CLng(rowKey)
        Dim uVals As Variant: uVals = updateVals(rowKey)
        Dim rowRng As Range: Set rowRng = shipTbl.ListRows(rowN).Range
        Dim rowArr As Variant: rowArr = rowRng.Value
        rowArr(1, 4) = uVals(0)
        If Not IsEmpty(uVals(1)) Then rowArr(1, 5) = uVals(1)
        If Not IsEmpty(uVals(2)) Then rowArr(1, 6) = uVals(2)
        rowArr(1, 7) = uVals(3)
        If Not IsEmpty(uVals(4)) Then rowArr(1, 9) = uVals(4)
        rowArr(1, 10) = uVals(5)
        rowArr(1, 11) = uVals(6)
        If Not IsEmpty(uVals(7)) Then rowArr(1, 12) = uVals(7)

        Dim leftPart As Variant
        ReDim leftPart(1 To 1, 1 To 7)
        Dim ci As Long
        For ci = 1 To 7: leftPart(1, ci) = rowArr(1, ci): Next ci
        rowRng.Resize(1, 7).Value = leftPart

        Dim rightPart As Variant
        ReDim rightPart(1 To 1, 1 To 4)
        For ci = 9 To 12: rightPart(1, ci - 8) = rowArr(1, ci): Next ci
        rowRng.Cells(1, 9).Resize(1, 4).Value = rightPart
    Next rowKey

    ' ---- Add new rows in bulk (a single Resize + array write. Since the
    ' Effective_Week column doesn't benefit from automatic formula-column
    ' duplication, the same expression as build_soh.py's
    ' week_index_formula_clamped is written explicitly) ----
    If newRecords.Count > 0 Then
        Dim oldRowCount As Long: oldRowCount = shipTbl.ListRows.Count
        Dim nWeeksCal As Long: nWeeksCal = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count
        shipTbl.Resize shipTbl.Range.Resize(shipTbl.Range.Rows.Count + newRecords.Count, shipTbl.Range.Columns.Count)
        Dim nNew As Long: nNew = newRecords.Count
        Dim outArr() As Variant
        ReDim outArr(1 To nNew, 1 To 12)
        Dim ni As Long: ni = 0
        Dim recKey As Variant
        For Each recKey In newRecords.Keys
            ni = ni + 1
            Dim rec As Variant: rec = newRecords(recKey)
            Dim absRow As Long: absRow = oldRowCount + 1 + ni  ' the actual sheet row number of this table's new row (1 header row + existing rows + ni)
            outArr(ni, 1) = rec(0)   ' Part Name
            outArr(ni, 2) = rec(1)   ' PO_No
            outArr(ni, 3) = Empty    ' Order_Date (hand-entered - not touched)
            outArr(ni, 4) = rec(2)   ' Confirmed_Qty
            outArr(ni, 5) = rec(3)   ' Latest_ETA
            outArr(ni, 6) = rec(4)   ' Received_Date
            outArr(ni, 7) = rec(5)   ' Status
            outArr(ni, 8) = "=IFERROR(MAX(1,MIN(" & nWeeksCal & ",INT((IF(F" & absRow & "="""",E" & absRow & ",F" & absRow & _
                ")-(DATE(Cal_Weeks!$B$1,1,1)-WEEKDAY(DATE(Cal_Weeks!$B$1,1,1),3)))/7)+1)),"""")"
            outArr(ni, 9) = rec(6)   ' Order_Month
            outArr(ni, 10) = rec(7)  ' Vessel
            outArr(ni, 11) = rec(8)  ' Container
            outArr(ni, 12) = rec(9)  ' Original_ETD
        Next recKey
        shipTbl.ListRows(oldRowCount + 1).Range.Resize(nNew, 12).Formula = outArr
    End If

    ' [Important] Application.Calculation is still xlCalculationManual, so
    ' neither the existing rows' Latest_ETA/Received_Date updates nor the
    ' new rows' Effective_Week formulas have been recalculated yet as
    ' things stand (SyncMaterialDetailOrders reading via
    ' DataBodyRange.Value would risk getting the stale cached values).
    ' Explicitly recalculate just the T_Shipments range (avoiding a
    ' whole-workbook recalculation).
    shipTbl.Range.Calculate

    ' After T_Shipments has been imported, sync Material_Detail's
    ' Order/PO_No rows to match the CSA Report's latest Status/ETA (see the
    ' comment at the top of this module).
    Dim mdChanged As Long, mdFrozen As Long
    Call SyncMaterialDetailOrders(thisWb, rmTbl, shipTbl, mdChanged, mdFrozen)

    ' Guard the cleanup step itself against failing with "object variable
    ' not set," even in the case where srcWb has already become Nothing
    ' (some automated process on the source file's side can close the
    ' workbook right after it's opened).
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim msg As String
    msg = "T_Shipments has been updated." & vbCrLf & "Added: " & added & ", updated: " & updated & vbCrLf & _
          "(Rows with the same PO No + material combination are overwritten. The Order_Date field is hand-entered, so it is never overwritten.)" & vbCrLf & vbCrLf & _
          "Material_Detail Order/PO_No auto-updates: " & mdChanged & " (of which finalized as TTAF Stock and excluded from calculation: " & mdFrozen & ")" & vbCrLf & _
          "(Cell comments have been left at each changed location)"
    If Len(unresolved) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Rows whose material code could not be matched and were not applied:" & vbCrLf & unresolved
    End If
    MsgBox msg, vbInformation
    Exit Sub

ErrHandler:
    ' [Important] On Error Resume Next automatically clears the Err object
    ' (a VBA quirk), so the error number/description must be saved into
    ' variables before any cleanup code runs. Skipping this means the
    ' MsgBox below always shows "(blank)" and the real cause of the error
    ' is never known (this actually happened and made troubleshooting
    ' impossible, hence the fix).
    Dim errNum2 As Long: errNum2 = Err.Number
    Dim errMsg2 As String: errMsg2 = Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "An error occurred during the refresh: (" & errNum2 & ") " & errMsg2, vbCritical
End Sub

' Indexes M_RawMaterials' Part Name (=RM_Code) itself by normalized text.
' Distinct from the TTAF_Code/Description matching
' (BuildTTAFCodeAndDescIndex on the RefreshData_StockActuals side): since
' CSA Product Code is already spelled almost identically to RM_Code,
' matching RM_Codes directly against each other is more reliable.
Private Sub BuildRMCodeIndex(rmTbl As ListObject, rmCodeIdx As Object)
    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    If rmN > 0 Then
        Dim rmNames As Variant: rmNames = rmTbl.ListColumns(1).DataBodyRange.Value
        Dim i As Long
        For i = 1 To rmN
            Dim k As String: k = NormalizeText(CStr(rmNames(i, 1)))
            If Len(k) > 0 And Not rmCodeIdx.Exists(k) Then rmCodeIdx(k) = CStr(rmNames(i, 1))
        Next i
    End If
End Sub

' Handles known cases individually where the CSA Report side's name
' spelling differs significantly from M_RawMaterials' official Part Name
' (cases that simple symbol-variant or 0/O spelling handling can't
' absorb). The key is the CSA Product Code side's spelling run through
' NormalizeText; the value is M_RawMaterials' official Part Name (this too
' is run through NormalizeText on the RefreshShipments side and matched
' against rmCodeIdx). When a new case is found, it can be handled just by
' adding one line here.
Private Function BuildKnownAliasIndex() As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbTextCompare
    ' The CSA Report's "PET FILM(900*1000)" is registered in
    ' M_RawMaterials as "Ester Film" (we were initially told it was "PP
    ' Film", but confirmed this is the correct one).
    idx(NormalizeText("PET FILM(900*1000)")) = "Ester Film"
    ' The material whose CSA Product Code in the CSA Report is "TPP-1469"
    ' is correctly registered in M_RawMaterials as "CHEM-1850". This has
    ' been confirmed to be a spelling mistake on the TTAF side when
    ' generating the CSA Report, and TTAF has already been asked to fix
    ' it, but (1) during the period before the fix takes effect, and (2)
    ' even after the fix, until this PO drops out of the CSA Report, past
    ' rows may keep the old spelling - so this is kept here to accept both
    ' spellings (once the TTAF side is fixed to "CHEM-1850", the exact
    ' match on the rmCodeIdx side resolves it correctly on its own, so
    ' there is no need to remove this alias entry afterward).
    idx(NormalizeText("TPP-1469")) = "CHEM-1850"
    Set BuildKnownAliasIndex = idx
End Function

' Builds, once, an index of T_Shipments' composite key (material name + PO
' number + container + OriginalETD + appearance-order sequence number) ->
' row number. Uses exactly the same composite-key generation logic as the
' RefreshShipments side, re-numbering the sequence from T_Shipments' own
' current row order (see the comment at the top of this module).
Private Function BuildShipmentRowIndex(tbl As ListObject) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim n As Long: n = tbl.ListRows.Count
    If n > 0 Then
        Dim data As Variant
        data = tbl.DataBodyRange.Value  ' read all 12 columns (Part Name...Original_ETD) together
        Dim seqCounter As Object: Set seqCounter = CreateObject("Scripting.Dictionary")
        seqCounter.CompareMode = vbTextCompare
        Dim i As Long
        For i = 1 To n
            Dim partNameV As String: partNameV = Trim(CStr(data(i, 1)))
            Dim poNoV As String: poNoV = Trim(CStr(data(i, 2)))
            Dim containerV As String: containerV = Trim(CStr(data(i, 11)))
            Dim origEtdV As Variant: origEtdV = data(i, 12)
            Dim baseKey As String
            baseKey = partNameV & "|" & poNoV & "|" & containerV & "|" & DateKeyStr(origEtdV)
            Dim seq As Long
            If seqCounter.Exists(baseKey) Then
                seq = seqCounter(baseKey) + 1
            Else
                seq = 1
            End If
            seqCounter(baseKey) = seq
            idx(baseKey & "|" & seq) = i
        Next i
    End If
    Set BuildShipmentRowIndex = idx
End Function

' Converts a date (or Empty) into a stable string for use in a composite
' key (uses the serial value's integer part as-is, so it's unaffected by
' regional date-display differences).
Private Function DateKeyStr(v As Variant) As String
    If IsEmpty(v) Then
        DateKeyStr = ""
    ElseIf IsDate(v) Then
        DateKeyStr = CStr(CLng(CDate(v)))
    Else
        DateKeyStr = ""
    End If
End Function

' Columns: Part Name(1), PO_No(2), Order_Date(3, hand-entered - not
' touched), Confirmed_Qty(4), Latest_ETA(5), Received_Date(6), Status(7),
' Effective_Week(8, formula), Order_Month(9), Vessel(10), Container(11),
' Original_ETD(12).
' The Effective_Week (column 8) formula is explicitly generated using the
' same expression as build_soh.py's week_index_formula_clamped when
' RefreshShipments writes new rows in bulk.
' Vessel/Container/Original_ETD are used in the composite key that
' uniquely distinguishes each row of a split shipment (see
' BuildShipmentRowIndex/DateKeyStr).

' Automatically updates Material_Detail's Order row (Planned, kg) and
' PO_No row (right below the Order row) to match T_Shipments' latest
' Status/Effective_Week (see the comment at the top of this module). For
' each material, only the 2 rows (Order row, PO_No row) are read in bulk
' -> rewritten in memory -> written back in bulk (no cell-by-cell
' reads/writes - this isn't a heavyweight table so it's unlikely to matter
' much either way, but this keeps it consistent with the established
' design approach).
Private Sub SyncMaterialDetailOrders(thisWb As Workbook, rmTbl As ListObject, shipTbl As ListObject, _
        ByRef changedCells As Long, ByRef frozenCells As Long)
    changedCells = 0
    frozenCells = 0
    Dim mdSheet As Worksheet
    On Error Resume Next
    Set mdSheet = thisWb.Sheets("Material_Detail")
    On Error GoTo 0
    If mdSheet Is Nothing Then Exit Sub

    Dim shipN As Long: shipN = shipTbl.ListRows.Count
    If shipN = 0 Then Exit Sub
    Dim shipData As Variant: shipData = shipTbl.DataBodyRange.Value
    ' shipData columns: 1=Part Name,2=PO_No,3=Order_Date,4=Confirmed_Qty,5=Latest_ETA,
    '                    6=Received_Date,7=Status,8=Effective_Week,9=Order_Month

    ' "material name|PO number" -> list of shipData row numbers for that combination (a split shipment produces multiple rows)
    Dim byMatPo As Object: Set byMatPo = CreateObject("Scripting.Dictionary")
    byMatPo.CompareMode = vbTextCompare
    Dim i As Long
    For i = 1 To shipN
        Dim sPoNo As String: sPoNo = Trim(CStr(shipData(i, 2)))
        If Len(sPoNo) = 0 Then GoTo NextShipRow
        Dim mpKey As String: mpKey = Trim(CStr(shipData(i, 1))) & "|" & sPoNo
        Dim lst As Object
        If byMatPo.Exists(mpKey) Then
            Set lst = byMatPo(mpKey)
        Else
            Set lst = CreateObject("Scripting.Dictionary")
            byMatPo.Add mpKey, lst
        End If
        lst.Add lst.Count, i
NextShipRow:
    Next i
    If byMatPo.Count = 0 Then Exit Sub

    ' Index M_RawMaterials' LeadTimeWeeks by material name once, so it can
    ' be looked up (used for the provisional forecast when ETA is not yet
    ' set (TBC)).
    Dim ltIdx As Object: Set ltIdx = CreateObject("Scripting.Dictionary")
    ltIdx.CompareMode = vbTextCompare
    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    If rmN > 0 Then
        Dim rmNameLt As Variant: rmNameLt = rmTbl.ListColumns(1).DataBodyRange.Resize(rmN, 1).Value
        Dim rmLtCol As Variant: rmLtCol = rmTbl.ListColumns("LeadTimeWeeks").DataBodyRange.Value
        Dim li As Long
        For li = 1 To rmN
            Dim ltk As String: ltk = Trim(CStr(rmNameLt(li, 1)))
            If Len(ltk) > 0 And Not ltIdx.Exists(ltk) Then
                Dim ltv As Double: ltv = 0
                If IsNumeric(rmLtCol(li, 1)) Then ltv = CDbl(rmLtCol(li, 1))
                ltIdx(ltk) = ltv
            End If
        Next li
    End If

    ' Collect the "Order" row position for each block on Material_Detail once (the PO_No row is the one right below it).
    Dim lastRow As Long: lastRow = mdSheet.Cells(mdSheet.Rows.Count, 2).End(xlUp).Row
    Dim orderRowByMat As Object: Set orderRowByMat = CreateObject("Scripting.Dictionary")
    orderRowByMat.CompareMode = vbTextCompare
    Dim curMatCode As String: curMatCode = ""
    Dim r As Long
    For r = MD_HEADER_ROW + 1 To lastRow
        Dim colAVal As String: colAVal = Trim(CStr(mdSheet.Cells(r, 1).Value))
        If Len(colAVal) > 0 Then curMatCode = colAVal
        If Len(curMatCode) > 0 And Trim(CStr(mdSheet.Cells(r, 2).Value)) = "Order (Planned, kg)" Then
            If Not orderRowByMat.Exists(curMatCode) Then orderRowByMat.Add curMatCode, r
        End If
    Next r
    If orderRowByMat.Count = 0 Then Exit Sub

    Dim nWeeks As Long: nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count
    Dim lastWeekCol As Long: lastWeekCol = MD_WEEK_START_COL + nWeeks - 1

    ' Extract just the "material name" side of the shipment info, dedupe it, and process each material only once.
    Dim matNamesSeen As Object: Set matNamesSeen = CreateObject("Scripting.Dictionary")
    matNamesSeen.CompareMode = vbTextCompare
    Dim mpKeyOuter As Variant
    For Each mpKeyOuter In byMatPo.Keys
        Dim sepPos As Long: sepPos = InStr(CStr(mpKeyOuter), "|")
        Dim matName As String: matName = Left(CStr(mpKeyOuter), sepPos - 1)
        If matNamesSeen.Exists(matName) Then GoTo NextMatKey
        matNamesSeen.Add matName, True
        If Not orderRowByMat.Exists(matName) Then GoTo NextMatKey  ' a material with no block on Material_Detail

        Dim orderRow As Long: orderRow = orderRowByMat(matName)
        Dim poRow As Long: poRow = orderRow + 1
        Dim blockArr As Variant
        blockArr = mdSheet.Range(mdSheet.Cells(orderRow, MD_WEEK_START_COL), mdSheet.Cells(poRow, lastWeekCol)).Value
        ' blockArr(1,w)=Order row value, blockArr(2,w)=PO_No row value (w=1..nWeeks)

        ' Collect the not-yet-finalized (no "[DONE]" marker) PO numbers
        ' that appear in this material's PO_No row, week by week, and
        ' group them by PO number.
        Dim activeByPo As Object: Set activeByPo = CreateObject("Scripting.Dictionary")
        activeByPo.CompareMode = vbTextCompare
        Dim w As Long
        For w = 1 To nWeeks
            Dim poCellVal As String: poCellVal = Trim(CStr(blockArr(2, w)))
            If Len(poCellVal) > 0 And InStr(poCellVal, "[DONE]") = 0 Then
                Dim poLst As Object
                If activeByPo.Exists(poCellVal) Then
                    Set poLst = activeByPo(poCellVal)
                Else
                    Set poLst = CreateObject("Scripting.Dictionary")
                    activeByPo.Add poCellVal, poLst
                End If
                poLst.Add poLst.Count, w
            End If
        Next w
        If activeByPo.Count = 0 Then GoTo NextMatKey

        Dim blockChanged As Boolean: blockChanged = False
        Dim commentsToAdd As Object: Set commentsToAdd = CreateObject("Scripting.Dictionary")  ' week -> comment text

        Dim poKey As Variant
        For Each poKey In activeByPo.Keys
            Dim thisMpKey As String: thisMpKey = matName & "|" & CStr(poKey)
            If Not byMatPo.Exists(thisMpKey) Then GoTo NextPoKey  ' this PO number's shipment info isn't in the CSA Report yet

            Dim shipRowsForPo As Object: Set shipRowsForPo = byMatPo(thisMpKey)
            Dim existingWeeks As Object: Set existingWeeks = activeByPo(poKey)

            ' --- Build what this PO number's "correct state" should be, from the CSA Report's latest info ---
            Dim newWeeks As Object: Set newWeeks = CreateObject("Scripting.Dictionary")  ' week -> Array(qty, frozen)
            Dim siKey As Variant
            For Each siKey In shipRowsForPo.Keys
                Dim shipRow As Long: shipRow = shipRowsForPo(siKey)
                Dim qty As Double: qty = 0
                If IsNumeric(shipData(shipRow, 4)) Then qty = CDbl(shipData(shipRow, 4))
                Dim statusText As String: statusText = Trim(CStr(shipData(shipRow, 7)))
                Dim targetWeek As Long: targetWeek = 0
                If IsNumeric(shipData(shipRow, 8)) Then
                    targetWeek = CLng(shipData(shipRow, 8))
                Else
                    ' ETA not yet set (TBC). Calculate a provisional week
                    ' from Order_Month + LeadTime_Weeks (using the middle
                    ' of the month - the 15th - as the anchor, so it
                    ' doesn't skew toward either the start or end of the
                    ' month).
                    Dim orderMonthVal As Variant: orderMonthVal = shipData(shipRow, 9)
                    If IsDate(orderMonthVal) Then
                        Dim ltWeeks As Double: ltWeeks = 0
                        If ltIdx.Exists(matName) Then ltWeeks = ltIdx(matName)
                        Dim provDate As Date: provDate = DateSerial(Year(CDate(orderMonthVal)), Month(CDate(orderMonthVal)), 15) + ltWeeks * 7
                        targetWeek = WeekIndexForDate(thisWb, provDate)
                    End If
                End If
                If targetWeek > 0 Then
                    Dim frozenFlag As Boolean: frozenFlag = (statusText = "TTAF Stock")
                    ' Multiple shipment rows for the same PO number (a
                    ' split shipment) can land on the same week, so sum
                    ' rather than overwrite. frozen (treated as finalized)
                    ' only becomes true once every row summed into that
                    ' week has become TTAF Stock (if even one is still in
                    ' transit, that week is not yet excluded from
                    ' Incoming's calculation).
                    If newWeeks.Exists(targetWeek) Then
                        Dim existingNW As Variant: existingNW = newWeeks(targetWeek)
                        newWeeks(targetWeek) = Array(CDbl(existingNW(0)) + qty, CBool(existingNW(1)) And frozenFlag)
                    Else
                        newWeeks(targetWeek) = Array(qty, frozenFlag)
                    End If
                End If
            Next siKey
            If newWeeks.Count = 0 Then GoTo NextPoKey

            ' --- Compare the "correct state" against the current state, week by week ---
            ' (The earlier approach of "clear the whole PO at once -> rewrite
            ' it all at once" had a bug: in a case where part of a split
            ' shipment was already finalized [DONE] with the remainder
            ' still in transit, even weeks that were already finalized and
            ' completely unchanged got cleared and rewritten every time,
            ' making the comment look like an unrelated week move and
            ' causing frozenCells to count the same finalization
            ' repeatedly. So only what actually differs from the current
            ' content is rewritten, week by week.)
            ' weeksToClear: currently-active (not-yet-finalized) weeks that don't exist in the new state (= a source of a move)
            Dim weeksToClear As Object: Set weeksToClear = CreateObject("Scripting.Dictionary")
            Dim ewk As Variant
            For Each ewk In existingWeeks.Items
                Dim ewkL As Long: ewkL = CLng(ewk)
                If Not newWeeks.Exists(ewkL) Then weeksToClear(ewkL) = True
            Next ewk
            ' weeksToWrite: only entries in the new state that actually differ from the current cell content (quantity, finalized marker)
            Dim weeksToWrite As Object: Set weeksToWrite = CreateObject("Scripting.Dictionary")  ' week -> Array(qty,frozen,poText)
            Dim wk As Variant
            For Each wk In newWeeks.Keys
                Dim nwk As Long: nwk = CLng(wk)
                Dim info As Variant: info = newWeeks(wk)
                Dim wantQty As Double: wantQty = CDbl(info(0))
                Dim wantFrozen As Boolean: wantFrozen = CBool(info(1))
                Dim wantPoText As String: wantPoText = CStr(poKey)
                If wantFrozen Then wantPoText = wantPoText & " [DONE]"

                Dim curQty As Double: curQty = 0
                If IsNumeric(blockArr(1, nwk)) Then curQty = CDbl(blockArr(1, nwk))
                Dim curPoText As String: curPoText = Trim(CStr(blockArr(2, nwk)))

                If Abs(curQty - wantQty) > 0.0001 Or curPoText <> wantPoText Then
                    weeksToWrite(nwk) = Array(wantQty, wantFrozen, wantPoText)
                End If
            Next wk
            If weeksToClear.Count = 0 And weeksToWrite.Count = 0 Then GoTo NextPoKey

            ' --- Clear the source weeks (weeksToClear) (operating only on the material block's in-memory array) ---
            Dim oldWeeksSummary As String: oldWeeksSummary = ""
            Dim cwk2 As Variant
            For Each cwk2 In weeksToClear.Keys
                Dim ow As Long: ow = CLng(cwk2)
                If Len(oldWeeksSummary) > 0 Then oldWeeksSummary = oldWeeksSummary & ", "
                oldWeeksSummary = oldWeeksSummary & "week " & ow & " (" & Format(blockArr(1, ow), "0") & "kg)"
                blockArr(1, ow) = Empty
                blockArr(2, ow) = Empty
            Next cwk2

            Dim changeNote As String
            If Len(oldWeeksSummary) > 0 Then
                changeNote = "PO#" & poKey & ": auto-updated from " & oldWeeksSummary & " (" & Format(Date, "yyyy-mm-dd") & ")"
            Else
                changeNote = "PO#" & poKey & ": auto-updated (" & Format(Date, "yyyy-mm-dd") & ")"
            End If
            Dim wwk As Variant
            For Each wwk In weeksToWrite.Keys
                Dim nw As Long: nw = CLng(wwk)
                Dim winfo As Variant: winfo = weeksToWrite(wwk)
                blockArr(1, nw) = CDbl(winfo(0))
                blockArr(2, nw) = CStr(winfo(2))
                If CBool(winfo(1)) Then frozenCells = frozenCells + 1
                commentsToAdd(nw) = changeNote
            Next wwk

            blockChanged = True
            changedCells = changedCells + 1
NextPoKey:
        Next poKey

        If blockChanged Then
            mdSheet.Range(mdSheet.Cells(orderRow, MD_WEEK_START_COL), mdSheet.Cells(poRow, lastWeekCol)).Value = blockArr
            Dim cwKey As Variant
            For Each cwKey In commentsToAdd.Keys
                Dim commentCol As Long: commentCol = MD_WEEK_START_COL + CLng(cwKey) - 1
                Dim targetCell As Range: Set targetCell = mdSheet.Cells(orderRow, commentCol)
                On Error Resume Next
                targetCell.Comment.Delete
                targetCell.AddComment CStr(commentsToAdd(cwKey))
                targetCell.Comment.Shape.TextFrame.AutoSize = True
                On Error GoTo 0
            Next cwKey
        End If
NextMatKey:
    Next mpKeyOuter
End Sub
