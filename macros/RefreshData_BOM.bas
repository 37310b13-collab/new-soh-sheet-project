Attribute VB_Name = "RefreshData_BOM"
Option Explicit

' ============================================================================
' RefreshData_BOM module
'
'   RefreshBOM : When you select "Raw Material - Look Up", updates M_BOM
'                (the usage rates). Targets 4 sheets:
'                  - Slurry Data Base / Powder Data Base / Solution ...
'                    rows are transcribed as-is. Column A = Intermediate,
'                    column D = RM Code, column M = usage per batch.
'                    RM Code can hold not just an M_RawMaterials RM_Code
'                    (e.g. CHEM-1030) but also another intermediate code
'                    (SOL-SCH, TPP-103, etc., for cases where a Slurry/
'                    Powder uses a Solution or another Powder as an
'                    ingredient) - these are written into M_BOM as-is
'                    regardless (they're simply ignored on the
'                    WeeklyConsumption side, with no ill effect).
'                  - Catalyst Data Base ... only uses Substrate rows (rows
'                    where column H "Substrate SC" is filled in) (chemical/
'                    slurry reference rows already have their usage rate
'                    calculated on the Slurry side, so picking them up here
'                    too would double-count them). Corning-supplied
'                    Substrate (column E Description contains "CORNING")
'                    is excluded. The intermediate name is the shortened
'                    product code extracted from column A's Catalyst name
'                    (e.g. "18461-0Q110-1st COAT" -> "0Q110"), RM_Code is
'                    column H, and the quantity is column F (usage per one
'                    catalyst unit, since catalysts are tracked by unit
'                    count on the Production_Plan side).
'                When a new material x intermediate combination is found,
'                the corresponding intermediate breakdown rows (No. of
'                batches / Usage (kg)) are also automatically added to that
'                material's block on Material_Detail
'                (SyncMaterialDetailIntermediates). This means a material
'                just added by AddMaterial (a mini block with no breakdown
'                rows yet, since it isn't in the BOM) ends up looking the
'                same as existing materials after just one run of
'                RefreshBOM (no need to run AddMaterial twice). The same
'                automatic sync happens when an existing material starts
'                being used in a new intermediate.
'
' [Bug already fixed in the past] Production_Plan's pass-through formulas (the
' SUMPRODUCT+M_BOM[PPGridRow] back-calculation formula placed in the row
' of an intermediate that has no production plan of its own and is only
' ever used as an ingredient of other intermediates) used to have a
' structural flaw where, for that intermediate's own recipe row,
' INDEX(Production_Plan[#Data],M_BOM[PPGridRow],...) would also get evaluated,
' creating an actual "path that references the cell itself," even though
' it was ultimately multiplied by 0 and had no effect on the result
' (testing in LibreOffice silently treated it as 0 so this went unnoticed,
' but real Excel detects it as a circular reference). A one-time migration
' macro (FixPassthroughCircularRefs) that rewrote the existing pass-through
' formulas in Production_Plan into a safe form guarded with IF(...,NA(),...) has
' already been run, so it has been removed from the VBA code. Newly
' generated pass-through formulas (from build_soh.py / when RefreshBOM
' runs) are built with this guard from the start.
'
'   FixTheoreticalStockMonthlyReset : TheoreticalStock (theoretical
'                stock) used to never reset from the moment operations
'                began (T_OpeningStock), rolling forward forever without
'                ever looking at actuals, so error would accumulate
'                without bound over a long enough period. This macro
'                rewrites the formulas so that, only in the first week a
'                new month begins, it re-syncs from the previous week's
'                actual stock (Stock), addressing the request to
'                "reset every month and only see this month's plan-vs-
'                actual gap." Week 1 (which starts from T_OpeningStock) is
'                excluded. Running it any number of times produces the
'                same result, so it's safe to re-run.
'
' For the overall design rationale (performance, DataBodyRange,
' DisplayAlerts, etc.), see the comment at the top of the
' RefreshData_Utilities module.
' ============================================================================

Sub RefreshBOM()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , _
        "Please select Raw Material - Look Up")
    If srcPath = False Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' For a file with "read-only recommended" set, leaving DisplayAlerts=True
    ' causes a confirmation dialog to appear during Workbooks.Open, which
    ' destabilizes processing while waiting for a response (this was the
    ' cause of a bug where srcWb ultimately failed to be obtained correctly)
    ' - so suppress it.
    Application.DisplayAlerts = False

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(CStr(srcPath), ReadOnly:=True, UpdateLinks:=False)

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim bomTbl As ListObject: Set bomTbl = thisWb.Sheets("M_BOM").ListObjects("M_BOM")
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")
    Dim ppGrid As ListObject: Set ppGrid = thisWb.Sheets("Production_Plan").ListObjects("Production_Plan")

    ' Build each name set once: one to check whether an RM_Code (Part Name)
    ' exists, and one to determine "is this a known intermediate name"
    ' (used when resolving the trailing-"s" spelling variant) while
    ' extracting Catalyst short codes.
    Dim rmCodeSet As Object: Set rmCodeSet = BuildNameIndex(rmTbl, "Part Name")
    Dim knownIntermediates As Object: Set knownIntermediates = BuildNameIndex(ppGrid, "Intermediate")

    ' To count the diff (updated/added/removed), hold the pre-update
    ' M_BOM content as a (pk->True) set (used only for this count, not for
    ' the actual data update that follows).
    Dim oldPairs As Object: Set oldPairs = CreateObject("Scripting.Dictionary")
    oldPairs.CompareMode = vbTextCompare
    Dim oldN As Long: oldN = bomTbl.ListRows.Count
    If oldN > 0 Then
        Dim oldArr As Variant: oldArr = bomTbl.ListColumns(1).DataBodyRange.Resize(oldN, 2).Value
        Dim oi As Long
        For oi = 1 To oldN
            oldPairs(CStr(oldArr(oi, 1)) & "|" & CStr(oldArr(oi, 2))) = True
        Next oi
    End If

    ' Assemble the new BOM content entirely in an in-memory Dictionary,
    ' without writing anything to the sheet yet
    ' (pk="Intermediate|RM_Code" -> Array(Intermediate, RM_Code, Qty)).
    ' Since Look Up is this system's single source of truth for BOM
    ' information, M_BOM is ultimately replaced in full with this content
    ' (adding rows one at a time via ListRows.Add triggers a dependency
    ' recheck on every add across the large number of formulas that
    ' reference M_BOM - WeeklyConsumption, Material_Detail, Production_Plan's
    ' pass-through formulas - which was the cause of a "freeze" bug once
    ' the row count passed roughly a thousand; so the write is done as a
    ' single block write at the end instead).
    Dim newRows As Object: Set newRows = CreateObject("Scripting.Dictionary")
    newRows.CompareMode = vbTextCompare
    Dim unresolved As String: unresolved = ""

    Dim flatSheets As Variant: flatSheets = Array("Slurry Data Base", "Powder Data Base", "Solution")
    Dim sIdx As Integer
    For sIdx = LBound(flatSheets) To UBound(flatSheets)
        Dim shName As String: shName = CStr(flatSheets(sIdx))
        If SheetExists(srcWb, shName) Then
            Call ProcessLookupFlatSheet(srcWb.Sheets(shName), newRows, rmCodeSet, knownIntermediates, unresolved)
        End If
    Next sIdx
    If SheetExists(srcWb, "Catalyst Data Base") Then
        Call ProcessLookupCatalystSheet(srcWb.Sheets("Catalyst Data Base"), newRows, rmCodeSet, _
            knownIntermediates, unresolved)
    End If

    ' Guard the cleanup step itself against failing with "object variable
    ' not set," even in the case where srcWb has already become Nothing
    ' (some automated process on the source file's side can close the
    ' workbook right after it's opened).
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False

    Dim added As Long, updated As Long, removedCount As Long
    added = 0: updated = 0: removedCount = 0
    Dim k As Variant
    For Each k In newRows.Keys
        If oldPairs.Exists(k) Then updated = updated + 1 Else added = added + 1
    Next k
    For Each k In oldPairs.Keys
        If Not newRows.Exists(k) Then removedCount = removedCount + 1
    Next k

    ' Replaces the whole of M_BOM with a single block write. The
    ' PPGridRow column (column D) is normally a formula column that
    ' Excel's Table feature auto-duplicates on ListRows.Add, but since Add
    ' isn't used here, it's explicitly built as a formula string instead
    ' (identical to build_soh.py's generation pattern).
    Dim n As Long: n = newRows.Count
    If Not bomTbl.DataBodyRange Is Nothing Then bomTbl.DataBodyRange.Delete
    If n > 0 Then
        Dim outArr() As Variant
        ReDim outArr(1 To n, 1 To 4)
        Dim idx As Long: idx = 0
        For Each k In newRows.Keys
            idx = idx + 1
            Dim rec As Variant: rec = newRows(k)
            outArr(idx, 1) = rec(0)
            outArr(idx, 2) = rec(1)
            outArr(idx, 3) = rec(2)
            outArr(idx, 4) = "=IFERROR(MATCH($A" & (idx + 1) & ",Production_Plan[Intermediate],0),99999)"
        Next k
        bomTbl.Resize bomTbl.Range.Resize(n + 1, 4)
        bomTbl.DataBodyRange.Formula = outArr
    End If

    ' Syncs Material_Detail's intermediate breakdown rows (No. of batches /
    ' Usage (kg)) to match the now-updated M_BOM content. Since the M_BOM
    ' update itself is already committed above, an error in this sync step
    ' does not lose the changes already applied to M_BOM/Production_Plan/
    ' WeeklyConsumption (caught individually with On Error Resume Next; on
    ' failure it's just reported in the message).
    Dim addedDetailRows As Long: addedDetailRows = 0
    Dim syncErrMsg As String: syncErrMsg = ""
    On Error Resume Next
    Call SyncMaterialDetailIntermediates(bomTbl, addedDetailRows)
    If Err.Number <> 0 Then
        syncErrMsg = Err.Description
        Err.Clear
    End If
    On Error GoTo ErrHandler

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim msg As String
    msg = "M_BOM has been updated." & vbCrLf & "Updated: " & updated & ", added: " & added & _
          ", removed: " & removedCount
    If addedDetailRows > 0 Then
        msg = msg & vbCrLf & "Intermediate breakdown row pairs added to Material_Detail: " & addedDetailRows
    End If
    If Len(syncErrMsg) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "(Note) An error occurred while automatically adding" & vbCrLf & _
              "Material_Detail's breakdown rows: " & syncErrMsg & vbCrLf & "The M_BOM update itself completed successfully."
    End If
    If Len(unresolved) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Codes found in neither M_RawMaterials nor Production_Plan" & vbCrLf & _
              "(possibly an unregistered new material - please check whether it was missed in AddMaterial):" & vbCrLf & unresolved
    End If
    msg = msg & vbCrLf & vbCrLf & "Since WeeklyConsumption references M_BOM directly, newly added combinations" & vbCrLf & "are also reflected automatically right away."
    MsgBox msg, vbInformation
    Exit Sub

ErrHandler:
    ' [Important] On Error Resume Next automatically clears the Err object
    ' (a VBA quirk), so the error number/description must be saved into
    ' variables before any cleanup code runs. Skipping this means the
    ' MsgBox below always shows "(blank)" and the real cause of the error
    ' is never known (this actually happened and made troubleshooting
    ' impossible, hence the fix).
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "An error occurred during the refresh: (" & errNum & ") " & errMsg, vbCritical
End Sub

' Rewrites TheoreticalStock's (theoretical stock) formulas so that,
' only in the first week a new month begins, they re-sync from
' Stock (actual stock). Every other week still rolls forward as
' before ("previous week's theoretical stock + incoming - consumption"),
' unchanged (week 1, which starts from T_OpeningStock, is unaffected and
' excluded). It used to never reset from the moment operations began,
' letting error accumulate without bound - this addresses the request to
' "reset once a month and only see this month's plan-vs-actual gap."
' Running it any number of times produces the same result, so it's safe
' to re-run. Workbooks newly generated by build_soh.py are built with this
' formula from the start.
Sub FixTheoreticalStockMonthlyReset()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim theoTbl As ListObject: Set theoTbl = thisWb.Sheets("TheoreticalStock").ListObjects("TheoreticalStock")
    Dim dataRange As Range: Set dataRange = theoTbl.DataBodyRange
    If dataRange Is Nothing Then
        MsgBox "TheoreticalStock has no data rows.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim firstDataRow As Long: firstDataRow = dataRange.Row
    Dim firstDataCol As Long: firstDataCol = dataRange.Column
    Dim nRows As Long: nRows = dataRange.Rows.Count
    Dim nCols As Long: nCols = dataRange.Columns.Count  ' column 1=Part Name, column 2=Description, column 3=Week1, column 4=Week2...

    Dim allFormulas As Variant: allFormulas = dataRange.Formula

    Dim r As Long, c As Long, fixedCells As Long: fixedCells = 0
    For r = 1 To nRows
        Dim actualRow As Long: actualRow = firstDataRow + r - 1
        For c = 4 To nCols  ' column 3 (week 1) starts from T_OpeningStock and is excluded
            Dim w As Long: w = c - 2
            Dim curCol As String: curCol = ColLetter(firstDataCol + c - 1)
            Dim prevCol As String: prevCol = ColLetter(firstDataCol + c - 2)
            Dim monthChanged As String
            monthChanged = "INDEX(Cal_Weeks[MonthYearLabel]," & w & ")<>INDEX(Cal_Weeks[MonthYearLabel]," & (w - 1) & ")"
            Dim theoPrior As String
            theoPrior = "IF(" & monthChanged & ",'Stock'!" & prevCol & actualRow & "," & prevCol & actualRow & ")"
            allFormulas(r, c) = "=" & theoPrior & "+'Incoming'!" & curCol & actualRow & "-'WeeklyConsumption'!" & curCol & actualRow
            fixedCells = fixedCells + 1
        Next c
    Next r

    dataRange.Formula = allFormulas

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "TheoreticalStock has been switched to the monthly-reset formula." & vbCrLf & _
           "Cells rewritten: " & fixedCells, vbInformation
    Exit Sub

ErrHandler:
    Dim errNum2 As Long: errNum2 = Err.Number
    Dim errMsg2 As String: errMsg2 = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred during the fix: (" & errNum2 & ") " & errMsg2, vbCritical
End Sub

' Common processing for the Slurry Data Base/Powder Data Base/Solution
' sheets. Rows are imported as-is: column A = Intermediate, column D = RM
' Code, column M = usage per batch. RM Code may be an M_RawMaterials Part
' Name directly, or another intermediate code (SOL-SCH, etc.) - either way
' it's written into M_BOM without distinction (the WeeklyConsumption side
' only looks at Part Names that actually exist in M_RawMaterials, so
' intermediate-code rows are ignored with no ill effect).
Private Sub ProcessLookupFlatSheet(sh As Worksheet, newRows As Object, _
        rmCodeSet As Object, knownIntermediates As Object, ByRef unresolved As String)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If lastRow > 5000 Then lastRow = 5000  ' guard against an outlier value
    If lastRow < 2 Then Exit Sub

    ' Read the whole sheet into an array just once (calling
    ' .Cells(r,c).Value on every loop iteration is slow).
    Dim data As Variant
    data = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, 13)).Value  ' through column M (13)

    Dim r As Long
    For r = 2 To lastRow
        Dim inter As String: inter = Trim(CStr(data(r, 1)))
        Dim rmCode As String: rmCode = Trim(CStr(data(r, 4)))
        Dim v As Variant: v = data(r, 13)
        If Len(inter) > 0 And Len(rmCode) > 0 And IsNumeric(v) Then
            If CDbl(v) <> 0 Then
                newRows(inter & "|" & rmCode) = Array(inter, rmCode, CDbl(v))
                If Not rmCodeSet.Exists(rmCode) And Not knownIntermediates.Exists(rmCode) Then
                    If InStr(unresolved, rmCode) = 0 Then unresolved = unresolved & rmCode & "; "
                End If
            End If
        End If
    Next r
End Sub

' Only uses the Catalyst Data Base sheet's Substrate rows (rows where
' column H "Substrate SC" is filled in). Corning-supplied Substrate
' (column E Description contains "CORNING") is excluded. The intermediate
' name is the shortened product code extracted from column A's Catalyst
' name, RM_Code is column H, and the quantity is column F (per one
' catalyst unit).
Private Sub ProcessLookupCatalystSheet(sh As Worksheet, newRows As Object, _
        rmCodeSet As Object, knownIntermediates As Object, ByRef unresolved As String)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If lastRow > 5000 Then lastRow = 5000
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, 8)).Value  ' through column H (8)

    Dim r As Long
    For r = 2 To lastRow
        Dim subSC As String: subSC = Trim(CStr(data(r, 8)))
        If Len(subSC) = 0 Then GoTo NextRow
        Dim desc As String: desc = UCase(Trim(CStr(data(r, 5))))
        If InStr(desc, "CORNING") > 0 Then GoTo NextRow  ' Corning-supplied is excluded
        Dim catName As String: catName = Trim(CStr(data(r, 1)))
        Dim v As Variant: v = data(r, 6)
        If Len(catName) = 0 Or Not IsNumeric(v) Then GoTo NextRow
        If CDbl(v) = 0 Then GoTo NextRow
        Dim code As String: code = CatalystShortCode(catName, knownIntermediates)
        newRows(code & "|" & subSC) = Array(code, subSC, CDbl(v))
        If Not rmCodeSet.Exists(subSC) Then
            If InStr(unresolved, subSC) = 0 Then unresolved = unresolved & subSC & "; "
        End If
NextRow:
    Next r
End Sub

' Extracts a shortened product code (e.g. "0Q110") from a Catalyst name
' (e.g. "18461-0Q110-1st COAT"). Strips the "18461-" prefix, takes the
' token up to the first space/hyphen, and if it starts with "O" (letter O)
' replaces it with "0" (zero) (a spelling variant in the source data). If
' it still isn't found in the intermediate master and ends in "s", strips
' the trailing "s" and tries again (e.g. "0T420s" -> "0T420").
Private Function CatalystShortCode(catName As String, knownIntermediates As Object) As String
    Dim s As String: s = catName
    If Left(s, 6) = "18461-" Then s = Mid(s, 7)
    Dim i As Long, ch As String, code As String
    code = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = " " Or ch = "-" Then Exit For
        code = code & ch
    Next i
    If Left(code, 1) = "O" Then code = "0" & Mid(code, 2)
    If Not knownIntermediates.Exists(code) Then
        Dim tailCh As String: tailCh = Right(code, 1)
        If (tailCh = "s" Or tailCh = "S") And knownIntermediates.Exists(Left(code, Len(code) - 1)) Then
            code = Left(code, Len(code) - 1)
        End If
    End If
    CatalystShortCode = code
End Function

Private Function SheetExists(wb As Workbook, sName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets(sName)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function

' Called after RefreshBOM runs. Looks at the updated M_BOM content
' (material code x intermediate combinations) and, for any intermediate
' that doesn't yet have breakdown rows (No. of batches / Usage (kg)) in
' its material's block on Material_Detail, inserts that pair of 2 rows
' right before the "Total Usage (kg)/week" row.
Private Sub SyncMaterialDetailIntermediates(bomTbl As ListObject, ByRef addedPairs As Long)
    addedPairs = 0
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    If Not SheetExists(thisWb, "Material_Detail") Then Exit Sub
    If Not SheetExists(thisWb, "Cal_Weeks") Then Exit Sub
    Dim sh As Worksheet: Set sh = thisWb.Sheets("Material_Detail")
    Dim nWeeks As Long: nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count
    If nWeeks <= 0 Then Exit Sub

    ' Assemble M_BOM into a Dictionary of "material code -> list of
    ' intermediate names that use that material (in order of appearance,
    ' duplicates removed)" (a single array read).
    Dim byMat As Object: Set byMat = CreateObject("Scripting.Dictionary")
    byMat.CompareMode = vbTextCompare
    Dim bomN As Long: bomN = bomTbl.ListRows.Count
    If bomN = 0 Then Exit Sub
    Dim bomData As Variant
    bomData = bomTbl.ListColumns(1).DataBodyRange.Resize(bomN, 2).Value  ' 1=Intermediate, 2=Part Name
    Dim bi As Long
    For bi = 1 To bomN
        Dim interN As String: interN = Trim(CStr(bomData(bi, 1)))
        Dim partN As String: partN = Trim(CStr(bomData(bi, 2)))
        If Len(interN) > 0 And Len(partN) > 0 Then
            If Not byMat.Exists(partN) Then
                Dim seenDict As Object: Set seenDict = CreateObject("Scripting.Dictionary")
                seenDict.CompareMode = vbTextCompare
                byMat.Add partN, seenDict
            End If
            If Not byMat(partN).Exists(interN) Then byMat(partN).Add interN, True
        End If
    Next bi
    If byMat.Count = 0 Then Exit Sub

    ' The template row pair for copying formatting (the first "No. of
    ' batches" row pair found in the sheet). Held as a Range object, so
    ' Excel automatically tracks its target row position even as row
    ' insertions happen elsewhere on this sheet from here on (standard VBA
    ' object-reference behavior: inserting another row right below it
    ' doesn't shift it).
    Dim lastRowScan As Long: lastRowScan = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row
    Dim templateRow As Long: templateRow = 0
    Dim tr As Long
    For tr = MD_HEADER_ROW + 1 To lastRowScan
        If Trim(CStr(sh.Cells(tr, 3).Value)) = "No. of batches" Then
            templateRow = tr
            Exit For
        End If
    Next tr
    Dim templateRows As Range
    If templateRow > 0 Then Set templateRows = sh.Rows(templateRow & ":" & (templateRow + 1))

    Dim r As Long: r = MD_HEADER_ROW + 1
    Do While r <= lastRowScan
        Dim rmCode As String: rmCode = Trim(CStr(sh.Cells(r, 1).Value))
        If Len(rmCode) = 0 Then
            r = r + 1
        Else
            Dim headerRow As Long: headerRow = r
            Dim existing As Object: Set existing = CreateObject("Scripting.Dictionary")
            existing.CompareMode = vbTextCompare
            Dim rr As Long: rr = headerRow + 1
            Dim sumRow As Long: sumRow = 0
            Do While rr <= lastRowScan
                Dim lbl As String: lbl = Trim(CStr(sh.Cells(rr, 2).Value))
                If lbl = "Total Usage (kg)/week" Then
                    sumRow = rr
                    Exit Do
                End If
                If Trim(CStr(sh.Cells(rr, 3).Value)) = "No. of batches" Then existing(lbl) = True
                rr = rr + 1
            Loop
            If sumRow > 0 And byMat.Exists(rmCode) Then
                Dim allInter As Object: Set allInter = byMat(rmCode)
                Dim k As Variant, insertAt As Long: insertAt = sumRow
                For Each k In allInter.Keys
                    If Not existing.Exists(CStr(k)) Then
                        Call InsertIntermediateRowPair(sh, insertAt, headerRow, CStr(k), nWeeks, templateRows)
                        addedPairs = addedPairs + 1
                        insertAt = insertAt + 2
                        lastRowScan = lastRowScan + 2
                    End If
                Next k
                r = insertAt  ' insertAt is the new position of what was originally sumRow (the total-usage row).
                               ' Column A is blank there, so the outer loop's blank-row skip carries on to the next block as-is
            Else
                If sumRow > 0 Then r = sumRow + 1 Else r = headerRow + 1
            End If
        End If
    Loop
End Sub

' Inserts, at the specified position (insertAtRow) on Material_Detail, the
' 2 breakdown rows for one intermediate (No. of batches / Usage (kg)).
' Formatting is duplicated from templateRows (an existing breakdown row pair).
Private Sub InsertIntermediateRowPair(sh As Worksheet, insertAtRow As Long, headerRow As Long, _
        interName As String, nWeeks As Long, templateRows As Range)
    sh.Rows(insertAtRow & ":" & (insertAtRow + 1)).Insert Shift:=xlDown
    Dim batchesRow As Long: batchesRow = insertAtRow
    Dim usageRow As Long: usageRow = insertAtRow + 1
    Dim helperCol As Long: helperCol = MD_WEEK_START_COL + nWeeks
    Dim helperColLetter As String: helperColLetter = ColLetter(helperCol)

    sh.Cells(batchesRow, 2).Value = interName
    sh.Cells(batchesRow, 3).Value = "No. of batches"
    sh.Cells(batchesRow, helperCol).Value = "=IFERROR(MATCH($B" & batchesRow & ",Production_Plan[Intermediate],0),99999)"
    sh.Cells(batchesRow, helperCol).Font.Size = 8
    sh.Cells(batchesRow, helperCol).Font.Color = RGB(128, 128, 128)

    sh.Cells(usageRow, 2).Value = "Usage (kg)"
    sh.Cells(usageRow, 3).Value = "=SUMIFS(M_BOM[RM_Qty_Per_Batch],M_BOM[Intermediate],$B" & batchesRow & _
        ",M_BOM[Part Name],$A" & headerRow & ")"

    Dim w As Long, col As Long, wc As String
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        wc = ColLetter(col)
        sh.Cells(batchesRow, col).Value = _
            "=IFERROR(INDEX(Production_Plan[#Data],$" & helperColLetter & batchesRow & "," & (w + 1) & "),0)"
        sh.Cells(usageRow, col).Value = "=$C" & usageRow & "*" & wc & batchesRow
    Next w

    If Not templateRows Is Nothing Then
        On Error Resume Next
        templateRows.Copy
        sh.Rows(batchesRow & ":" & usageRow).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        On Error GoTo 0
    End If
End Sub
