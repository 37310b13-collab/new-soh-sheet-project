Attribute VB_Name = "RefreshData_MaterialMgmt"
Option Explicit

' ============================================================================
' RefreshData_MaterialMgmt module
'
' AddMaterial / RemoveMaterial / RemoveIntermediate
'
' Macros for adding/removing materials (raw materials) and intermediates
' (produced product codes) using only Excel (VBA), without Python.
'
'   AddMaterial : Adds a new material (a TTAF-supplied item) to the
'                 system. Entering the Part Name (RM_Code), Description,
'                 Supplier, Category, and TTAF_Code in sequence via
'                 InputBox adds the necessary rows in bulk to the bottom
'                 of M_RawMaterials, WeeklyConsumption, Incoming,
'                 Stock, TheoreticalStock, T_OpeningStock,
'                 CSAstock, TTAFstock, Dashboard, Material_Detail,
'                 and the matching PO_Draft_* sheet. Right after adding,
'                 since there's no usage history in M_BOM yet, its block
'                 on Material_Detail is a mini block with no intermediate
'                 breakdown (totals row only). After running RefreshBOM
'                 (the RefreshData_BOM module), once an intermediate that
'                 actually uses this material is found, the breakdown rows
'                 are added automatically too. Adds 2 rows per material to
'                 Dashboard, "Theoretical Stock" and "Actual Stock" (see
'                 section 5.6).
'   RemoveMaterial : Removes a material no longer in use from the system.
'                 Entering the Part Name (RM_Code) via InputBox deletes the
'                 matching row from every sheet AddMaterial adds to. Data
'                 in T_Shipments, CSAstock_Log/
'                 TTAFstock_Log, and M_BOM is not deleted (kept as
'                 history - adding the same Part Name again with
'                 AddMaterial automatically reconnects it).
'   RemoveIntermediate : Removes a discontinued intermediate (a finished
'                 product code) from the system. Entering the intermediate
'                 name via InputBox deletes the matching rows from Production_Plan
'                 and M_BOM, and the matching breakdown rows on
'                 Material_Detail (No. of batches / Usage (kg), from every
'                 material block that uses this intermediate). Raw
'                 material-side data (T_Shipments, T_OpeningStock,
'                 actuals logs, M_RawMaterials) is not
'                 deleted.
'                 (Note) The "add" side for intermediates is already
'                 automated. RefreshWeeklyBatches/RefreshBOM automatically
'                 add rows to Production_Plan/M_BOM/Material_Detail every time they
'                 find a new intermediate in the production plan/usage
'                 rate table, so a dedicated "AddIntermediate" macro isn't
'                 needed.
'
' [Where AddMaterial inserts] Rather than inserting in the middle of
' existing rows, it always adds at the very bottom of each sheet, since
' inserting in the middle carries a much larger risk of shifting existing
' rows.
'
' [A note on safety (RemoveIntermediate)] Deleting rows from Production_Plan/M_BOM
' does not affect the calculations of other intermediates or other
' materials. The reason: every formula on WeeklyConsumption/Material_Detail
' that references Production_Plan/M_BOM is built with MATCH/structured references
' (the TableName[ColumnName] form), never a hard-coded row number (MATCH
' recalculates the new row position every time after a deletion, and
' structured references automatically follow the table's row count after
' deletion).
'
' [Important caution] None of these operations can be undone (Ctrl+Z may
' not undo them). We strongly recommend backing up (copying) the file
' before running any of them.
'
' For the overall design rationale, see the comment at the top of the
' RefreshData_Utilities module.
'
' [About material ordering] M_RawMaterials, Dashboard, Material_Detail,
' etc. are ordered as Substrate -> other Chemical -> Ester Film/PP Film ->
' TPZ-family (the same rule as build_soh.py's rm_master.sort()).
' AddMaterial doesn't always add at the very end - it inserts at the end
' of the group the new material belongs to, preserving this order (see
' ClassifyGroup/FindGroupInsertPosition/FindRowByCode).
'
' SyncPODraftCategories : When M_RawMaterials' Category/Origin_County is
'                 changed later, re-syncs the PO_Draft_* sheet assignment
'                 to match the actual values (see the comment right above
'                 each Sub for details).
'
' Note: one-time migration macros (FixOpeningStockColumnReference,
' SetupOrderManagementMigration, SetupSubstratePODraftByCountry, etc.)
' have been removed from this module once their respective migrations
' were completed. See docs/SOH_System_Guide.md and the git change history
' for what was done.
' ============================================================================

' Determines, from rmCode/descVal/categoryVal, which group a material
' belongs to (0=Substrate, 1=other Chemical, 2=Ester Film/PP Film,
' 3=TPZ-family). Follows the exact same rule as build_soh.py's
' _rm_sort_group() (match the priority order of the checks exactly too:
' Ester Film/PP Film always take priority as group 2 regardless of
' Category, and the TPZ-family always takes priority as group 3
' regardless of Category).
Private Function ClassifyGroup(rmCode As String, descVal As String, categoryVal As String) As Long
    Dim codeUpper As String: codeUpper = UCase(Trim(rmCode))
    Dim descUpper As String: descUpper = UCase(Trim(descVal))
    If codeUpper = "ESTER FILM" Or codeUpper = "PP FILM" Then
        ClassifyGroup = 2
    ElseIf InStr(descUpper, "TPZ") > 0 Or InStr(descUpper, "TZP") > 0 Then
        ClassifyGroup = 3
    ElseIf Trim(categoryVal) = "Substrate" Then
        ClassifyGroup = 0
    Else
        ClassifyGroup = 1
    End If
End Function

' Returns the position (a 1-based ListRows position) within M_RawMaterials
' where a material in targetGroup should be inserted. Inserting "right
' after the last row of the same group" preserves the existing order
' (Substrate -> other Chemical -> Ester Film/PP Film -> TPZ-family). If
' that group has no rows at all, returns the position right before the
' first row of the next group after it. If it should be added at the very
' end overall, returns ListRows.Count+1 (the caller determines "added at
' the end" vs. "inserted in the middle" by whether this value equals
' Count+1).
Private Function FindGroupInsertPosition(rmTbl As ListObject, targetGroup As Long) As Long
    Dim n As Long: n = rmTbl.ListRows.Count
    If n = 0 Then
        FindGroupInsertPosition = 1
        Exit Function
    End If
    Dim data As Variant
    data = rmTbl.ListColumns(1).DataBodyRange.Resize(n, 4).Value  ' Part Name, Description, Supplier, Category
    Dim i As Long, lastLE As Long: lastLE = 0
    For i = 1 To n
        If ClassifyGroup(CStr(data(i, 1)), CStr(data(i, 2)), CStr(data(i, 4))) <= targetGroup Then
            lastLE = i
        End If
    Next i
    FindGroupInsertPosition = lastLE + 1
End Function

' Searches column colIdx of a sheet, from startRow through the last used
' row (by that column), and returns the row number of the first row whose
' value exactly matches rmCode (compared with Trim), or 0 if not found.
Private Function FindRowByCode(sh As Worksheet, colIdx As Long, rmCode As String, startRow As Long) As Long
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, colIdx).End(xlUp).Row
    If lastRow < startRow Then
        FindRowByCode = 0
        Exit Function
    End If
    Dim r As Long
    For r = startRow To lastRow
        If Trim(CStr(sh.Cells(r, colIdx).Value)) = rmCode Then
            FindRowByCode = r
            Exit Function
        End If
    Next r
    FindRowByCode = 0
End Function

Sub AddMaterial()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    Dim rmCode As String
    rmCode = Trim(InputBox("Please enter the material code (Part Name) to add." & vbCrLf & "(e.g. CHEM-9999)", "Add Material"))
    If Len(rmCode) = 0 Then Exit Sub

    If Not IsMaterialCodeFree(rmTbl, rmCode) Then
        MsgBox "That material code is already registered: " & rmCode, vbExclamation
        Exit Sub
    End If

    Dim descVal As String
    descVal = Trim(InputBox("Please enter the material name (Description).", "Add Material"))
    If Len(descVal) = 0 Then Exit Sub

    Dim supplierVal As String
    supplierVal = Trim(InputBox("Please enter the supplier.", "Add Material", "TTAF"))

    Dim categoryVal As String
    categoryVal = Trim(InputBox("Please enter the category." & vbCrLf & _
        "Enter exactly one of: Chemical / Hazardous Chemical / Substrate.", _
        "Add Material", "Chemical"))
    If categoryVal <> "Chemical" And categoryVal <> "Hazardous Chemical" And categoryVal <> "Substrate" Then
        MsgBox "Please enter the category as one of: Chemical / Hazardous Chemical / Substrate." & vbCrLf & _
               "You entered: " & categoryVal, vbExclamation
        Exit Sub
    End If

    Dim ttafCodeVal As String
    ttafCodeVal = Trim(InputBox("Please enter the TTAF_Code (the part number on the TTAF side). Leave it blank if you don't know it.", "Add Material"))

    ' Origin_Country: for Substrate, the order form (PO_Draft) sheet
    ' depends on the country of origin (Japan/China/Poland), so this is
    ' only asked for Substrate (irrelevant for Chemical/Hazardous
    ' Chemical). Anything other than Japan/China/Poland (blank or
    ' anything else) won't appear on any PO_Draft_* sheet (there is no
    ' catch-all sheet by design - please re-enter the correct country of
    ' origin when it comes time to order).
    Dim originCountryVal As String: originCountryVal = ""
    If categoryVal = "Substrate" Then
        originCountryVal = Trim(InputBox("Please enter the country of origin (Origin_Country) as" & vbCrLf & _
            "one of ""Japan"", ""China"", or ""Poland"" (case-insensitive)." & vbCrLf & _
            "Anything other than these 3 (blank included) will not appear on any order-form (PO_Draft) sheet." & vbCrLf & _
            "If you're not ordering yet or don't know the country of origin, it's fine to leave this blank.", "Add Material"))
    End If

    If MsgBox("The material will be added with the following details." & vbCrLf & vbCrLf & _
              "Material code: " & rmCode & vbCrLf & "Material name: " & descVal & vbCrLf & _
              "Supplier: " & supplierVal & vbCrLf & "Category: " & categoryVal & vbCrLf & _
              "TTAF_Code: " & ttafCodeVal & vbCrLf & _
              IIf(categoryVal = "Substrate", "Country of origin: " & originCountryVal & vbCrLf, "") & vbCrLf & _
              "Is this OK?", vbYesNo + vbQuestion, "Confirm Add Material") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim nWeeks As Long
    nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count

    ' ---- Determine the insert position ----
    ' Materials are kept ordered as "Substrate -> other Chemical -> Ester
    ' Film/PP Film -> TPZ-family" (the same rule as build_soh.py's
    ' rm_master.sort()). Rather than adding at the very end, inserting
    ' right after the last row of the same group preserves this order for
    ' future additions too. anchorRmCode represents "what material was at
    ' the position the new row will be inserted into, before the insert"
    ' (an empty string means "add at the end of the group = the end of the
    ' whole workbook"). This must be read before actually adding the row to
    ' M_RawMaterials (after adding, that position would be the new row).
    Dim targetGroup As Long: targetGroup = ClassifyGroup(rmCode, descVal, categoryVal)
    Dim insertPos As Long: insertPos = FindGroupInsertPosition(rmTbl, targetGroup)
    Dim anchorRmCode As String: anchorRmCode = ""
    If insertPos <= rmTbl.ListRows.Count Then
        anchorRmCode = Trim(CStr(rmTbl.ListRows(insertPos).Range.Cells(1, 1).Value))
    End If

    ' ---- M_RawMaterials ----
    Dim newRmRow As ListRow: Set newRmRow = rmTbl.ListRows.Add(Position:=insertPos)
    newRmRow.Range.Cells(1, 1).Value = rmCode
    newRmRow.Range.Cells(1, 2).Value = descVal
    newRmRow.Range.Cells(1, 3).Value = supplierVal
    newRmRow.Range.Cells(1, 4).Value = categoryVal
    newRmRow.Range.Cells(1, 5).Value = "kg"
    newRmRow.Range.Cells(1, 6).Value = 0
    newRmRow.Range.Cells(1, 7).Value = 0
    newRmRow.Range.Cells(1, 8).Value = 4
    newRmRow.Range.Cells(1, 9).Value = ttafCodeVal
    newRmRow.Range.Cells(1, 10).Value = 0  ' FixedWeeklyConsumption. Edit it directly in Excel after adding if needed
    newRmRow.Range.Cells(1, 11).Value = originCountryVal

    ' ---- WeeklyConsumption / Incoming / Stock / TheoreticalStock / T_OpeningStock ----
    Dim reqTbl As ListObject: Set reqTbl = thisWb.Sheets("WeeklyConsumption").ListObjects("WeeklyConsumption")
    Dim inTbl As ListObject: Set inTbl = thisWb.Sheets("Incoming").ListObjects("Incoming")
    Dim stTbl As ListObject: Set stTbl = thisWb.Sheets("Stock").ListObjects("Stock")
    Dim theoTbl As ListObject: Set theoTbl = thisWb.Sheets("TheoreticalStock").ListObjects("TheoreticalStock")
    Dim osTbl As ListObject: Set osTbl = thisWb.Sheets("T_OpeningStock").ListObjects("T_OpeningStock")

    ' WeeklyConsumption/Incoming/Stock/TheoreticalStock must correspond
    ' exactly, row for row, with M_RawMaterials (since grow = the shared
    ' row number they reference directly). Insert with the same insertPos.
    ' T_OpeningStock uses MATCH lookups (order-independent), so it's fine
    ' to keep always adding it at the end.
    Dim reqRow As ListRow: Set reqRow = reqTbl.ListRows.Add(Position:=insertPos)
    Dim inRow As ListRow: Set inRow = inTbl.ListRows.Add(Position:=insertPos)
    Dim stRow As ListRow: Set stRow = stTbl.ListRows.Add(Position:=insertPos)
    Dim theoRow As ListRow: Set theoRow = theoTbl.ListRows.Add(Position:=insertPos)
    Dim osRow As ListRow: Set osRow = osTbl.ListRows.Add

    Dim grow As Long: grow = reqRow.Range.Row  ' the actual sheet row number for WeeklyConsumption/Incoming/Stock/TheoreticalStock (same across all 4 tables)

    Dim descFormula As String
    descFormula = "=IFERROR(INDEX(M_RawMaterials[Description],MATCH($A" & grow & ",M_RawMaterials[Part Name],0)),"""")"
    reqRow.Range.Cells(1, 1).Value = rmCode
    reqRow.Range.Cells(1, 2).Value = descFormula
    inRow.Range.Cells(1, 1).Value = rmCode
    inRow.Range.Cells(1, 2).Value = descFormula
    stRow.Range.Cells(1, 1).Value = rmCode
    stRow.Range.Cells(1, 2).Value = descFormula
    theoRow.Range.Cells(1, 1).Value = rmCode
    theoRow.Range.Cells(1, 2).Value = descFormula
    osRow.Range.Cells(1, 1).Value = rmCode
    osRow.Range.Cells(1, 2).Value = 0
    osRow.Range.Cells(1, 3).Value = Date

    Dim w As Long, col As Long, cl As String
    For w = 1 To nWeeks
        col = 2 + w
        ' In addition to BOM-based consumption, simply adds
        ' M_RawMaterials' FixedWeeklyConsumption (the same formula as
        ' build_soh.py's WeeklyConsumption generation logic).
        reqRow.Range.Cells(1, col).Value = _
            "=SUMPRODUCT((M_BOM[Part Name]=$A" & grow & ")*M_BOM[RM_Qty_Per_Batch]*" & _
            "IFERROR(INDEX(Production_Plan[#Data],M_BOM[PPGridRow]," & (w + 1) & "),0))" & _
            "+IFERROR(INDEX(M_RawMaterials[FixedWeeklyConsumption],MATCH($A" & grow & ",M_RawMaterials[Part Name],0)),0)"
        inRow.Range.Cells(1, col).Value = _
            "=SUMIFS(T_Shipments[Confirmed_Qty],T_Shipments[Part Name],$A" & grow & _
            ",T_Shipments[Effective_Week]," & w & ")"
    Next w

    ' ---- CSAstock / TTAFstock (a bordered grid, not a Table. Insert
    ' right before anchorRmCode's row so the order matches M_RawMaterials) ----
    Dim ssRowSelf As Long, ssRowTTAF As Long
    ssRowSelf = InsertOrAppendStockGridRow(thisWb.Sheets("CSAstock"), rmCode, nWeeks, "CSAstock_Log", "Self_Qty", anchorRmCode)
    ssRowTTAF = InsertOrAppendStockGridRow(thisWb.Sheets("TTAFstock"), rmCode, nWeeks, "TTAFstock_Log", "TTAF_Qty", anchorRmCode)
    If ssRowSelf <> ssRowTTAF Then
        MsgBox "Warning: CSAstock and TTAFstock's row numbers ended up different (" & ssRowSelf & " / " & ssRowTTAF & ")." & vbCrLf & _
               "Please check manually. Processing will continue.", vbExclamation
    End If
    Dim ssRow As Long: ssRow = ssRowSelf

    ' Stock's formula (priority order: sum of self+TTAF actuals >
    ' normal roll-forward)
    Dim priorExpr As String, normalExpr As String
    Dim hasSelf As String, selfVal As String, hasTTAF As String, ttafVal As String
    For w = 1 To nWeeks
        col = 2 + w
        cl = ColLetter(col)
        hasSelf = "('CSAstock'!" & cl & ssRow & "<>"""")"
        selfVal = "'CSAstock'!" & cl & ssRow
        hasTTAF = "('TTAFstock'!" & cl & ssRow & "<>"""")"
        ttafVal = "'TTAFstock'!" & cl & ssRow
        If w = 1 Then
            priorExpr = "IFERROR(INDEX(T_OpeningStock[OpeningQty],MATCH($A" & grow & ",T_OpeningStock[Part Name],0)),0)"
        Else
            priorExpr = ColLetter(col - 1) & grow
        End If
        normalExpr = priorExpr & "+'Incoming'!" & cl & grow & "-'WeeklyConsumption'!" & cl & grow
        stRow.Range.Cells(1, col).Value = _
            "=IF((" & hasSelf & ")*(" & hasTTAF & ")>0," & selfVal & "+" & ttafVal & "," & normalExpr & ")"

        ' TheoreticalStock: never looks at
        ' self/TTAF actuals - a pure "previous week + incoming -
        ' consumption" roll-forward, except that only in the first week a
        ' new month begins does it re-sync from the previous week's actual
        ' stock (Stock) (the same formula as
        ' FixTheoreticalStockMonthlyReset/build_soh.py - if this were left
        ' as the old "never reset" formula, a material added via
        ' AddMaterial would behave inconsistently with every other
        ' material, so this formula must always be used).
        Dim theoPriorExpr As String
        If w = 1 Then
            theoPriorExpr = "IFERROR(INDEX(T_OpeningStock[OpeningQty],MATCH($A" & grow & ",T_OpeningStock[Part Name],0)),0)"
        Else
            Dim monthChangedExpr As String
            monthChangedExpr = "INDEX(Cal_Weeks[MonthYearLabel]," & w & ")<>INDEX(Cal_Weeks[MonthYearLabel]," & (w - 1) & ")"
            theoPriorExpr = "IF(" & monthChangedExpr & ",'Stock'!" & ColLetter(col - 1) & grow & "," & ColLetter(col - 1) & grow & ")"
        End If
        theoRow.Range.Cells(1, col).Value = _
            "=" & theoPriorExpr & "+'Incoming'!" & cl & grow & "-'WeeklyConsumption'!" & cl & grow
    Next w

    ' ---- Dashboard (a bordered grid, not a Table. Insert right before anchorRmCode's row) ----
    Call AppendDashboardRow(thisWb.Sheets("Dashboard"), rmCode, nWeeks, ssRow, grow, anchorRmCode)

    ' ---- Material_Detail (insert the material's block right before
    ' anchorRmCode's block. No intermediate rows yet since it's not
    ' registered in the BOM) ----
    Call AppendMaterialDetailBlock(thisWb.Sheets("Material_Detail"), rmCode, descVal, nWeeks, ssRow, grow, anchorRmCode)

    ' ---- PO_Draft_{Category} (for Substrate, further branches on Origin_Country) ----
    Dim poSheetName As String: poSheetName = POSheetNameForMaterial(categoryVal, originCountryVal)
    If Len(poSheetName) > 0 Then
        Call AppendPODraftRow(thisWb.Sheets(poSheetName), rmCode, ttafCodeVal, nWeeks)
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Added material """ & rmCode & """." & vbCrLf & vbCrLf & _
           "Since this material isn't registered in M_BOM (usage rates) yet, its" & vbCrLf & _
           "Material_Detail block currently shows ""no intermediates used."" Running RefreshBOM" & vbCrLf & _
           "will automatically reflect it as soon as a combination with an intermediate that" & vbCrLf & _
           "actually uses it is found.", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred while adding the material: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "Some changes may have already been applied. Please check the sheets' state" & vbCrLf & _
           "(if you're unsure, close the file without saving and reopen it to return to the last saved state).", vbCritical
End Sub

Sub RemoveMaterial()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    Dim rmCode As String
    rmCode = Trim(InputBox("Please enter the material code (Part Name) to remove.", "Remove Material"))
    If Len(rmCode) = 0 Then Exit Sub

    Dim rmFoundRow As Long: rmFoundRow = FindMaterialRow(rmTbl, rmCode)
    If rmFoundRow = 0 Then
        MsgBox "Not found in M_RawMaterials: " & rmCode, vbExclamation
        Exit Sub
    End If
    ' Align to the actual registered spelling (case). Skipping Trim() here
    ' would let stray whitespace mixed into the M_RawMaterials cell carry
    ' over into rmCode, causing mismatches later in
    ' DeleteMatchingGridRow etc. (whose target cells are Trim()'d, but
    ' which assume rmCode itself was already cleaned up by the caller),
    ' making deletion fail on just some sheets (this actually happened as
    ' a reported bug).
    rmCode = Trim(CStr(rmTbl.ListRows(rmFoundRow).Range.Cells(1, 1).Value))
    Dim categoryVal As String: categoryVal = CStr(rmTbl.ListRows(rmFoundRow).Range.Cells(1, 4).Value)
    Dim originCountryVal As String: originCountryVal = CStr(rmTbl.ListRows(rmFoundRow).Range.Cells(1, 11).Value)

    If MsgBox("This will remove material """ & rmCode & """." & vbCrLf & _
              "The matching row will be deleted from every related sheet (M_RawMaterials," & vbCrLf & _
              "WeeklyConsumption, Incoming, Stock, TheoreticalStock," & vbCrLf & _
              "T_OpeningStock, CSAstock, TTAFstock, Dashboard, Material_Detail," & vbCrLf & _
              "and the matching PO_Draft). This cannot be undone." & vbCrLf & vbCrLf & _
              "(Past data for this material remaining in T_Shipments," & vbCrLf & _
              "the actuals logs, and M_BOM will not be deleted)" & vbCrLf & vbCrLf & "Are you sure?", _
              vbYesNo + vbExclamation, "Confirm Remove Material") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call DeleteMatchingTableRow(thisWb.Sheets("WeeklyConsumption").ListObjects("WeeklyConsumption"), rmCode)
    Call DeleteMatchingTableRow(thisWb.Sheets("Incoming").ListObjects("Incoming"), rmCode)
    Call DeleteMatchingTableRow(thisWb.Sheets("Stock").ListObjects("Stock"), rmCode)
    Call DeleteMatchingTableRow(thisWb.Sheets("TheoreticalStock").ListObjects("TheoreticalStock"), rmCode)
    Call DeleteMatchingTableRow(thisWb.Sheets("T_OpeningStock").ListObjects("T_OpeningStock"), rmCode)

    Call DeleteMatchingGridRow(thisWb.Sheets("CSAstock"), rmCode, 1)
    Call DeleteMatchingGridRow(thisWb.Sheets("TTAFstock"), rmCode, 1)
    Call DeleteMatchingGridRow(thisWb.Sheets("Dashboard"), rmCode, 1)

    Dim poSheetName As String: poSheetName = POSheetNameForMaterial(categoryVal, originCountryVal)
    If Len(poSheetName) > 0 Then
        Call DeleteMatchingGridRow(thisWb.Sheets(poSheetName), rmCode, 4)
    End If

    Call DeleteMaterialDetailBlock(thisWb.Sheets("Material_Detail"), rmCode)

    ' M_RawMaterials itself is deleted last, only after it's done being used as the search key for the other sheets
    rmTbl.ListRows(rmFoundRow).Delete

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Removed material """ & rmCode & """." & vbCrLf & vbCrLf & _
           "(Past data for this material remaining in T_Shipments," & vbCrLf & _
           "the actuals logs, and M_BOM has not been deleted. Delete it manually if needed.)", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred while removing the material: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "Some deletion may have already been applied. Please check the sheets' state" & vbCrLf & _
           "(if you're unsure, close the file without saving and reopen it to return to the last saved state).", vbCritical
End Sub

Sub RemoveIntermediate()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim ppGrid As ListObject: Set ppGrid = thisWb.Sheets("Production_Plan").ListObjects("Production_Plan")
    Dim bomTbl As ListObject: Set bomTbl = thisWb.Sheets("M_BOM").ListObjects("M_BOM")

    Dim interName As String
    interName = Trim(InputBox("Please enter the intermediate name to remove." & vbCrLf & _
        "It should match the name shown in the ""Intermediate"" column on the Production_Plan sheet" & vbCrLf & _
        "(case-insensitive).", "Remove Intermediate"))
    If Len(interName) = 0 Then Exit Sub

    Dim ppFoundRow As Long: ppFoundRow = FindMaterialRow(ppGrid, interName)
    If ppFoundRow = 0 Then
        MsgBox "No intermediate named """ & interName & """ was found in Production_Plan." & vbCrLf & _
               "Please check the exact name on the Production_Plan sheet.", vbExclamation
        Exit Sub
    End If
    ' Align to the actual registered spelling (case) (to absorb variation
    ' in how it was typed). Trim() is applied for the same reason as
    ' RemoveMaterial (to prevent stray whitespace mixed into the cell from
    ' causing a mismatch during the deletion steps that follow).
    Dim canonicalName As String
    canonicalName = Trim(CStr(ppGrid.ListRows(ppFoundRow).Range.Cells(1, 1).Value))

    Dim bomCount As Long: bomCount = CountMatchingTableRows(bomTbl, canonicalName)

    If MsgBox("This will remove intermediate """ & canonicalName & """." & vbCrLf & vbCrLf & _
              "What will be deleted:" & vbCrLf & _
              "- Production_Plan: the matching row (weekly batch counts), 1 row" & vbCrLf & _
              "- M_BOM: usage-rate rows for this intermediate, " & bomCount & " row(s)" & vbCrLf & _
              "- Material_Detail: this intermediate's breakdown rows (No. of batches / Usage (kg))," & vbCrLf & _
              "  from every material block that uses it" & vbCrLf & vbCrLf & _
              "Raw material-side data (T_Shipments, T_OpeningStock, actuals logs," & vbCrLf & _
              "M_RawMaterials, etc.) is never deleted." & vbCrLf & vbCrLf & _
              "This cannot be undone. Are you sure?", _
              vbYesNo + vbExclamation, "Confirm Remove Intermediate") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim removedDetailPairs As Long
    removedDetailPairs = DeleteIntermediateFromMaterialDetail(thisWb.Sheets("Material_Detail"), canonicalName)
    Call DeleteMatchingTableRow(bomTbl, canonicalName)
    ' Production_Plan itself is deleted last, only after it's done being used as the search key for FindMaterialRow above
    Call DeleteMatchingTableRow(ppGrid, canonicalName)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Removed intermediate """ & canonicalName & """." & vbCrLf & _
           "Breakdown row pairs removed from Material_Detail: " & removedDetailPairs & " (material block count)" & vbCrLf & vbCrLf & _
           "(T_Shipments, T_OpeningStock, actuals logs, and" & vbCrLf & _
           "M_RawMaterials have not been deleted)", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred while removing the intermediate: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "Some deletion may have already been applied. Please check the sheets' state" & vbCrLf & _
           "(if you're unsure, close the file without saving and reopen it to return to the last saved state).", vbCritical
End Sub

Private Function IsMaterialCodeFree(rmTbl As ListObject, rmCode As String) As Boolean
    IsMaterialCodeFree = True
    Dim n As Long: n = rmTbl.ListRows.Count
    Dim i As Long
    For i = 1 To n
        If UCase(Trim(CStr(rmTbl.ListRows(i).Range.Cells(1, 1).Value))) = UCase(Trim(rmCode)) Then
            IsMaterialCodeFree = False
            Exit Function
        End If
    Next i
End Function

Private Function FindMaterialRow(rmTbl As ListObject, rmCode As String) As Long
    FindMaterialRow = 0
    Dim n As Long: n = rmTbl.ListRows.Count
    Dim i As Long
    For i = 1 To n
        If UCase(Trim(CStr(rmTbl.ListRows(i).Range.Cells(1, 1).Value))) = UCase(Trim(rmCode)) Then
            FindMaterialRow = i
            Exit Function
        End If
    Next i
End Function

' Returns the matching PO_Draft_* sheet name from Category (and, for
' Substrate, Origin_Country). For Substrate, an Origin_Country of Poland
' routes to its own dedicated sheet (PO_Draft_Substrate_Poland), while
' Japan/China are routed together into a single shared sheet
' (PO_Draft_Substrate_JPN_CHN) (there is no catch-all sheet by design -
' the user's explicit decision: items that haven't been ordered yet or
' whose country of origin isn't confirmed don't need to appear on any
' PO_Draft. Entering the Origin_Country and running SyncPODraftCategories
' when it's time to order reflects it). The same split as
' build_soh.py's build_po_draft calls.
' Origin_Country is matched case-insensitively (to avoid a hand-typed
' spelling variant, e.g. "poland", silently failing to match and slipping
' off of every PO_Draft).
Public Function POSheetNameForMaterial(categoryVal As String, originCountryVal As String) As String
    Dim originKey As String: originKey = UCase(Trim(originCountryVal))
    Select Case categoryVal
        Case "Chemical": POSheetNameForMaterial = "PO_Draft_Chemical"
        Case "Hazardous Chemical": POSheetNameForMaterial = "PO_Draft_Hazardous"
        Case "Substrate"
            Select Case originKey
                Case "POLAND": POSheetNameForMaterial = "PO_Draft_Substrate_Poland"
                Case "JAPAN", "CHINA": POSheetNameForMaterial = "PO_Draft_Substrate_JPN_CHN"
                Case Else: POSheetNameForMaterial = ""
            End Select
        Case Else: POSheetNameForMaterial = ""
    End Select
End Function

' [Maintenance macro - safe to run anytime] Each row on
' PO_Draft_Chemical/_Hazardous/_Substrate_* was added to the matching
' sheet just once, based on M_RawMaterials[Category] (and, for Substrate,
' [Origin_Country]) as of the moment AddMaterial ran (or as of when
' build_soh.py ran) - it's not a mechanism that re-sorts itself
' automatically by reading a live formula. So if M_RawMaterials' Category
' or Origin_Country is changed afterward, the row on the PO_Draft_* side
' stays behind on the old sheet and doesn't move on its own (this bug was
' actually reported and found in practice). Running this macro treats
' M_RawMaterials' current Category/Origin_Country as authoritative,
' rescans every material across all PO_Draft_* sheets, and moves any
' inconsistent row (one left on the wrong sheet, or one not yet on any
' PO_Draft_*) to the correct sheet by deleting and re-creating it. A
' Substrate item whose country of origin is anything other than
' Japan/China/Poland (blank or otherwise) is intentionally left off every
' PO_Draft_* (there's no catch-all sheet by design - enter the
' Origin_Country when it's time to order, then run this macro). The order
' quantity itself (a reference to Material_Detail) doesn't depend on row
' position, so moving a row never loses order information. Rows already in
' the correct place are left completely untouched.
Sub SyncPODraftCategories()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")
    Dim nWeeks As Long: nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim poSheetNames As Variant
    poSheetNames = Array("PO_Draft_Chemical", "PO_Draft_Hazardous", _
        "PO_Draft_Substrate_Poland", "PO_Draft_Substrate_JPN_CHN")
    Const PO_HDR_TABLE_ROW As Long = 26  ' the same fixed value as build_soh.py's PO_HDR_UOM_FIRM_ROW (last header row - data starts right below it)

    ' Collect the material codes (column D) currently present on each PO_Draft_* sheet, tagged with "which sheet they're on".
    Dim currentSheetOf As Object: Set currentSheetOf = CreateObject("Scripting.Dictionary")
    currentSheetOf.CompareMode = vbTextCompare
    Dim si As Long
    For si = LBound(poSheetNames) To UBound(poSheetNames)
        Dim shName As String: shName = CStr(poSheetNames(si))
        Dim sh As Worksheet: Set sh = thisWb.Sheets(shName)
        Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 4).End(xlUp).Row
        Dim r As Long
        For r = PO_HDR_TABLE_ROW + 1 To lastRow
            Dim codeInSheet As String: codeInSheet = Trim(CStr(sh.Cells(r, 4).Value))
            If Len(codeInSheet) > 0 And Not currentSheetOf.Exists(codeInSheet) Then
                currentSheetOf(codeInSheet) = shName
            End If
        Next r
    Next si

    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    Dim movedCount As Long: movedCount = 0
    Dim addedCount As Long: addedCount = 0
    If rmN > 0 Then
        Dim rmData As Variant: rmData = rmTbl.ListColumns(1).DataBodyRange.Resize(rmN, 11).Value  ' Part Name,Description,Supplier,Category,UOM,...,TTAF_Code(col 9),FixedWeeklyConsumption(col 10),Origin_Country(col 11)
        Dim i As Long
        For i = 1 To rmN
            Dim rmCode As String: rmCode = Trim(CStr(rmData(i, 1)))
            Dim categoryVal As String: categoryVal = Trim(CStr(rmData(i, 4)))
            Dim ttafCodeVal As String: ttafCodeVal = Trim(CStr(rmData(i, 9)))
            Dim originCountryVal As String: originCountryVal = Trim(CStr(rmData(i, 11)))
            Dim correctSheet As String: correctSheet = POSheetNameForMaterial(categoryVal, originCountryVal)

            Dim existingSheet As String: existingSheet = ""
            If currentSheetOf.Exists(rmCode) Then existingSheet = CStr(currentSheetOf(rmCode))

            If existingSheet <> correctSheet Then
                If Len(existingSheet) > 0 Then
                    Call DeleteMatchingGridRow(thisWb.Sheets(existingSheet), rmCode, 4)
                    movedCount = movedCount + 1
                Else
                    addedCount = addedCount + 1
                End If
                If Len(correctSheet) > 0 Then
                    Call AppendPODraftRow(thisWb.Sheets(correctSheet), rmCode, ttafCodeVal, nWeeks)
                End If
            End If
        Next i
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Synced the PO_Draft_* sheet assignment to match M_RawMaterials' current Category." & vbCrLf & vbCrLf & _
           "Materials moved to a different sheet: " & movedCount & vbCrLf & _
           "Materials newly added (not on any PO_Draft before): " & addedCount, vbInformation
    Exit Sub

ErrHandler:
    Dim errNum5 As Long: errNum5 = Err.Number
    Dim errMsg5 As String: errMsg5 = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred during processing: (" & errNum5 & ") " & errMsg5, vbCritical
End Sub

' Adds a row for a new material to CSAstock/TTAFstock. Since these
' are a bordered grid rather than a Table, row addition is handled
' manually. If anchorRmCode (the code of the material that was at the same
' insert position on the M_RawMaterials side) is given, an Excel row
' insertion (Rows.Insert) is used to insert right before that material's
' row, preserving the same order as M_RawMaterials/Dashboard/etc.
' Rows.Insert automatically adjusts every other formula's row references
' throughout the workbook, so no existing row's formula needs to be
' rewritten. If anchorRmCode is an empty string (this material's group is
' the last group, with no material following it), it's added at the very
' bottom as before. Returns the row number actually added/inserted at.
Private Function InsertOrAppendStockGridRow(sh As Worksheet, rmCode As String, nWeeks As Long, logTableName As String, qtyColName As String, anchorRmCode As String) As Long
    Dim targetRow As Long, formatSrcRow As Long
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim anchorRow As Long: anchorRow = 0
    If Len(anchorRmCode) > 0 Then
        anchorRow = FindRowByCode(sh, 1, anchorRmCode, SS_TABLE_ROW + 1)
    End If
    If anchorRow > 0 Then
        sh.Rows(anchorRow).Insert Shift:=xlDown
        targetRow = anchorRow
        formatSrcRow = anchorRow + 1  ' the original anchor row, now shifted down 1 row by the insert - copy formatting from here
    Else
        targetRow = lastRow + 1
        formatSrcRow = lastRow
    End If

    sh.Cells(targetRow, 1).Value = rmCode
    sh.Cells(targetRow, 2).Value = _
        "=IFERROR(INDEX(M_RawMaterials[Description],MATCH($A" & targetRow & ",M_RawMaterials[Part Name],0)),"""")"
    Dim w As Long, col As Long
    For w = 1 To nWeeks
        col = 2 + w
        sh.Cells(targetRow, col).Value = _
            "=IF(COUNTIFS(" & logTableName & "[Part Name],$A" & targetRow & "," & logTableName & "[WeekIndex]," & w & ")=0,"""",SUMIFS(" & _
            logTableName & "[" & qtyColName & "]," & logTableName & "[Part Name],$A" & targetRow & "," & logTableName & "[WeekIndex]," & w & "))"
    Next w
    On Error Resume Next
    sh.Rows(formatSrcRow).Copy
    sh.Rows(targetRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0
    InsertOrAppendStockGridRow = targetRow
End Function

' Adds 2 rows for a new material to Dashboard (Theoretical Stock, then
' Actual Stock, in that order) (ssRow = the row number on the
' CSAstock/TTAFstock side, grow = the row number on the
' WeeklyConsumption/Incoming/Stock side). If anchorRmCode is given, an
' Excel row insertion is used to insert right before that material's row,
' preserving the same order as M_RawMaterials (an empty string means
' added at the very bottom as before).
' Since Part Name (column A) gets the same value written on both rows,
' RemoveMaterial's DeleteMatchingGridRow (which deletes by matching column
' A) removes both rows together (no special handling is needed on the
' deletion side). Column positions correspond to build_soh.py's
' LEFT_COLS/DASH_ROW_LABEL_COL/DASH_DIFF_COL: 1=Part Name,
' 2=Description, 3=Category, 4=SafetyStock_Min, 5=SafetyStock_Max,
' 6=Self Stock (Actual), 7=TTAF Stock (Actual), 8=Actual Week, 9=Row Type,
' 10=Variance (kg); the week data starts at column 11 (column K).
Private Sub AppendDashboardRow(sh As Worksheet, rmCode As String, nWeeks As Long, ssRow As Long, grow As Long, anchorRmCode As String)
    Dim theoRow As Long, actualRow As Long, formatSrcTheo As Long, formatSrcActual As Long
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim anchorRow As Long: anchorRow = 0
    If Len(anchorRmCode) > 0 Then
        anchorRow = FindRowByCode(sh, 1, anchorRmCode, DASH_DATA_START_ROW)
    End If
    If anchorRow > 0 Then
        sh.Rows(anchorRow & ":" & (anchorRow + 1)).Insert Shift:=xlDown
        theoRow = anchorRow
        actualRow = anchorRow + 1
        formatSrcTheo = anchorRow + 2   ' the original anchor pair's Theoretical Stock row, now shifted down 2 rows by the insert
        formatSrcActual = anchorRow + 3 ' likewise, the Actual Stock row
    Else
        theoRow = lastRow + 1
        actualRow = lastRow + 2
        formatSrcTheo = lastRow - 1
        formatSrcActual = lastRow
    End If
    Dim lastWeekCol As String: lastWeekCol = ColLetter(2 + nWeeks)

    Dim ssSelfRng As String: ssSelfRng = "'CSAstock'!$C$" & ssRow & ":$" & lastWeekCol & "$" & ssRow
    Dim ssTTAFRng As String: ssTTAFRng = "'TTAFstock'!$C$" & ssRow & ":$" & lastWeekCol & "$" & ssRow
    Dim ssLabelRng As String: ssLabelRng = "'CSAstock'!$C$" & SS_TABLE_ROW & ":$" & lastWeekCol & "$" & SS_TABLE_ROW
    ' Variance (kg) = Actual Stock (Stock) - Theoretical Stock
    ' (TheoreticalStock). Since theoretical stock resets at the start
    ' of each month, looking at the last week of the displayed period (2
    ' years out) would almost always show close to 0. So this compares
    ' using the same reference as the Actual Week column (column 8 - the
    ' most recent week that has self-stock actuals entered) (the same
    ' formula as build_soh.py).
    Dim gsLastCol As String: gsLastCol = ColLetter(2 + nWeeks)
    Dim lastActualColIdx As String
    lastActualColIdx = "LOOKUP(2,1/(" & ssSelfRng & "<>""""),COLUMN(" & ssSelfRng & "))"
    Dim diffFormula As String
    diffFormula = "=IFERROR(INDEX('Stock'!$A" & grow & ":$" & gsLastCol & grow & ",1," & lastActualColIdx & ")" & _
        "-INDEX('TheoreticalStock'!$A" & grow & ":$" & gsLastCol & grow & ",1," & lastActualColIdx & "),0)"

    Dim rowsArr(1 To 2) As Long
    rowsArr(1) = theoRow
    rowsArr(2) = actualRow
    Dim idx As Long, rr As Long
    For idx = 1 To 2
        rr = rowsArr(idx)
        sh.Cells(rr, 1).Value = rmCode
        sh.Cells(rr, 2).Value = "=IFERROR(INDEX(M_RawMaterials[Description],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),"""")"
        sh.Cells(rr, 3).Value = "=IFERROR(INDEX(M_RawMaterials[Category],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),"""")"
        sh.Cells(rr, 4).Value = "=IFERROR(INDEX(M_RawMaterials[SafetyStockMin],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),0)"
        sh.Cells(rr, 5).Value = "=IFERROR(INDEX(M_RawMaterials[SafetyStockMax],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),0)"
        sh.Cells(rr, 6).Value = "=IFERROR(LOOKUP(2,1/(" & ssSelfRng & "<>"""")," & ssSelfRng & "),"""")"
        sh.Cells(rr, 7).Value = "=IFERROR(LOOKUP(2,1/(" & ssTTAFRng & "<>"""")," & ssTTAFRng & "),"""")"
        sh.Cells(rr, 8).Value = "=IFERROR(LOOKUP(2,1/(" & ssSelfRng & "<>"""")," & ssLabelRng & "),"""")"
        sh.Cells(rr, 10).Value = diffFormula
    Next idx

    sh.Cells(theoRow, 9).Value = "Theoretical Stock"
    sh.Cells(theoRow, 9).Font.Italic = True
    sh.Cells(theoRow, 9).Font.Color = RGB(128, 128, 128)
    sh.Cells(actualRow, 9).Value = "Actual Stock"
    sh.Cells(actualRow, 9).Font.Bold = True

    Dim w As Long, col As Long
    For w = 1 To nWeeks
        col = 10 + w  ' Dashboard's week-data start column = 11 (column K)
        sh.Cells(theoRow, col).Value = "='TheoreticalStock'!" & ColLetter(2 + w) & grow
        sh.Cells(theoRow, col).Font.Italic = True
        sh.Cells(theoRow, col).Font.Color = RGB(128, 128, 128)
        sh.Cells(actualRow, col).Value = "='Stock'!" & ColLetter(2 + w) & grow
    Next w

    ' Format copy: when adding at the end, duplicate from the immediately
    ' preceding material pair; when inserting in the middle, duplicate from
    ' the anchor material's pair that got pushed down by the insert
    On Error Resume Next
    sh.Rows(formatSrcTheo).Copy
    sh.Rows(theoRow).PasteSpecial xlPasteFormats
    sh.Rows(formatSrcActual).Copy
    sh.Rows(actualRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

' Adds a new material's block to Material_Detail. Right after adding,
' since it isn't registered in M_BOM yet, there are no intermediate rows
' (No. of batches/Usage), so it's a mini block with just 7 rows (Total
' Usage - should be 0, TTAF Stock, Self Stock, Order, PO_No, Total Stock,
' and the comment). After RefreshBOM runs, once an intermediate that
' actually uses this material is found,
' SyncMaterialDetailIntermediates (in the RefreshData_BOM module, called
' automatically at the end of RefreshBOM) automatically adds the
' intermediate breakdown rows to this block (no need to run AddMaterial
' again). If anchorRmCode is given, this block is inserted right before
' that block (preserving the same order as M_RawMaterials). If it's an
' empty string, or the matching block can't be found, it's added at the
' bottom as before.
Private Sub AppendMaterialDetailBlock(sh As Worksheet, rmCode As String, descVal As String, nWeeks As Long, ssRow As Long, grow As Long, anchorRmCode As String)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row  ' based on column B (item)
    Dim headerRow As Long
    Dim formatSrcHeader As Long, formatSrcContentFirst As Long

    Dim anchorHeaderRow As Long: anchorHeaderRow = 0
    If Len(anchorRmCode) > 0 Then
        anchorHeaderRow = FindRowByCode(sh, 1, anchorRmCode, MD_HEADER_ROW + 1)
    End If

    If anchorHeaderRow > 0 Then
        ' Insert 8 rows' worth of space (1 blank + 1 header + 6 breakdown)
        ' right before the anchor block. The insert position differs
        ' depending on whether the anchor is the very first block (with no
        ' blank row right before it), but either way "after inserting, the
        ' original anchor header shifts down 8 rows" holds true, so the
        ' format-copy source row numbers (formatSrcHeader etc.) can be
        ' computed with the same formula regardless.
        Dim spacerExists As Boolean: spacerExists = (anchorHeaderRow > MD_HEADER_ROW + 1)
        Dim insertAt As Long
        If spacerExists Then
            insertAt = anchorHeaderRow - 1
        Else
            insertAt = anchorHeaderRow
        End If
        sh.Rows(insertAt & ":" & (insertAt + 7)).Insert Shift:=xlDown
        If spacerExists Then
            headerRow = insertAt + 1
        Else
            headerRow = insertAt
        End If
        formatSrcHeader = anchorHeaderRow + 8
        formatSrcContentFirst = anchorHeaderRow + 9
    Else
        headerRow = lastRow + 2  ' leave 1 blank row between this and the previous block
        formatSrcHeader = MD_HEADER_ROW + 1
        formatSrcContentFirst = lastRow - 5
    End If

    sh.Cells(headerRow, 1).Value = rmCode
    sh.Cells(headerRow, 2).Value = descVal

    Dim r As Long, w As Long, col As Long
    r = headerRow

    r = r + 1
    sh.Cells(r, 2).Value = "Total Usage (kg)/week"
    sh.Cells(r, 2).Font.Bold = True
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        sh.Cells(r, col).Value = "='WeeklyConsumption'!" & ColLetter(2 + w) & grow
    Next w

    r = r + 1
    sh.Cells(r, 2).Value = "TTAF Stock (Actual, kg)"
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        sh.Cells(r, col).Value = "='TTAFstock'!" & ColLetter(2 + w) & ssRow
    Next w

    r = r + 1
    sh.Cells(r, 2).Value = "Self Stock (Actual, kg)"
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        sh.Cells(r, col).Value = "='CSAstock'!" & ColLetter(2 + w) & ssRow
    Next w

    ' Order row: the field where the planned order quantity is entered
    ' directly, per material x week. The PO_Draft_* sheets locate this row
    ' via MATCH on the mdOrderHelperCol column (a hidden column that
    ' duplicates the Part Name only on this row), and transcribe whatever
    ' value is entered here as-is (the same design as build_soh.py's
    ' Material_Detail generation).
    r = r + 1
    sh.Cells(r, 2).Value = "Order (Planned, kg)"
    Dim mdOrderHelperCol As Long: mdOrderHelperCol = MD_WEEK_START_COL + nWeeks + 1
    sh.Cells(r, mdOrderHelperCol).Value = rmCode
    sh.Cells(r, mdOrderHelperCol).Font.Size = 8
    sh.Cells(r, mdOrderHelperCol).Font.Color = RGB(128, 128, 128)
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        sh.Cells(r, col).Value = Empty
        sh.Cells(r, col).Interior.Color = RGB(255, 242, 204)  ' equivalent to INPUT_FILL (FFF2CC)
    Next w

    ' PO_No row: right below the Order row, the field where the matching
    ' PO number is entered per material x week (fine to fill in later,
    ' once it's issued). RefreshShipments matches this row against the CSA
    ' Report's shipment rows and automatically moves/finalizes the Order
    ' row's cells depending on Status (the same layout as build_soh.py).
    r = r + 1
    sh.Cells(r, 2).Value = "PO_No"
    sh.Cells(r, 2).Font.Size = 9
    sh.Cells(r, 2).Font.Color = RGB(128, 128, 128)
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        sh.Cells(r, col).Value = Empty
        sh.Cells(r, col).Interior.Color = RGB(255, 242, 204)  ' equivalent to INPUT_FILL (FFF2CC)
        sh.Cells(r, col).Font.Size = 9
        sh.Cells(r, col).Font.Color = RGB(128, 128, 128)
    Next w

    r = r + 1
    sh.Cells(r, 2).Value = "Total Stock (End of Week, kg)"
    sh.Cells(r, 2).Font.Bold = True
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        sh.Cells(r, col).Value = "='Stock'!" & ColLetter(2 + w) & grow
    Next w

    ' Copy borders/formatting: when adding at the end, duplicate from the
    ' first existing header row and the last 6 rows of the immediately
    ' preceding block; when inserting in the middle, duplicate from the
    ' anchor block's own header row and its first 6 rows, pushed down by
    ' the insert (block length varies by material, but the header row
    ' followed by 6 rows is always in the same order).
    On Error Resume Next
    sh.Rows(formatSrcHeader).Copy
    sh.Rows(headerRow).PasteSpecial xlPasteFormats
    sh.Rows(formatSrcContentFirst & ":" & (formatSrcContentFirst + 5)).Copy
    sh.Rows((headerRow + 1) & ":" & r).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0

    ' The MOQ input field (column C of the header row) is for manual
    ' entry, so it's reset individually on top of the copied formatting
    With sh.Cells(headerRow, 3)
        .Value = Empty
        .Font.Bold = False
    End With
    On Error Resume Next
    sh.Cells(headerRow, 3).AddComment "Please enter MOQ (Minimum Order Quantity) - manual entry is fine"
    On Error GoTo 0
End Sub

' Adds a new material's row to the bottom of the matching category's
' PO_Draft_* sheet. The reference to the base-week cell is obtained via
' BaseWeekRef (RefreshData_Utilities). It used to hard-code "$P$7"
' directly, which meant that on a sheet whose base-week cell had been
' moved off P7 (e.g. moved to P13 via a manual layout change), only rows
' newly added by AddMaterial or SyncPODraftCategories would reference the
' old (empty) cell.
Public Sub AppendPODraftRow(sh As Worksheet, rmCode As String, ttafCodeVal As String, nWeeks As Long)
    ' Determines the last row based on column E (UOM) rather than column D
    ' (CSA Code). In the new layout, B20:B26/C20:C26/D20:D26 are header
    ' cells merged vertically, so with zero data rows, End(xlUp) on column
    ' D would land on the top of that merge (row 20, partway through the
    ' header) and write the new row into the middle of the header - a bug.
    ' Column E is never merged, and both the last header row (the UOM
    ' label) and every data row (kg) always carry a value there, so this
    ' issue never happens (the old layout also has the same UOM
    ' header/kg in column E, so this works correctly under either layout).
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 5).End(xlUp).Row
    Dim newRow As Long: newRow = lastRow + 1
    Dim bwRefExpr As String: bwRefExpr = BaseWeekRef(sh)

    ' Formatting is copied from the previous row (lastRow) first, and the
    ' values/formulas are written only "after" that. If the order were
    ' reversed, and the copy source happened to be the header row (row 26,
    ' where H26:K26/L26:T26 are merged for the "Firm"/"Forecast" labels),
    ' that merge would get copied too, and writing separate values/
    ' formulas into what was now one merged cell would leave only the last
    ' one written (K27 or T27) intact, overwriting all the others - a bug.
    ' (This happened when adding the first item to an empty PO_Draft_*
    ' sheet, and since each subsequent addition copies formatting from
    ' "the row before it," the merge propagated down through every
    ' following row like a chain.) As a fix, right after the format copy,
    ' the week-data columns (H through T) are always explicitly unmerged
    ' before any value/formula is written.
    sh.Rows(lastRow).Copy
    sh.Rows(newRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    Const PO_FIRST_WEEK_COL_ As Long = 6
    Const PO_N_WEEKS_ As Long = 13
    If sh.Range(sh.Cells(newRow, PO_FIRST_WEEK_COL_), sh.Cells(newRow, PO_FIRST_WEEK_COL_ + PO_N_WEEKS_ - 1)).MergeCells Then
        sh.Range(sh.Cells(newRow, PO_FIRST_WEEK_COL_), sh.Cells(newRow, PO_FIRST_WEEK_COL_ + PO_N_WEEKS_ - 1)).UnMerge
    End If

    sh.Cells(newRow, 2).Value = "=IFERROR(INDEX(M_RawMaterials[Description],MATCH(""" & rmCode & """,M_RawMaterials[Part Name],0)),"""")"
    sh.Cells(newRow, 3).Value = ttafCodeVal
    sh.Cells(newRow, 4).Value = rmCode
    sh.Cells(newRow, 5).Value = "kg"

    Const PO_FIRST_WEEK_COL As Long = 6
    Const PO_N_WEEKS As Long = 13
    Dim w As Long, col As Long
    ' The order quantity is not auto-calculated from Stock - it's
    ' transcribed as-is from whatever value was manually entered into
    ' Material_Detail's Order (Planned, kg) row (the same formula as
    ' build_soh.py's build_po_draft()).
    Dim mdOrderHelperCol As Long: mdOrderHelperCol = MD_WEEK_START_COL + nWeeks + 1
    Dim mdOrderHelperColLetter As String: mdOrderHelperColLetter = ColLetter(mdOrderHelperCol)
    Dim mdWeekFirstColLetter As String: mdWeekFirstColLetter = ColLetter(MD_WEEK_START_COL)
    Dim mdWeekLastColLetter As String: mdWeekLastColLetter = ColLetter(MD_WEEK_START_COL + nWeeks - 1)
    Dim mdOrderMatch As String
    mdOrderMatch = "MATCH($D" & newRow & ",Material_Detail!$" & mdOrderHelperColLetter & ":$" & mdOrderHelperColLetter & ",0)"
    For w = 1 To PO_N_WEEKS
        col = PO_FIRST_WEEK_COL + w - 1
        With sh.Cells(newRow, col)
            .Value = "=IFERROR(INDEX(Material_Detail!$" & mdWeekFirstColLetter & ":$" & mdWeekLastColLetter & _
                "," & mdOrderMatch & "," & bwRefExpr & "+" & (w - 1) & "),0)"
            ' The number format that hides the digit for a week with 0 order quantity (makes it look blank). The value itself stays 0.
            .NumberFormat = "0;-0;;@"
        End With
    Next w
    Dim totalCol As Long: totalCol = PO_FIRST_WEEK_COL + PO_N_WEEKS
    sh.Cells(newRow, totalCol).Value = "=SUM(" & ColLetter(PO_FIRST_WEEK_COL) & newRow & ":" & ColLetter(PO_FIRST_WEEK_COL + PO_N_WEEKS - 1) & newRow & ")"

    ' The Firm (weeks 1-4)/Forecast (weeks 5-13) color-coding is left to
    ' the conditional formatting set up once at sheet-setup time
    ' (SetupPODraftLetterheadLayout/build_soh.py) (colored only when the
    ' order quantity is non-zero) rather than a direct fill. Since that's
    ' already set up over a generously wide row range, nothing needs to be
    ' set individually for this row.
End Sub

' Deletes, from a Table (ListObject), the row matching the given material code.
Private Sub DeleteMatchingTableRow(tbl As ListObject, rmCode As String)
    Dim n As Long: n = tbl.ListRows.Count
    Dim i As Long
    Dim key As String: key = UCase(Trim(rmCode))
    For i = n To 1 Step -1
        If UCase(Trim(CStr(tbl.ListRows(i).Range.Cells(1, 1).Value))) = key Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

' Deletes, from a bordered grid that doesn't use the Table feature
' (CSAstock/TTAFstock/Dashboard/PO_Draft_*), the row matching the
' given material code. nameCol = the column number that holds the material code.
Private Sub DeleteMatchingGridRow(sh As Worksheet, rmCode As String, nameCol As Long)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, nameCol).End(xlUp).Row
    Dim r As Long
    Dim key As String: key = UCase(Trim(rmCode))
    For r = lastRow To 1 Step -1
        If UCase(Trim(CStr(sh.Cells(r, nameCol).Value))) = key Then
            sh.Rows(r).Delete
        End If
    Next r
End Sub

' Deletes, on Material_Detail, the given material's whole block in one
' shot (from its header row through right before the next material's
' header row, including any blank separator row). Does nothing if there's
' no block (e.g. because it was never registered in the BOM).
Private Sub DeleteMaterialDetailBlock(sh As Worksheet, rmCode As String)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row
    Dim startRow As Long: startRow = 0
    Dim r As Long
    For r = MD_HEADER_ROW + 1 To lastRow
        If UCase(Trim(CStr(sh.Cells(r, 1).Value))) = UCase(rmCode) Then
            startRow = r
            Exit For
        End If
    Next r
    If startRow = 0 Then Exit Sub

    Dim endRow As Long: endRow = lastRow
    For r = startRow + 1 To lastRow
        If Len(Trim(CStr(sh.Cells(r, 1).Value))) > 0 Then
            endRow = r - 1
            Exit For
        End If
    Next r
    sh.Rows(startRow & ":" & endRow).Delete
End Sub

' Counts the number of rows in a Table matching column 1 (for
' RemoveIntermediate, the "Intermediate" column of Production_Plan/M_BOM) (used
' only to show the count to the user in the confirmation dialog before the
' actual deletion).
Private Function CountMatchingTableRows(tbl As ListObject, keyVal As String) As Long
    Dim n As Long: n = tbl.ListRows.Count
    Dim i As Long, cnt As Long: cnt = 0
    Dim key As String: key = UCase(Trim(keyVal))
    For i = 1 To n
        If UCase(Trim(CStr(tbl.ListRows(i).Range.Cells(1, 1).Value))) = key Then cnt = cnt + 1
    Next i
    CountMatchingTableRows = cnt
End Function

' Scans every block on Material_Detail and deletes the breakdown row pair
' (No. of batches / Usage (kg)) for the given intermediate name wherever
' found. If one intermediate is used across multiple material blocks, it's
' deleted from every matching block (returns the number of pairs
' deleted). This is the deletion-side counterpart to RefreshData_BOM's
' SyncMaterialDetailIntermediates (the addition side), and uses the same
' block-scanning approach (reading cells live, staying consistent on the
' fly even as deletion shifts row positions).
Private Function DeleteIntermediateFromMaterialDetail(sh As Worksheet, interName As String) As Long
    Dim removedPairs As Long: removedPairs = 0
    Dim lastRowScan As Long: lastRowScan = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row

    Dim r As Long: r = MD_HEADER_ROW + 1
    Do While r <= lastRowScan
        Dim rmCode As String: rmCode = Trim(CStr(sh.Cells(r, 1).Value))
        If Len(rmCode) = 0 Then
            r = r + 1
        Else
            Dim headerRow As Long: headerRow = r
            Dim rr As Long: rr = headerRow + 1
            Dim sumRow As Long: sumRow = 0
            Do While rr <= lastRowScan
                Dim lbl As String: lbl = Trim(CStr(sh.Cells(rr, 2).Value))
                If lbl = "Total Usage (kg)/week" Then
                    sumRow = rr
                    Exit Do
                End If
                If Trim(CStr(sh.Cells(rr, 3).Value)) = "No. of batches" And StrComp(lbl, interName, vbTextCompare) = 0 Then
                    sh.Rows(rr & ":" & (rr + 1)).Delete
                    removedPairs = removedPairs + 1
                    lastRowScan = lastRowScan - 2
                    ' rr stays as-is (deletion shifts the next row up into this position)
                Else
                    rr = rr + 1
                End If
            Loop
            If sumRow > 0 Then r = sumRow + 1 Else r = headerRow + 1
        End If
    Loop
    DeleteIntermediateFromMaterialDetail = removedPairs
End Function
