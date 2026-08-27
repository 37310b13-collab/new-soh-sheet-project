Attribute VB_Name = "RefreshData_StockActuals"
Option Explicit

' ============================================================================
' RefreshData_StockActuals module
'
'   RefreshSelfStock : When you select "Raw materials daily check" (the
'                      sheet where our own warehouse's physical count is
'                      recorded; the filename includes a date in
'                      DD.MM.YYYY format), adds/updates that week's actuals
'                      in T_SelfStock. Reads the "CHEMICAL SOH" sheet
'                      (column B = CSA/code, columns C and D = stock -
'                      in the Chemical section C=WH, D=Floor, but in the
'                      Substrate/Consumables section the meaning of the
'                      columns swaps to C=Floor, D=WH; either way the total
'                      stock is simply C+D). Since the check is done every
'                      Monday morning, the date in the filename is "that
'                      week's Monday (or the next business day if it's a
'                      holiday)," but this actually represents stock as of
'                      the end of the previous week. So 7 days are
'                      subtracted before determining the target week (the
'                      same idea as RefreshTTAFStock - see the
'                      RefreshTTAFStock explanation below for details).
'   RefreshTTAFStock : When you select the "CSA Report", adds/updates that
'                      week's actuals in T_TTAFStock from its "Stock
'                      invoiced to CSA" sheet (column A = TTAF PART NUMBER,
'                      column C = Part No, column D = Description, column
'                      F = stock quantity; the header is row 4, data starts
'                      at row 5). Since this reads the hand-entered raw
'                      data directly, it is unaffected by anyone forgetting
'                      to refresh a pivot table (the old " COUNT SHEET
'                      SOH"/"PIVOT SOH TTAF"). The target week is
'                      determined from the date in cell F4 (independent of
'                      the filename). Materials are matched by TTAF_Code
'                      first; if not found, by Part No (M_RawMaterials'
'                      Part Name is often entered there as-is, more
'                      reliable than Description - also tries reading the
'                      TTAF-side "0" (zero)/"O" (letter O) spelling
'                      variants, e.g. CSA Report's Part No column has
'                      "0JN" while M_RawMaterials' official spelling is
'                      "OJN"); only if still not found is Description's
'                      normalized text matched.
'
' Both macros overwrite the value if the (raw material, week) combination
' already exists, and add a new row if not (importing multiple times
' within the same week still collapses into one row; import order doesn't
' matter).
'
' [About the two-layer structure of T_SelfStock/T_TTAFStock]
' RefreshSelfStock/RefreshTTAFStock never write to the visible
' T_SelfStock/T_TTAFStock sheets at all. They write to the hidden
' T_SelfStock_Log/T_TTAFStock_Log (a raw log keyed by count date), and the
' visible sheet is built entirely from formulas (a material x week grid)
' that recompute from that log every time. The WeekIndex column on the
' _Log sheet side is a formula column calculated automatically from the
' Date column (RefreshSelfStock/RefreshTTAFStock only write Date, never
' WeekIndex). This ensures that advancing Cal_Weeks!B1 (AnchorYear) never
' causes already-recorded actuals to be misdisplayed as "another week's
' data." BuildStockRowIndex/UpsertStockRowIndexed's matching key is
' (RM_Code, that week's Monday = MondayOfWeek, calculated from the real
' date). Keying on Monday ensures that multiple imports within the same
' week overwrite a single row (previously Date itself was the key, which
' caused rows to pile up on every daily import).
'
' For the overall design rationale (performance, DataBodyRange,
' DisplayAlerts, etc.), see the comment at the top of the
' RefreshData_Utilities module.
' ============================================================================

Sub RefreshSelfStock()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , _
        "Please select the Raw materials daily check (self stock) file")
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
    ' The date in the filename is "the Monday the count was done (or the
    ' next business day if it's a holiday)," but that value represents
    ' stock as of the end of the previous week. So 7 days are subtracted
    ' before determining the week number (even if a holiday shifted it off
    ' Monday, shifting back exactly one week still lands correctly within
    ' the previous week's range - the same idea as RefreshTTAFStock).
    Dim reportDate As Date: reportDate = ExtractDateFromName(CStr(srcPath)) - 7

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim selfTbl As ListObject: Set selfTbl = thisWb.Sheets("T_SelfStock_Log").ListObjects("T_SelfStock_Log")
    Dim wIdx As Long: wIdx = WeekIndexForDate(thisWb, reportDate)
    Dim selfIdx As Object: Set selfIdx = BuildStockRowIndex(selfTbl)
    ' Determines target rows by whether column C (CSA code) is an actual
    ' Part Name in M_RawMaterials (it used to be limited to chemicals only
    ' via Left(code,4)="CHEM", which meant Substrate rows - short codes
    ' like OJN, 7EP, etc. - were never reflected at all, a bug).
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")
    Dim rmCodeSet As Object: Set rmCodeSet = BuildNameIndex(rmTbl, "Part Name")

    ' Uses the "CHEMICAL SOH" sheet directly rather than the "Stock" sheet's
    ' column J (Total). The "Stock" sheet's mechanism differs by material
    ' type - chemicals use a column K+L formula, Substrate uses a VLOOKUP
    ' referencing a different sheet - and on top of that, both were found
    ' to not recalculate right after Workbooks.Open, so they'd be read with
    ' the source file's stale cached value as of when it was last saved.
    ' The "CHEMICAL SOH" sheet has the measured values for chemicals,
    ' Substrate, Ester Film/Original Towel/PP Film, etc. all as common raw
    ' data in column B (CSA/code), column C, and column D (in the Chemical
    ' section C=WH, D=Floor; in later sections the meaning swaps to
    ' C=Floor, D=WH, but either way the total stock is simply the sum of
    ' columns C+D, so there's no need to distinguish sections).
    ' Note that MAT codes (18456-xxxxx) are not on this sheet, only on the
    ' "Stock" sheet. MAT is outside the scope of self-stock management, so
    ' it's left unhandled.
    Dim sh As Worksheet: Set sh = srcWb.Sheets("CHEMICAL SOH")
    ' Reading the sheet cell by cell is slow, so read a generous range
    ' (rows 5-300, columns A-D) as a single array read.
    Dim data As Variant
    data = sh.Range(sh.Cells(5, 1), sh.Cells(300, 4)).Value
    Dim r As Long, added As Long, updated As Long
    added = 0: updated = 0
    For r = 1 To (300 - 5 + 1)
        Dim code As String
        code = Trim(CStr(data(r, 2)))
        If Len(code) > 0 Then
            Dim matchedCode As String: matchedCode = ResolveSelfStockCode(rmTbl, rmCodeSet, code)
            If Len(matchedCode) > 0 Then
                Dim vC As Double: vC = 0
                Dim vD As Double: vD = 0
                If IsNumeric(data(r, 3)) Then vC = CDbl(data(r, 3))
                If IsNumeric(data(r, 4)) Then vD = CDbl(data(r, 4))
                Call UpsertStockRowIndexed(selfTbl, selfIdx, matchedCode, reportDate, vC + vD, added, updated)
            End If
        End If
    Next r

    ' Guard the cleanup step itself against failing with "object variable
    ' not set," even in the case where srcWb has already become Nothing
    ' (some automated process on the source file's side can close the
    ' workbook right after it's opened).
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "T_SelfStock has been updated." & vbCrLf & "Target week: " & wIdx & " (" & Format(reportDate, "yyyy-mm-dd") & ")" & vbCrLf & _
           "Added: " & added & ", updated: " & updated & vbCrLf & _
           "(Multiple actuals within the same week are collapsed into one row. The grid-view T_SelfStock sheet updates automatically.)", vbInformation
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

Sub RefreshTTAFStock()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , _
        "Please select the CSA Report (TTAF stock) file")
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

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim ttafTbl As ListObject: Set ttafTbl = thisWb.Sheets("T_TTAFStock_Log").ListObjects("T_TTAFStock_Log")
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    ' Uses the "Stock invoiced to CSA" sheet. Column A = TTAF PART NUMBER,
    ' column C = Part No, column D = Description, column F = stock
    ' quantity. The header is row 4, data starts at row 5. Since this is
    ' hand-entered raw data, there's no risk of someone forgetting to
    ' refresh a pivot table like " COUNT SHEET SOH"/"PIVOT SOH TTAF".
    Dim sh As Worksheet: Set sh = srcWb.Sheets("Stock invoiced to CSA")
    ' The date in F4 is "the Monday the report arrived (or the next
    ' business day if it's a holiday)," but that value represents stock as
    ' of the close of business the previous Friday. So 7 days are
    ' subtracted before determining the week number (Monday and Friday both
    ' fall within the same Excel Mon-Sun week, so -7 or -3 would give the
    ' same week-number result either way; -7 is used so the date itself
    ' lines up with Monday, the start of the week, for consistency with
    ' other actuals like T_SelfStock). The "TTAF count(dd.mm.yyyy)" header
    ' is sometimes merged across F3:F4 etc., and a merged cell only carries
    ' a value on its top-left (anchor) cell. To avoid reading a blank when
    ' .Cells(4,6) isn't the anchor, .MergeArea.Cells(1,1) is used to always
    ' fetch the anchor cell's value.
    Dim reportDate As Date: reportDate = ExtractDDMMYYYYFromText(sh.Cells(4, 6).MergeArea.Cells(1, 1).Value) - 7
    Dim wIdx As Long: wIdx = WeekIndexForDate(thisWb, reportDate)
    Dim ttafIdx As Object: Set ttafIdx = BuildStockRowIndex(ttafTbl)

    Dim ttafCodeIdx As Object: Set ttafCodeIdx = CreateObject("Scripting.Dictionary")
    Dim descIdx As Object: Set descIdx = CreateObject("Scripting.Dictionary")
    Dim rmNameIdx As Object: Set rmNameIdx = CreateObject("Scripting.Dictionary")
    Call BuildTTAFCodeAndDescIndex(rmTbl, ttafCodeIdx, descIdx, rmNameIdx)

    ' Reading the sheet cell by cell is slow, so read a generously sized
    ' range as a single array read, then scan it (column A = TTAF PART
    ' NUMBER, column C = Part No, column D = Description, column F = stock
    ' quantity).
    Const MAX_ROWS As Long = 2000
    Dim data As Variant
    data = sh.Range(sh.Cells(5, 1), sh.Cells(MAX_ROWS, 6)).Value

    Dim r As Long, added As Long, updated As Long, unresolved As String
    added = 0: updated = 0: unresolved = ""
    For r = 1 To (MAX_ROWS - 5 + 1)
        Dim ttafCodeRaw As String: ttafCodeRaw = Trim(CStr(data(r, 1)))
        If Len(ttafCodeRaw) = 0 Then GoTo NextRow
        Dim v As Variant: v = data(r, 6)
        If Not IsNumeric(v) Then GoTo NextRow

        Dim partNoRaw As String: partNoRaw = Trim(CStr(data(r, 3)))
        Dim descRaw As String: descRaw = Trim(CStr(data(r, 4)))
        Dim matchedPart As String
        matchedPart = ResolveTTAFPart(ttafCodeIdx, rmNameIdx, descIdx, ttafCodeRaw, partNoRaw, descRaw)

        If Len(matchedPart) = 0 Then
            If InStr(unresolved, ttafCodeRaw) = 0 Then
                unresolved = unresolved & ttafCodeRaw & " (" & descRaw & "); "
            End If
            GoTo NextRow
        End If

        Call UpsertStockRowIndexed(ttafTbl, ttafIdx, matchedPart, reportDate, CDbl(v), added, updated)
NextRow:
    Next r

    ' Guard the cleanup step itself against failing with "object variable
    ' not set," even in the case where srcWb has already become Nothing
    ' (some automated process on the source file's side can close the
    ' workbook right after it's opened).
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim msg As String
    msg = "T_TTAFStock has been updated." & vbCrLf & "Target week: " & wIdx & " (" & Format(reportDate, "yyyy-mm-dd") & ")" & vbCrLf & _
          "Added: " & added & ", updated: " & updated & vbCrLf & _
          "(Multiple actuals within the same week are collapsed into one row. The grid-view T_TTAFStock sheet updates automatically.)"
    If Len(unresolved) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Rows that could not be matched by either TTAF_Code or material name and were not applied:" & vbCrLf & unresolved
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
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "An error occurred during the refresh: (" & errNum & ") " & errMsg, vbCritical
End Sub

' Builds, once, indexes from M_RawMaterials: TTAF_Code (normalized) ->
' Part Name, Description (normalized) -> Part Name, and Part Name
' (case/leading-trailing-space insensitive) -> Part Name. A shared routine
' used by RefreshTTAFStock. Since a Dictionary is an object (passed by
' reference), the caller's ttafCodeIdx/descIdx/rmNameIdx are updated
' directly without needing an explicit ByRef.
Private Sub BuildTTAFCodeAndDescIndex(rmTbl As ListObject, ttafCodeIdx As Object, descIdx As Object, rmNameIdx As Object)
    rmNameIdx.CompareMode = vbTextCompare
    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    If rmN > 0 Then
        Dim rmNameDesc As Variant
        rmNameDesc = rmTbl.ListColumns(1).DataBodyRange.Resize(rmN, 2).Value  ' Part Name, Description
        Dim rmTtafCode As Variant
        rmTtafCode = rmTbl.ListColumns(9).DataBodyRange.Value                 ' TTAF_Code
        Dim i As Long
        For i = 1 To rmN
            Dim partName As String: partName = Trim(CStr(rmNameDesc(i, 1)))
            Dim tKeyBuild As String: tKeyBuild = NormalizeText(CStr(rmTtafCode(i, 1)))
            If Len(tKeyBuild) > 0 And Not ttafCodeIdx.Exists(tKeyBuild) Then ttafCodeIdx(tKeyBuild) = partName
            Dim dKeyBuild As String: dKeyBuild = NormalizeText(CStr(rmNameDesc(i, 2)))
            If Len(dKeyBuild) > 0 And Not descIdx.Exists(dKeyBuild) Then descIdx(dKeyBuild) = partName
            If Len(partName) > 0 And Not rmNameIdx.Exists(partName) Then rmNameIdx(partName) = partName
        Next i
    End If

    ' Manual alias for a known spelling variant (the same idea as
    ' RefreshData_Shipments' BuildKnownAliasIndex). In the "Stock invoiced
    ' to CSA" sheet's TTAF PART NUMBER column (column A), ND TAC
    ' (CHEM-1280) alone uses the string "NDTAC" instead of a numeric
    ' TTAF_Code like every other material (83988002202, matching the CSA
    ' Report's Shipping Schedule) - an input inconsistency on the TTAF
    ' side. Currently this causes no real harm because it's still caught
    ' by the Description ("ND TAC") match, but this alias is added so it
    ' also matches directly via TTAF_Code, in case the Description
    ' spelling changes in the future.
    Dim aliasKey As String: aliasKey = NormalizeText("NDTAC")
    If Not ttafCodeIdx.Exists(aliasKey) Then ttafCodeIdx(aliasKey) = "CHEM-1280"
End Sub

' Used by RefreshSelfStock. If the "CHEMICAL SOH" sheet's column B
' (CSA/code) exactly matches an M_RawMaterials Part Name, returns it as-is.
' If not found, tries both directions of the "0" (zero)/"O" (letter O)
' spelling variant (e.g. "0JN" vs "OJN" - repeatedly found across several
' TTAF-side files such as daily check and CSA Report). If a match is
' found, returns M_RawMaterials' official spelling (Part Name) (writing
' the source file's raw spelling as-is would disagree with the official
' spelling used in existing T_SelfStock_Log rows and other sheets, and the
' grid view would fail to recognize it as that material).
' Returns an empty string if nothing matches at all.
Private Function ResolveSelfStockCode(rmTbl As ListObject, rmCodeSet As Object, codeRaw As String) As String
    If rmCodeSet.Exists(codeRaw) Then
        ResolveSelfStockCode = Trim(CStr(rmTbl.ListRows(rmCodeSet(codeRaw)).Range.Cells(1, 1).Value))
        Exit Function
    End If
    Dim zeroToO As String: zeroToO = Replace(codeRaw, "0", "O")
    If zeroToO <> codeRaw And rmCodeSet.Exists(zeroToO) Then
        ResolveSelfStockCode = Trim(CStr(rmTbl.ListRows(rmCodeSet(zeroToO)).Range.Cells(1, 1).Value))
        Exit Function
    End If
    Dim oToZero As String: oToZero = Replace(codeRaw, "O", "0")
    If oToZero <> codeRaw And rmCodeSet.Exists(oToZero) Then
        ResolveSelfStockCode = Trim(CStr(rmTbl.ListRows(rmCodeSet(oToZero)).Range.Cells(1, 1).Value))
        Exit Function
    End If
    ResolveSelfStockCode = ""
End Function

' Matches by TTAF_Code first; if not found, matches by Part No (column C -
' M_RawMaterials' Part Name is often entered there as-is, more reliable
' than Description); only if still not found does it match by
' Description's (material name's) normalized text. Returns an empty string
' if nothing matches at all.
Private Function ResolveTTAFPart(ttafCodeIdx As Object, rmNameIdx As Object, descIdx As Object, _
        ttafCodeRaw As String, partNoRaw As String, descRaw As String) As String
    Dim tKey As String: tKey = NormalizeText(ttafCodeRaw)
    If Len(tKey) > 0 And ttafCodeIdx.Exists(tKey) Then
        ResolveTTAFPart = ttafCodeIdx(tKey)
        Exit Function
    End If
    If Len(partNoRaw) > 0 And rmNameIdx.Exists(partNoRaw) Then
        ResolveTTAFPart = rmNameIdx(partNoRaw)
        Exit Function
    End If
    ' The "0" (zero)/"O" (letter O) spelling variant is repeatedly found in
    ' the TTAF-side source data (e.g. CSA Report's Part No column has
    ' "0JN" while M_RawMaterials' official spelling is "OJN"). If an exact
    ' match isn't found, try reading it as either spelling.
    If Len(partNoRaw) > 0 Then
        Dim zeroToO As String: zeroToO = Replace(partNoRaw, "0", "O")
        If zeroToO <> partNoRaw And rmNameIdx.Exists(zeroToO) Then
            ResolveTTAFPart = rmNameIdx(zeroToO)
            Exit Function
        End If
        Dim oToZero As String: oToZero = Replace(partNoRaw, "O", "0")
        If oToZero <> partNoRaw And rmNameIdx.Exists(oToZero) Then
            ResolveTTAFPart = rmNameIdx(oToZero)
            Exit Function
        End If
    End If
    Dim dKey As String: dKey = NormalizeText(descRaw)
    If descIdx.Exists(dKey) Then
        ResolveTTAFPart = descIdx(dKey)
        Exit Function
    End If
    ResolveTTAFPart = ""
End Function

' Cell F4 shows the header text "TTAF count(dd.mm.yyyy)", but if this is
' purely a display effect of the cell's format (a custom number format),
' the actual .Value can be a real date serial number. In that case,
' stringifying it with CStr() would produce a different, locale-dependent
' form (e.g. 6/29/2026) that wouldn't match the DD.MM.YYYY regex. So this
' first checks with IsDate() whether it's a real date value, and if so
' uses it directly (without going through a string). When it does need to
' read it as a string, it's first normalized to half-width digits with
' StrConv(..., vbNarrow), in case it was entered in full-width digits
' (e.g. 29.06.2026 in full-width characters), before matching.
Private Function ExtractDDMMYYYYFromText(cellValue As Variant) As Date
    If IsDate(cellValue) Then
        ExtractDDMMYYYYFromText = CDate(cellValue)
        Exit Function
    End If
    Dim text As String: text = StrConv(CStr(cellValue), vbNarrow)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d{1,2})\.(\d{1,2})\.(\d{4})"
    Dim m As Object
    Set m = re.Execute(text)
    If m.Count = 0 Then
        Err.Raise vbObjectError + 1, , "Could not read the TTAF count date (DD.MM.YYYY): " & CStr(cellValue)
    End If
    ExtractDDMMYYYYFromText = DateSerial(CInt(m(0).SubMatches(2)), CInt(m(0).SubMatches(1)), CInt(m(0).SubMatches(0)))
End Function

Private Function ExtractDateFromName(path As String) As Date
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d{2})\.(\d{2})\.(\d{4})"
    Dim m As Object
    Set m = re.Execute(path)
    If m.Count = 0 Then
        Err.Raise vbObjectError + 1, , "Could not read a date (DD.MM.YYYY) from the filename."
    End If
    ExtractDateFromName = DateSerial(CInt(m(0).SubMatches(2)), CInt(m(0).SubMatches(1)), CInt(m(0).SubMatches(0)))
End Function

' Computes "that week's Monday" from a date using real calendar arithmetic
' (a pure date calculation with no dependency at all on Cal_Weeks!B1's
' AnchorYear). Used as the matching key for T_SelfStock_Log/T_TTAFStock_Log
' so that importing multiple times within the same week overwrites a
' single row (previously Date itself was the key, which caused rows to
' pile up on every daily import).
Private Function MondayOfWeek(d As Date) As Date
    MondayOfWeek = d - Weekday(d, vbMonday) + 1
End Function

' Builds, once, an index of tbl's (T_SelfStock_Log/T_TTAFStock_Log)
' (RM_Code, that week's Monday) -> row number. Columns are RM_Code(1),
' Date(2), WeekIndex(3, a formula auto-calculated from Date), Qty(4).
' Keying on "that week's Monday" (calculated from the real calendar date,
' independent of AnchorYear) satisfies two goals at once: (1) collapsing
' multiple imports within the same week into one row, and (2) ensuring
' that changing AnchorYear never misdisplays already-recorded actuals as
' "another week's data." To avoid scanning every row cell-by-cell every
' time UpsertStockRow is called, the Dictionary is built once up front
' from a single Range read (the benefit grows as the table grows month
' over month). Dates are stringified via CLng(the serial value) so the
' result is independent of the regional date display format.
Private Function BuildStockRowIndex(tbl As ListObject) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim n As Long: n = tbl.ListRows.Count
    If n > 0 Then
        Dim data As Variant
        data = tbl.ListColumns(1).DataBodyRange.Resize(n, 2).Value  ' read columns 1,2 (RM_Code,Date) together
        Dim i As Long
        For i = 1 To n
            idx(CStr(data(i, 1)) & "|" & CStr(CLng(MondayOfWeek(CDate(data(i, 2)))))) = i
        Next i
    End If
    Set BuildStockRowIndex = idx
End Function

' WeekIndex (column 3) is a formula column, so no value is written here
' (it recalculates automatically whenever Date changes).
' [Important] This used to assume that "adding a new row makes Excel's
' Table feature automatically duplicate the same formula as the existing
' rows," but that behavior only applies when a row is added by hand below
' the table in the UI - a row added via VBA's ListRows.Add is not always
' auto-duplicated (an actual reported bug: a row added via VBA to
' T_SelfStock_Log/T_TTAFStock_Log had its WeekIndex column left entirely
' blank, formula and all, so the grid side's SUMIFS/COUNTIFS couldn't find
' that row and T_SelfStock/T_TTAFStock showed nothing). So for a new row,
' the previous row's WeekIndex formula is explicitly copied (copying via
' FormulaR1C1 means the relative reference - the part pointing at its own
' Date cell - automatically adjusts to match the destination row).
' If there's a second (or later) import within the same week, both Date
' and Qty are overwritten with the latest values (so the record left
' behind is always for the most recent count date within that week).
Private Sub UpsertStockRowIndexed(tbl As ListObject, idx As Object, code As String, d As Date, v As Double, ByRef added As Long, ByRef updated As Long)
    Dim key As String: key = code & "|" & CStr(CLng(MondayOfWeek(d)))
    If idx.Exists(key) Then
        Dim rowN As Long: rowN = idx(key)
        tbl.ListRows(rowN).Range.Cells(1, 2).Value = d
        tbl.ListRows(rowN).Range.Cells(1, 4).Value = v
        updated = updated + 1
    Else
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add
        newRow.Range.Cells(1, 1).Value = code
        newRow.Range.Cells(1, 2).Value = d
        newRow.Range.Cells(1, 4).Value = v
        If newRow.Index > 1 Then
            newRow.Range.Cells(1, 3).FormulaR1C1 = tbl.ListRows(newRow.Index - 1).Range.Cells(1, 3).FormulaR1C1
        End If
        idx(key) = newRow.Index
        added = added + 1
    End If
End Sub
