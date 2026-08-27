Attribute VB_Name = "RefreshData_PODraft"
Option Explicit

' ============================================================================
' RefreshData_PODraft module
'
' SetupPODraftLetterheadLayout : [One-time migration macro]
'   Takes the letterhead layout manually built out on the PO_Draft_Hazardous
'   sheet (TO/FROM/CC, Order Date/Issue Month/Firm Month, Revision, base
'   week (WeekIndex), Firm/Forecast color-coding, SafetyStock/CurrentStock
'   placed outside the print area) and:
'     (1) fixes bugs that remained on PO_Draft_Hazardous itself, then
'     (2) duplicates the same layout onto the three sheets
'         PO_Draft_Chemical, PO_Draft_Substrate_JPN_CHN, and
'         PO_Draft_Substrate_Poland.
'
'   Bugs fixed in (1):
'     - The month/week header (row 20) used fixed-width merged cells
'       (weeks 1-4 / 5-8 / 9-13); when the actual number of weeks in a
'       month varies (e.g. November 2026 spans 5 weeks), that month's
'       header could disappear (this was actually reported as a bug where
'       the year/month vanished on the Forecast side).
'       -> Stop merging cells; go back to build_soh.py's original approach
'       of "one cell per week, shown only when it differs from the
'       previous week's month" (always displays correctly regardless of
'       how long the month is).
'     - When the base-week cell was manually moved from P7 to P13, the
'       formulas in some existing materials' rows (ND TAC/CHEM-1280, etc.)
'       were left behind still referencing the old $P$7, so only those
'       materials' order quantities stayed stuck at 0 and never updated
'       (because they kept referencing the now-blank P7 cell).
'       -> Unify all direct $P$7/$P$13 references in every data row to the
'       sheet-local named range "BaseWeek". From now on, moving the base
'       week cell anywhere only requires fixing where the name points, and
'       rows newly added by AddMaterial/SyncPODraftCategories automatically
'       follow suit (see AppendPODraftRow in RefreshData_MaterialMgmt.bas
'       and BaseWeekRef in RefreshData_Utilities).
'     - The Firm (weeks 1-4)/Forecast (weeks 5-13) order-quantity cells
'       were missing the requested red/green color-coding.
'     - SafetyStock/CurrentStock (columns F/G) were kept out of print by
'       hiding the columns, but unhiding them to check the values also
'       made them show up in print (in practice this made "show them again
'       only when I want to check for reference" unworkable).
'       -> Unhide the columns so they're always visible, and instead
'       exclude them from the print area itself (Excel's print area can
'       specify multiple rectangles, so specific columns can be excluded
'       without relying on hiding).
'
'   When duplicating PO_Draft_Chemical/_Substrate_JPN_CHN/_Substrate_Poland
'   in (2):
'     - The TO/FROM/CC fields are not copied as-is from PO_Draft_Hazardous
'       (since the contact/supplier may differ by category, they're reset
'       to placeholder text after duplication - please enter the real
'       recipient by hand after duplicating).
'     - Revision and base week (WeekIndex) carry over from each sheet's own
'       pre-duplication values.
'     - The data rows (material list) are not reused from before
'       duplication - they are rebuilt based on M_RawMaterials' current
'       Category/Origin_Country (the same logic as POSheetNameForMaterial;
'       the same idea as SyncPODraftCategories, so the latest state as of
'       the duplication time is reflected).
'     - The logo image/decorations (banner, etc.) are duplicated
'       automatically by Excel's sheet-copy feature, but if the banner has
'       a hand-typed date string, please check and fix the content on each
'       sheet after duplicating.
'
'   Both steps skip parts that are already migrated, so it's safe to run
'   more than once by mistake.
' ============================================================================

Private Const PO_HDR_ROW As Long = 26        ' last header row (data rows start right below). Corresponds to build_soh.py's PO_HDR_UOM_FIRM_ROW
Private Const PO_DATA_START_ROW As Long = 27 ' Corresponds to build_soh.py's PO_DATA_START_ROW
Private Const PO_TITLE_ROW As Long = 17      ' Corresponds to build_soh.py's PO_TITLE_ROW (rows 17-18 merged)
Private Const PO_MONTHYEAR_ROW As Long = 20  ' Corresponds to build_soh.py's PO_HDR_MONTHYEAR_ROW
Private Const PO_FIRST_WEEK_COL As Long = 8  ' column H
Private Const PO_N_WEEKS As Long = 13
Private Const PO_BASEWEEK_ADDR As String = "$P$13"
Private Const PO_BASEWEEK_ROW As Long = 13   ' the row number of PO_BASEWEEK_ADDR (used to exclude it from the print area)
Private Const PO_REVISION_ADDR As String = "$P$11"

Sub SetupPODraftLetterheadLayout()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook

    Dim hazFixed As Boolean
    hazFixed = FixHazardousPODraftLayout(thisWb)

    Dim c1 As Boolean, c2 As Boolean, c3 As Boolean
    c1 = ClonePODraftLetterheadIfNeeded(thisWb, "PO_Draft_Chemical", "Chemicals : TTAF Supply")
    c2 = ClonePODraftLetterheadIfNeeded(thisWb, "PO_Draft_Substrate_JPN_CHN", "Substrates (Japan / China)")
    c3 = ClonePODraftLetterheadIfNeeded(thisWb, "PO_Draft_Substrate_Poland", "Substrates (Poland)")

    If Not hazFixed And Not c1 And Not c2 And Not c3 Then
        MsgBox "Already migrated (all PO_Draft_* sheets are already using the new layout).", vbInformation
        Exit Sub
    End If

    MsgBox "The PO_Draft_* sheet layout migration is complete." & vbCrLf & vbCrLf & _
           "PO_Draft_Hazardous: " & IIf(hazFixed, "fixed (month/week header, base-week reference, color-coding, print area)", "was already fixed") & vbCrLf & _
           "PO_Draft_Chemical: " & IIf(c1, "switched to the new layout", "already on the new layout") & vbCrLf & _
           "PO_Draft_Substrate_JPN_CHN: " & IIf(c2, "switched to the new layout", "already on the new layout") & vbCrLf & _
           "PO_Draft_Substrate_Poland: " & IIf(c3, "switched to the new layout", "already on the new layout") & vbCrLf & vbCrLf & _
           "The TO/FROM/CC fields on the 3 newly created sheets are still placeholder text. Please enter" & vbCrLf & _
           "the real recipient/our company name (PO_Draft_Hazardous's content was not copied)." & vbCrLf & vbCrLf & _
           "The logo image/banner have been duplicated from PO_Draft_Hazardous. If the banner has a" & vbCrLf & _
           "hand-typed date, please check and fix the content on each sheet.", vbInformation
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred during processing: (" & Err.Number & ") " & Err.Description & vbCrLf & vbCrLf & _
           "Some changes may have already been applied. If you're unsure, close without saving and reopen.", vbCritical
End Sub

' Makes the sheet-local (local-scope) named ranges "BaseWeek"/"PORevision"
' always point at that sheet's own cells (overwrites the target if the name
' already exists; creates it if not). RefersTo must always explicitly
' include the sheet name (in the form "='" & sh.Name & "'!..."). Omitting
' the sheet name can cause Names.Add to end up pointing at whatever sheet
' happened to be active at the time (this actually caused runtime error
' 1004) - as a countermeasure, this function is designed to always be
' called and overwrite the target correctly, regardless of whether
' migration has already happened.
Private Sub EnsureLocalBaseWeekNames(sh As Worksheet)
    Dim bwName As Name
    On Error Resume Next
    Set bwName = sh.Names("BaseWeek")
    On Error GoTo 0
    If bwName Is Nothing Then
        sh.Names.Add Name:="BaseWeek", RefersTo:="='" & sh.Name & "'!" & PO_BASEWEEK_ADDR
    Else
        bwName.RefersTo = "='" & sh.Name & "'!" & PO_BASEWEEK_ADDR
    End If

    Dim revName As Name
    On Error Resume Next
    Set revName = sh.Names("PORevision")
    On Error GoTo 0
    If revName Is Nothing Then
        sh.Names.Add Name:="PORevision", RefersTo:="='" & sh.Name & "'!" & PO_REVISION_ADDR
    Else
        revName.RefersTo = "='" & sh.Name & "'!" & PO_REVISION_ADDR
    End If
End Sub

' Fixes the bugs on PO_Draft_Hazardous itself. If already fixed (the
' month/week header (row 20) merge is already undone), just double-check/
' fix the named ranges' targets as a precaution, then return False and skip
' the rest.
Private Function FixHazardousPODraftLayout(thisWb As Workbook) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = thisWb.Sheets("PO_Draft_Hazardous")
    On Error GoTo 0
    If sh Is Nothing Then
        FixHazardousPODraftLayout = False
        Exit Function
    End If

    ' Rebuild the named ranges every run regardless of whether migration
    ' already happened (if the previous run stopped partway through with an
    ' error, the names might not be set up correctly).
    Call EnsureLocalBaseWeekNames(sh)

    If Not sh.Cells(PO_MONTHYEAR_ROW, PO_FIRST_WEEK_COL).MergeCells Then
        ' The month/week header is already unmerged = the body of this function already ran
        FixHazardousPODraftLayout = False
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ---- Unify every data row's direct $P$7/$P$13 references to the
    ' BaseWeek name (column G = stock lookup, columns H-T = order quantity) ----
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 5).End(xlUp).Row  ' based on column E (UOM)
    Dim r As Long, c As Long
    If lastRow >= PO_DATA_START_ROW Then
        For r = PO_DATA_START_ROW To lastRow
            Dim gCell As Range: Set gCell = sh.Cells(r, 7)
            If Len(gCell.Formula) > 0 Then
                gCell.Formula = Replace(Replace(gCell.Formula, "$P$7", "BaseWeek"), "$P$13", "BaseWeek")
            End If
            For c = PO_FIRST_WEEK_COL To PO_FIRST_WEEK_COL + PO_N_WEEKS - 1
                Dim wCell As Range: Set wCell = sh.Cells(r, c)
                If Len(wCell.Formula) > 0 Then
                    wCell.Formula = Replace(Replace(wCell.Formula, "$P$7", "BaseWeek"), "$P$13", "BaseWeek")
                End If
            Next c
        Next r
    End If

    ' ---- Stop merging the month/week header (row 20) and rewrite it as a per-week formula ----
    Dim col As Long
    For col = PO_FIRST_WEEK_COL To PO_FIRST_WEEK_COL + PO_N_WEEKS - 1
        If sh.Cells(PO_MONTHYEAR_ROW, col).MergeCells Then
            sh.Cells(PO_MONTHYEAR_ROW, col).MergeArea.UnMerge
        End If
    Next col
    Dim w As Long
    For w = 1 To PO_N_WEEKS
        col = PO_FIRST_WEEK_COL + w - 1
        Dim f As String
        If w = 1 Then
            f = "=INDEX(Cal_Weeks[MonthYearLabel],BaseWeek)"
        Else
            f = "=IF(INDEX(Cal_Weeks[MonthYearLabel],BaseWeek+" & (w - 1) & ")<>INDEX(Cal_Weeks[MonthYearLabel],BaseWeek+" & (w - 2) & "),INDEX(Cal_Weeks[MonthYearLabel],BaseWeek+" & (w - 1) & "),"""")"
        End If
        With sh.Cells(PO_MONTHYEAR_ROW, col)
            .Formula = f
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(191, 191, 191)
            If w <= 4 Then
                .Interior.Color = RGB(255, 193, 193)  ' Firm: FFC1C1
            Else
                .Interior.Color = RGB(235, 241, 222)  ' Forecast: EBF1DE
            End If
        End With
    Next w

    ' ---- The Firm (weeks 1-4)/Forecast (weeks 5-13) order-quantity cell
    ' color-coding uses conditional formatting rather than direct fill (per
    ' the request that a week with 0 order quantity show no color and no
    ' number). Applied over a row range (500 rows) well beyond the actual
    ' data, so rows added later by AddMaterial etc. are covered too. The
    ' number format is also set to hide 0.
    Call ApplyPODraftZeroHiddenFormatting(sh)

    ' ---- Unhide SafetyStock/CurrentStock (columns F/G) and exclude them from the print area ----
    sh.Columns("F:G").Hidden = False
    sh.Columns("F").ColumnWidth = 12
    sh.Columns("G").ColumnWidth = 12
    Dim printLastRow As Long
    printLastRow = IIf(lastRow >= PO_DATA_START_ROW, lastRow, PO_HDR_ROW)
    Call SetPODraftPrintArea(sh, printLastRow)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    FixHazardousPODraftLayout = True
End Function

' Duplicates the (fixed) structure of PO_Draft_Hazardous to rebuild the
' targetSheetName sheet with the new layout. If it's already on the new
' layout (the title row (rows 17-18) is already merged), returns False and
' does nothing.
Private Function ClonePODraftLetterheadIfNeeded(thisWb As Workbook, targetSheetName As String, titleText As String) As Boolean
    Dim oldSh As Worksheet
    On Error Resume Next
    Set oldSh = thisWb.Sheets(targetSheetName)
    On Error GoTo 0
    If oldSh Is Nothing Then
        ClonePODraftLetterheadIfNeeded = False
        Exit Function
    End If
    If oldSh.Cells(PO_TITLE_ROW, 2).MergeCells Then
        ' The title row is already merged = already migrated to the new layout
        ClonePODraftLetterheadIfNeeded = False
        Exit Function
    End If

    Dim hazSh As Worksheet
    On Error Resume Next
    Set hazSh = thisWb.Sheets("PO_Draft_Hazardous")
    On Error GoTo 0
    If hazSh Is Nothing Then
        MsgBox "The PO_Draft_Hazardous sheet was not found.", vbExclamation
        ClonePODraftLetterheadIfNeeded = False
        Exit Function
    End If
    If Not hazSh.Cells(PO_TITLE_ROW, 2).MergeCells Then
        MsgBox "PO_Draft_Hazardous must be fixed to the new layout first." & vbCrLf & _
               "Please run SetupPODraftLetterheadLayout again.", vbExclamation
        ClonePODraftLetterheadIfNeeded = False
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Read the values to carry over from the existing sheet (old layout)
    ' before deleting it (in the old layout, Revision=P5, base week=P7).
    Dim oldRevision As Variant: oldRevision = oldSh.Range("P5").Value
    If Not IsNumeric(oldRevision) Then oldRevision = "00"
    Dim oldBaseWeek As Variant: oldBaseWeek = oldSh.Range("P7").Value
    If Not IsNumeric(oldBaseWeek) Then oldBaseWeek = hazSh.Range(PO_BASEWEEK_ADDR).Value

    Dim oldIdx As Long: oldIdx = oldSh.Index
    Application.DisplayAlerts = False
    oldSh.Delete
    Application.DisplayAlerts = True

    hazSh.Copy After:=thisWb.Sheets(thisWb.Sheets.Count)
    Dim newSh As Worksheet: Set newSh = thisWb.Sheets(thisWb.Sheets.Count)
    newSh.Name = targetSheetName
    ' Move it back near its original position (to preserve the visual sheet order; harmless to ignore if this fails)
    On Error Resume Next
    newSh.Move Before:=thisWb.Sheets(Application.WorksheetFunction.Min(oldIdx, thisWb.Sheets.Count))
    On Error GoTo 0

    ' A sheet copy does not always duplicate named ranges correctly, so
    ' explicitly rebuild them to point at the destination (newSh) itself
    ' (see EnsureLocalBaseWeekNames).
    Call EnsureLocalBaseWeekNames(newSh)

    ' ---- Letterhead: reset the title and TO/FROM/CC to placeholder text
    ' (don't carry over Hazardous's real recipient as-is, since the
    ' contact/supplier may differ by category) ----
    newSh.Range("B8").Value = "TO: (Enter supplier / TTAF contact name)"
    newSh.Range("B9").Value = "     (Company name)"
    newSh.Range("B11").Value = "CC: (Enter if needed)"
    newSh.Range("B13").Value = "FROM: (Issuer name)"
    newSh.Range("B14").Value = "     (Our company name)"
    newSh.Range("B17").Value = titleText
    newSh.Range(PO_BASEWEEK_ADDR).Value = oldBaseWeek
    newSh.Range(PO_REVISION_ADDR).Value = oldRevision

    ' ---- Delete the source (Hazardous) sheet's data rows and rebuild the
    ' rows for materials that belong on this sheet, based on
    ' M_RawMaterials' current Category/Origin_Country (the same criteria as
    ' SyncPODraftCategories; see POSheetNameForMaterial) ----
    Dim oldLastRow As Long: oldLastRow = newSh.Cells(newSh.Rows.Count, 5).End(xlUp).Row
    If oldLastRow >= PO_DATA_START_ROW Then
        newSh.Rows(PO_DATA_START_ROW & ":" & oldLastRow).Delete
    End If

    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")
    Dim nWeeks As Long: nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count
    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    Dim addedItems As Long: addedItems = 0
    If rmN > 0 Then
        Dim rmData As Variant: rmData = rmTbl.ListColumns(1).DataBodyRange.Resize(rmN, 11).Value
        Dim i As Long
        For i = 1 To rmN
            Dim rmCode As String: rmCode = Trim(CStr(rmData(i, 1)))
            Dim catVal As String: catVal = Trim(CStr(rmData(i, 4)))
            Dim ttafCodeVal As String: ttafCodeVal = Trim(CStr(rmData(i, 9)))
            Dim originVal As String: originVal = Trim(CStr(rmData(i, 11)))
            If POSheetNameForMaterial(catVal, originVal) = targetSheetName Then
                Call AppendPODraftRow(newSh, rmCode, ttafCodeVal, nWeeks)
                addedItems = addedItems + 1
            End If
        Next i
    End If
    If addedItems = 0 Then
        newSh.Cells(PO_DATA_START_ROW, 2).Value = "(No matching items)"
    End If

    ' ---- Update the print area to match the actual number of data rows ----
    Dim finalLastRow As Long
    finalLastRow = IIf(addedItems > 0, PO_DATA_START_ROW + addedItems - 1, PO_HDR_ROW)
    Call SetPODraftPrintArea(newSh, finalLastRow)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ClonePODraftLetterheadIfNeeded = True
End Function

' [Maintenance macro - safe to run anytime] (Re)applies, to all four
' PO_Draft_* sheets, the formatting that hides weeks with 0 order quantity
' (conditional formatting / number format) and the setting that excludes
' the base-week cell from the print area. Since
' SetupPODraftLetterheadLayout skips sheets already migrated to the new
' layout, run this macro directly if you want to add/update just this
' formatting after migration. Safe to run any number of times.
Sub ApplyPODraftZeroHiddenFormattingToAllSheets()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Application.ScreenUpdating = False

    Dim sheetNames As Variant
    sheetNames = Array("PO_Draft_Chemical", "PO_Draft_Hazardous", "PO_Draft_Substrate_JPN_CHN", "PO_Draft_Substrate_Poland")
    Dim si As Long, n As Long: n = 0
    For si = LBound(sheetNames) To UBound(sheetNames)
        Dim sh As Worksheet
        On Error Resume Next
        Set sh = Nothing
        Set sh = thisWb.Sheets(CStr(sheetNames(si)))
        On Error GoTo ErrHandler
        If Not sh Is Nothing Then
            Call ApplyPODraftZeroHiddenFormatting(sh)
            Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 5).End(xlUp).Row
            If lastRow < PO_HDR_ROW Then lastRow = PO_HDR_ROW
            Call SetPODraftPrintArea(sh, lastRow)
            n = n + 1
        End If
    Next si

    Application.ScreenUpdating = True
    MsgBox "Applied to " & n & " sheet(s). Please save and verify.", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred during processing: (" & Err.Number & ") " & Err.Description, vbCritical
End Sub

' Applies conditional-formatting color to the Firm (weeks 1-4)/Forecast
' (weeks 5-13) order-quantity cells (colored only when the order quantity
' is non-zero; a week of 0 gets no color and shows no number). Applied over
' a range (500 rows) well beyond the actual data row count, so rows added
' later by AddMaterial etc. are automatically covered too (no need to set
' up conditional formatting individually every time a row is added). Any
' existing conditional formatting is cleared before being reapplied, so
' it's safe to run any number of times.
Private Sub ApplyPODraftZeroHiddenFormatting(sh As Worksheet)
    Dim cfLastRow As Long: cfLastRow = PO_DATA_START_ROW + 500
    Dim firmFirstCol As Long: firmFirstCol = PO_FIRST_WEEK_COL
    Dim firmLastCol As Long: firmLastCol = PO_FIRST_WEEK_COL + 3
    Dim forecastFirstCol As Long: forecastFirstCol = PO_FIRST_WEEK_COL + 4
    Dim forecastLastCol As Long: forecastLastCol = PO_FIRST_WEEK_COL + PO_N_WEEKS - 1

    Dim allWeeksRng As Range
    Set allWeeksRng = sh.Range(sh.Cells(PO_DATA_START_ROW, firmFirstCol), sh.Cells(cfLastRow, forecastLastCol))
    allWeeksRng.NumberFormat = "0;-0;;@"
    ' Reset any cells that may have direct color left over from an earlier
    ' version (direct fill) or an emergency-recovery macro like
    ' FixMergedPODraftDataRows. Conditional formatting falls back to the
    ' original (directly-specified) formatting when its condition doesn't
    ' match, so leftover direct color would keep showing even for a 0 week.
    allWeeksRng.Interior.ColorIndex = xlNone
    allWeeksRng.Font.ColorIndex = xlAutomatic

    Dim firmRng As Range: Set firmRng = sh.Range(sh.Cells(PO_DATA_START_ROW, firmFirstCol), sh.Cells(cfLastRow, firmLastCol))
    firmRng.FormatConditions.Delete
    Dim firmAnchor As String: firmAnchor = ColLetter(firmFirstCol) & PO_DATA_START_ROW
    Dim fc1 As FormatCondition
    Set fc1 = firmRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(" & firmAnchor & "<>0," & firmAnchor & "<>"""")")
    fc1.Interior.Color = RGB(255, 193, 193)  ' Firm: FFC1C1
    fc1.Font.Color = RGB(192, 0, 0)          ' dark red: C00000

    Dim forecastRng As Range: Set forecastRng = sh.Range(sh.Cells(PO_DATA_START_ROW, forecastFirstCol), sh.Cells(cfLastRow, forecastLastCol))
    forecastRng.FormatConditions.Delete
    Dim forecastAnchor As String: forecastAnchor = ColLetter(forecastFirstCol) & PO_DATA_START_ROW
    Dim fc2 As FormatCondition
    Set fc2 = forecastRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(" & forecastAnchor & "<>0," & forecastAnchor & "<>"""")")
    fc2.Interior.Color = RGB(235, 241, 222)  ' Forecast: EBF1DE
    fc2.Font.Color = RGB(0, 97, 0)           ' dark green: 006100
End Sub

' Sets the print area for a PO_Draft_* sheet. In addition to
' SafetyStock/CurrentStock (columns F/G), the base-week (WeekIndex) input
' field (columns N/P of row PO_BASEWEEK_ROW) is also an internal-use cell
' not needed when printing/issuing the order form, so the H:U column print
' area is split around that row to exclude it.
Private Sub SetPODraftPrintArea(sh As Worksheet, lastRow As Long)
    sh.PageSetup.PrintArea = _
        "$A$1:$E$" & lastRow & "," & _
        "$H$1:$U$" & (PO_BASEWEEK_ROW - 1) & "," & _
        "$H$" & (PO_BASEWEEK_ROW + 1) & ":$U$" & lastRow
End Sub
