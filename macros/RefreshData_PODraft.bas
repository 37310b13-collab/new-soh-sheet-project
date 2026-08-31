Attribute VB_Name = "RefreshData_PODraft"
Option Explicit

' ============================================================================
' RefreshData_PODraft module
'
' Ongoing maintenance for the PO_Draft_* sheets' letterhead layout (TO/FROM/
' CC, Order Date/Issue Month/Firm Month, Revision, base week (WeekIndex)
' via the sheet-local named ranges "BaseWeek"/"PORevision", Firm/Forecast
' color-coding, print area). The one-time migration that originally built
' this layout out on PO_Draft_Hazardous and cloned it onto the other three
' PO_Draft_* sheets (SetupPODraftLetterheadLayout) has been run on the live
' file and removed from this module - see git history if the details of
' that migration are ever needed again.
'
' ApplyPODraftZeroHiddenFormattingToAllSheets : [Maintenance macro - safe to
'   run anytime] Re-applies the "hide weeks with 0 order quantity"
'   conditional formatting and the print area (which excludes the base-week
'   input row) to all four PO_Draft_* sheets. Useful after any manual
'   layout tweak, or after AddMaterial has added rows beyond what the
'   conditional formatting range already covers.
' ============================================================================

Private Const PO_HDR_ROW As Long = 26        ' last header row (data rows start right below). Corresponds to build_soh.py's PO_HDR_UOM_FIRM_ROW
Private Const PO_DATA_START_ROW As Long = 27 ' Corresponds to build_soh.py's PO_DATA_START_ROW
Private Const PO_FIRST_WEEK_COL As Long = 6  ' column F (SafetyStock/CurrentStock columns F/G were deleted - see RemovePODraftStockColumns)
Private Const PO_N_WEEKS As Long = 13
Private Const PO_BASEWEEK_ROW As Long = 13   ' the row holding the base-week (WeekIndex) input cell (used to exclude it from the print area)

' [Maintenance macro - safe to run anytime] (Re)applies, to all four
' PO_Draft_* sheets, the formatting that hides weeks with 0 order quantity
' (conditional formatting / number format) and the setting that excludes
' the base-week cell from the print area. Safe to run any number of times.
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

' Sets the print area for a PO_Draft_* sheet. The base-week (WeekIndex)
' input field (row PO_BASEWEEK_ROW) is an internal-use cell not needed when
' printing/issuing the order form, so the print area is split into two
' blocks (above/below that row) to exclude just that one row. (SafetyStock/
' CurrentStock no longer exist as separate columns - see
' RemovePODraftStockColumns - so there is no column-range gap to exclude
' any more; the print area now runs the full width from column A.)
Private Sub SetPODraftPrintArea(sh As Worksheet, lastRow As Long)
    Dim printColEnd As String: printColEnd = ColLetter(PO_FIRST_WEEK_COL + PO_N_WEEKS)  ' Total column
    sh.PageSetup.PrintArea = _
        "$A$1:$" & printColEnd & "$" & (PO_BASEWEEK_ROW - 1) & "," & _
        "$A$" & (PO_BASEWEEK_ROW + 1) & ":$" & printColEnd & "$" & lastRow
End Sub
