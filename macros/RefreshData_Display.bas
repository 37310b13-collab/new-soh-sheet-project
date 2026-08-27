Attribute VB_Name = "RefreshData_Display"
Option Explicit

' ============================================================================
' RefreshData_Display module
'
' HideInactiveIntermediates / ShowAllIntermediates / JumpToSelectedWeek
'
' None of these are data-refresh macros - they only operate on appearance
' (row show/hide, window scroll position).
'
' The Material_Detail sheet is laid out as vertical blocks per material
' (RM_Code), each with an "Intermediate name / No. of batches" row and a
' "Usage (kg)" row. For intermediates with no production planned for a
' while (six months, a year, etc.), this lets the user collapse those rows
' with a single button click, for a given period, to reduce the amount of
' scrolling. The material's total usage, TTAF stock (actual), self stock
' (actual), and total stock (end of week) rows are never hidden and always
' stay visible (so the stock level itself can always be checked at a glance).
'
' How it decides: for the weeks from "this week" through the
' specified number of months, if that intermediate's "No. of batches" row
' is all 0 (or blank), both of its rows (No. of batches row + Usage (kg)
' row) are hidden. If even one week is non-zero, it stays visible. Every
' run first re-shows all rows before re-evaluating, so re-running with a
' different period always produces a result matching the newly specified
' condition.
'
' Row position is determined by the label text (column C = "No. of
' batches"), so it automatically tracks row-count changes made on the
' build_soh.py side. However, the week-data start column
' (MD_WEEK_START_COL = column E) and header row (MD_HEADER_ROW = row 6,
' defined in RefreshData_Utilities) must be kept in sync with
' build_soh.py's WEEK_START_COL/MD_TABLE_ROW values (if the sheet layout
' changes, update these too).
'
' [Caution] HideInactiveIntermediates/ShowAllIntermediates need to be
' assigned to buttons on Material_Detail by hand (openpyxl can't create
' buttons automatically). JumpToSelectedWeek does nothing on its own once
' imported - it also needs a Worksheet_Change wired into each of the
' Dashboard/Material_Detail/T_SelfStock/T_TTAFStock sheets' own code
' modules. See the comment at the top of RefreshData_Utilities and
' docs/SOH_System_Guide.md for the exact steps/code.
' ============================================================================

Sub HideInactiveIntermediates()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets("Material_Detail")
    On Error GoTo 0
    If sh Is Nothing Then
        MsgBox "The Material_Detail sheet was not found.", vbExclamation
        Exit Sub
    End If

    Dim monthsStr As String
    monthsStr = InputBox("Hide intermediates with 0 batches every week, starting from this week, for how many months?" & _
        vbCrLf & "(e.g. 6 -> hide intermediates with no production planned for 6 months, 12 -> 1 year)" & _
        vbCrLf & "Please enter a number.", "Period to hide (months)", "6")
    If monthsStr = "" Then Exit Sub  ' cancelled
    If Not IsNumeric(monthsStr) Then
        MsgBox "Please enter a number.", vbExclamation
        Exit Sub
    End If
    Dim months As Double: months = CDbl(monthsStr)
    If months <= 0 Then
        MsgBox "Please enter a number of 1 or greater.", vbExclamation
        Exit Sub
    End If

    Dim nWeeks As Long
    nWeeks = wb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count

    Dim curWeek As Long: curWeek = WeekIndexForDate(wb, Date)
    Dim thresholdWeeks As Long
    thresholdWeeks = CLng(Application.WorksheetFunction.RoundUp(months * 52 / 12, 0))

    Dim endWeek As Long: endWeek = curWeek + thresholdWeeks - 1
    If endWeek > nWeeks Then endWeek = nWeeks
    Dim startCol As Long: startCol = MD_WEEK_START_COL + curWeek - 1
    Dim endCol As Long: endCol = MD_WEEK_START_COL + endWeek - 1
    If endCol < startCol Then
        MsgBox "No target weeks were found.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row

    ' Re-show all rows first, then re-evaluate (so re-running with a different period always matches the specified condition)
    If lastRow >= MD_HEADER_ROW Then sh.Rows(MD_HEADER_ROW & ":" & lastRow).Hidden = False

    Dim r As Long, hiddenCount As Long, shownCount As Long
    r = MD_HEADER_ROW + 1
    Do While r <= lastRow
        If sh.Cells(r, 3).Value = "No. of batches" Then
            Dim vals As Variant
            vals = sh.Range(sh.Cells(r, startCol), sh.Cells(r, endCol)).Value
            Dim allZero As Boolean: allZero = True
            Dim c As Long
            For c = 1 To UBound(vals, 2)
                If IsNumeric(vals(1, c)) Then
                    If CDbl(vals(1, c)) <> 0 Then
                        allZero = False
                        Exit For
                    End If
                End If
            Next c
            If allZero Then
                sh.Rows(r).Hidden = True
                sh.Rows(r + 1).Hidden = True   ' also hide the matching "Usage (kg)" row
                hiddenCount = hiddenCount + 1
            Else
                shownCount = shownCount + 1
            End If
            r = r + 2   ' skip past the No. of batches row + Usage (kg) row pair
        Else
            r = r + 1
        End If
    Loop

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Done." & vbCrLf & _
           "Hid intermediates with 0 batches every week for " & thresholdWeeks & " weeks starting from this week (week " & curWeek & ")." & vbCrLf & _
           "Hidden: " & hiddenCount & " / Shown: " & shownCount & vbCrLf & _
           "(The material name row and stock-related rows always stay visible)", vbInformation
End Sub

' Common routine, meant to be called when C1 (selected week) is entered on
' Dashboard/Material_Detail/T_SelfStock/T_TTAFStock. Rather than copying the
' selected week's value into a separate pinned column, this simply scrolls
' the window horizontally so the "real week-data column" always appears
' right next to the label column (right after the frozen pane). Because
' nothing is ever duplicated, there is no chance for numbers to disagree
' between sheets, Grid_Stock, etc.
' Called from each sheet's own Worksheet_Change with the target sheet, the
' resolved week-number cell (F1), and the week-data start column (9 =
' column I for Dashboard, 4 = column D for Material_Detail, 2 = column B
' for T_SelfStock/T_TTAFStock).
Public Sub JumpToSelectedWeek(sh As Worksheet, weekIndexCell As String, weekStartCol As Long)
    Dim wIdx As Variant
    wIdx = sh.Range(weekIndexCell).Value
    If Not IsNumeric(wIdx) Then Exit Sub
    Dim targetCol As Long
    targetCol = weekStartCol + CLng(wIdx) - 1
    On Error Resume Next
    ActiveWindow.ScrollColumn = targetCol
    On Error GoTo 0
End Sub

Sub ShowAllIntermediates()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets("Material_Detail")
    On Error GoTo 0
    If sh Is Nothing Then
        MsgBox "The Material_Detail sheet was not found.", vbExclamation
        Exit Sub
    End If
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row
    If lastRow >= MD_HEADER_ROW Then sh.Rows(MD_HEADER_ROW & ":" & lastRow).Hidden = False
    MsgBox "All intermediate rows are now shown again.", vbInformation
End Sub
