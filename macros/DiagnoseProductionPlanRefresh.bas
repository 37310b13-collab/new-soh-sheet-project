Attribute VB_Name = "DiagnoseProductionPlanRefresh"
Option Explicit

' ============================================================================
' DiagnoseProductionPlanRefresh module - READ-ONLY DIAGNOSTIC
'
' [Why] RefreshWeeklyBatches (RefreshData_ProductionPlan.bas) is throwing
' runtime error 9 ("Subscript out of range"), and WeeklyConsumption is
' showing 0 for BOM-based materials. Both symptoms point at one of
' RefreshWeeklyBatches's several Sheets()/ListObjects()/ListColumns()
' name-based lookups failing - but the macro's own error handler only
' reports the raw VBA error number, not WHICH specific lookup failed, and
' it aborts at the FIRST failure so later problems never even get a
' chance to show up in the same run.
'
' This macro re-runs every one of those lookups individually, each wrapped
' in its own error trap, and reports PASS/FAIL for all of them in a single
' MsgBox - so every problem shows up at once instead of one per run. It
' makes NO changes whatsoever (doesn't open any file, doesn't write any
' cell) - completely safe to run anytime, as many times as needed.
' ============================================================================

Sub DiagnoseProductionPlanRefresh()
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim report As String: report = ""
    Dim allOk As Boolean: allOk = True

    Dim ppGrid As ListObject
    Dim calWeeks As ListObject
    Dim solTbl As ListObject

    ' ---- 1. Production_Plan sheet + table ----
    Dim ppSheet As Worksheet
    On Error Resume Next
    Set ppSheet = Nothing
    Set ppSheet = thisWb.Sheets("Production_Plan")
    On Error GoTo 0
    If ppSheet Is Nothing Then
        report = report & "[FAIL] Sheets(""Production_Plan"") - sheet not found" & vbCrLf
        allOk = False
    Else
        report = report & "[OK]   Sheets(""Production_Plan"") found" & vbCrLf
        On Error Resume Next
        Set ppGrid = Nothing
        Set ppGrid = ppSheet.ListObjects("Production_Plan")
        On Error GoTo 0
        If ppGrid Is Nothing Then
            Dim ppTblNames As String: ppTblNames = ""
            Dim t1 As ListObject
            For Each t1 In ppSheet.ListObjects
                If Len(ppTblNames) > 0 Then ppTblNames = ppTblNames & ", "
                ppTblNames = ppTblNames & """" & t1.Name & """"
            Next t1
            report = report & "[FAIL] Production_Plan sheet's Table is not named ""Production_Plan"" - " & _
                "actual table(s) on this sheet: " & IIf(Len(ppTblNames) = 0, "(none)", ppTblNames) & vbCrLf
            allOk = False
        Else
            report = report & "[OK]   Table ""Production_Plan"" found (" & ppGrid.ListRows.Count & " rows, " & _
                ppGrid.ListColumns.Count & " columns)" & vbCrLf
            On Error Resume Next
            Dim interColIdx As Long: interColIdx = 0
            interColIdx = ppGrid.ListColumns("Intermediate").Index
            On Error GoTo 0
            If interColIdx = 0 Then
                Dim ppColNames As String: ppColNames = ""
                Dim ci As Long
                For ci = 1 To ppGrid.ListColumns.Count
                    If Len(ppColNames) > 0 Then ppColNames = ppColNames & ", "
                    ppColNames = ppColNames & """" & ppGrid.ListColumns(ci).Name & """"
                    If ci >= 5 Then
                        ppColNames = ppColNames & ", ..."
                        Exit For
                    End If
                Next ci
                report = report & "[FAIL] Table ""Production_Plan"" has no column named ""Intermediate"" - " & _
                    "first columns are: " & ppColNames & vbCrLf
                allOk = False
            Else
                report = report & "[OK]   Column ""Intermediate"" found (position " & interColIdx & ")" & vbCrLf
            End If
        End If
    End If

    ' ---- 2. Cal_Weeks sheet + table + WeekStart column ----
    Dim cwSheet As Worksheet
    On Error Resume Next
    Set cwSheet = Nothing
    Set cwSheet = thisWb.Sheets("Cal_Weeks")
    On Error GoTo 0
    If cwSheet Is Nothing Then
        report = report & "[FAIL] Sheets(""Cal_Weeks"") - sheet not found" & vbCrLf
        allOk = False
    Else
        On Error Resume Next
        Set calWeeks = Nothing
        Set calWeeks = cwSheet.ListObjects("Cal_Weeks")
        On Error GoTo 0
        If calWeeks Is Nothing Then
            report = report & "[FAIL] Cal_Weeks sheet's Table is not named ""Cal_Weeks""" & vbCrLf
            allOk = False
        Else
            report = report & "[OK]   Table ""Cal_Weeks"" found (" & calWeeks.ListRows.Count & " rows)" & vbCrLf
            On Error Resume Next
            Dim wsColIdx As Long: wsColIdx = 0
            wsColIdx = calWeeks.ListColumns("WeekStart").Index
            On Error GoTo 0
            If wsColIdx = 0 Then
                report = report & "[FAIL] Table ""Cal_Weeks"" has no column named ""WeekStart""" & vbCrLf
                allOk = False
            Else
                report = report & "[OK]   Column ""WeekStart"" found (position " & wsColIdx & ")" & vbCrLf
            End If
        End If
    End If

    ' ---- 3. Control_Panel sheet + T_SolutionNames table ----
    Dim cpSheet As Worksheet
    On Error Resume Next
    Set cpSheet = Nothing
    Set cpSheet = thisWb.Sheets("Control_Panel")
    On Error GoTo 0
    If cpSheet Is Nothing Then
        report = report & "[FAIL] Sheets(""Control_Panel"") - sheet not found" & vbCrLf
        allOk = False
    Else
        report = report & "[OK]   Sheets(""Control_Panel"") found" & vbCrLf
        On Error Resume Next
        Set solTbl = Nothing
        Set solTbl = cpSheet.ListObjects("T_SolutionNames")
        On Error GoTo 0
        If solTbl Is Nothing Then
            Dim cpTblNames As String: cpTblNames = ""
            Dim t2 As ListObject
            For Each t2 In cpSheet.ListObjects
                If Len(cpTblNames) > 0 Then cpTblNames = cpTblNames & ", "
                cpTblNames = cpTblNames & """" & t2.Name & """"
            Next t2
            report = report & "[FAIL] Control_Panel sheet has no Table named ""T_SolutionNames"" - " & _
                "actual table(s) on this sheet: " & IIf(Len(cpTblNames) = 0, "(none)", cpTblNames) & vbCrLf
            allOk = False
        Else
            report = report & "[OK]   Table ""T_SolutionNames"" found (" & solTbl.ListRows.Count & " rows)" & vbCrLf
        End If
    End If

    ' ---- 4. M_BOM sheet + table + PPGridRow column (used by RefreshBOM/WeeklyConsumption, checked here too since it's adjacent) ----
    Dim bomSheet As Worksheet
    On Error Resume Next
    Set bomSheet = Nothing
    Set bomSheet = thisWb.Sheets("M_BOM")
    On Error GoTo 0
    If bomSheet Is Nothing Then
        report = report & "[FAIL] Sheets(""M_BOM"") - sheet not found" & vbCrLf
        allOk = False
    Else
        Dim bomTbl As ListObject
        On Error Resume Next
        Set bomTbl = Nothing
        Set bomTbl = bomSheet.ListObjects("M_BOM")
        On Error GoTo 0
        If bomTbl Is Nothing Then
            report = report & "[FAIL] M_BOM sheet's Table is not named ""M_BOM""" & vbCrLf
            allOk = False
        Else
            report = report & "[OK]   Table ""M_BOM"" found (" & bomTbl.ListRows.Count & " rows)" & vbCrLf
            On Error Resume Next
            Dim ppgrIdx As Long: ppgrIdx = 0
            ppgrIdx = bomTbl.ListColumns("PPGridRow").Index
            On Error GoTo 0
            If ppgrIdx = 0 Then
                report = report & "[FAIL] Table ""M_BOM"" has no column named ""PPGridRow""" & vbCrLf
                allOk = False
            Else
                report = report & "[OK]   Column ""PPGridRow"" found (position " & ppgrIdx & ")" & vbCrLf
                ' Sample the first row's PPGridRow value - if every row shows 99999,
                ' that confirms the Production_Plan[Intermediate] structured reference
                ' is failing at the WORKSHEET FORMULA level even though the VBA-side
                ' checks above passed (e.g. a stale cached formula, or a genuinely
                ' different problem worth reporting separately).
                If bomTbl.ListRows.Count > 0 Then
                    Dim sampleVal As Variant
                    sampleVal = bomTbl.ListRows(1).Range.Cells(1, ppgrIdx).Value
                    report = report & "       Sample M_BOM row 1 PPGridRow value = " & CStr(sampleVal) & _
                        IIf(IsNumeric(sampleVal) And CDbl(sampleVal) = 99999, " <- 99999 means Production_Plan[Intermediate] MATCH is failing on the worksheet", "") & vbCrLf
                End If
            End If
        End If
    End If

    MsgBox "Diagnostic complete (no changes were made)." & vbCrLf & vbCrLf & report & vbCrLf & _
           IIf(allOk, "Everything checked out OK - if RefreshWeeklyBatches still errors, the problem is " & _
               "likely in the SOURCE file (Powder & Slurry & Pgm Plan) itself, not in this workbook. " & _
               "Please report the exact error number/text again along with this report.", _
               "At least one [FAIL] above is the root cause - please share this exact report."), _
           IIf(allOk, vbInformation, vbExclamation)
End Sub
