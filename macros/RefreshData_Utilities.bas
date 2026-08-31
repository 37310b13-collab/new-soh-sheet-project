Attribute VB_Name = "RefreshData_Utilities"
Option Explicit

Public Const MD_HEADER_ROW As Long = 6       ' Material_Detail: header row. Corresponds to build_soh.py's MD_TABLE_ROW
Public Const MD_WEEK_START_COL As Long = 4   ' Material_Detail: week-data start column (column D). Corresponds to build_soh.py's WEEK_START_COL
Public Const SS_TABLE_ROW As Long = 5        ' CSAstock/TTAFstock: header row (week labels). Corresponds to build_soh.py's SS_TABLE_ROW
Public Const DASH_DATA_START_ROW As Long = 7 ' Dashboard: start row of material data. Corresponds to build_soh.py's DATA_START_ROW

' ============================================================================
' SOH management workbook VBA macros - overview (this comment is written only in the RefreshData_Utilities module)
'
' Purpose: complete the monthly data refresh entirely within Excel (VBA), without any external environment like Python.
'
' [Module layout] All macros used to live in a single large standard module
' (RefreshData.bas); for maintainability they are now split by function into
' the 8 modules below. Macro names and behavior are unchanged from before the split.
'
'   RefreshData_Utilities      : This module. Public Consts (MD_HEADER_ROW etc.) and
'                                 helper functions shared across multiple modules
'                                 (BuildNameIndex/NormalizeText/
'                                 WeekIndexForDate/ColLetter). All other modules
'                                 depend on this one.
'   RefreshData_ProductionPlan : RefreshWeeklyBatches (imports "Powder & Slurry & Pgm Plan"
'                                 and updates Production_Plan)
'   RefreshData_BOM             : RefreshBOM (imports "Raw Material - Look Up"
'                                 and updates M_BOM / Material_Detail's breakdown rows)
'   RefreshData_StockActuals    : RefreshSelfStock / RefreshTTAFStock (import self/TTAF actual stock)
'   RefreshData_Shipments       : RefreshShipments (imports the CSA Report's Shipping Schedule)
'   RefreshData_Display         : HideInactiveIntermediates / ShowAllIntermediates /
'                                 JumpToSelectedWeek (display toggles only; never changes data)
'   RefreshData_MaterialMgmt    : AddMaterial / RemoveMaterial / RemoveIntermediate
'                                 (add/remove materials and intermediates)
'   RefreshData_PODraft         : ApplyPODraftZeroHiddenFormattingToAllSheets and
'                                 other PO_Draft_* letterhead-layout maintenance macros
'
' For a detailed explanation of each macro, see the comment at the top of the
' module where it is implemented. Every macro only updates the sheet(s) it is
' responsible for. It never touches content you entered by hand during normal
' operation, such as T_Shipments, T_OpeningStock, or SafetyStock_Qty.
'
' [Caution] JumpToSelectedWeek does not do anything on its own once
' imported - it must additionally be wired up via a Worksheet_Change
' event pasted directly into the code module of each of the "Dashboard",
' "Material_Detail", "CSAstock", and "TTAFstock" sheets (a standard
' module's code never fires from a cell edit). The routine itself lives
' in the RefreshData_Display module; see docs/SOH_System_Guide.md for the
' exact Worksheet_Change code to paste into each sheet.
'
' [About performance] Every Refresh* macro reads its target range as a single
' array (Range.Value) once, instead of reading an external file's cells one
' at a time via .Cells(r,c).Value, and afterward only touches the in-memory
' array. Likewise, writes to Production_Plan (intermediate name -> row number) and
' CSAstock_Log/TTAFstock_Log ((RM_Code, week's Monday) -> row number)
' build a Dictionary once at the start of the run and look values up in it,
' instead of calling .Find() or scanning every row on each write (see
' BuildNameIndex in this module and BuildStockRowIndex in
' RefreshData_StockActuals). This is a countermeasure against a real bug that
' was reported multiple times, where Excel actually force-quit (caused by
' cell-by-cell reads/writes or a full-row scan on every write becoming
' extremely slow from the accumulated COM round-trips).
'
' [About writing to M_BOM] RefreshBOM first assembles the new BOM content
' read from Look Up into an in-memory Dictionary, then replaces the whole of
' M_BOM in a single block write (Range.Formula = a 2D array). It used to add
' rows one at a time via ListRows.Add, but because M_BOM is referenced by a
' very large number of formulas (WeeklyConsumption, Material_Detail, Production_Plan's
' pass-through formulas, etc.), adding a single row triggered a dependency
' recheck across all of them, and this was reported to slow to a near-freeze
' once the row count passed roughly a thousand - hence this design.
'
' [Why existing rows are not updated via .DataBodyRange] Running
' RefreshBOM/RefreshWeeklyBatches was reported to sometimes raise the error
' "Object variable or With block variable not set" (error 91). The cause is
' a known Excel/VBA quirk where ListObject.DataBodyRange can unreliably
' return Nothing (especially right after adding a new row with
' ListRows.Add). New-row code already used the return value of .ListRows.Add
' (newRow.Range) in each Sub, which avoided this problem, but the code path
' that updates existing rows still used the unstable
' .DataBodyRange.Cells(...) form. Everything has now been unified to use
' .ListRows(rowNumber).Range.Cells(...), which works reliably even
' immediately after a row is added.
'
' [Why DisplayAlerts is suppressed] Even after the .DataBodyRange fix, a
' case was reported where srcWb became Nothing right after opening the
' source file, and the same error (91) recurred at the srcWb.Close cleanup
' step. The source files (Powder & Slurry & Pgm Plan, Raw Material - Look
' Up, etc.) are confirmed to trigger a "read-only recommended" confirmation
' dialog when opened manually; with Application.DisplayAlerts left True,
' this dialog is believed to also appear during VBA's Workbooks.Open call
' and stall execution (waiting for a response instead of moving to the next
' line, or the object coming back in an unexpected state). Application.
' DisplayAlerts is now set to False before Workbooks.Open to suppress such
' dialogs, and restored to True during cleanup. As a precaution, the
' srcWb.Close call itself is still guarded with If Not srcWb Is Nothing Then.
'
' [About the two-layer structure of CSAstock/TTAFstock]
' RefreshSelfStock/RefreshTTAFStock never write to the visible
' CSAstock/TTAFstock sheets at all. They write to the hidden
' CSAstock_Log/TTAFstock_Log (a raw log keyed by count date), and the
' visible sheet is built entirely from formulas (a material x week grid)
' that recompute from that log every time. The WeekIndex column on the
' _Log sheet side is a formula column calculated automatically from the
' Date column (RefreshSelfStock/RefreshTTAFStock only write Date, never
' WeekIndex). This ensures that advancing Cal_Weeks!B1 (AnchorYear) never
' causes already-recorded actuals to be misdisplayed as "another week's
' data." See the top of the RefreshData_StockActuals module for details.
'
' [Note: when an entirely new substrate/Cat code is added]
'   RefreshWeeklyBatches automatically adds rows to Production_Plan and M_BOM, but
'   it does not add new substrate codes to M_RawMaterials (the raw material
'   master) via VBA.
'   If you notice a new substrate code (one that doesn't appear on the TTAF
'   stock actuals sheet or Dashboard), please add a row to M_RawMaterials
'   manually (RM_Code, TTAF_Code, Description, Supplier, Category="Substrate").
'
' [Note] This environment cannot actually run and verify VBA. Please test it
'        in your own Excel. If you hit an error, let us know what it says.
' ============================================================================

' Builds a Dictionary of value -> row number for the given column (colName)
' of tbl, once. The old implementation used .Find(), whose default Excel
' behavior is a case-insensitive search. To preserve that same behavior,
' this uses a Dictionary with CompareMode=vbTextCompare (case-insensitive)
' (treating case as significant would fail to recognize spelling variants
' such as "TSP-049" vs "tsp-049" as the same intermediate, leading to a
' separate bug where duplicate rows kept accumulating on every run).
Public Function BuildNameIndex(tbl As ListObject, colName As String) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbTextCompare
    Dim n As Long: n = tbl.ListRows.Count
    If n = 1 Then
        Dim colPos As Long: colPos = tbl.ListColumns(colName).Index
        idx(Trim(CStr(tbl.ListRows(1).Range.Cells(1, colPos).Value))) = 1
    ElseIf n > 1 Then
        Dim data As Variant
        data = tbl.ListColumns(colName).DataBodyRange.Value
        Dim i As Long
        For i = 1 To n
            ' Without Trim(), if a cell on the M_RawMaterials side has stray
            ' leading/trailing whitespace mixed in, comparing it against the
            ' caller's already-Trim()'d string will fail to match, silently
            ' dropping just that row from the lookup (this actually happened
            ' as a bug where Ester Film/Original Towel/PP Film were not
            ' reflected by RefreshSelfStock).
            Dim k As String: k = Trim(CStr(data(i, 1)))
            If Not idx.Exists(k) Then idx(k) = i
        Next i
    End If
    Set BuildNameIndex = idx
End Function

' Normalization (keep only alphanumerics, uppercased) used to absorb
' spelling variants (case, symbols, spacing differences) in material/
' intermediate names. Used in common across multiple modules -
' RefreshData_BOM, RefreshData_StockActuals, RefreshData_Shipments - when
' matching an external file's names against the M_RawMaterials/Production_Plan side.
Public Function NormalizeText(s As String) As String
    Dim i As Long, ch As String, result As String
    s = UCase(s)
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            result = result & ch
        End If
    Next i
    NormalizeText = result
End Function

' Looks up Cal_Weeks by date to get the WeekIndex. Shared routine used by
' both RefreshData_StockActuals (determining the target week for actuals
' import) and RefreshData_Display (HideInactiveIntermediates' "this week"
' determination).
Public Function WeekIndexForDate(wb As Workbook, d As Date) As Long
    Dim calTbl As ListObject: Set calTbl = wb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks")
    Dim n As Long: n = calTbl.ListRows.Count
    If n > 0 Then
        ' Columns: 1=WeekIndex, 2=WeekStart, 6=WeekEnd. Read all columns together in one array read.
        Dim data As Variant
        data = calTbl.DataBodyRange.Value
        Dim i As Long
        For i = 1 To n
            If d >= CDate(data(i, 2)) And d <= CDate(data(i, 6)) Then
                WeekIndexForDate = CLng(data(i, 1))
                Exit Function
            End If
        Next i
    End If
    WeekIndexForDate = 1 ' fall back to Week1 if not found
End Function

' Converts a column number (e.g. 28) to a column letter (e.g. AB). A pure
' calculation with no dependency on any worksheet. Shared routine used by
' both RefreshData_BOM (InsertIntermediateRowPair) and
' RefreshData_MaterialMgmt (the Append*/Delete* family).
Public Function ColLetter(colNum As Long) As String
    Dim s As String, n As Long, r As Long
    n = colNum
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr(65 + r) & s
        n = (n - r - 1) \ 26
    Loop
    ColLetter = s
End Function

' Returns the reference string to use for the base-week cell in PO_Draft_*
' sheet formulas (order quantity / stock lookups). If the sheet-local
' (local-scope) named range "BaseWeek" is defined, its name is returned
' as-is (so even if the cell moves, e.g. from P7 to P13, only the name's
' target needs fixing and formulas never need to be rewritten. Because
' RefreshData_MaterialMgmt.bas's AppendPODraftRow used to hard-code $P$7
' directly, any existing sheet whose base-week cell had been moved, e.g.
' manually from P7 to P13, would only have its newly-appended rows break).
' A sheet with no "BaseWeek" name at all (e.g. one built by hand outside
' the normal build_soh.py/migration flow) falls back to $P$7 as before.
Public Function BaseWeekRef(sh As Worksheet) As String
    On Error Resume Next
    Dim nm As Name: Set nm = sh.Names("BaseWeek")
    On Error GoTo 0
    If Not nm Is Nothing Then
        BaseWeekRef = "BaseWeek"
    Else
        BaseWeekRef = "$P$7"
    End If
End Function
