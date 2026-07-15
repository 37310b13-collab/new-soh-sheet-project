Attribute VB_Name = "RefreshData"
Option Explicit

' ============================================================================
' RefreshData モジュール
'
' 目的: Python等の外部環境を使わず、Excel(VBA)だけで毎月のデータ更新を完結させる。
'   RefreshWeeklyBatches : 「Powder & Slurry & Pgm Plan」(毎月改版)を選択すると、
'                          PP_Grid（生産計画バッチ数）を更新する。
'   RefreshBOM           : 「Usage from Production Engineering」を選択すると、
'                          M_BOM（原単位）を更新する。
'
' どちらも、T_Shipments・T_OpeningStock・T_StockCount・SafetyStock_Qty等、
' 運用中に手入力した内容には一切触れません（PP_GridとM_BOMだけを更新します）。
'
' 【注意】この環境ではVBAを実際に実行して検証できません。貴社のExcelで動作確認を
'        お願いします。エラーが出た場合は内容を教えてください。
'
' 【導入方法】
'   1. SOH_Master.xlsm（マクロ有効ブックとして保存）を開く
'   2. Alt+F11 → 挿入 → 標準モジュール → このファイルの中身を貼り付けて保存
'   3. Alt+F8 でマクロ一覧から RefreshWeeklyBatches / RefreshBOM を実行、
'      または任意のシートに図形を挿入して「マクロの登録」で割り当てる
' ============================================================================

Sub RefreshWeeklyBatches()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel ファイル (*.xlsx),*.xlsx", , _
        "Powder & Slurry & Pgm Plan（最新版）を選択してください")
    If srcPath = False Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(CStr(srcPath), ReadOnly:=True, UpdateLinks:=False)

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim ppGrid As ListObject: Set ppGrid = thisWb.Sheets("PP_Grid").ListObjects("PP_Grid")
    Dim calWeeks As ListObject: Set calWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks")

    ' WeekStart(日付のシリアル値) -> PP_Grid内の列番号(データ範囲内, 1=Intermediate,2=Week1,...)
    Dim weekColByDate As Object: Set weekColByDate = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To calWeeks.ListRows.Count
        Dim wkDate As Variant
        wkDate = calWeeks.ListColumns("WeekStart").DataBodyRange.Cells(i, 1).Value
        weekColByDate(CLng(CDate(wkDate))) = i + 1
    Next i

    Dim updatedCells As Long, newInterRows As Long
    updatedCells = 0
    newInterRows = 0

    Dim sh As Worksheet
    For Each sh In srcWb.Worksheets
        Dim snLower As String: snLower = LCase(sh.Name)
        If InStr(snLower, "substr") = 0 And InStr(snLower, "gpf") = 0 Then
            Call ProcessMaterialSheet(sh, ppGrid, weekColByDate, updatedCells, newInterRows)
        End If
    Next sh

    srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "PP_Grid を更新しました。" & vbCrLf & _
           "更新セル数: " & updatedCells & vbCrLf & _
           "新規追加した中間体: " & newInterRows & vbCrLf & vbCrLf & _
           "(参考) 全く新しい中間体×原材料の組み合わせがある場合、M_BOM側にも" & vbCrLf & _
           "RefreshBOMで追加してください。それでもCalc_Demandに反映されない場合はご連絡ください。", _
           vbInformation
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "更新処理でエラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Sub ProcessMaterialSheet(sh As Worksheet, ppGrid As ListObject, _
        weekColByDate As Object, ByRef updatedCells As Long, ByRef newInterRows As Long)
    Dim usedRows As Long, usedCols As Long
    usedRows = sh.UsedRange.Rows.Count
    usedCols = sh.UsedRange.Columns.Count
    If usedRows < 5 Then Exit Sub

    ' row2: 週初日(日付), col D(4)以降
    Dim dateCols As Object: Set dateCols = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 4 To usedCols
        If IsDate(sh.Cells(2, c).Value) Then
            dateCols(c) = CLng(CDate(sh.Cells(2, c).Value))
        End If
    Next c
    If dateCols.Count < 3 Then Exit Sub

    ' row3 col B: 材料コード(CHEM-xxxx)であることを確認（材料シートのみ処理）
    Dim matCode As String
    matCode = CStr(sh.Cells(3, 2).Value)
    If Left(matCode, 4) <> "CHEM" Then Exit Sub

    Dim r As Long
    For r = 4 To usedRows
        Dim lbl As String
        lbl = LCase(Trim(CStr(sh.Cells(r, 3).Value)))
        If Left(lbl, 14) = "no. of batches" Then
            Dim rawInter As String
            rawInter = Trim(CStr(sh.Cells(r, 2).Value))
            If Len(rawInter) > 0 Then
                Dim ppRowIndex As Long
                ppRowIndex = FindOrAddIntermediateRow(ppGrid, rawInter, newInterRows)

                Dim keyVariant As Variant
                For Each keyVariant In dateCols.Keys
                    If weekColByDate.Exists(dateCols(keyVariant)) Then
                        Dim colIdx As Long
                        colIdx = weekColByDate(dateCols(keyVariant))
                        Dim v As Double
                        v = 0
                        If IsNumeric(sh.Cells(r, keyVariant).Value) Then v = sh.Cells(r, keyVariant).Value
                        ppGrid.DataBodyRange.Cells(ppRowIndex, colIdx).Value = v
                        updatedCells = updatedCells + 1
                    End If
                Next keyVariant
            End If
        End If
    Next r
End Sub

Private Function FindOrAddIntermediateRow(ppGrid As ListObject, interName As String, ByRef newInterRows As Long) As Long
    Dim foundCell As Range
    Set foundCell = ppGrid.ListColumns("Intermediate").DataBodyRange.Find(What:=interName, LookAt:=xlWhole)
    If foundCell Is Nothing Then
        Dim newRow As ListRow
        Set newRow = ppGrid.ListRows.Add
        newRow.Range.Cells(1, 1).Value = interName
        newInterRows = newInterRows + 1
        FindOrAddIntermediateRow = newRow.Index
    Else
        FindOrAddIntermediateRow = foundCell.Row - ppGrid.HeaderRowRange.Row
    End If
End Function

Sub RefreshBOM()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel ファイル (*.xlsx),*.xlsx", , _
        "Usage from Production Engineering を選択してください")
    If srcPath = False Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(CStr(srcPath), ReadOnly:=True, UpdateLinks:=False)

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim bomTbl As ListObject: Set bomTbl = thisWb.Sheets("M_BOM").ListObjects("M_BOM")
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    ' 材料名(正規化) -> RM_Code
    Dim descIndex As Object: Set descIndex = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To rmTbl.ListRows.Count
        Dim code As String, desc As String, key As String
        code = CStr(rmTbl.ListColumns("RM_Code").DataBodyRange.Cells(i, 1).Value)
        desc = CStr(rmTbl.ListColumns("Description").DataBodyRange.Cells(i, 1).Value)
        key = NormalizeText(desc)
        If Len(key) > 0 And Not descIndex.Exists(key) Then descIndex(key) = code
    Next i

    ' (Intermediate|RM_Code) -> M_BOM内の行番号(データ範囲内)
    Dim pairIndex As Object: Set pairIndex = CreateObject("Scripting.Dictionary")
    For i = 1 To bomTbl.ListRows.Count
        Dim inter As String, rmc As String
        inter = CStr(bomTbl.ListColumns("Intermediate").DataBodyRange.Cells(i, 1).Value)
        rmc = CStr(bomTbl.ListColumns("RM_Code").DataBodyRange.Cells(i, 1).Value)
        pairIndex(inter & "|" & rmc) = i
    Next i

    Dim updated As Long, added As Long, unresolved As String
    updated = 0: added = 0: unresolved = ""

    Dim sheetNames As Variant: sheetNames = Array("Slurry", "Powder")
    Dim sIdx As Integer
    For sIdx = LBound(sheetNames) To UBound(sheetNames)
        Dim shName As String: shName = CStr(sheetNames(sIdx))
        If SheetExists(srcWb, shName) Then
            Call ProcessBomSheet(srcWb.Sheets(shName), bomTbl, descIndex, pairIndex, updated, added, unresolved)
        End If
    Next sIdx

    srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "M_BOM を更新しました。" & vbCrLf & "更新: " & updated & " 件、新規追加: " & added & " 件"
    If Len(unresolved) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "原材料コードが見つからず未反映の材料名:" & vbCrLf & unresolved
    End If
    msg = msg & vbCrLf & vbCrLf & "新規追加した組み合わせは、Calc_Demandの再生成が必要な場合があります。"
    MsgBox msg, vbInformation
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "更新処理でエラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Sub ProcessBomSheet(sh As Worksheet, bomTbl As ListObject, descIndex As Object, _
        pairIndex As Object, ByRef updated As Long, ByRef added As Long, ByRef unresolved As String)
    Dim lastRow As Long, lastCol As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    lastCol = sh.Cells(2, sh.Columns.Count).End(xlToLeft).Column

    Dim r As Long, c As Long
    For r = 4 To lastRow
        Dim interName As String
        interName = Trim(CStr(sh.Cells(r, 1).Value))
        If Len(interName) > 0 Then
            For c = 2 To lastCol
                Dim matName As String: matName = Trim(CStr(sh.Cells(2, c).Value))
                Dim v As Variant: v = sh.Cells(r, c).Value
                If Len(matName) > 0 And IsNumeric(v) Then
                    If CDbl(v) <> 0 Then
                        Dim mkey As String: mkey = NormalizeText(matName)
                        If descIndex.Exists(mkey) Then
                            Dim rmCode As String: rmCode = descIndex(mkey)
                            Dim pk As String: pk = interName & "|" & rmCode
                            If pairIndex.Exists(pk) Then
                                Dim rowN As Long: rowN = pairIndex(pk)
                                bomTbl.ListColumns("RM_Qty_Per_Batch").DataBodyRange.Cells(rowN, 1).Value = CDbl(v)
                                updated = updated + 1
                            Else
                                Dim newRow As ListRow
                                Set newRow = bomTbl.ListRows.Add
                                newRow.Range.Cells(1, 1).Value = interName
                                newRow.Range.Cells(1, 2).Value = rmCode
                                newRow.Range.Cells(1, 3).Value = CDbl(v)
                                pairIndex(pk) = newRow.Index
                                added = added + 1
                            End If
                        Else
                            If InStr(unresolved, matName) = 0 Then unresolved = unresolved & matName & "; "
                        End If
                    End If
                End If
            Next c
        End If
    Next r
End Sub

Private Function NormalizeText(s As String) As String
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

Private Function SheetExists(wb As Workbook, sName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets(sName)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function
