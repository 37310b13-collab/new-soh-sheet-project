Attribute VB_Name = "RefreshData"
Option Explicit

' ============================================================================
' RefreshData モジュール
'
' 目的: Python等の外部環境を使わず、Excel(VBA)だけで毎月のデータ更新を完結させる。
'   RefreshWeeklyBatches : 「Powder & Slurry & Pgm Plan」(毎月改版)を選択すると、
'                          PP_Grid（生産計画バッチ数）を更新する。化学原料シート
'                          (シート全体で1材料, row3のB列がCHEM-)とsubstrate/フィルム等の
'                          シート(1シート内に複数ブロックが並ぶ場合を含む)の両方に対応。
'                          substrateブロックの"Usage per day"(完成品1個あたり使用量)は
'                          M_BOMにも自動反映する。
'   RefreshBOM           : 「Usage from Production Engineering」を選択すると、
'                          M_BOM（化学原料の原単位）を更新する。
'   RefreshSelfStock      : 「Raw materials daily check」(自社倉庫の現物確認)を選択すると、
'                          T_SelfStock（自社在庫実績）を更新する。
'   RefreshTTAFStock      : 「CSA Report」を選択すると、T_TTAFStock（TTAF在庫実績）を更新する。
'
' いずれも、それぞれ対応するシートだけを更新します。T_Shipments・T_OpeningStock・
' T_StockCount・SafetyStock_Qty等、運用中に手入力した内容には一切触れません。
'
' 【注意: 完全に新しいsubstrate/Catコードが増えた場合】
'   RefreshWeeklyBatchesはPP_GridとM_BOMには自動で行を追加しますが、
'   M_RawMaterials(原材料マスタ)への新規substrateコードの追加はVBAでは行いません。
'   新しいsubstrateコード(TTAF在庫実績シートやDashboardに現れないコード)に気づいたら、
'   M_RawMaterialsシートに手動で1行追加してください(RM_Code, TTAF_Code, Description,
'   Supplier, Category="Substrate")。
'
' 【注意】この環境ではVBAを実際に実行して検証できません。貴社のExcelで動作確認を
'        お願いします。エラーが出た場合は内容を教えてください。
'
' 【導入方法】
'   1. SOH_Master.xlsm（マクロ有効ブックとして保存）を開く
'   2. Alt+F11 → 「ファイル」→「ファイルのインポート」→ このファイル(RefreshData.bas)を選択
'      （コピー＆貼り付けで導入する場合は、1行目の Attribute VB_Name = "..." を必ず削除してから
'       貼り付けてください。この行は貼り付けでは使えず、含めるとコンパイルエラーになります）
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
    Dim bomTbl As ListObject: Set bomTbl = thisWb.Sheets("M_BOM").ListObjects("M_BOM")

    ' WeekStart(日付のシリアル値) -> PP_Grid内の列番号(データ範囲内, 1=Intermediate,2=Week1,...)
    Dim weekColByDate As Object: Set weekColByDate = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To calWeeks.ListRows.Count
        Dim wkDate As Variant
        wkDate = calWeeks.ListColumns("WeekStart").DataBodyRange.Cells(i, 1).Value
        weekColByDate(CLng(CDate(wkDate))) = i + 1
    Next i

    Dim updatedCells As Long, newInterRows As Long, bomUpdated As Long
    updatedCells = 0
    newInterRows = 0
    bomUpdated = 0

    ' 化学原料シート(row3のB列がCHEM-で始まり、row2のD列以降に週初日が3つ以上ある、
    ' シート全体で1材料)か、substrate/フィルム等のシート(row2以降のどこかにコード+週初日の
    ' ヘッダーが複数ブロック並ぶ)かを、Pythonの抽出スクリプトと同じロジックでシートごとに判定する。
    Dim sh As Worksheet
    For Each sh In srcWb.Worksheets
        Dim usedCols As Long: usedCols = sh.UsedRange.Columns.Count
        Dim topDateCount As Long: topDateCount = 0
        Dim cc As Long
        For cc = 4 To usedCols
            If IsDate(sh.Cells(2, cc).Value) Then topDateCount = topDateCount + 1
        Next cc
        Dim topCode As String: topCode = CStr(sh.Cells(3, 2).Value)

        If topDateCount >= 3 And Left(topCode, 4) = "CHEM" Then
            Call ProcessMaterialSheet(sh, ppGrid, weekColByDate, updatedCells, newInterRows)
        Else
            Call ProcessSubstrateSheet(sh, ppGrid, bomTbl, weekColByDate, updatedCells, newInterRows, bomUpdated)
        End If
    Next sh

    srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "PP_Grid / M_BOM を更新しました。" & vbCrLf & _
           "更新セル数(PP_Grid): " & updatedCells & vbCrLf & _
           "新規追加した中間体/完成品(Cat)コード: " & newInterRows & vbCrLf & _
           "更新/追加したsubstrate原単位(M_BOM): " & bomUpdated & vbCrLf & vbCrLf & _
           "(参考) 全く新しい中間体×原材料の組み合わせがある場合、M_BOM側にも" & vbCrLf & _
           "RefreshBOMで追加してください（Grid_RequirementはM_BOM/PP_Gridを直接参照するため、" & vbCrLf & _
           "追加した行はその場で自動反映されます）。", _
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

' substrate/フィルム等のシート用: 1シート内にコード+週初日ヘッダーを持つブロックが
' 複数(または1つ)並んでいる構造に対応する。各ブロックは
'   header行   : (空欄), SSコード等, (空欄), 週初日(日付, D列以降)
'   header+1行 : (空欄), 説明, TTAF品番, 週番号
'   header+2行~: (空欄), 完成品コード(Cat)+"No. of batches"+週次バッチ数(=受注数量)、
'                (空欄), "Usage per day"+1個あたり使用量+週次使用量 …
' の並び。完成品コード(Cat)が空欄の場合(Ester Film/PP Filmのように1ブロック=1品目のシート)は、
' ブロック自身のコード(SSコード等)を中間体名として使う。
Private Sub ProcessSubstrateSheet(sh As Worksheet, ppGrid As ListObject, bomTbl As ListObject, _
        weekColByDate As Object, ByRef updatedCells As Long, ByRef newInterRows As Long, ByRef bomUpdated As Long)
    Dim usedRows As Long, usedCols As Long
    usedRows = sh.UsedRange.Rows.Count
    usedCols = sh.UsedRange.Columns.Count
    If usedRows < 5 Then Exit Sub

    Dim ri As Long
    ri = 1
    Do While ri <= usedRows - 3
        Dim dateCols As Object: Set dateCols = CreateObject("Scripting.Dictionary")
        Dim c As Long
        For c = 4 To usedCols
            If IsDate(sh.Cells(ri, c).Value) Then dateCols(c) = CLng(CDate(sh.Cells(ri, c).Value))
        Next c

        Dim blockCode As String: blockCode = Trim(CStr(sh.Cells(ri, 2).Value))
        If dateCols.Count >= 3 And Len(blockCode) > 0 And Len(blockCode) <= 12 And Left(blockCode, 4) <> "CHEM" Then
            ' ブロック発見。データはheader+2行目から
            Dim r As Long: r = ri + 2
            Dim currentInter As String: currentInter = ""
            Do While r <= usedRows
                ' 次のブロックのヘッダー(コード+日付)に到達したら終了
                Dim nextDateCount As Long: nextDateCount = 0
                Dim c2 As Long
                For c2 = 4 To usedCols
                    If IsDate(sh.Cells(r, c2).Value) Then nextDateCount = nextDateCount + 1
                Next c2
                Dim nextCode As String: nextCode = Trim(CStr(sh.Cells(r, 2).Value))
                If nextDateCount >= 3 And Len(nextCode) > 0 And Len(nextCode) <= 12 And Left(nextCode, 4) <> "CHEM" Then
                    Exit Do
                End If

                Dim lbl3 As String: lbl3 = LCase(Trim(CStr(sh.Cells(r, 3).Value)))
                Dim lbl2 As String: lbl2 = LCase(Trim(CStr(sh.Cells(r, 2).Value)))

                If Left(lbl3, 14) = "no. of batches" Then
                    Dim rawInter As String: rawInter = Trim(CStr(sh.Cells(r, 2).Value))
                    If Len(rawInter) = 0 Then rawInter = blockCode
                    currentInter = rawInter
                    If Len(currentInter) > 0 Then
                        Dim ppRowIndex As Long
                        ppRowIndex = FindOrAddIntermediateRow(ppGrid, currentInter, newInterRows)
                        Dim keyVariant As Variant
                        For Each keyVariant In dateCols.Keys
                            If weekColByDate.Exists(dateCols(keyVariant)) Then
                                Dim colIdx As Long: colIdx = weekColByDate(dateCols(keyVariant))
                                Dim v As Double: v = 0
                                If IsNumeric(sh.Cells(r, keyVariant).Value) Then v = sh.Cells(r, keyVariant).Value
                                ppGrid.DataBodyRange.Cells(ppRowIndex, colIdx).Value = v
                                updatedCells = updatedCells + 1
                            End If
                        Next keyVariant
                    End If
                ElseIf lbl2 = "usage per day" And Len(currentInter) > 0 Then
                    Dim rateVal As Variant: rateVal = sh.Cells(r, 3).Value
                    If IsNumeric(rateVal) Then
                        Call UpsertBomRow(bomTbl, currentInter, blockCode, CDbl(rateVal), bomUpdated)
                    End If
                ElseIf lbl2 = "total weekly usage" Or lbl3 = "total weekly usage" Then
                    r = r + 1
                    Exit Do
                End If
                r = r + 1
            Loop
            ri = r
        Else
            ri = ri + 1
        End If
    Loop
End Sub

Private Sub UpsertBomRow(bomTbl As ListObject, interName As String, rmCode As String, rate As Double, ByRef bomUpdated As Long)
    Dim i As Long
    For i = 1 To bomTbl.ListRows.Count
        If bomTbl.ListColumns("Intermediate").DataBodyRange.Cells(i, 1).Value = interName And _
           bomTbl.ListColumns("RM_Code").DataBodyRange.Cells(i, 1).Value = rmCode Then
            bomTbl.ListColumns("RM_Qty_Per_Batch").DataBodyRange.Cells(i, 1).Value = rate
            bomUpdated = bomUpdated + 1
            Exit Sub
        End If
    Next i
    Dim newRow As ListRow
    Set newRow = bomTbl.ListRows.Add
    newRow.Range.Cells(1, 1).Value = interName
    newRow.Range.Cells(1, 2).Value = rmCode
    newRow.Range.Cells(1, 3).Value = rate
    bomUpdated = bomUpdated + 1
End Sub

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
    msg = msg & vbCrLf & vbCrLf & "新規追加した組み合わせも、Grid_RequirementがM_BOMを直接参照しているため" & vbCrLf & "その場で自動反映されます。"
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

' ============================================================================
' RefreshSelfStock / RefreshTTAFStock
'
'   RefreshSelfStock : 「Raw materials daily check」(自社倉庫の現物確認シート、ファイル名に
'                      DD.MM.YYYY形式の日付を含む)を選択すると、T_SelfStockにその週の実績を
'                      追加/更新する。
'   RefreshTTAFStock : 「CSA Report」(ファイル名に日付を含む)を選択すると、その中の
'                      "PIVOT SOH TTAF"シートからT_TTAFStockにその週の実績を追加/更新する。
'
' どちらも対象週は、ファイル名から読み取った日付をCal_Weeksと照合して自動判定します。
' 既存の(原材料コード, 週)の組み合わせがあれば値を上書き、無ければ新しい行として末尾に追加します
' （Dashboardの「直近実績」表示は、各原材料コードについてテーブル内で最後に見つかった行を
' 採用する仕組みのため、必ず日付が新しい順に取り込んでください。過去のファイルを後から
' 取り込むと最新表示がずれる可能性があります）。
' ============================================================================

Sub RefreshSelfStock()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel ファイル (*.xlsx),*.xlsx", , _
        "Raw materials daily check（自社在庫）ファイルを選択してください")
    If srcPath = False Then Exit Sub

    Dim srcWb As Workbook
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set srcWb = Workbooks.Open(CStr(srcPath), ReadOnly:=True, UpdateLinks:=False)
    Dim reportDate As Date: reportDate = ExtractDateFromName(CStr(srcPath))

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim selfTbl As ListObject: Set selfTbl = thisWb.Sheets("T_SelfStock").ListObjects("T_SelfStock")
    Dim wIdx As Long: wIdx = WeekIndexForDate(thisWb, reportDate)

    Dim sh As Worksheet: Set sh = srcWb.Sheets("Stock")
    Dim r As Long, added As Long, updated As Long
    added = 0: updated = 0
    For r = 9 To 200
        Dim code As String
        code = Trim(CStr(sh.Cells(r, 3).Value))
        If Left(code, 4) = "CHEM" Then
            Dim v As Variant: v = sh.Cells(r, 10).Value
            If IsNumeric(v) Then
                Call UpsertStockRow(selfTbl, code, wIdx, reportDate, CDbl(v), added, updated)
            End If
        End If
    Next r

    srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "T_SelfStock を更新しました。" & vbCrLf & "対象週: " & wIdx & " (" & Format(reportDate, "yyyy-mm-dd") & ")" & vbCrLf & _
           "追加: " & added & " 件、更新: " & updated & " 件", vbInformation
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "更新処理でエラーが発生しました: " & Err.Description, vbCritical
End Sub

Sub RefreshTTAFStock()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel ファイル (*.xlsx),*.xlsx", , _
        "CSA Report（TTAF在庫）ファイルを選択してください")
    If srcPath = False Then Exit Sub

    Dim srcWb As Workbook
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set srcWb = Workbooks.Open(CStr(srcPath), ReadOnly:=True, UpdateLinks:=False)
    Dim reportDate As Date: reportDate = ExtractDateFromName(CStr(srcPath))

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim ttafTbl As ListObject: Set ttafTbl = thisWb.Sheets("T_TTAFStock").ListObjects("T_TTAFStock")
    Dim wIdx As Long: wIdx = WeekIndexForDate(thisWb, reportDate)

    Dim sh As Worksheet: Set sh = srcWb.Sheets("PIVOT SOH TTAF")
    Dim r As Long, added As Long, updated As Long
    added = 0: updated = 0
    r = 5
    Do While True
        Dim code As Variant
        code = sh.Cells(r, 1).Value
        If IsEmpty(code) Or CStr(code) = "Grand Total" Then Exit Do
        If Left(CStr(code), 4) = "CHEM" Then
            Dim v As Variant: v = sh.Cells(r, 4).Value
            If IsNumeric(v) Then
                Call UpsertStockRow(ttafTbl, CStr(code), wIdx, reportDate, CDbl(v), added, updated)
            End If
        End If
        r = r + 1
    Loop

    srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "T_TTAFStock を更新しました。" & vbCrLf & "対象週: " & wIdx & " (" & Format(reportDate, "yyyy-mm-dd") & ")" & vbCrLf & _
           "追加: " & added & " 件、更新: " & updated & " 件", vbInformation
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "更新処理でエラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Function ExtractDateFromName(path As String) As Date
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d{2})\.(\d{2})\.(\d{4})"
    Dim m As Object
    Set m = re.Execute(path)
    If m.Count = 0 Then
        Err.Raise vbObjectError + 1, , "ファイル名から日付(DD.MM.YYYY)を読み取れませんでした。"
    End If
    ExtractDateFromName = DateSerial(CInt(m(0).SubMatches(2)), CInt(m(0).SubMatches(1)), CInt(m(0).SubMatches(0)))
End Function

Private Function WeekIndexForDate(wb As Workbook, d As Date) As Long
    Dim calTbl As ListObject: Set calTbl = wb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks")
    Dim i As Long
    For i = 1 To calTbl.ListRows.Count
        Dim wkStart As Date, wkEnd As Date
        wkStart = calTbl.ListColumns("WeekStart").DataBodyRange.Cells(i, 1).Value
        wkEnd = calTbl.ListColumns("WeekEnd").DataBodyRange.Cells(i, 1).Value
        If d >= wkStart And d <= wkEnd Then
            WeekIndexForDate = calTbl.ListColumns("WeekIndex").DataBodyRange.Cells(i, 1).Value
            Exit Function
        End If
    Next i
    WeekIndexForDate = 1 ' 見つからない場合はWeek1にフォールバック
End Function

Private Sub UpsertStockRow(tbl As ListObject, code As String, wIdx As Long, d As Date, v As Double, ByRef added As Long, ByRef updated As Long)
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If tbl.ListColumns(1).DataBodyRange.Cells(i, 1).Value = code And _
           tbl.ListColumns(2).DataBodyRange.Cells(i, 1).Value = wIdx Then
            tbl.ListColumns(3).DataBodyRange.Cells(i, 1).Value = d
            tbl.ListColumns(4).DataBodyRange.Cells(i, 1).Value = v
            updated = updated + 1
            Exit Sub
        End If
    Next i
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    newRow.Range.Cells(1, 1).Value = code
    newRow.Range.Cells(1, 2).Value = wIdx
    newRow.Range.Cells(1, 3).Value = d
    newRow.Range.Cells(1, 4).Value = v
    added = added + 1
End Sub
