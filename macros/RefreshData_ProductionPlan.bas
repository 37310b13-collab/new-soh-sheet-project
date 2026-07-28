Attribute VB_Name = "RefreshData_ProductionPlan"
Option Explicit

' ============================================================================
' RefreshData_ProductionPlan モジュール
'
'   RefreshWeeklyBatches : 「Powder & Slurry & Pgm Plan」(毎月改版)を選択すると、
'                          PP_Grid（生産計画バッチ数）を更新する。化学原料シート
'                          (シート全体で1材料, row3のB列がCHEM-)とsubstrate/フィルム等の
'                          シート(1シート内に複数ブロックが並ぶ場合を含む)の両方に対応。
'                          substrateブロックの"Usage per day"(完成品1個あたり使用量)は
'                          M_BOMにも自動反映する。新しい中間体が見つかれば自動的に
'                          PP_Gridに行を追加する(FindOrAddIntermediateRow)。
'
' 全体の設計方針(パフォーマンス・DataBodyRange・DisplayAlerts等)はRefreshData_Utilities
' モジュール冒頭のコメントを参照してください。
' ============================================================================

Sub RefreshWeeklyBatches()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel ファイル (*.xlsx),*.xlsx", , _
        "Powder & Slurry & Pgm Plan（最新版）を選択してください")
    If srcPath = False Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' 「読み取り専用を推奨」設定のファイルだと、DisplayAlerts=Trueのままでは
    ' Workbooks.Open時に確認ダイアログが表示され、応答待ちで処理が不安定になる
    ' (最終的にsrcWbが正しく取得できない不具合の原因になっていた)ため抑制する。
    Application.DisplayAlerts = False

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(CStr(srcPath), ReadOnly:=True, UpdateLinks:=False)

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim ppGrid As ListObject: Set ppGrid = thisWb.Sheets("PP_Grid").ListObjects("PP_Grid")
    Dim calWeeks As ListObject: Set calWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks")
    Dim bomTbl As ListObject: Set bomTbl = thisWb.Sheets("M_BOM").ListObjects("M_BOM")

    ' WeekStart(日付のシリアル値) -> PP_Grid内の列番号(データ範囲内, 1=Intermediate,2=Week1,...)
    ' Cal_Weeksは104行程度なので影響は小さいが、他の箇所と同じく1回の配列読み込みに統一する。
    Dim weekColByDate As Object: Set weekColByDate = CreateObject("Scripting.Dictionary")
    Dim calN As Long: calN = calWeeks.ListRows.Count
    If calN > 0 Then
        Dim calWeekStartData As Variant
        calWeekStartData = calWeeks.ListColumns("WeekStart").DataBodyRange.Value
        Dim i As Long
        For i = 1 To calN
            weekColByDate(CLng(CDate(calWeekStartData(i, 1)))) = i + 1
        Next i
    End If

    ' PP_Grid(Intermediate->行番号)とM_BOM(Intermediate|RM_Code->行番号)のインデックスを
    ' 実行の最初に1回だけ作る。以前はPP_Gridは.Find()を中間体ごとに毎回、M_BOMは
    ' UpsertBomRow呼び出しのたびに全711行をセル単位でスキャンしており、これが
    ' 「RefreshWeeklyBatches実行時に必ず強制終了する」不具合の直接の原因だった
    ' （前回の見直しで見つけていながら直し忘れていた箇所）。
    Dim ppIdx As Object: Set ppIdx = BuildNameIndex(ppGrid, "Intermediate")
    Dim bomIdx As Object: Set bomIdx = BuildPairIndex(bomTbl)

    Dim updatedCells As Long, newInterRows As Long, bomUpdated As Long
    updatedCells = 0
    newInterRows = 0
    bomUpdated = 0

    ' 化学原料シート(row3のB列がCHEM-で始まり、row2のD列以降に週初日が3つ以上ある、
    ' シート全体で1材料)か、substrate/フィルム等のシート(row2以降のどこかにコード+週初日の
    ' ヘッダーが複数ブロック並ぶ)かを、Pythonの抽出スクリプトと同じロジックでシートごとに判定する。
    '
    ' 【重要】シートのセルを1つずつ.Cells(r,c).Valueで読むと、Excelの内部的な「使用範囲
    ' (UsedRange)」が実データより大きく認識されているケース(書式設定の残骸等でよくある)や、
    ' シート数(40枚超)と掛け合わさった際に、COM通信の呼び出し回数が数十万〜数百万回に
    ' 達し、Excelが長時間「応答なし」になり強制終了したように見える不具合の原因となる。
    ' そのため、各シートのデータは1回だけ配列にまとめて読み込み(sh.Range(...).Value)、
    ' 以降はその配列(メモリ上)だけを参照する。また使用範囲の異常な肥大化に備え上限も設ける。
    Const MAX_SCAN_ROWS As Long = 500
    Const MAX_SCAN_COLS As Long = 200

    Dim sh As Worksheet
    For Each sh In srcWb.Worksheets
        Dim usedRowsTop As Long, usedColsTop As Long
        usedRowsTop = sh.UsedRange.Rows.Count
        usedColsTop = sh.UsedRange.Columns.Count
        If usedRowsTop > MAX_SCAN_ROWS Then usedRowsTop = MAX_SCAN_ROWS
        If usedColsTop > MAX_SCAN_COLS Then usedColsTop = MAX_SCAN_COLS
        If usedRowsTop < 3 Or usedColsTop < 4 Then GoTo NextSheet

        Dim row2Data As Variant
        row2Data = sh.Range(sh.Cells(2, 1), sh.Cells(2, usedColsTop)).Value
        Dim topDateCount As Long: topDateCount = 0
        Dim cc As Long
        For cc = 4 To usedColsTop
            If IsDate(row2Data(1, cc)) Then topDateCount = topDateCount + 1
        Next cc
        Dim topCode As String: topCode = CStr(sh.Cells(3, 2).Value)

        If topDateCount >= 3 And Left(topCode, 4) = "CHEM" Then
            Call ProcessMaterialSheet(sh, usedRowsTop, usedColsTop, ppGrid, ppIdx, weekColByDate, updatedCells, newInterRows)
        Else
            Call ProcessSubstrateSheet(sh, usedRowsTop, usedColsTop, ppGrid, ppIdx, bomTbl, bomIdx, weekColByDate, updatedCells, newInterRows, bomUpdated)
        End If
NextSheet:
    Next sh

    ' srcWbが既にNothingになっているケース(取込元ファイル側の自動処理等で、開いた
    ' 直後にワークブックが閉じられてしまう場合がある)でも、後始末処理自体が
    ' 「オブジェクト変数が設定されていません」で落ちないようにガードする。
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

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
    ' 【重要】On Error Resume Next はErr オブジェクトを自動的にクリアしてしまう(VBAの仕様)ため、
    ' 後始末処理より前に、エラー番号・内容を必ず変数へ退避しておく。これを怠ると、
    ' 下のMsgBoxが常に「(空欄)」を表示してしまい、本当のエラー原因が一切分からなくなる
    ' (実際にこの不具合が発生し、原因調査ができない状態になっていたため修正)。
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "更新処理でエラーが発生しました: (" & errNum & ") " & errMsg, vbCritical
End Sub

Private Sub ProcessMaterialSheet(sh As Worksheet, usedRows As Long, usedCols As Long, ppGrid As ListObject, ppIdx As Object, _
        weekColByDate As Object, ByRef updatedCells As Long, ByRef newInterRows As Long)
    If usedRows < 5 Then Exit Sub

    ' シート全体を1回だけ配列として読み込み、以降はメモリ上の配列(data)だけを参照する
    ' （.Cells(r,c).Valueをループ内で毎回呼ぶとCOM通信が積み重なり非常に遅くなるため）
    Dim data As Variant
    data = sh.Range(sh.Cells(1, 1), sh.Cells(usedRows, usedCols)).Value

    ' row2: 週初日(日付), col D(4)以降
    Dim dateCols As Object: Set dateCols = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 4 To usedCols
        If IsDate(data(2, c)) Then
            dateCols(c) = CLng(CDate(data(2, c)))
        End If
    Next c
    If dateCols.Count < 3 Then Exit Sub

    ' row3 col B: 材料コード(CHEM-xxxx)であることを確認（材料シートのみ処理）
    Dim matCode As String
    matCode = CStr(data(3, 2))
    If Left(matCode, 4) <> "CHEM" Then Exit Sub

    Dim r As Long
    For r = 4 To usedRows
        Dim lbl As String
        lbl = LCase(Trim(CStr(data(r, 3))))
        If Left(lbl, 14) = "no. of batches" Then
            Dim rawInter As String
            rawInter = Trim(CStr(data(r, 2)))
            If Len(rawInter) > 0 Then
                Dim ppRowIndex As Long
                ppRowIndex = FindOrAddIntermediateRow(ppGrid, ppIdx, rawInter, newInterRows)

                Dim keyVariant As Variant
                For Each keyVariant In dateCols.Keys
                    If weekColByDate.Exists(dateCols(keyVariant)) Then
                        Dim colIdx As Long
                        colIdx = weekColByDate(dateCols(keyVariant))
                        Dim v As Double
                        v = 0
                        If IsNumeric(data(r, keyVariant)) Then v = data(r, keyVariant)
                        ppGrid.ListRows(ppRowIndex).Range.Cells(1, colIdx).Value = v
                        updatedCells = updatedCells + 1
                    End If
                Next keyVariant
            End If
        End If
    Next r
End Sub

' ppIdx(中間体名->PP_Grid内の行番号)を使い、.Find()を毎回呼ばずに済ませる。
' 新規追加時はppIdxにもその場で登録し、同じ実行内で同じ中間体が再度出てきても
' 重複追加せず既存行を再利用できるようにする。
Private Function FindOrAddIntermediateRow(ppGrid As ListObject, ppIdx As Object, interName As String, ByRef newInterRows As Long) As Long
    If ppIdx.Exists(interName) Then
        FindOrAddIntermediateRow = ppIdx(interName)
    Else
        Dim newRow As ListRow
        Set newRow = ppGrid.ListRows.Add
        newRow.Range.Cells(1, 1).Value = interName
        newInterRows = newInterRows + 1
        FindOrAddIntermediateRow = newRow.Index
        ppIdx(interName) = newRow.Index
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
Private Sub ProcessSubstrateSheet(sh As Worksheet, usedRows As Long, usedCols As Long, ppGrid As ListObject, ppIdx As Object, _
        bomTbl As ListObject, bomIdx As Object, _
        weekColByDate As Object, ByRef updatedCells As Long, ByRef newInterRows As Long, ByRef bomUpdated As Long)
    If usedRows < 5 Then Exit Sub

    ' シート全体を1回だけ配列として読み込み、以降はメモリ上の配列(data)だけを参照する
    ' （.Cells(r,c).Valueをループ内で毎回呼ぶとCOM通信が積み重なり非常に遅くなるため）
    Dim data As Variant
    data = sh.Range(sh.Cells(1, 1), sh.Cells(usedRows, usedCols)).Value

    Dim ri As Long
    ri = 1
    Do While ri <= usedRows - 3
        Dim dateCols As Object: Set dateCols = CreateObject("Scripting.Dictionary")
        Dim c As Long
        For c = 4 To usedCols
            If IsDate(data(ri, c)) Then dateCols(c) = CLng(CDate(data(ri, c)))
        Next c

        Dim blockCode As String: blockCode = Trim(CStr(data(ri, 2)))
        If dateCols.Count >= 3 And Len(blockCode) > 0 And Len(blockCode) <= 12 And Left(blockCode, 4) <> "CHEM" Then
            ' ブロック発見。データはheader+2行目から
            Dim r As Long: r = ri + 2
            Dim currentInter As String: currentInter = ""
            Do While r <= usedRows
                ' 次のブロックのヘッダー(コード+日付)に到達したら終了
                Dim nextDateCount As Long: nextDateCount = 0
                Dim c2 As Long
                For c2 = 4 To usedCols
                    If IsDate(data(r, c2)) Then nextDateCount = nextDateCount + 1
                Next c2
                Dim nextCode As String: nextCode = Trim(CStr(data(r, 2)))
                If nextDateCount >= 3 And Len(nextCode) > 0 And Len(nextCode) <= 12 And Left(nextCode, 4) <> "CHEM" Then
                    Exit Do
                End If

                Dim lbl3 As String: lbl3 = LCase(Trim(CStr(data(r, 3))))
                Dim lbl2 As String: lbl2 = LCase(Trim(CStr(data(r, 2))))

                If Left(lbl3, 14) = "no. of batches" Then
                    Dim rawInter As String: rawInter = Trim(CStr(data(r, 2)))
                    If Len(rawInter) = 0 Then rawInter = blockCode
                    currentInter = rawInter
                    If Len(currentInter) > 0 Then
                        Dim ppRowIndex As Long
                        ppRowIndex = FindOrAddIntermediateRow(ppGrid, ppIdx, currentInter, newInterRows)
                        Dim keyVariant As Variant
                        For Each keyVariant In dateCols.Keys
                            If weekColByDate.Exists(dateCols(keyVariant)) Then
                                Dim colIdx As Long: colIdx = weekColByDate(dateCols(keyVariant))
                                Dim v As Double: v = 0
                                If IsNumeric(data(r, keyVariant)) Then v = data(r, keyVariant)
                                ppGrid.ListRows(ppRowIndex).Range.Cells(1, colIdx).Value = v
                                updatedCells = updatedCells + 1
                            End If
                        Next keyVariant
                    End If
                ElseIf lbl2 = "usage per day" And Len(currentInter) > 0 Then
                    Dim rateVal As Variant: rateVal = data(r, 3)
                    If IsNumeric(rateVal) Then
                        Call UpsertBomRow(bomTbl, bomIdx, currentInter, blockCode, CDbl(rateVal), bomUpdated)
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

' bomIdx("Intermediate|RM_Code"->行番号)を使い、呼び出すたびにM_BOM全行をセル単位で
' スキャンするのを避ける（この線形スキャンが「RefreshWeeklyBatches実行時に必ず強制終了する」
' 不具合の直接の原因だった）。新規追加時はbomIdxにもその場で登録する。
Private Sub UpsertBomRow(bomTbl As ListObject, bomIdx As Object, interName As String, rmCode As String, rate As Double, ByRef bomUpdated As Long)
    Dim key As String: key = interName & "|" & rmCode
    If bomIdx.Exists(key) Then
        Dim rowN As Long: rowN = bomIdx(key)
        bomTbl.ListRows(rowN).Range.Cells(1, 3).Value = rate
        bomUpdated = bomUpdated + 1
    Else
        Dim newRow As ListRow
        Set newRow = bomTbl.ListRows.Add
        newRow.Range.Cells(1, 1).Value = interName
        newRow.Range.Cells(1, 2).Value = rmCode
        newRow.Range.Cells(1, 3).Value = rate
        bomIdx(key) = newRow.Index
        bomUpdated = bomUpdated + 1
    End If
End Sub
