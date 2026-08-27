Attribute VB_Name = "RefreshData_ProductionPlan"
Option Explicit

' ============================================================================
' RefreshData_ProductionPlan モジュール
'
'   RefreshWeeklyBatches : 「Powder & Slurry & Pgm Plan」(毎月改版)を選択すると、
'                          PP_Grid（生産計画バッチ数）を更新する。ファイルは単一シートで、
'                          B列=中間体/完成品(Catalyst)/Solutionの名前、週初日ヘッダー行(C列
'                          以降に日付が並ぶ行)を自動検出し、それ以降の行が名前+週次数量。
'                          1バッチあたりの原単位(M_BOM)はこのファイルには含まれておらず、
'                          RefreshBOM（Raw Material - Look Up）側の担当。
'
'   【行の種類の自動判定について】中間体(Slurry/Powder)なのか完成品(Catalyst)なのか
'   Solutionなのかは、名前の接頭辞と「Solution名リスト」(Control_PanelシートのT_SolutionNames
'   テーブル)で機械的に判定する。TSP-/TPP-/TSZ-/TVS-/VSP-で始まれば中間体、T_SolutionNames
'   に載っている名前(20P・SH等の略称)ならSolution(略称は自動的に正式名SOL-xxxへ変換)、
'   どちらでもなければ完成品(Catalyst)として扱う(消去法)。行の追加・削除や新しいSolutionの
'   追加があっても、この判定はどこにも行番号を持たないため、ファイル側の行が増減しても
'   マクロ側のメンテナンスは一切不要（新しいSolutionだけはT_SolutionNamesに1行追加で対応）。
'
' 全体の設計方針(パフォーマンス・DataBodyRange・DisplayAlerts等)はRefreshData_Utilities
' モジュール冒頭のコメントを参照してください。
' ============================================================================

Sub RefreshWeeklyBatches()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel ファイル (*.xlsx),*.xlsx", , _
        "Powder & Slurry & Pgm Plan（最新版）を選択してください")
    If srcPath = False Then Exit Sub

    Dim srcWb As Workbook
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' 「読み取り専用を推奨」設定のファイルだと、DisplayAlerts=Trueのままでは
    ' Workbooks.Open時に確認ダイアログが表示され、応答待ちで処理が不安定になる
    ' (最終的にsrcWbが正しく取得できない不具合の原因になっていた)ため抑制する。
    Application.DisplayAlerts = False

    Set srcWb = Workbooks.Open(CStr(srcPath), ReadOnly:=True, UpdateLinks:=False)

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim ppGrid As ListObject: Set ppGrid = thisWb.Sheets("PP_Grid").ListObjects("PP_Grid")
    Dim calWeeks As ListObject: Set calWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks")

    ' WeekStart(日付のシリアル値) -> PP_Grid内の列番号(データ範囲内, 1=Intermediate,2=Week1,...)
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

    Dim ppIdx As Object: Set ppIdx = BuildNameIndex(ppGrid, "Intermediate")

    ' Solution名リスト(略称->正式名)を1回だけ読み込む。CompareMode=vbTextCompareで
    ' 大文字小文字の違い(20p/20P等)を吸収する。
    Dim solutionAlias As Object: Set solutionAlias = CreateObject("Scripting.Dictionary")
    solutionAlias.CompareMode = vbTextCompare
    Dim solTbl As ListObject: Set solTbl = thisWb.Sheets("Control_Panel").ListObjects("T_SolutionNames")
    Dim solN As Long: solN = solTbl.ListRows.Count
    If solN > 0 Then
        Dim solData As Variant: solData = solTbl.ListColumns(1).DataBodyRange.Resize(solN, 2).Value
        Dim si As Long
        For si = 1 To solN
            Dim aliasKey As String: aliasKey = Trim(CStr(solData(si, 1)))
            If Len(aliasKey) > 0 Then solutionAlias(aliasKey) = Trim(CStr(solData(si, 2)))
        Next si
    End If

    ' 対象シートを探す(単一シート形式。念のため複数シートあっても構造(週初日の見出し行)が
    ' 合うものを自動判定する。1つのファイルにつき対象は1シートのみ)。
    Dim foundSheet As Worksheet
    Dim sh As Worksheet
    For Each sh In srcWb.Worksheets
        If IsWeeklyPlanSheet(sh) Then
            Set foundSheet = sh
            Exit For
        End If
    Next sh
    If foundSheet Is Nothing Then
        Err.Raise vbObjectError + 2, , "「Powder & Slurry & Pgm Plan」の週次データが見つかりませんでした。ファイル形式をご確認ください。"
    End If

    Const MAX_SCAN_ROWS As Long = 500
    Const MAX_SCAN_COLS As Long = 200
    Dim usedRows As Long, usedCols As Long
    usedRows = foundSheet.UsedRange.Rows.Count
    usedCols = foundSheet.UsedRange.Columns.Count
    If usedRows > MAX_SCAN_ROWS Then usedRows = MAX_SCAN_ROWS
    If usedCols > MAX_SCAN_COLS Then usedCols = MAX_SCAN_COLS

    ' シート全体を1回だけ配列として読み込み、以降はメモリ上の配列(data)だけを参照する
    ' （.Cells(r,c).Valueをループ内で毎回呼ぶとCOM通信が積み重なり非常に遅くなるため）
    Dim data As Variant
    data = foundSheet.Range(foundSheet.Cells(1, 1), foundSheet.Cells(usedRows, usedCols)).Value

    Dim hdrRow As Long: hdrRow = FindDateHeaderRow(data, usedRows, usedCols)
    If hdrRow = 0 Then
        Err.Raise vbObjectError + 3, , "週初日の見出し行が見つかりませんでした。ファイル形式をご確認ください。"
    End If

    Dim dateCols As Object: Set dateCols = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 3 To usedCols
        If IsDate(data(hdrRow, c)) Then dateCols(c) = CLng(CDate(data(hdrRow, c)))
    Next c

    Dim updatedCells As Long, newInterRows As Long
    updatedCells = 0
    newInterRows = 0

    ' 【重要】PP_Gridの行には2種類ある。①このファイルに直接載っている中間体/完成品/Solution
    ' (通常の行。週次バッチ数は値として書き込む)と、②M_BOM経由で他の中間体から算出する
    ' 「パススルー中間体」(週次バッチ数はSUMPRODUCT数式。このファイルには行として登場しない)。
    ' そのため、テーブル全体を配列で読み込んで書き戻す方式(以前試した版)は、パススルー
    ' 中間体の行の数式まで「その時点の計算結果の値」で上書きしてしまい、数式を破壊する
    ' 事故になる。書き込みは、このファイルに実際に登場した行だけに絞る必要がある。
    '
    ' また、新規中間体行の追加(ListRows.Add)を1件ずつループで呼ぶと、PP_Gridのように
    ' 多数の数式から参照される重量級テーブルでは1件ごとに依存関係の再チェックが走り、
    ' 新規行が多い場合(名称の突き合わせ漏れ等で本来一致するはずの行が軒並み「新規」と
    ' 判定された場合を含む)にExcelが応答なしになる(以前のRefreshBOMのM_BOMと同じ不具合)。
    ' そのため新規行はまず件数をすべて数えてから、テーブルを1回のResizeでまとめて拡張し、
    ' 中間体名の列だけ1回の配列書き込みで埋める。
    Dim canonNames As Object: Set canonNames = CreateObject("Scripting.Dictionary")
    Dim newNames As Object: Set newNames = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = hdrRow + 1 To usedRows
        Dim rawName As String: rawName = Trim(CStr(data(r, 2)))
        If Len(rawName) = 0 Then GoTo NextRowNames

        Dim canonName As String
        If solutionAlias.Exists(rawName) Then
            canonName = solutionAlias(rawName)
        Else
            canonName = rawName
        End If
        canonNames(r) = canonName

        If Not ppIdx.Exists(canonName) And Not newNames.Exists(canonName) Then
            newNames(canonName) = True
        End If
NextRowNames:
    Next r

    If newNames.Count > 0 Then
        Dim oldRowCount As Long: oldRowCount = ppGrid.ListRows.Count
        ppGrid.Resize ppGrid.Range.Resize(ppGrid.Range.Rows.Count + newNames.Count, ppGrid.Range.Columns.Count)
        Dim newNameArr() As Variant
        ReDim newNameArr(1 To newNames.Count, 1 To 1)
        Dim ni As Long: ni = 0
        Dim nameKey As Variant
        For Each nameKey In newNames.Keys
            ni = ni + 1
            newNameArr(ni, 1) = nameKey
            ppIdx(CStr(nameKey)) = oldRowCount + ni
            newInterRows = newInterRows + 1
        Next nameKey
        ppGrid.ListColumns(1).DataBodyRange.Resize(newNames.Count, 1).Offset(oldRowCount, 0).Value = newNameArr
    End If

    ' このファイルに登場した行だけを対象に、行ごと(セルごとではなく)にまとめて読み込み→
    ' メモリ上で該当週だけ書き換え→行ごと書き戻す(パススルー中間体の行には一切触れない)。
    Dim touchedRows As Object: Set touchedRows = CreateObject("Scripting.Dictionary")
    For r = hdrRow + 1 To usedRows
        If Not canonNames.Exists(r) Then GoTo NextRow
        Dim ppRowIndex As Long: ppRowIndex = ppIdx(canonNames(r))
        If Not touchedRows.Exists(ppRowIndex) Then
            touchedRows(ppRowIndex) = ppGrid.ListRows(ppRowIndex).Range.Value
        End If
        Dim rowArr As Variant: rowArr = touchedRows(ppRowIndex)

        Dim keyVariant As Variant
        For Each keyVariant In dateCols.Keys
            If weekColByDate.Exists(dateCols(keyVariant)) Then
                Dim colIdx As Long: colIdx = weekColByDate(dateCols(keyVariant))
                Dim v As Double: v = 0
                If IsNumeric(data(r, keyVariant)) Then v = data(r, keyVariant)
                rowArr(1, colIdx) = v
                updatedCells = updatedCells + 1
            End If
        Next keyVariant
        touchedRows(ppRowIndex) = rowArr
NextRow:
    Next r

    Dim rowKey As Variant
    For Each rowKey In touchedRows.Keys
        ppGrid.ListRows(CLng(rowKey)).Range.Value = touchedRows(rowKey)
    Next rowKey

    ' srcWbが既にNothingになっているケース(取込元ファイル側の自動処理等で、開いた
    ' 直後にワークブックが閉じられてしまう場合がある)でも、後始末処理自体が
    ' 「オブジェクト変数が設定されていません」で落ちないようにガードする。
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "PP_Grid を更新しました。" & vbCrLf & _
           "更新セル数: " & updatedCells & vbCrLf & _
           "新規追加した中間体/完成品(Cat)/Solutionコード: " & newInterRows & vbCrLf & vbCrLf & _
           "（参考）1バッチあたりの原単位(M_BOM)はこのマクロの対象外です。全く新しい" & vbCrLf & _
           "組み合わせがある場合はRefreshBOMで追加してください。", _
           vbInformation
    Exit Sub

ErrHandler:
    ' 【重要】On Error Resume Next はErr オブジェクトを自動的にクリアしてしまう(VBAの仕様)ため、
    ' 後始末処理より前に、エラー番号・内容を必ず変数へ退避しておく。これを怠ると、
    ' 下のMsgBoxが常に「(空欄)」を表示してしまい、本当のエラー原因が一切分からなくなる。
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "更新処理でエラーが発生しました: (" & errNum & ") " & errMsg, vbCritical
End Sub

' シートの先頭付近(10行以内)に、C列以降に日付が5つ以上並ぶ行があるかどうかで
' 「週次計画シート」らしいかを判定する(軽い範囲だけを見るので、40シート程度あっても
' パフォーマンス上の問題にはならない)。
Private Function IsWeeklyPlanSheet(sh As Worksheet) As Boolean
    Dim r As Long, c As Long
    Dim maxC As Long: maxC = Application.WorksheetFunction.Min(20, sh.UsedRange.Columns.Count)
    Dim maxR As Long: maxR = Application.WorksheetFunction.Min(10, sh.UsedRange.Rows.Count)
    For r = 1 To maxR
        Dim dateCount As Long: dateCount = 0
        For c = 3 To maxC
            If IsDate(sh.Cells(r, c).Value) Then dateCount = dateCount + 1
        Next c
        If dateCount >= 5 Then
            IsWeeklyPlanSheet = True
            Exit Function
        End If
    Next r
    IsWeeklyPlanSheet = False
End Function

' 配列化済みのシートデータ(data)から、週初日ヘッダー行(C列以降に日付が5つ以上並ぶ行)を探す。
Private Function FindDateHeaderRow(data As Variant, usedRows As Long, usedCols As Long) As Long
    Dim r As Long, c As Long
    Dim maxR As Long: maxR = Application.WorksheetFunction.Min(10, usedRows)
    Dim maxC As Long: maxC = Application.WorksheetFunction.Min(20, usedCols)
    For r = 1 To maxR
        Dim dateCount As Long: dateCount = 0
        For c = 3 To maxC
            If IsDate(data(r, c)) Then dateCount = dateCount + 1
        Next c
        If dateCount >= 5 Then
            FindDateHeaderRow = r
            Exit Function
        End If
    Next r
    FindDateHeaderRow = 0
End Function
