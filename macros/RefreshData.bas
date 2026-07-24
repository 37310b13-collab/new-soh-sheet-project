Attribute VB_Name = "RefreshData"
Option Explicit

Public Const MD_HEADER_ROW As Long = 6       ' Material_Detail: ヘッダー行。build_soh.pyのMD_TABLE_ROWと対応
Public Const MD_WEEK_START_COL As Long = 4   ' Material_Detail: 週データ開始列(D列)。build_soh.pyのWEEK_START_COLと対応

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
'   RefreshTTAFStock      : 「CSA Report」を選択すると、その中の「 COUNT SHEET SOH」シートから
'                          T_TTAFStock（TTAF在庫実績）を更新する。TTAF_Code優先、見つからなければ
'                          Description(材料名)で照合する。
'   HideInactiveIntermediates : Material_Detailで、指定した月数の間ずっと生産予定(バッチ数)が
'                          0の中間体の行をボタン一つで非表示にする（材料の在庫関連の行は
'                          常に表示されたまま）。
'   ShowAllIntermediates  : HideInactiveIntermediatesで非表示にした行を全て再表示する。
'   JumpToSelectedWeek    : Dashboard/Material_Detail/T_SelfStock/T_TTAFStockのC1(選択週)を
'                          入力したときに、実際の週データ列(複製ではない本物の列)が固定ペイン
'                          のすぐ右に来るようウィンドウを横スクロールする。呼び出しには
'                          対象の4シートそれぞれのシートモジュールにWorksheet_Changeを1行ずつ
'                          設置する必要があります（下記【導入方法】参照。標準モジュールの
'                          機能だけではシートの変更を検知できないための対応です）。
'
' いずれも、それぞれ対応するシートだけを更新します。T_Shipments・T_OpeningStock・
' T_StockCount・SafetyStock_Qty等、運用中に手入力した内容には一切触れません。
'
' 【パフォーマンスについて】どのRefresh*マクロも、外部ファイルのシートを1セルずつ.Cells(r,c).Value
' で読む代わりに、対象範囲を1回だけ配列として読み込み(Range.Value)、以降はメモリ上の配列だけを
' 参照する設計にしています。また、PP_Grid(中間体名->行番号)・M_BOM(Intermediate|RM_Code->行番号)・
' T_SelfStock_Log/T_TTAFStock_Log((RM_Code,週の月曜日)->行番号)への書き込みも、呼び出すたびに
' .Find()や全行スキャンをする代わりに、実行の最初にDictionaryを1回だけ作って参照する設計です
' （BuildNameIndex/BuildPairIndex/BuildStockRowIndex）。これは実際にExcelが強制終了する不具合
' (1セルずつの読み書きや毎回の全行スキャンがCOM通信の積み重ねで極めて遅くなることが原因)が
' 複数回報告されたことを受けての対策です。
'
' 【既存行の更新に.DataBodyRangeを使わない理由】RefreshBOM/RefreshWeeklyBatchesの実行時に
' 「(91) オブジェクト変数または With ブロック変数が設定されていません」というエラーが報告され
' ました。原因は、Excel/VBAの既知のクセとして、ListObject.DataBodyRangeが(特にListRows.Add
' で新しい行を追加した直後など)不安定にNothingを返すことがあるためです。新規行の追加時は元々
' 各Sub内で.ListRows.Add後の戻り値(newRow.Range)を使っておりこの問題を回避できていましたが、
' 既存行の値を更新する側だけ.DataBodyRange.Cells(...)という不安定な書き方が残っていました。
' すべて.ListRows(行番号).Range.Cells(...)という、行追加直後でも安定して動く書き方に統一
' しています。
'
' 【DisplayAlertsを抑制している理由】.DataBodyRangeの修正後も、取込元ファイルを開いた直後の
' srcWbが「Nothing」になり、後始末のsrcWb.Closeで同じ(91)エラーが再発するケースが報告されま
' した。取込元ファイル(Powder & Slurry & Pgm Plan、Usage from Production Engineering等)は
' 手動で開く際に「読み取り専用を推奨」の確認ダイアログが出るファイルであることが確認できて
' おり、Application.DisplayAlerts=Trueのままだと、VBAのWorkbooks.Open実行時にもこのダイアログ
' が表示されて処理が止まる(応答待ちのまま次の行に進めない、または想定外の状態でオブジェクトが
' 返る)ことが原因と考えられます。Workbooks.Open前にApplication.DisplayAlerts=Falseを設定して
' このようなダイアログを抑制し、後始末時にTrueへ戻すようにしました。念のため、srcWb.Close自体
' も引き続きIf Not srcWb Is Nothing Thenでガードしています。
'
' 【T_SelfStock/T_TTAFStockの二層構造について】RefreshSelfStock/RefreshTTAFStockは、目に
' 見えるT_SelfStock/T_TTAFStockシートには一切書き込みません。書き込み先は非表示の
' T_SelfStock_Log/T_TTAFStock_Log（実施日ベースの生ログ）で、目に見える方のシートは
' そこから毎回計算し直す数式(材料×週のグリッド)だけで組み立てられています。
' _Logシート側のWeekIndex列はDate列から自動計算される数式列です（RefreshSelfStock/
' RefreshTTAFStockはDateだけを書き込み、WeekIndexは書き込みません）。Cal_Weeks!B1
' (AnchorYear)を進めても、記録済みの実績データが「別の週のデータ」として誤表示される
' ことがないようにするためです。BuildStockRowIndex/UpsertStockRowIndexedの突合キーは、
' (RM_Code, その週の月曜日=MondayOfWeekで実日付から計算)です。月曜日をキーにしている
' のは、同じ週内に複数回取り込んでも1行に上書きされるようにするためです（以前はDateその
' ものをキーにしていたため、日次で取り込むたびに行が積み上がる不具合がありました）。
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
'
' 【選択週の自動スクロール(JumpToSelectedWeek)を有効にする場合の追加手順（任意）】
'   標準モジュール(このファイル)へのインポートだけでは動きません。以下のコードを
'   「Dashboard」シート・「Material_Detail」シートそれぞれの“シート自身のコードモジュール”に
'   直接貼り付けてください（VBEのプロジェクトエクスプローラーでシート名をダブルクリックすると
'   開きます。標準モジュールに貼り付けても発火しません）。
'
'   ' --- Dashboardシートのコードモジュールに貼り付け ---
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Intersect(Target, Me.Range("C1")) Is Nothing Then Exit Sub
'       Call JumpToSelectedWeek(Me, "F1", 9)   ' 9 = I列(週データ開始列)
'   End Sub
'
'   ' --- Material_Detailシートのコードモジュールに貼り付け ---
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Intersect(Target, Me.Range("C1")) Is Nothing Then Exit Sub
'       Call JumpToSelectedWeek(Me, "F1", 4)   ' 4 = D列(週データ開始列)
'   End Sub
'
'   ' --- T_SelfStockシートのコードモジュールに貼り付け ---
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Intersect(Target, Me.Range("C1")) Is Nothing Then Exit Sub
'       Call JumpToSelectedWeek(Me, "F1", 2)   ' 2 = B列(週データ開始列)
'   End Sub
'
'   ' --- T_TTAFStockシートのコードモジュールに貼り付け ---
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Intersect(Target, Me.Range("C1")) Is Nothing Then Exit Sub
'       Call JumpToSelectedWeek(Me, "F1", 2)   ' 2 = B列(週データ開始列)
'   End Sub
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

' tblの1・2列目から「1列目の値|2列目の値」->行番号のDictionaryを1回だけ作る。
' M_BOM(Intermediate|RM_Code)・T_SelfStock/T_TTAFStock(RM_Code|WeekIndex)など、
' 「先頭2列の組み合わせで行を特定する」テーブル全般に使う汎用ヘルパー。
Private Function BuildPairIndex(tbl As ListObject) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim n As Long: n = tbl.ListRows.Count
    If n > 0 Then
        Dim data As Variant
        data = tbl.ListColumns(1).DataBodyRange.Resize(n, 2).Value
        Dim i As Long
        For i = 1 To n
            idx(CStr(data(i, 1)) & "|" & CStr(data(i, 2))) = i
        Next i
    End If
    Set BuildPairIndex = idx
End Function

' tblの指定した1列(colName)の値->行番号のDictionaryを1回だけ作る。
' 旧実装は.Find()を使っており、Excelの既定の挙動として大文字/小文字を区別しない
' 検索だった。同じ挙動を保つため、CompareMode=vbTextCompareで大文字/小文字を
' 区別しないDictionaryにしている（区別してしまうと、表記ゆれ(TSP-049 と tsp-049等)を
' 同じ中間体として扱えず、実行のたびに重複行が増えていく別の不具合につながるため）。
Private Function BuildNameIndex(tbl As ListObject, colName As String) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbTextCompare
    Dim n As Long: n = tbl.ListRows.Count
    If n = 1 Then
        Dim colPos As Long: colPos = tbl.ListColumns(colName).Index
        idx(CStr(tbl.ListRows(1).Range.Cells(1, colPos).Value)) = 1
    ElseIf n > 1 Then
        Dim data As Variant
        data = tbl.ListColumns(colName).DataBodyRange.Value
        Dim i As Long
        For i = 1 To n
            Dim k As String: k = CStr(data(i, 1))
            If Not idx.Exists(k) Then idx(k) = i
        Next i
    End If
    Set BuildNameIndex = idx
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

Sub RefreshBOM()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel ファイル (*.xlsx),*.xlsx", , _
        "Usage from Production Engineering を選択してください")
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
    Dim bomTbl As ListObject: Set bomTbl = thisWb.Sheets("M_BOM").ListObjects("M_BOM")
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    ' 材料名(正規化) -> RM_Code。RM_Code・Descriptionは隣接列(1,2列目)なので1回の配列読み込みで済む
    ' （セルを1つずつ.Cells(i,1).Valueで読むより大幅に速い）。
    Dim descIndex As Object: Set descIndex = CreateObject("Scripting.Dictionary")
    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    If rmN > 0 Then
        Dim rmData As Variant
        rmData = rmTbl.ListColumns(1).DataBodyRange.Resize(rmN, 2).Value  ' Part Name, Description
        Dim i As Long
        For i = 1 To rmN
            Dim key As String: key = NormalizeText(CStr(rmData(i, 2)))
            If Len(key) > 0 And Not descIndex.Exists(key) Then descIndex(key) = CStr(rmData(i, 1))
        Next i
    End If

    ' (Intermediate|RM_Code) -> M_BOM内の行番号(データ範囲内)
    Dim pairIndex As Object: Set pairIndex = BuildPairIndex(bomTbl)

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

    ' srcWbが既にNothingになっているケース(取込元ファイル側の自動処理等で、開いた
    ' 直後にワークブックが閉じられてしまう場合がある)でも、後始末処理自体が
    ' 「オブジェクト変数が設定されていません」で落ちないようにガードする。
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim msg As String
    msg = "M_BOM を更新しました。" & vbCrLf & "更新: " & updated & " 件、新規追加: " & added & " 件"
    If Len(unresolved) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "原材料コードが見つからず未反映の材料名:" & vbCrLf & unresolved
    End If
    msg = msg & vbCrLf & vbCrLf & "新規追加した組み合わせも、Grid_RequirementがM_BOMを直接参照しているため" & vbCrLf & "その場で自動反映されます。"
    MsgBox msg, vbInformation
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

Private Sub ProcessBomSheet(sh As Worksheet, bomTbl As ListObject, descIndex As Object, _
        pairIndex As Object, ByRef updated As Long, ByRef added As Long, ByRef unresolved As String)
    Dim lastRow As Long, lastCol As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    lastCol = sh.Cells(2, sh.Columns.Count).End(xlToLeft).Column
    If lastRow > 3000 Then lastRow = 3000  ' 異常値対策(通常は数十〜数百行)
    If lastCol > 500 Then lastCol = 500
    If lastRow < 4 Or lastCol < 2 Then Exit Sub

    ' シート全体を1回だけ配列として読み込む（.Cells(r,c).Valueをループ内で毎回呼ぶと遅いため）
    Dim data As Variant
    data = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, lastCol)).Value

    Dim r As Long, c As Long
    For r = 4 To lastRow
        Dim interName As String
        interName = Trim(CStr(data(r, 1)))
        If Len(interName) > 0 Then
            For c = 2 To lastCol
                Dim matName As String: matName = Trim(CStr(data(2, c)))
                Dim v As Variant: v = data(r, c)
                If Len(matName) > 0 And IsNumeric(v) Then
                    If CDbl(v) <> 0 Then
                        Dim mkey As String: mkey = NormalizeText(matName)
                        If descIndex.Exists(mkey) Then
                            Dim rmCode As String: rmCode = descIndex(mkey)
                            Dim pk As String: pk = interName & "|" & rmCode
                            If pairIndex.Exists(pk) Then
                                Dim rowN As Long: rowN = pairIndex(pk)
                                bomTbl.ListRows(rowN).Range.Cells(1, 3).Value = CDbl(v)
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
'                      追加/更新する。対象週はファイル名から読み取った日付で判定する。
'   RefreshTTAFStock : 「CSA Report」を選択すると、その中の「 COUNT SHEET SOH」シート
'                      (先頭に半角スペースあり。A列=TTAF PART NUMBER、D列=Description、
'                      H列1行目=対象日、H列2行目以降=Total SOH)からT_TTAFStockにその週の
'                      実績を追加/更新する。対象週はH1セルの日付で判定する(ファイル名には
'                      依存しない)。材料の照合はTTAF_Codeを優先し、見つからなければ
'                      Descriptionの正規化テキストで照合する。
'
' どちらも、既存の(原材料, 週)の組み合わせがあれば値を上書き、無ければ新しい行として追加する
' (同じ週内に複数回取り込んでも1行にまとまる。取り込む順序は問わない)。
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
    ' 「読み取り専用を推奨」設定のファイルだと、DisplayAlerts=Trueのままでは
    ' Workbooks.Open時に確認ダイアログが表示され、応答待ちで処理が不安定になる
    ' (最終的にsrcWbが正しく取得できない不具合の原因になっていた)ため抑制する。
    Application.DisplayAlerts = False

    Set srcWb = Workbooks.Open(CStr(srcPath), ReadOnly:=True, UpdateLinks:=False)
    Dim reportDate As Date: reportDate = ExtractDateFromName(CStr(srcPath))

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim selfTbl As ListObject: Set selfTbl = thisWb.Sheets("T_SelfStock_Log").ListObjects("T_SelfStock_Log")
    Dim wIdx As Long: wIdx = WeekIndexForDate(thisWb, reportDate)
    Dim selfIdx As Object: Set selfIdx = BuildStockRowIndex(selfTbl)

    Dim sh As Worksheet: Set sh = srcWb.Sheets("Stock")
    ' シートを1セルずつ読むと遅くなるため、対象範囲(9〜200行, A〜J列)を1回だけ配列で読み込む
    Dim data As Variant
    data = sh.Range(sh.Cells(9, 1), sh.Cells(200, 10)).Value
    Dim r As Long, added As Long, updated As Long
    added = 0: updated = 0
    For r = 1 To (200 - 9 + 1)
        Dim code As String
        code = Trim(CStr(data(r, 3)))
        If Left(code, 4) = "CHEM" Then
            Dim v As Variant: v = data(r, 10)
            If IsNumeric(v) Then
                Call UpsertStockRowIndexed(selfTbl, selfIdx, code, reportDate, CDbl(v), added, updated)
            End If
        End If
    Next r

    ' srcWbが既にNothingになっているケース(取込元ファイル側の自動処理等で、開いた
    ' 直後にワークブックが閉じられてしまう場合がある)でも、後始末処理自体が
    ' 「オブジェクト変数が設定されていません」で落ちないようにガードする。
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "T_SelfStock を更新しました。" & vbCrLf & "対象週: " & wIdx & " (" & Format(reportDate, "yyyy-mm-dd") & ")" & vbCrLf & _
           "追加: " & added & " 件、更新: " & updated & " 件" & vbCrLf & _
           "（同じ週内の実績は1件にまとめられます。グリッド表示のT_SelfStockシートは自動で反映されます）", vbInformation
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

Sub RefreshTTAFStock()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel ファイル (*.xlsx),*.xlsx", , _
        "CSA Report（TTAF在庫）ファイルを選択してください")
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
    Dim ttafTbl As ListObject: Set ttafTbl = thisWb.Sheets("T_TTAFStock_Log").ListObjects("T_TTAFStock_Log")
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    ' 「 COUNT SHEET SOH」シート(先頭に半角スペースあり)を使う。A列2行目=TTAF PART NUMBER、
    ' D列2行目=Description、H列1行目=対象日(DD.MM.YYYY)、H列2行目=Total SOH、データは3行目から。
    ' 以前使っていた「PIVOT SOH TTAF」シートは、CHEM-で始まるコードしか拾えず(substrate等が
    ' 漏れる)、日付もファイル名から読み取っていたため、こちらのシートの方が網羅的かつ確実。
    Dim sh As Worksheet: Set sh = srcWb.Sheets(" COUNT SHEET SOH")
    ' H1の日付は「レポートが届いた月曜日(祝日の場合は翌営業日)」だが、その数値は前週の
    ' 在庫を表す。そのため7日引いてから週Noを判定する(祝日で月曜以外の日になっていても、
    ' ちょうど1週間前にずらすだけなので、前週の範囲内に正しく収まる)。
    Dim reportDate As Date: reportDate = ExtractDDMMYYYYFromText(CStr(sh.Cells(1, 8).Value)) - 7
    Dim wIdx As Long: wIdx = WeekIndexForDate(thisWb, reportDate)
    Dim ttafIdx As Object: Set ttafIdx = BuildStockRowIndex(ttafTbl)

    Dim ttafCodeIdx As Object: Set ttafCodeIdx = CreateObject("Scripting.Dictionary")
    Dim descIdx As Object: Set descIdx = CreateObject("Scripting.Dictionary")
    Call BuildTTAFCodeAndDescIndex(rmTbl, ttafCodeIdx, descIdx)

    ' シートを1セルずつ読むと遅くなるため、余裕を持った範囲を1回だけ配列で読み込んでから走査する
    ' (A列=TTAF PART NUMBER, D列=Description, H列=Total SOH)。
    Const MAX_ROWS As Long = 2000
    Dim data As Variant
    data = sh.Range(sh.Cells(3, 1), sh.Cells(MAX_ROWS, 8)).Value

    Dim r As Long, added As Long, updated As Long, unresolved As String
    added = 0: updated = 0: unresolved = ""
    For r = 1 To (MAX_ROWS - 3 + 1)
        Dim ttafCodeRaw As String: ttafCodeRaw = Trim(CStr(data(r, 1)))
        If Len(ttafCodeRaw) = 0 Then GoTo NextRow
        Dim v As Variant: v = data(r, 8)
        If Not IsNumeric(v) Then GoTo NextRow

        Dim descRaw As String: descRaw = Trim(CStr(data(r, 4)))
        Dim matchedPart As String
        matchedPart = ResolveTTAFPart(ttafCodeIdx, descIdx, ttafCodeRaw, descRaw)

        If Len(matchedPart) = 0 Then
            If InStr(unresolved, ttafCodeRaw) = 0 Then
                unresolved = unresolved & ttafCodeRaw & " (" & descRaw & "); "
            End If
            GoTo NextRow
        End If

        Call UpsertStockRowIndexed(ttafTbl, ttafIdx, matchedPart, reportDate, CDbl(v), added, updated)
NextRow:
    Next r

    ' srcWbが既にNothingになっているケース(取込元ファイル側の自動処理等で、開いた
    ' 直後にワークブックが閉じられてしまう場合がある)でも、後始末処理自体が
    ' 「オブジェクト変数が設定されていません」で落ちないようにガードする。
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim msg As String
    msg = "T_TTAFStock を更新しました。" & vbCrLf & "対象週: " & wIdx & " (" & Format(reportDate, "yyyy-mm-dd") & ")" & vbCrLf & _
          "追加: " & added & " 件、更新: " & updated & " 件" & vbCrLf & _
          "（同じ週内の実績は1件にまとめられます。グリッド表示のT_TTAFStockシートは自動で反映されます）"
    If Len(unresolved) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "TTAF_Code・材料名のどちらでも照合できず未反映の行:" & vbCrLf & unresolved
    End If
    MsgBox msg, vbInformation
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

' M_RawMaterialsから、TTAF_Code(正規化済み)->Part Name、Description(正規化済み)->Part Nameの
' インデックスを1回だけ作る。RefreshTTAFStockが使う共通処理。
' Dictionaryはオブジェクト(参照渡し)なので、ByRefを明示しなくても呼び出し元のtafCodeIdx/descIdxに
' そのまま反映される。
Private Sub BuildTTAFCodeAndDescIndex(rmTbl As ListObject, ttafCodeIdx As Object, descIdx As Object)
    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    If rmN > 0 Then
        Dim rmNameDesc As Variant
        rmNameDesc = rmTbl.ListColumns(1).DataBodyRange.Resize(rmN, 2).Value  ' Part Name, Description
        Dim rmTtafCode As Variant
        rmTtafCode = rmTbl.ListColumns(9).DataBodyRange.Value                 ' TTAF_Code
        Dim i As Long
        For i = 1 To rmN
            Dim tKeyBuild As String: tKeyBuild = NormalizeText(CStr(rmTtafCode(i, 1)))
            If Len(tKeyBuild) > 0 And Not ttafCodeIdx.Exists(tKeyBuild) Then ttafCodeIdx(tKeyBuild) = CStr(rmNameDesc(i, 1))
            Dim dKeyBuild As String: dKeyBuild = NormalizeText(CStr(rmNameDesc(i, 2)))
            If Len(dKeyBuild) > 0 And Not descIdx.Exists(dKeyBuild) Then descIdx(dKeyBuild) = CStr(rmNameDesc(i, 1))
        Next i
    End If
End Sub

' TTAF_Codeでの照合を優先し、見つからない場合だけDescription(材料名)の正規化テキストで照合する。
' どちらでも見つからなければ空文字を返す。
Private Function ResolveTTAFPart(ttafCodeIdx As Object, descIdx As Object, ttafCodeRaw As String, descRaw As String) As String
    Dim tKey As String: tKey = NormalizeText(ttafCodeRaw)
    If Len(tKey) > 0 And ttafCodeIdx.Exists(tKey) Then
        ResolveTTAFPart = ttafCodeIdx(tKey)
        Exit Function
    End If
    Dim dKey As String: dKey = NormalizeText(descRaw)
    If descIdx.Exists(dKey) Then
        ResolveTTAFPart = descIdx(dKey)
        Exit Function
    End If
    ResolveTTAFPart = ""
End Function

Private Function ExtractDDMMYYYYFromText(text As String) As Date
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d{2})\.(\d{2})\.(\d{4})"
    Dim m As Object
    Set m = re.Execute(text)
    If m.Count = 0 Then
        Err.Raise vbObjectError + 1, , "お届け予定日(DD.MM.YYYY)を読み取れませんでした: " & text
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
        Err.Raise vbObjectError + 1, , "ファイル名から日付(DD.MM.YYYY)を読み取れませんでした。"
    End If
    ExtractDateFromName = DateSerial(CInt(m(0).SubMatches(2)), CInt(m(0).SubMatches(1)), CInt(m(0).SubMatches(0)))
End Function

Private Function WeekIndexForDate(wb As Workbook, d As Date) As Long
    Dim calTbl As ListObject: Set calTbl = wb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks")
    Dim n As Long: n = calTbl.ListRows.Count
    If n > 0 Then
        ' 列: 1=WeekIndex, 2=WeekStart, 6=WeekEnd。1回の配列読み込みで全列まとめて取得する。
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
    WeekIndexForDate = 1 ' 見つからない場合はWeek1にフォールバック
End Function

' 日付から「その週の月曜日」を実際の暦計算で求める(Cal_Weeks!B1のAnchorYearには一切
' 依存しない、純粋な日付演算)。T_SelfStock_Log/T_TTAFStock_Logの突合キーに使うことで、
' 同じ週内に何度取り込んでも1行に上書きされるようにする(以前はDateそのものをキーに
' していたため、日次で取り込むたびに行が積み上がっていた)。
Private Function MondayOfWeek(d As Date) As Date
    MondayOfWeek = d - Weekday(d, vbMonday) + 1
End Function

' tbl(T_SelfStock_Log/T_TTAFStock_Log)の(RM_Code, その週の月曜日)->行番号のインデックスを
' 1回だけ作る。列は RM_Code(1), Date(2), WeekIndex(3, Dateから自動計算される数式), Qty(4)。
' キーを「その週の月曜日」(実際の暦日から計算。AnchorYearには依存しない)にしているのは、
' ①同じ週内の複数回の取り込みを1行にまとめるため、②AnchorYearを変更しても記録済みの
' 実績データが「別の週のデータ」として誤表示されないようにするため、の両方を同時に満たす。
' UpsertStockRowが呼ばれるたびに全行をセル単位でスキャンしていたのを避けるため、
' 事前に1回のRange読み込みでDictionaryを構築しておく（テーブルが月々増えるほど効果が大きい）。
' 日付を文字列化する際はCLng(シリアル値)を経由し、地域の日付表示形式に左右されないようにする。
Private Function BuildStockRowIndex(tbl As ListObject) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim n As Long: n = tbl.ListRows.Count
    If n > 0 Then
        Dim data As Variant
        data = tbl.ListColumns(1).DataBodyRange.Resize(n, 2).Value  ' 1,2列目(RM_Code,Date)をまとめて読む
        Dim i As Long
        For i = 1 To n
            idx(CStr(data(i, 1)) & "|" & CStr(CLng(MondayOfWeek(CDate(data(i, 2)))))) = i
        Next i
    End If
    Set BuildStockRowIndex = idx
End Function

' WeekIndex(3列目)は数式列のためここでは値を書き込まない(Dateが変われば自動的に再計算される)。
' 【重要】以前は「新規行を追加すればExcelのテーブル機能が既存行と同じ数式を自動的に複製する」
' という前提だったが、これはUI上でテーブルの下に手で行を追加した場合の挙動であり、
' VBAのListRows.Addで追加した行には自動複製されないことがある(実際に報告された不具合:
' T_SelfStock_Log/T_TTAFStock_LogにVBAで追加した行のWeekIndex列が数式ごと空欄のままになり、
' グリッド側のSUMIFS/COUNTIFSが該当行を見つけられず、T_SelfStock/T_TTAFStockに何も
' 表示されなくなっていた)。そのため、新規行では直前行のWeekIndexの数式を明示的にコピーする
' (FormulaR1C1でコピーすることで、相対参照(自分自身のDateセルを指す部分)はコピー先の行に
' 合わせて自動調整される)。
' 同じ週内で2回目以降の取り込みがあった場合は、Date・Qtyの両方を最新の値で上書きする
' (その週内で一番新しい実施日の記録が残るようにするため)。
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

' ============================================================================
' HideInactiveIntermediates / ShowAllIntermediates
'
' Material_Detailシートは、材料(RM_Code)ごとに「中間体名／No. of batches」
' 「使用量(kg)」の行が縦に並ぶブロック構成になっている。将来しばらく(半年・1年など)
' 生産予定の無い中間体の行を、ユーザー指定期間に応じてボタン一つで折りたたみ、
' 見た目のスクロール量を減らせるようにする。材料の合計使用量・TTAF在庫実績・
' 自社在庫実績・合計在庫(週末時点)の行は、非表示の対象にせず常に見える状態を保つ
' （在庫量そのものは常に把握できるようにするため）。
'
' 判定方法: 「今週」から指定月数分の週について、その中間体の「No. of batches」行が
' 全て0(または空白)であれば、その中間体の2行(No. of batches行＋使用量(kg)行)を隠す。
' 1つでも0以外の週があれば表示したままにする。実行のたびにまず全行を再表示してから
' 判定し直すため、期間を変えて何度実行しても結果は指定した条件どおりになる。
'
' 行位置はラベル文字列(列C="No. of batches")で判定しており、build_soh.py側で
' 行数が増減しても自動的に追従する。ただし週データの開始列(MD_WEEK_START_COL=E列)
' とヘッダー行(MD_HEADER_ROW=6行目)は、build_soh.pyのWEEK_START_COL/MD_TABLE_ROWと
' 値を合わせる必要がある(シート構成を変更した場合はここも合わせて変更すること)。
'
' 【ボタンの割り当て方(手動での一度だけの作業。openpyxlではボタンを自動作成できないため)】
'   1. Material_Detailシートを開く
'   2. 「挿入」タブ →「図形」等で好きな形の図形を1~2個描く(例:「非表示にする」「全部表示」)
'   3. 図形を右クリック →「マクロの登録」→ HideInactiveIntermediates (もう1つには
'      ShowAllIntermediates)を選択
'   4. お好みでシート上部の空いている場所(A1付近など)に配置する
' ============================================================================

Sub HideInactiveIntermediates()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets("Material_Detail")
    On Error GoTo 0
    If sh Is Nothing Then
        MsgBox "Material_Detailシートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim monthsStr As String
    monthsStr = InputBox("今週から何ヶ月間、全週バッチ数0の中間体を非表示にしますか？" & _
        vbCrLf & "（例: 6 → 半年間生産予定の無い中間体を非表示、12 → 1年間）" & _
        vbCrLf & "半角数字で入力してください。", "非表示にする期間(月数)", "6")
    If monthsStr = "" Then Exit Sub  ' キャンセル
    If Not IsNumeric(monthsStr) Then
        MsgBox "数値を入力してください。", vbExclamation
        Exit Sub
    End If
    Dim months As Double: months = CDbl(monthsStr)
    If months <= 0 Then
        MsgBox "1以上の数値を入力してください。", vbExclamation
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
        MsgBox "対象週が見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row

    ' まず全行を再表示してから判定する(期間を変えて再実行しても常に指定条件どおりになるように)
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
                sh.Rows(r + 1).Hidden = True   ' 対応する「使用量(kg)」行も一緒に隠す
                hiddenCount = hiddenCount + 1
            Else
                shownCount = shownCount + 1
            End If
            r = r + 2   ' No. of batches行＋使用量(kg)行のペア分をスキップ
        Else
            r = r + 1
        End If
    Loop

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "完了しました。" & vbCrLf & _
           "今週(week " & curWeek & ")から" & thresholdWeeks & "週間、全週バッチ数0の中間体を非表示にしました。" & vbCrLf & _
           "非表示: " & hiddenCount & " 件 / 表示中: " & shownCount & " 件" & vbCrLf & _
           "（材料名の行・在庫関連の行は常に表示されます）", vbInformation
End Sub

' Dashboard/Material_Detail/T_SelfStock/T_TTAFStockのC1(選択週)を入力したときに呼ばれる
' 想定の共通処理。選択週の値を別セルに複製する(ピン留め列)方式は廃止し、代わりに
' 「本物の週データ列」が常にラベル列のすぐ右(固定ペインの直後)に見えるよう、ウィンドウを
' 横スクロールするだけにしている。値の複製が一切ないため、各シート・Grid_Stock等の間で
' 数字が食い違う余地がない。
' 呼び出し元のシート自身のWorksheet_Changeから、対象シート・週No解決済みセル(F1)・
' 週データ開始列(Dashboardは9=I列、Material_Detailは4=D列、T_SelfStock/T_TTAFStockは
' 2=B列)を渡して呼び出す。
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
        MsgBox "Material_Detailシートが見つかりません。", vbExclamation
        Exit Sub
    End If
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row
    If lastRow >= MD_HEADER_ROW Then sh.Rows(MD_HEADER_ROW & ":" & lastRow).Hidden = False
    MsgBox "すべての中間体行を再表示しました。", vbInformation
End Sub
