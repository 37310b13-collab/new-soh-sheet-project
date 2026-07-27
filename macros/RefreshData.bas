Attribute VB_Name = "RefreshData"
Option Explicit

Public Const MD_HEADER_ROW As Long = 6       ' Material_Detail: ヘッダー行。build_soh.pyのMD_TABLE_ROWと対応
Public Const MD_WEEK_START_COL As Long = 4   ' Material_Detail: 週データ開始列(D列)。build_soh.pyのWEEK_START_COLと対応
Public Const SS_TABLE_ROW As Long = 5        ' T_SelfStock/T_TTAFStock: 見出し行(週ラベル)。build_soh.pyのSS_TABLE_ROWと対応

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
'   RefreshShipments : 「CSA Report」を選択すると、その中の「Shipping Schedule」シートを
'                          丸ごと取り込み、T_Shipments（発注〜着荷の実績・予定）を更新する。
'                          PO No単位ではなく全件まとめて取り込む(発注から着荷まで4〜6ヶ月かかり、
'                          常に複数件のPOが並行して進むため)。
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
'   AddMaterial           : 新しい材料(TTAF供給品)をシステムに追加する。InputBoxで
'                          Part Name(RM_Code)・Description・Supplier・Category・TTAF_Code
'                          を順に入力すると、M_RawMaterials・Grid_Requirement・Grid_Incoming・
'                          Grid_Stock・T_OpeningStock・T_SelfStock・T_TTAFStock・Dashboard・
'                          Material_Detail・対応するPO_Draft_*シートの一番下に、必要な行を
'                          まとめて追加する。追加直後はまだM_BOMに使用実績が無いため、
'                          Material_Detailのブロックは中間体の内訳が無いミニブロック(合計欄のみ)
'                          になる。RefreshBOM実行後、実際にこの材料を使う中間体が見つかれば
'                          Grid_Requirement経由で合計使用量に自動反映される。
'   RemoveMaterial        : 使わなくなった材料をシステムから削除する。InputBoxでPart Name
'                          (RM_Code)を入力すると、AddMaterialが追加する全シートから該当行を
'                          削除する。T_Shipments・T_PlannedOrders・T_StockCount・
'                          T_SelfStock_Log/T_TTAFStock_Log・M_BOMのデータは削除しない
'                          (履歴として残すため。再度AddMaterialで同じPart Nameを追加すれば
'                          自動的に再びつながる)。取り消せない操作のため、実行前にファイルの
'                          バックアップを取ることを強く推奨します。
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

' CSA Reportの「Shipping Schedule」シートを丸ごと取り込み、T_Shipmentsを更新する。
' 発注から着荷まで4〜6ヶ月かかり、常に複数件のPOが並行して進むため、PO Noを1件ずつ指定する
' 方式ではなく、RefreshTTAFStock等と同じ「ファイルを選ぶだけ」で全件まとめて更新する方式にした。
' 列: D=CSA Product Code(材料コード)、G=CSA Order firm month(発注月)、I=CSA PO No.、
' N=Confirmed Order Qty、P=Latest ETA、S=Received At TTAF、T=Status。
' CSA Product CodeはTTAF PART NUMBERと異なり、既にこちらのRM_Codeとほぼ同じ表記のため、
' TTAF_Code/Description経由の照合(ResolveTTAFPart)ではなく、RM_Code同士を直接
' (大文字小文字・前後空白を無視して)照合する。
Sub RefreshShipments()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel ファイル (*.xlsx),*.xlsx", , _
        "CSA Report（Shipping Schedule取込み）ファイルを選択してください")
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
    Dim shipTbl As ListObject: Set shipTbl = thisWb.Sheets("T_Shipments").ListObjects("T_Shipments")
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    Dim sh As Worksheet: Set sh = srcWb.Sheets("Shipping Schedule")

    Dim rmCodeIdx As Object: Set rmCodeIdx = CreateObject("Scripting.Dictionary")
    Call BuildRMCodeIndex(rmTbl, rmCodeIdx)

    Dim shipIdx As Object: Set shipIdx = BuildShipmentRowIndex(shipTbl)

    ' シートを1セルずつ読むと遅くなるため、余裕を持った範囲を1回だけ配列で読み込んでから走査する。
    Const MAX_ROWS As Long = 3000
    Dim data As Variant
    data = sh.Range(sh.Cells(2, 1), sh.Cells(MAX_ROWS, 20)).Value

    Dim r As Long, added As Long, updated As Long, unresolved As String
    added = 0: updated = 0: unresolved = ""
    For r = 1 To (MAX_ROWS - 2 + 1)
        Dim rmCodeRaw As String: rmCodeRaw = Trim(CStr(data(r, 4)))
        Dim poNo As String: poNo = Trim(CStr(data(r, 9)))
        If Len(rmCodeRaw) = 0 Or Len(poNo) = 0 Then GoTo NextRow

        Dim kRow As String: kRow = NormalizeText(rmCodeRaw)
        Dim matchedPart As String: matchedPart = ""
        If rmCodeIdx.Exists(kRow) Then matchedPart = rmCodeIdx(kRow)

        If Len(matchedPart) = 0 Then
            If InStr(unresolved, rmCodeRaw) = 0 Then unresolved = unresolved & rmCodeRaw & "; "
            GoTo NextRow
        End If

        Dim qty As Variant: qty = data(r, 14)
        If Not IsNumeric(qty) Then qty = 0

        Dim eta As Variant: eta = data(r, 16)
        If Not IsDate(eta) Then eta = Empty

        Dim receivedDate As Variant: receivedDate = data(r, 19)
        If Not IsDate(receivedDate) Then receivedDate = Empty

        Dim statusVal As String: statusVal = Trim(CStr(data(r, 20)))

        Dim orderMonth As Variant: orderMonth = data(r, 7)
        If Not IsDate(orderMonth) Then orderMonth = Empty

        Call UpsertShipmentRow(shipTbl, shipIdx, matchedPart, poNo, CDbl(qty), eta, receivedDate, statusVal, orderMonth, added, updated)
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
    msg = "T_Shipments を更新しました。" & vbCrLf & "追加: " & added & " 件、更新: " & updated & " 件" & vbCrLf & _
          "（PO No＋材料の組み合わせが同じ行は上書きされます。Order_Date欄は手入力のため上書きしません）"
    If Len(unresolved) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "材料コードが見つからず未反映の行:" & vbCrLf & unresolved
    End If
    MsgBox msg, vbInformation
    Exit Sub

ErrHandler:
    ' 【重要】On Error Resume Next はErr オブジェクトを自動的にクリアしてしまう(VBAの仕様)ため、
    ' 後始末処理より前に、エラー番号・内容を必ず変数へ退避しておく。これを怠ると、
    ' 下のMsgBoxが常に「(空欄)」を表示してしまい、本当のエラー原因が一切分からなくなる
    ' (実際にこの不具合が発生し、原因調査ができない状態になっていたため修正)。
    Dim errNum2 As Long: errNum2 = Err.Number
    Dim errMsg2 As String: errMsg2 = Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "更新処理でエラーが発生しました: (" & errNum2 & ") " & errMsg2, vbCritical
End Sub

' M_RawMaterialsのPart Name(=RM_Code)自体を正規化テキストでインデックス化する。
' TTAF_Code/Descriptionでの照合(BuildTTAFCodeAndDescIndex)とは別物: CSA Product Codeは
' 既にRM_Codeとほぼ同じ表記のため、RM_Code同士を直接照合する方が確実。
Private Sub BuildRMCodeIndex(rmTbl As ListObject, rmCodeIdx As Object)
    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    If rmN > 0 Then
        Dim rmNames As Variant: rmNames = rmTbl.ListColumns(1).DataBodyRange.Value
        Dim i As Long
        For i = 1 To rmN
            Dim k As String: k = NormalizeText(CStr(rmNames(i, 1)))
            If Len(k) > 0 And Not rmCodeIdx.Exists(k) Then rmCodeIdx(k) = CStr(rmNames(i, 1))
        Next i
    End If
End Sub

' T_Shipmentsの(Part Name, PO_No)->行番号のインデックスを1回だけ作る。
Private Function BuildShipmentRowIndex(tbl As ListObject) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim n As Long: n = tbl.ListRows.Count
    If n > 0 Then
        Dim data As Variant
        data = tbl.ListColumns(1).DataBodyRange.Resize(n, 2).Value  ' Part Name, PO_No
        Dim i As Long
        For i = 1 To n
            idx(CStr(data(i, 1)) & "|" & CStr(data(i, 2))) = i
        Next i
    End If
    Set BuildShipmentRowIndex = idx
End Function

' 列: Part Name(1), PO_No(2), Order_Date_発注日(3, 手入力のため触れない), Confirmed_Qty(4),
' Latest_ETA(5), Received_Date(6), Status(7), Effective_Week(8, 数式), Order_Month(9)。
' Effective_Week(8列目)は数式列のため、新規行ではT_SelfStock_Log等と同様に直前行の数式を
' FormulaR1C1でコピーする(VBAのListRows.Add経由では計算列の自動複製が効かないことがあるため)。
Private Sub UpsertShipmentRow(tbl As ListObject, idx As Object, partName As String, poNo As String, qty As Double, _
        eta As Variant, receivedDate As Variant, status As String, orderMonth As Variant, ByRef added As Long, ByRef updated As Long)
    Dim key As String: key = partName & "|" & poNo
    If idx.Exists(key) Then
        Dim rowN As Long: rowN = idx(key)
        tbl.ListRows(rowN).Range.Cells(1, 4).Value = qty
        If Not IsEmpty(eta) Then tbl.ListRows(rowN).Range.Cells(1, 5).Value = eta
        If Not IsEmpty(receivedDate) Then tbl.ListRows(rowN).Range.Cells(1, 6).Value = receivedDate
        tbl.ListRows(rowN).Range.Cells(1, 7).Value = status
        If Not IsEmpty(orderMonth) Then tbl.ListRows(rowN).Range.Cells(1, 9).Value = orderMonth
        updated = updated + 1
    Else
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add
        newRow.Range.Cells(1, 1).Value = partName
        newRow.Range.Cells(1, 2).Value = poNo
        newRow.Range.Cells(1, 4).Value = qty
        If Not IsEmpty(eta) Then newRow.Range.Cells(1, 5).Value = eta
        If Not IsEmpty(receivedDate) Then newRow.Range.Cells(1, 6).Value = receivedDate
        newRow.Range.Cells(1, 7).Value = status
        If Not IsEmpty(orderMonth) Then newRow.Range.Cells(1, 9).Value = orderMonth
        If newRow.Index > 1 Then
            newRow.Range.Cells(1, 8).FormulaR1C1 = tbl.ListRows(newRow.Index - 1).Range.Cells(1, 8).FormulaR1C1
        End If
        idx(key) = newRow.Index
        added = added + 1
    End If
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

' ============================================================================
' AddMaterial / RemoveMaterial
'
' 材料(原材料)をPythonを使わず、Excel(VBA)だけで追加・削除するためのマクロです。
' 関係する全シート(M_RawMaterials, Grid_Requirement, Grid_Incoming, Grid_Stock,
' T_OpeningStock, T_SelfStock, T_TTAFStock, Dashboard, Material_Detail,
' 該当カテゴリのPO_Draft_*)に、対応する行を追加/削除します。
'
' 【追加(AddMaterial)の位置】既存行の途中に割り込ませるのではなく、必ず各シートの
' 一番下に追加します。途中に割り込ませると既存行がずれるリスクが大きいためです。
'
' 【重要な注意点】
' ・追加した材料はまだM_BOM(原単位)に登録されていないため、Material_Detailでは
'   「使用中間体なし」の状態で追加されます。RefreshBOMを実行すると、その材料が
'   実際に使われている中間体との組み合わせが見つかり次第、自動的にM_BOM経由で
'   反映されます(Grid_RequirementがM_BOMを直接参照するため)。
' ・削除(RemoveMaterial)は、T_Shipments・T_PlannedOrders・T_StockCount・
'   実績ログ(T_SelfStock_Log/T_TTAFStock_Log)・M_BOMに残っているその材料の
'   過去データは削除しません(誤って必要なデータを失わないための安全策です)。
'   不要であれば手動で削除してください。
' ・どちらの操作も元に戻せません(Ctrl+Zでは戻せない場合があります)。実行前に
'   ファイルのバックアップ(コピー)を取っておくことを強くおすすめします。
' ============================================================================

Sub AddMaterial()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    Dim rmCode As String
    rmCode = Trim(InputBox("追加する材料コード(Part Name)を入力してください。" & vbCrLf & "（例: CHEM-9999）", "材料の追加"))
    If Len(rmCode) = 0 Then Exit Sub

    If Not IsMaterialCodeFree(rmTbl, rmCode) Then
        MsgBox "その材料コードは既に登録されています: " & rmCode, vbExclamation
        Exit Sub
    End If

    Dim descVal As String
    descVal = Trim(InputBox("材料名(Description)を入力してください。", "材料の追加"))
    If Len(descVal) = 0 Then Exit Sub

    Dim supplierVal As String
    supplierVal = Trim(InputBox("仕入先(Supplier)を入力してください。", "材料の追加", "TTAF"))

    Dim categoryVal As String
    categoryVal = Trim(InputBox("カテゴリを入力してください。" & vbCrLf & _
        "Chemical / Hazardous Chemical / Substrate のいずれかを、そのまま入力してください。", _
        "材料の追加", "Chemical"))
    If categoryVal <> "Chemical" And categoryVal <> "Hazardous Chemical" And categoryVal <> "Substrate" Then
        MsgBox "カテゴリは Chemical / Hazardous Chemical / Substrate のいずれかで入力してください。" & vbCrLf & _
               "入力値: " & categoryVal, vbExclamation
        Exit Sub
    End If

    Dim ttafCodeVal As String
    ttafCodeVal = Trim(InputBox("TTAF_Code(TTAF側の部品番号)を入力してください。分からなければ空欄のままでOKです。", "材料の追加"))

    If MsgBox("以下の内容で材料を追加します。" & vbCrLf & vbCrLf & _
              "材料コード: " & rmCode & vbCrLf & "材料名: " & descVal & vbCrLf & _
              "仕入先: " & supplierVal & vbCrLf & "カテゴリ: " & categoryVal & vbCrLf & _
              "TTAF_Code: " & ttafCodeVal & vbCrLf & vbCrLf & _
              "よろしいですか？", vbYesNo + vbQuestion, "材料の追加の確認") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim nWeeks As Long
    nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count

    ' ---- M_RawMaterials ----
    Dim newRmRow As ListRow: Set newRmRow = rmTbl.ListRows.Add
    newRmRow.Range.Cells(1, 1).Value = rmCode
    newRmRow.Range.Cells(1, 2).Value = descVal
    newRmRow.Range.Cells(1, 3).Value = supplierVal
    newRmRow.Range.Cells(1, 4).Value = categoryVal
    newRmRow.Range.Cells(1, 5).Value = "kg"
    newRmRow.Range.Cells(1, 6).Value = 0
    newRmRow.Range.Cells(1, 7).Value = 0
    newRmRow.Range.Cells(1, 8).Value = 4
    newRmRow.Range.Cells(1, 9).Value = ttafCodeVal

    ' ---- Grid_Requirement / Grid_Incoming / Grid_Stock / T_OpeningStock ----
    Dim reqTbl As ListObject: Set reqTbl = thisWb.Sheets("Grid_Requirement").ListObjects("Grid_Requirement")
    Dim inTbl As ListObject: Set inTbl = thisWb.Sheets("Grid_Incoming").ListObjects("Grid_Incoming")
    Dim stTbl As ListObject: Set stTbl = thisWb.Sheets("Grid_Stock").ListObjects("Grid_Stock")
    Dim osTbl As ListObject: Set osTbl = thisWb.Sheets("T_OpeningStock").ListObjects("T_OpeningStock")

    Dim reqRow As ListRow: Set reqRow = reqTbl.ListRows.Add
    Dim inRow As ListRow: Set inRow = inTbl.ListRows.Add
    Dim stRow As ListRow: Set stRow = stTbl.ListRows.Add
    Dim osRow As ListRow: Set osRow = osTbl.ListRows.Add

    Dim grow As Long: grow = reqRow.Range.Row  ' Grid_Requirement/Incoming/Stockの実シート行番号(3表とも同じ)

    reqRow.Range.Cells(1, 1).Value = rmCode
    inRow.Range.Cells(1, 1).Value = rmCode
    stRow.Range.Cells(1, 1).Value = rmCode
    osRow.Range.Cells(1, 1).Value = rmCode
    osRow.Range.Cells(1, 2).Value = 0
    osRow.Range.Cells(1, 3).Value = Date

    Dim w As Long, col As Long, cl As String
    For w = 1 To nWeeks
        col = 1 + w
        reqRow.Range.Cells(1, col).Value = _
            "=SUMPRODUCT((M_BOM[Part Name]=$A" & grow & ")*M_BOM[RM_Qty_Per_Batch]*" & _
            "IFERROR(INDEX(PP_Grid[#Data],M_BOM[PPGridRow]," & (w + 1) & "),0))"
        inRow.Range.Cells(1, col).Value = _
            "=SUMIFS(T_Shipments[Confirmed_Qty],T_Shipments[Part Name],$A" & grow & _
            ",T_Shipments[Effective_Week]," & w & ")"
    Next w

    ' ---- T_SelfStock / T_TTAFStock (テーブルではない罫線グリッド。一番下に追加) ----
    Dim ssRowSelf As Long, ssRowTTAF As Long
    ssRowSelf = AppendStockGridRow(thisWb.Sheets("T_SelfStock"), rmCode, nWeeks, "T_SelfStock_Log", "Self_Qty")
    ssRowTTAF = AppendStockGridRow(thisWb.Sheets("T_TTAFStock"), rmCode, nWeeks, "T_TTAFStock_Log", "TTAF_Qty")
    If ssRowSelf <> ssRowTTAF Then
        MsgBox "警告: T_SelfStockとT_TTAFStockの行番号がずれました(" & ssRowSelf & " / " & ssRowTTAF & ")。" & vbCrLf & _
               "手動で確認してください。処理は続行します。", vbExclamation
    End If
    Dim ssRow As Long: ssRow = ssRowSelf

    ' Grid_Stockの数式(手動棚卸 > 自社+TTAF実績の合計 > 通常のロールフォワード、の優先順位)
    Dim priorExpr As String, normalExpr As String
    Dim hasCount As String, countVal As String, hasSelf As String, selfVal As String, hasTTAF As String, ttafVal As String
    For w = 1 To nWeeks
        col = 1 + w
        cl = ColLetter(col)
        hasCount = "COUNTIFS(T_StockCount[Part Name],$A" & grow & ",T_StockCount[WeekIndex]," & w & ")"
        countVal = "SUMIFS(T_StockCount[CountedQty],T_StockCount[Part Name],$A" & grow & ",T_StockCount[WeekIndex]," & w & ")"
        hasSelf = "('T_SelfStock'!" & cl & ssRow & "<>"""")"
        selfVal = "'T_SelfStock'!" & cl & ssRow
        hasTTAF = "('T_TTAFStock'!" & cl & ssRow & "<>"""")"
        ttafVal = "'T_TTAFStock'!" & cl & ssRow
        If w = 1 Then
            priorExpr = "IFERROR(INDEX(T_OpeningStock[Opening_Qty],MATCH($A" & grow & ",T_OpeningStock[Part Name],0)),0)"
        Else
            priorExpr = ColLetter(col - 1) & grow
        End If
        normalExpr = priorExpr & "+'Grid_Incoming'!" & cl & grow & "-'Grid_Requirement'!" & cl & grow
        stRow.Range.Cells(1, col).Value = _
            "=IF(" & hasCount & ">0," & countVal & ",IF((" & hasSelf & ")*(" & hasTTAF & ")>0," & selfVal & "+" & ttafVal & "," & normalExpr & "))"
    Next w

    ' ---- Dashboard (テーブルではない罫線グリッド。一番下に追加) ----
    Call AppendDashboardRow(thisWb.Sheets("Dashboard"), rmCode, nWeeks, ssRow, grow)

    ' ---- Material_Detail (材料のブロックを一番下に追加。BOM未登録のため中間体行はまだ無い) ----
    Call AppendMaterialDetailBlock(thisWb.Sheets("Material_Detail"), rmCode, descVal, nWeeks, ssRow, grow)

    ' ---- PO_Draft_{Category} ----
    Dim poSheetName As String: poSheetName = POSheetNameForCategory(categoryVal)
    If Len(poSheetName) > 0 Then
        Call AppendPODraftRow(thisWb.Sheets(poSheetName), rmCode, ttafCodeVal)
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "材料「" & rmCode & "」を追加しました。" & vbCrLf & vbCrLf & _
           "この材料はまだM_BOM(原単位)に登録されていないため、Material_Detailでは" & vbCrLf & _
           "「使用中間体なし」の状態です。RefreshBOMを実行すると、実際に使われている" & vbCrLf & _
           "中間体との組み合わせが見つかり次第、自動的に反映されます。", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "材料の追加中にエラーが発生しました: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "途中まで反映されている可能性があります。シートの状態を確認してください" & vbCrLf & _
           "(心配な場合は、保存せずにファイルを閉じて開き直せば、直前の保存状態に戻せます)。", vbCritical
End Sub

Sub RemoveMaterial()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    Dim rmCode As String
    rmCode = Trim(InputBox("削除する材料コード(Part Name)を入力してください。", "材料の削除"))
    If Len(rmCode) = 0 Then Exit Sub

    Dim rmFoundRow As Long: rmFoundRow = FindMaterialRow(rmTbl, rmCode)
    If rmFoundRow = 0 Then
        MsgBox "M_RawMaterialsに見つかりませんでした: " & rmCode, vbExclamation
        Exit Sub
    End If
    ' 実際に登録されている表記(大文字小文字)に揃える
    rmCode = CStr(rmTbl.ListRows(rmFoundRow).Range.Cells(1, 1).Value)
    Dim categoryVal As String: categoryVal = CStr(rmTbl.ListRows(rmFoundRow).Range.Cells(1, 4).Value)

    If MsgBox("材料「" & rmCode & "」を削除します。" & vbCrLf & _
              "関係する全シート(M_RawMaterials・Grid_Requirement・Grid_Incoming・Grid_Stock・" & vbCrLf & _
              "T_OpeningStock・T_SelfStock・T_TTAFStock・Dashboard・Material_Detail・" & vbCrLf & _
              "該当するPO_Draft)から該当行を削除します。この操作は元に戻せません。" & vbCrLf & vbCrLf & _
              "（T_Shipments・T_PlannedOrders・T_StockCount・実績ログ・M_BOMに残っている" & vbCrLf & _
              "この材料の過去データは削除されません）" & vbCrLf & vbCrLf & "よろしいですか？", _
              vbYesNo + vbExclamation, "材料の削除の確認") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call DeleteMatchingTableRow(thisWb.Sheets("Grid_Requirement").ListObjects("Grid_Requirement"), rmCode)
    Call DeleteMatchingTableRow(thisWb.Sheets("Grid_Incoming").ListObjects("Grid_Incoming"), rmCode)
    Call DeleteMatchingTableRow(thisWb.Sheets("Grid_Stock").ListObjects("Grid_Stock"), rmCode)
    Call DeleteMatchingTableRow(thisWb.Sheets("T_OpeningStock").ListObjects("T_OpeningStock"), rmCode)

    Call DeleteMatchingGridRow(thisWb.Sheets("T_SelfStock"), rmCode, 1)
    Call DeleteMatchingGridRow(thisWb.Sheets("T_TTAFStock"), rmCode, 1)
    Call DeleteMatchingGridRow(thisWb.Sheets("Dashboard"), rmCode, 1)

    Dim poSheetName As String: poSheetName = POSheetNameForCategory(categoryVal)
    If Len(poSheetName) > 0 Then
        Call DeleteMatchingGridRow(thisWb.Sheets(poSheetName), rmCode, 4)
    End If

    Call DeleteMaterialDetailBlock(thisWb.Sheets("Material_Detail"), rmCode)

    ' M_RawMaterials自体は、他シートの検索キーとして使い終わってから最後に削除する
    rmTbl.ListRows(rmFoundRow).Delete

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "材料「" & rmCode & "」を削除しました。" & vbCrLf & vbCrLf & _
           "（T_Shipments・T_PlannedOrders・T_StockCount・実績ログ・M_BOMに残っている" & vbCrLf & _
           "この材料の過去データは削除されていません。必要であれば手動で削除してください）", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "材料の削除中にエラーが発生しました: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "途中まで削除されている可能性があります。シートの状態を確認してください" & vbCrLf & _
           "(心配な場合は、保存せずにファイルを閉じて開き直せば、直前の保存状態に戻せます)。", vbCritical
End Sub

' 列番号(例:28)を列名(例:AB)に変換する。ワークシートに依存しない純粋な計算。
Private Function ColLetter(colNum As Long) As String
    Dim s As String, n As Long, r As Long
    n = colNum
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr(65 + r) & s
        n = (n - r - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Function IsMaterialCodeFree(rmTbl As ListObject, rmCode As String) As Boolean
    IsMaterialCodeFree = True
    Dim n As Long: n = rmTbl.ListRows.Count
    Dim i As Long
    For i = 1 To n
        If UCase(Trim(CStr(rmTbl.ListRows(i).Range.Cells(1, 1).Value))) = UCase(rmCode) Then
            IsMaterialCodeFree = False
            Exit Function
        End If
    Next i
End Function

Private Function FindMaterialRow(rmTbl As ListObject, rmCode As String) As Long
    FindMaterialRow = 0
    Dim n As Long: n = rmTbl.ListRows.Count
    Dim i As Long
    For i = 1 To n
        If UCase(Trim(CStr(rmTbl.ListRows(i).Range.Cells(1, 1).Value))) = UCase(rmCode) Then
            FindMaterialRow = i
            Exit Function
        End If
    Next i
End Function

Private Function POSheetNameForCategory(categoryVal As String) As String
    Select Case categoryVal
        Case "Chemical": POSheetNameForCategory = "PO_Draft_Chemical"
        Case "Hazardous Chemical": POSheetNameForCategory = "PO_Draft_Hazardous"
        Case "Substrate": POSheetNameForCategory = "PO_Draft_Substrate"
        Case Else: POSheetNameForCategory = ""
    End Select
End Function

' T_SelfStock/T_TTAFStockの一番下に新しい材料の行を追加し、実際に追加した行番号を返す。
' テーブル機能を使わない罫線グリッドのため、直前行の罫線・縞模様をコピーして体裁を揃える。
Private Function AppendStockGridRow(sh As Worksheet, rmCode As String, nWeeks As Long, logTableName As String, qtyColName As String) As Long
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim newRow As Long: newRow = lastRow + 1
    sh.Cells(newRow, 1).Value = rmCode
    Dim w As Long, col As Long
    For w = 1 To nWeeks
        col = 1 + w
        sh.Cells(newRow, col).Value = _
            "=IF(COUNTIFS(" & logTableName & "[Part Name],$A" & newRow & "," & logTableName & "[WeekIndex]," & w & ")=0,"""",SUMIFS(" & _
            logTableName & "[" & qtyColName & "]," & logTableName & "[Part Name],$A" & newRow & "," & logTableName & "[WeekIndex]," & w & "))"
    Next w
    sh.Rows(lastRow).Copy
    sh.Rows(newRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    AppendStockGridRow = newRow
End Function

' Dashboardの一番下に新しい材料の行を追加する(ssRow=T_SelfStock/T_TTAFStock側の行番号、
' grow=Grid_Requirement/Incoming/Stock側の行番号)。
Private Sub AppendDashboardRow(sh As Worksheet, rmCode As String, nWeeks As Long, ssRow As Long, grow As Long)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim newRow As Long: newRow = lastRow + 1
    Dim lastWeekCol As String: lastWeekCol = ColLetter(1 + nWeeks)

    sh.Cells(newRow, 1).Value = rmCode
    sh.Cells(newRow, 2).Value = "=IFERROR(INDEX(M_RawMaterials[Description],MATCH($A" & newRow & ",M_RawMaterials[Part Name],0)),"""")"
    sh.Cells(newRow, 3).Value = "=IFERROR(INDEX(M_RawMaterials[Category],MATCH($A" & newRow & ",M_RawMaterials[Part Name],0)),"""")"
    sh.Cells(newRow, 4).Value = "=IFERROR(INDEX(M_RawMaterials[基準在庫下限_要入力],MATCH($A" & newRow & ",M_RawMaterials[Part Name],0)),0)"
    sh.Cells(newRow, 5).Value = "=IFERROR(INDEX(M_RawMaterials[基準在庫上限_要入力],MATCH($A" & newRow & ",M_RawMaterials[Part Name],0)),0)"
    Dim ssSelfRng As String: ssSelfRng = "'T_SelfStock'!$B$" & ssRow & ":$" & lastWeekCol & "$" & ssRow
    Dim ssTTAFRng As String: ssTTAFRng = "'T_TTAFStock'!$B$" & ssRow & ":$" & lastWeekCol & "$" & ssRow
    Dim ssLabelRng As String: ssLabelRng = "'T_SelfStock'!$B$" & SS_TABLE_ROW & ":$" & lastWeekCol & "$" & SS_TABLE_ROW
    sh.Cells(newRow, 6).Value = "=IFERROR(LOOKUP(2,1/(" & ssSelfRng & "<>"""")," & ssSelfRng & "),"""")"
    sh.Cells(newRow, 7).Value = "=IFERROR(LOOKUP(2,1/(" & ssTTAFRng & "<>"""")," & ssTTAFRng & "),"""")"
    sh.Cells(newRow, 8).Value = "=IFERROR(LOOKUP(2,1/(" & ssSelfRng & "<>"""")," & ssLabelRng & "),"""")"

    Dim w As Long, col As Long
    For w = 1 To nWeeks
        col = 8 + w  ' Dashboardの週データ開始列=9(I列)
        sh.Cells(newRow, col).Value = "='Grid_Stock'!" & ColLetter(1 + w) & grow
    Next w

    sh.Rows(lastRow).Copy
    sh.Rows(newRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

' Material_Detailの一番下に新しい材料のブロックを追加する。追加直後はM_BOMに未登録のため、
' 中間体の行(No. of batches/使用量)は無く、「合計使用量(0のはず)・TTAF在庫・自社在庫・
' Order・合計在庫・注記」の6行だけのミニブロックになる。RefreshBOM実行後、この材料が
' 実際に使われている中間体が見つかれば、Grid_Requirementの合計使用量には反映されるが、
' このブロックに中間体の内訳行を追加するには、このマクロの再実行(削除して追加し直す)か、
' 手動での行追加が必要（内訳表示は必須ではなく、合計在庫の計算自体には影響しません）。
Private Sub AppendMaterialDetailBlock(sh As Worksheet, rmCode As String, descVal As String, nWeeks As Long, ssRow As Long, grow As Long)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row  ' B列(項目)基準
    Dim headerRow As Long: headerRow = lastRow + 2  ' 直前ブロックとの間に空白行を1行はさむ

    sh.Cells(headerRow, 1).Value = rmCode
    sh.Cells(headerRow, 2).Value = descVal

    Dim r As Long, w As Long, col As Long
    r = headerRow

    r = r + 1
    sh.Cells(r, 2).Value = "合計使用量(kg)/週"
    sh.Cells(r, 2).Font.Bold = True
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        sh.Cells(r, col).Value = "='Grid_Requirement'!" & ColLetter(1 + w) & grow
    Next w

    r = r + 1
    sh.Cells(r, 2).Value = "TTAF在庫(実績,kg)"
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        sh.Cells(r, col).Value = "='T_TTAFStock'!" & ColLetter(1 + w) & ssRow
    Next w

    r = r + 1
    sh.Cells(r, 2).Value = "自社在庫(実績,kg)"
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        sh.Cells(r, col).Value = "='T_SelfStock'!" & ColLetter(1 + w) & ssRow
    Next w

    r = r + 1
    sh.Cells(r, 2).Value = "Order(発注予定,kg)"
    Dim qtyExpr As String, monthExpr As String
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        qtyExpr = "SUMIFS(T_PlannedOrders[EffectiveQty],T_PlannedOrders[Part Name],$A" & headerRow & ",T_PlannedOrders[WeekIndex]," & w & ")"
        monthExpr = "SUMPRODUCT(MAX((T_PlannedOrders[Part Name]=$A" & headerRow & ")*(T_PlannedOrders[WeekIndex]=" & w & _
                    ")*(T_PlannedOrders[EffectiveQty]>0)*T_PlannedOrders[Order_Month]))"
        sh.Cells(r, col).Value = "=IF(" & qtyExpr & "=0,"""","& qtyExpr & "&"" (""&TEXT(" & monthExpr & ",""m月"")&""発注)"")"
    Next w

    r = r + 1
    sh.Cells(r, 2).Value = "合計在庫(週末時点,kg)"
    sh.Cells(r, 2).Font.Bold = True
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        sh.Cells(r, col).Value = "='Grid_Stock'!" & ColLetter(1 + w) & grow
    Next w

    r = r + 1
    sh.Cells(r, 2).Value = "（発注の目安はDashboardの基準在庫[下限/上限]と色分けを参照）"
    sh.Cells(r, 2).Font.Italic = True
    sh.Cells(r, 2).Font.Color = RGB(128, 128, 128)

    ' 罫線・書式のコピー: ヘッダー行は既存の任意のヘッダー行(最初の材料の行)から、
    ' 残り6行(合計使用量〜注記)は直前ブロックの末尾6行から複製する
    ' (ブロックの長さは材料によって違うが、末尾6行の並びは常に同じ順序のため)。
    On Error Resume Next
    Dim firstHeaderRow As Long: firstHeaderRow = MD_HEADER_ROW + 1
    sh.Rows(firstHeaderRow).Copy
    sh.Rows(headerRow).PasteSpecial xlPasteFormats
    sh.Rows((lastRow - 5) & ":" & lastRow).Copy
    sh.Rows((headerRow + 1) & ":" & r).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0

    ' MOQ入力欄(ヘッダー行のC列)は手入力用のため、コピーされた書式の上から個別に設定し直す
    With sh.Cells(headerRow, 3)
        .Value = Empty
        .Font.Bold = False
    End With
    On Error Resume Next
    sh.Cells(headerRow, 3).AddComment "MOQ(最小発注量)を入力してください（手書きでOK）"
    On Error GoTo 0
End Sub

' 該当カテゴリのPO_Draft_*シートの一番下に新しい材料の行を追加する。
Private Sub AppendPODraftRow(sh As Worksheet, rmCode As String, ttafCodeVal As String)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 4).End(xlUp).Row  ' D列(Part Name)基準
    Dim newRow As Long: newRow = lastRow + 1
    ' Grid_Stock内の行位置はMATCHで毎回動的に求める(材料の追加・削除で行位置がずれても
    ' 数式側が自動的に正しい行を追従できるようにするため)。
    Dim growMatch As String: growMatch = "MATCH($D" & newRow & ",Grid_Stock[Part Name],0)"

    sh.Cells(newRow, 2).Value = "=IFERROR(INDEX(M_RawMaterials[Description],MATCH(""" & rmCode & """,M_RawMaterials[Part Name],0)),"""")"
    sh.Cells(newRow, 3).Value = ttafCodeVal
    sh.Cells(newRow, 4).Value = rmCode
    sh.Cells(newRow, 5).Value = "kg"
    sh.Cells(newRow, 6).Value = "=IFERROR(INDEX(M_RawMaterials[基準在庫下限_要入力],MATCH(""" & rmCode & """,M_RawMaterials[Part Name],0)),0)"
    sh.Cells(newRow, 7).Value = "=INDEX(Grid_Stock[#Data]," & growMatch & ",$P$7)"

    Const PO_FIRST_WEEK_COL As Long = 8
    Const PO_N_WEEKS As Long = 13
    Dim w As Long, col As Long
    For w = 1 To PO_N_WEEKS
        col = PO_FIRST_WEEK_COL + w - 1
        sh.Cells(newRow, col).Value = "=MAX(0,$F" & newRow & "-INDEX(Grid_Stock[#Data]," & growMatch & ",$P$7+" & (w - 1) & "))"
    Next w
    Dim totalCol As Long: totalCol = PO_FIRST_WEEK_COL + PO_N_WEEKS
    sh.Cells(newRow, totalCol).Value = "=SUM(" & ColLetter(PO_FIRST_WEEK_COL) & newRow & ":" & ColLetter(PO_FIRST_WEEK_COL + PO_N_WEEKS - 1) & newRow & ")"

    sh.Rows(lastRow).Copy
    sh.Rows(newRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

' テーブル(ListObject)から、指定した材料コードに一致する行を削除する。
Private Sub DeleteMatchingTableRow(tbl As ListObject, rmCode As String)
    Dim n As Long: n = tbl.ListRows.Count
    Dim i As Long
    For i = n To 1 Step -1
        If UCase(Trim(CStr(tbl.ListRows(i).Range.Cells(1, 1).Value))) = UCase(rmCode) Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

' テーブル機能を使わない罫線グリッド(T_SelfStock/T_TTAFStock/Dashboard/PO_Draft_*)から、
' 指定した材料コードに一致する行を削除する。nameCol=材料コードが入っている列番号。
Private Sub DeleteMatchingGridRow(sh As Worksheet, rmCode As String, nameCol As Long)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, nameCol).End(xlUp).Row
    Dim r As Long
    For r = lastRow To 1 Step -1
        If UCase(Trim(CStr(sh.Cells(r, nameCol).Value))) = UCase(rmCode) Then
            sh.Rows(r).Delete
        End If
    Next r
End Sub

' Material_Detailの、指定した材料のブロック(ヘッダー行から次の材料のヘッダー行の直前まで、
' 空白の区切り行を含む)をまとめて削除する。BOM未登録等でブロックが無い場合は何もしない。
Private Sub DeleteMaterialDetailBlock(sh As Worksheet, rmCode As String)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row
    Dim startRow As Long: startRow = 0
    Dim r As Long
    For r = MD_HEADER_ROW + 1 To lastRow
        If UCase(Trim(CStr(sh.Cells(r, 1).Value))) = UCase(rmCode) Then
            startRow = r
            Exit For
        End If
    Next r
    If startRow = 0 Then Exit Sub

    Dim endRow As Long: endRow = lastRow
    For r = startRow + 1 To lastRow
        If Len(Trim(CStr(sh.Cells(r, 1).Value))) > 0 Then
            endRow = r - 1
            Exit For
        End If
    Next r
    sh.Rows(startRow & ":" & endRow).Delete
End Sub
