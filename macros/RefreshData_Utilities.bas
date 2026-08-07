Attribute VB_Name = "RefreshData_Utilities"
Option Explicit

Public Const MD_HEADER_ROW As Long = 6       ' Material_Detail: ヘッダー行。build_soh.pyのMD_TABLE_ROWと対応
Public Const MD_WEEK_START_COL As Long = 4   ' Material_Detail: 週データ開始列(D列)。build_soh.pyのWEEK_START_COLと対応
Public Const SS_TABLE_ROW As Long = 5        ' T_SelfStock/T_TTAFStock: 見出し行(週ラベル)。build_soh.pyのSS_TABLE_ROWと対応
Public Const DASH_DATA_START_ROW As Long = 7 ' Dashboard: 材料データの開始行。build_soh.pyのDATA_START_ROWと対応

' ============================================================================
' SOH管理ブック VBAマクロ 全体像（このコメントはRefreshData_Utilitiesモジュールにのみ書かれています）
'
' 目的: Python等の外部環境を使わず、Excel(VBA)だけで毎月のデータ更新を完結させる。
'
' 【モジュール構成】以前は1つの巨大な標準モジュール(RefreshData.bas)にすべてのマクロが
' 入っていましたが、メンテナンス性のため機能ごとに以下の7モジュールへ分割しています。
' マクロ名や動作は分割前と一切変わりません。
'
'   RefreshData_Utilities      : このモジュール。Public Const(MD_HEADER_ROW等)と、
'                                 複数モジュールから共通で使うヘルパー関数
'                                 (BuildNameIndex/BuildPairIndex/NormalizeText/
'                                 WeekIndexForDate/ColLetter)。他の全モジュールが
'                                 これに依存するため、最初にインポートしてください。
'   RefreshData_ProductionPlan : RefreshWeeklyBatches（「Powder & Slurry & Pgm Plan」を
'                                 取り込み、PP_Gridを更新）
'   RefreshData_BOM             : RefreshBOM（「Raw Material - Look Up」を
'                                 取り込み、M_BOM・Material_Detailの内訳行を更新）
'   RefreshData_StockActuals    : RefreshSelfStock・RefreshTTAFStock（自社/TTAF在庫実績の取込み）
'   RefreshData_Shipments       : RefreshShipments（CSA ReportのShipping Scheduleを取込み）
'   RefreshData_Display         : HideInactiveIntermediates・ShowAllIntermediates・
'                                 JumpToSelectedWeek（表示の切り替えのみ。データは変更しない）
'   RefreshData_MaterialMgmt    : AddMaterial・RemoveMaterial・RemoveIntermediate
'                                 （材料・中間体の追加/削除）
'
' 各マクロの詳しい説明は、それぞれが実装されているモジュールの冒頭コメントを参照してください。
' いずれのマクロも、それぞれ対応するシートだけを更新します。T_Shipments・T_OpeningStock・
' T_StockCount・SafetyStock_Qty等、運用中に手入力した内容には一切触れません。
'
' 【導入方法(全モジュール共通)】
'   1. SOH_Master.xlsm（マクロ有効ブックとして保存）を開く
'   2. Alt+F11 → 「ファイル」→「ファイルのインポート」→ macros/フォルダの7つの.basファイルを
'      すべて選択してインポートする（1ファイルずつでも、複数選択して一括でも構いません。
'      インポート順序は結果に影響しません）
'      （コピー＆貼り付けで導入する場合は、各ファイルの1行目の Attribute VB_Name = "..." を
'       必ず削除してから貼り付けてください。この行は貼り付けでは使えず、含めるとコンパイル
'       エラーになります。また、貼り付け先は各モジュール名に対応する標準モジュールを
'       個別に挿入してください）
'   3. Alt+F8 でマクロ一覧から実行したいマクロ(RefreshWeeklyBatches等)を選択、
'      または任意のシートに図形を挿入して「マクロの登録」で割り当てる
'
' 【選択週の自動スクロール(JumpToSelectedWeek)を有効にする場合の追加手順（任意）】
'   標準モジュールへのインポートだけでは動きません。以下のコードを
'   「Dashboard」シート・「Material_Detail」シート・「T_SelfStock」シート・「T_TTAFStock」シート
'   それぞれの“シート自身のコードモジュール”に直接貼り付けてください（VBEのプロジェクト
'   エクスプローラーでシート名をダブルクリックすると開きます。標準モジュールに貼り付けても
'   発火しません）。JumpToSelectedWeek本体はRefreshData_Displayモジュールにあります。
'
'   ' --- Dashboardシートのコードモジュールに貼り付け ---
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Intersect(Target, Me.Range("C1")) Is Nothing Then Exit Sub
'       Call JumpToSelectedWeek(Me, "F1", 11)   ' 11 = K列(週データ開始列)
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
'
' 【パフォーマンスについて】どのRefresh*マクロも、外部ファイルのシートを1セルずつ.Cells(r,c).Value
' で読む代わりに、対象範囲を1回だけ配列として読み込み(Range.Value)、以降はメモリ上の配列だけを
' 参照する設計にしています。また、PP_Grid(中間体名->行番号)・M_BOM(Intermediate|RM_Code->行番号)・
' T_SelfStock_Log/T_TTAFStock_Log((RM_Code,週の月曜日)->行番号)への書き込みも、呼び出すたびに
' .Find()や全行スキャンをする代わりに、実行の最初にDictionaryを1回だけ作って参照する設計です
' （このモジュールのBuildNameIndex/BuildPairIndex、RefreshData_StockActualsのBuildStockRowIndex）。
' これは実際にExcelが強制終了する不具合(1セルずつの読み書きや毎回の全行スキャンがCOM通信の
' 積み重ねで極めて遅くなることが原因)が複数回報告されたことを受けての対策です。
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
' した。取込元ファイル(Powder & Slurry & Pgm Plan、Raw Material - Look Up等)は
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
' ことがないようにするためです。詳細はRefreshData_StockActualsモジュール冒頭を参照。
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
' ============================================================================

' tblの1・2列目から「1列目の値|2列目の値」->行番号のDictionaryを1回だけ作る。
' M_BOM(Intermediate|RM_Code)・T_Shipments(Part Name|PO_No)など、
' 「先頭2列の組み合わせで行を特定する」テーブル全般に使う汎用ヘルパー。
' RefreshData_ProductionPlan・RefreshData_BOMの両方から使われる共通処理。
Public Function BuildPairIndex(tbl As ListObject) As Object
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
Public Function BuildNameIndex(tbl As ListObject, colName As String) As Object
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

' 材料名・中間体名等の表記ゆれ(大文字小文字・記号・スペースの違い)を吸収するための
' 正規化(英数字だけを大文字で残す)。RefreshData_BOM・RefreshData_StockActuals・
' RefreshData_Shipmentsの複数モジュールで、外部ファイルの名称とM_RawMaterials/PP_Grid側の
' 名称を突き合わせる際に共通して使う。
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

' 日付からCal_Weeksを引いてWeekIndexを求める。RefreshData_StockActuals(実績取込みの
' 対象週判定)とRefreshData_Display(HideInactiveIntermediatesの「今週」判定)の両方から
' 使われる共通処理。
Public Function WeekIndexForDate(wb As Workbook, d As Date) As Long
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

' 列番号(例:28)を列名(例:AB)に変換する。ワークシートに依存しない純粋な計算。
' RefreshData_BOM(InsertIntermediateRowPair)とRefreshData_MaterialMgmt(Append*/Delete*系)の
' 両方から使われる共通処理。
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
