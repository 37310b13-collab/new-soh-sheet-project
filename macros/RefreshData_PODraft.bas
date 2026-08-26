Attribute VB_Name = "RefreshData_PODraft"
Option Explicit

' ============================================================================
' RefreshData_PODraft モジュール
'
' 【導入方法】Alt+F11 → 「ファイル」→「ファイルのインポート」→ このファイルを選択
' （コピー＆貼り付けで導入する場合は、1行目の Attribute VB_Name = "..." を必ず削除してから
'  貼り付けてください）
'
' SetupPODraftLetterheadLayout : 【一度だけ実行する移行用マクロ】
'   PO_Draft_Hazardousシートに手動で作り込んだレターヘッド付きレイアウト(TO/FROM/CC・
'   Order Date/Issue Month/Firm Month・Revision・基準週(WeekIndex)・Firm/Forecastの
'   色分け・SafetyStock/CurrentStockを印刷範囲外に配置)を、
'     ①PO_Draft_Hazardous自身に残っていた不具合を修正した上で、
'     ②PO_Draft_Chemical・PO_Draft_Substrate_JPN_CHN・PO_Draft_Substrate_Polandの
'       3シートにも同じレイアウトを複製する。
'
'   ①で修正する不具合:
'     ・月/週見出し(20行目)が固定幅(1〜4週目/5〜8週目/9〜13週目)のセル結合になっており、
'       月によって実際の週数が変わる(例: 2026年11月は5週にまたがる)と、その月の見出しが
'       消えてしまうことがあった(実際にForecast側の年月が消える不具合として報告された)。
'       → セル結合をやめ、「週ごとに1セル、直前の週と月が違う時だけ表示する」という
'       build_soh.py本来の方式に戻す(月の長さに関わらず常に正しく表示される)。
'     ・基準週セルをP7からP13へ手動で移動した際、既存材料の一部の行(ND TAC/CHEM-1280等)の
'       数式が古い$P$7参照のまま取り残されており、その材料だけ発注数量が常に0のまま
'       更新されなくなっていた(空欄のP7セルを参照し続けるため)。
'       → 全データ行の$P$7・$P$13の直接参照を、シート固有の名前付き範囲"BaseWeek"に
'       統一する。以後、基準週セルをどこに動かしても名前の参照先を直すだけで済み、
'       AddMaterial/SyncPODraftCategoriesで新規に追加される行も自動的に追従する
'       (RefreshData_MaterialMgmt.basのAppendPODraftRow・RefreshData_UtilitiesのBaseWeekRef
'       を参照)。
'     ・Firm(1〜4週目)/Forecast(5〜13週目)の発注数量セルに、依頼のあった赤/緑の色分けが
'       付いていなかった。
'     ・SafetyStock/CurrentStock(F/G列)は列を非表示にする形で印刷対象から外していたが、
'       非表示を解除すると印刷にも写ってしまう(参照用に確認したいときだけ表示し直す、
'       という運用が事実上できていなかった)。
'       → 列の非表示を解除して常に見える状態に戻した上で、印刷範囲そのものから除外する
'       (Excelの印刷範囲は複数の矩形を指定でき、非表示に頼らず狙った列だけ除外できる)。
'
'   ②でPO_Draft_Chemical/_Substrate_JPN_CHN/_Substrate_Polandを複製する際:
'     ・TO/FROM/CC欄はPO_Draft_Hazardousの内容をそのままコピーしない(カテゴリによって
'       担当者・取引先が異なる可能性があるため、複製後は仮の文字列に戻す。実際の宛先は
'       複製後に手入力してください)。
'     ・Revision・基準週(WeekIndex)は、複製前の各シート自身の値をそのまま引き継ぐ。
'     ・データ行(材料一覧)は複製前の内容を使い回さず、M_RawMaterialsの現在の
'       Category・Origin_Countryを基準に作り直す(POSheetNameForMaterialと同じ判定。
'       SyncPODraftCategoriesと同じ考え方で、複製時点の最新の状態が反映される)。
'     ・ロゴ画像・デコレーション(バナー等)はExcelのシートコピー機能により自動的に
'       複製されるが、バナーに日付文字列が手入力されている場合は、複製後に各シートで
'       内容を確認・修正してください。
'
'   どちらも、既に移行済みの部分はスキップするため、誤って複数回実行しても安全。
' ============================================================================

Private Const PO_HDR_ROW As Long = 26        ' 見出し最終行(この直下からデータ行)。build_soh.pyのPO_HDR_UOM_FIRM_ROWと対応
Private Const PO_DATA_START_ROW As Long = 27 ' build_soh.pyのPO_DATA_START_ROWと対応
Private Const PO_TITLE_ROW As Long = 17      ' build_soh.pyのPO_TITLE_ROWと対応(17〜18行目を結合)
Private Const PO_MONTHYEAR_ROW As Long = 20  ' build_soh.pyのPO_HDR_MONTHYEAR_ROWと対応
Private Const PO_FIRST_WEEK_COL As Long = 8  ' H列
Private Const PO_N_WEEKS As Long = 13
Private Const PO_BASEWEEK_ADDR As String = "$P$13"
Private Const PO_BASEWEEK_ROW As Long = 13   ' PO_BASEWEEK_ADDRの行番号(印刷範囲の除外に使う)
Private Const PO_REVISION_ADDR As String = "$P$11"

Sub SetupPODraftLetterheadLayout()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook

    Dim hazFixed As Boolean
    hazFixed = FixHazardousPODraftLayout(thisWb)

    Dim c1 As Boolean, c2 As Boolean, c3 As Boolean
    c1 = ClonePODraftLetterheadIfNeeded(thisWb, "PO_Draft_Chemical", "Chemicals : TTAF Supply")
    c2 = ClonePODraftLetterheadIfNeeded(thisWb, "PO_Draft_Substrate_JPN_CHN", "Substrates (Japan / China)")
    c3 = ClonePODraftLetterheadIfNeeded(thisWb, "PO_Draft_Substrate_Poland", "Substrates (Poland)")

    If Not hazFixed And Not c1 And Not c2 And Not c3 Then
        MsgBox "既に移行済みです(すべてのPO_Draft_*シートが新しいレイアウトになっています)。", vbInformation
        Exit Sub
    End If

    MsgBox "PO_Draft_*シートのレイアウト移行が完了しました。" & vbCrLf & vbCrLf & _
           "PO_Draft_Hazardous: " & IIf(hazFixed, "修正しました(月/週見出し・基準週参照・色分け・印刷範囲)", "既に修正済みでした") & vbCrLf & _
           "PO_Draft_Chemical: " & IIf(c1, "新レイアウトへ変更しました", "既に新レイアウトでした") & vbCrLf & _
           "PO_Draft_Substrate_JPN_CHN: " & IIf(c2, "新レイアウトへ変更しました", "既に新レイアウトでした") & vbCrLf & _
           "PO_Draft_Substrate_Poland: " & IIf(c3, "新レイアウトへ変更しました", "既に新レイアウトでした") & vbCrLf & vbCrLf & _
           "新しく作成された3シートのTO/FROM/CC欄は仮の文字列のままです。実際の宛先・自社名を" & vbCrLf & _
           "入力してください(PO_Draft_Hazardousの内容はコピーしていません)。" & vbCrLf & vbCrLf & _
           "ロゴ画像・バナーはPO_Draft_Hazardousから複製されています。バナーに日付が手入力されて" & vbCrLf & _
           "いる場合は、各シートで内容を確認・修正してください。", vbInformation
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました: (" & Err.Number & ") " & Err.Description & vbCrLf & vbCrLf & _
           "途中まで反映されている可能性があります。心配な場合は保存せずに閉じて開き直してください。", vbCritical
End Sub

' シート固有(ローカルスコープ)の名前付き範囲"BaseWeek"・"PORevision"を、必ずそのシート自身の
' セルを指すように作る(既にあれば参照先を上書きする。無ければ新規作成する)。RefersToには
' 必ずシート名を明示的に含める("='" & sh.Name & "'!..." の形)。シート名を省略すると、
' Names.Add実行時にたまたまアクティブだった別のシートを指してしまうことがあり
' (実際にこの不具合で実行時エラー1004が発生した)、この関数はその対策として、既に
' 移行済みかどうかに関わらず毎回呼び出して参照先を正しく上書きする設計にしている。
Private Sub EnsureLocalBaseWeekNames(sh As Worksheet)
    Dim bwName As Name
    On Error Resume Next
    Set bwName = sh.Names("BaseWeek")
    On Error GoTo 0
    If bwName Is Nothing Then
        sh.Names.Add Name:="BaseWeek", RefersTo:="='" & sh.Name & "'!" & PO_BASEWEEK_ADDR
    Else
        bwName.RefersTo = "='" & sh.Name & "'!" & PO_BASEWEEK_ADDR
    End If

    Dim revName As Name
    On Error Resume Next
    Set revName = sh.Names("PORevision")
    On Error GoTo 0
    If revName Is Nothing Then
        sh.Names.Add Name:="PORevision", RefersTo:="='" & sh.Name & "'!" & PO_REVISION_ADDR
    Else
        revName.RefersTo = "='" & sh.Name & "'!" & PO_REVISION_ADDR
    End If
End Sub

' PO_Draft_Hazardous自身の不具合を修正する。既に修正済み(月/週見出し(20行目)の結合が
' 既に解除されている)場合は、名前付き範囲の参照先だけ念のため確認・修正した上で、
' Falseを返して残りの処理はスキップする。
Private Function FixHazardousPODraftLayout(thisWb As Workbook) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = thisWb.Sheets("PO_Draft_Hazardous")
    On Error GoTo 0
    If sh Is Nothing Then
        FixHazardousPODraftLayout = False
        Exit Function
    End If

    ' 名前付き範囲は、既に移行済みかどうかに関わらず毎回作り直す(前回実行が
    ' 途中でエラー停止していた場合、名前が正しく設定されていない可能性があるため)。
    Call EnsureLocalBaseWeekNames(sh)

    If Not sh.Cells(PO_MONTHYEAR_ROW, PO_FIRST_WEEK_COL).MergeCells Then
        ' 月/週見出しが既に結合解除済み=既にこの関数の本体を実行済み
        FixHazardousPODraftLayout = False
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ---- 全データ行の$P$7/$P$13直接参照をBaseWeek名に統一する(G列=在庫参照、
    ' H〜T列=発注数量) ----
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 5).End(xlUp).Row  ' E列(UOM)基準
    Dim r As Long, c As Long
    If lastRow >= PO_DATA_START_ROW Then
        For r = PO_DATA_START_ROW To lastRow
            Dim gCell As Range: Set gCell = sh.Cells(r, 7)
            If Len(gCell.Formula) > 0 Then
                gCell.Formula = Replace(Replace(gCell.Formula, "$P$7", "BaseWeek"), "$P$13", "BaseWeek")
            End If
            For c = PO_FIRST_WEEK_COL To PO_FIRST_WEEK_COL + PO_N_WEEKS - 1
                Dim wCell As Range: Set wCell = sh.Cells(r, c)
                If Len(wCell.Formula) > 0 Then
                    wCell.Formula = Replace(Replace(wCell.Formula, "$P$7", "BaseWeek"), "$P$13", "BaseWeek")
                End If
            Next c
        Next r
    End If

    ' ---- 月/週見出し(20行目)の固定結合をやめ、週ごとの数式に書き換える ----
    Dim col As Long
    For col = PO_FIRST_WEEK_COL To PO_FIRST_WEEK_COL + PO_N_WEEKS - 1
        If sh.Cells(PO_MONTHYEAR_ROW, col).MergeCells Then
            sh.Cells(PO_MONTHYEAR_ROW, col).MergeArea.UnMerge
        End If
    Next col
    Dim w As Long
    For w = 1 To PO_N_WEEKS
        col = PO_FIRST_WEEK_COL + w - 1
        Dim f As String
        If w = 1 Then
            f = "=INDEX(Cal_Weeks[MonthYearLabel],BaseWeek)"
        Else
            f = "=IF(INDEX(Cal_Weeks[MonthYearLabel],BaseWeek+" & (w - 1) & ")<>INDEX(Cal_Weeks[MonthYearLabel],BaseWeek+" & (w - 2) & "),INDEX(Cal_Weeks[MonthYearLabel],BaseWeek+" & (w - 1) & "),"""")"
        End If
        With sh.Cells(PO_MONTHYEAR_ROW, col)
            .Formula = f
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(191, 191, 191)
            If w <= 4 Then
                .Interior.Color = RGB(255, 193, 193)  ' Firm: FFC1C1
            Else
                .Interior.Color = RGB(235, 241, 222)  ' Forecast: EBF1DE
            End If
        End With
    Next w

    ' ---- Firm(1〜4週目)/Forecast(5〜13週目)の発注数量セルの色分けは、直接の塗りつぶし
    ' ではなく条件付き書式にする(発注数量が0の週は色を付けない・数字も表示しない、
    ' という要望のため)。将来AddMaterial等で追加される行も対象に含まれるよう、
    ' 実データより十分広い行範囲(500行分)に対して設定する。数値の表示形式も、
    ' 0を表示しない書式にする。
    Call ApplyPODraftZeroHiddenFormatting(sh)

    ' ---- SafetyStock/CurrentStock(F/G列)の非表示を解除し、印刷範囲から除外する ----
    sh.Columns("F:G").Hidden = False
    sh.Columns("F").ColumnWidth = 12
    sh.Columns("G").ColumnWidth = 12
    Dim printLastRow As Long
    printLastRow = IIf(lastRow >= PO_DATA_START_ROW, lastRow, PO_HDR_ROW)
    Call SetPODraftPrintArea(sh, printLastRow)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    FixHazardousPODraftLayout = True
End Function

' PO_Draft_Hazardous(修正済み)の構造を複製して、targetSheetNameのシートを新レイアウトに
' 作り直す。既に新レイアウト(タイトル行(17〜18行目)が既に結合されている)ならFalseを返して
' 何もしない。
Private Function ClonePODraftLetterheadIfNeeded(thisWb As Workbook, targetSheetName As String, titleText As String) As Boolean
    Dim oldSh As Worksheet
    On Error Resume Next
    Set oldSh = thisWb.Sheets(targetSheetName)
    On Error GoTo 0
    If oldSh Is Nothing Then
        ClonePODraftLetterheadIfNeeded = False
        Exit Function
    End If
    If oldSh.Cells(PO_TITLE_ROW, 2).MergeCells Then
        ' タイトル行が既に結合済み=既に新レイアウトへ移行済み
        ClonePODraftLetterheadIfNeeded = False
        Exit Function
    End If

    Dim hazSh As Worksheet
    On Error Resume Next
    Set hazSh = thisWb.Sheets("PO_Draft_Hazardous")
    On Error GoTo 0
    If hazSh Is Nothing Then
        MsgBox "PO_Draft_Hazardousシートが見つかりません。", vbExclamation
        ClonePODraftLetterheadIfNeeded = False
        Exit Function
    End If
    If Not hazSh.Cells(PO_TITLE_ROW, 2).MergeCells Then
        MsgBox "PO_Draft_Hazardousが先に新レイアウトへ修正されている必要があります。" & vbCrLf & _
               "SetupPODraftLetterheadLayoutをもう一度実行し直してください。", vbExclamation
        ClonePODraftLetterheadIfNeeded = False
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 既存シート(旧レイアウト)から引き継ぐ値を、削除する前に読み取っておく
    ' (旧レイアウトではRevision=P5、基準週=P7)。
    Dim oldRevision As Variant: oldRevision = oldSh.Range("P5").Value
    If Not IsNumeric(oldRevision) Then oldRevision = "00"
    Dim oldBaseWeek As Variant: oldBaseWeek = oldSh.Range("P7").Value
    If Not IsNumeric(oldBaseWeek) Then oldBaseWeek = hazSh.Range(PO_BASEWEEK_ADDR).Value

    Dim oldIdx As Long: oldIdx = oldSh.Index
    Application.DisplayAlerts = False
    oldSh.Delete
    Application.DisplayAlerts = True

    hazSh.Copy After:=thisWb.Sheets(thisWb.Sheets.Count)
    Dim newSh As Worksheet: Set newSh = thisWb.Sheets(thisWb.Sheets.Count)
    newSh.Name = targetSheetName
    ' 元あった位置の近くへ移動する(見た目の並び順を保つため。失敗しても実害は無いので無視する)
    On Error Resume Next
    newSh.Move Before:=thisWb.Sheets(Application.WorksheetFunction.Min(oldIdx, thisWb.Sheets.Count))
    On Error GoTo 0

    ' シートコピーで名前付き範囲が正しく複製されているとは限らないため、コピー先(newSh)
    ' 自身を指すように明示的に作り直す(EnsureLocalBaseWeekNames参照)。
    Call EnsureLocalBaseWeekNames(newSh)

    ' ---- レターヘッド: タイトル・TO/FROM/CCは仮の文字列に戻す(Hazardousの実際の宛先を
    ' そのまま複製しない。カテゴリによって担当者・取引先が異なる可能性があるため) ----
    newSh.Range("B8").Value = "TO：（サプライヤー／TTAF担当者名を入力）"
    newSh.Range("B9").Value = "　　　　　（会社名）"
    newSh.Range("B11").Value = "CC：（必要であれば入力）"
    newSh.Range("B13").Value = "FROM：（発行者名）"
    newSh.Range("B14").Value = "　　　　　（自社名）"
    newSh.Range("B17").Value = titleText
    newSh.Range(PO_BASEWEEK_ADDR).Value = oldBaseWeek
    newSh.Range(PO_REVISION_ADDR).Value = oldRevision

    ' ---- 複製元(Hazardous)のデータ行を削除し、M_RawMaterialsの現在のCategory・
    ' Origin_Countryを基準に、このシートに載るべき材料の行を作り直す
    ' (SyncPODraftCategoriesと同じ判定基準。POSheetNameForMaterial参照) ----
    Dim oldLastRow As Long: oldLastRow = newSh.Cells(newSh.Rows.Count, 5).End(xlUp).Row
    If oldLastRow >= PO_DATA_START_ROW Then
        newSh.Rows(PO_DATA_START_ROW & ":" & oldLastRow).Delete
    End If

    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")
    Dim nWeeks As Long: nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count
    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    Dim addedItems As Long: addedItems = 0
    If rmN > 0 Then
        Dim rmData As Variant: rmData = rmTbl.ListColumns(1).DataBodyRange.Resize(rmN, 11).Value
        Dim i As Long
        For i = 1 To rmN
            Dim rmCode As String: rmCode = Trim(CStr(rmData(i, 1)))
            Dim catVal As String: catVal = Trim(CStr(rmData(i, 4)))
            Dim ttafCodeVal As String: ttafCodeVal = Trim(CStr(rmData(i, 9)))
            Dim originVal As String: originVal = Trim(CStr(rmData(i, 11)))
            If POSheetNameForMaterial(catVal, originVal) = targetSheetName Then
                Call AppendPODraftRow(newSh, rmCode, ttafCodeVal, nWeeks)
                addedItems = addedItems + 1
            End If
        Next i
    End If
    If addedItems = 0 Then
        newSh.Cells(PO_DATA_START_ROW, 2).Value = "(該当品目なし)"
    End If

    ' ---- 印刷範囲を、実際のデータ行数に合わせて更新する ----
    Dim finalLastRow As Long
    finalLastRow = IIf(addedItems > 0, PO_DATA_START_ROW + addedItems - 1, PO_HDR_ROW)
    Call SetPODraftPrintArea(newSh, finalLastRow)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ClonePODraftLetterheadIfNeeded = True
End Function

' 【いつでも実行してよいメンテナンス用マクロ】4つのPO_Draft_*シートすべてに、
' 発注数量0の週を非表示にする書式(条件付き書式・数値の表示形式)と、基準週セルを
' 印刷範囲から除外する設定を(再)適用する。SetupPODraftLetterheadLayoutは、既に
' 新レイアウトへ移行済みのシートをスキップしてしまうため、移行後にこの書式だけを
' 追加・更新したい場合はこのマクロを直接実行する。何度実行しても安全。
Sub ApplyPODraftZeroHiddenFormattingToAllSheets()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Application.ScreenUpdating = False

    Dim sheetNames As Variant
    sheetNames = Array("PO_Draft_Chemical", "PO_Draft_Hazardous", "PO_Draft_Substrate_JPN_CHN", "PO_Draft_Substrate_Poland")
    Dim si As Long, n As Long: n = 0
    For si = LBound(sheetNames) To UBound(sheetNames)
        Dim sh As Worksheet
        On Error Resume Next
        Set sh = Nothing
        Set sh = thisWb.Sheets(CStr(sheetNames(si)))
        On Error GoTo ErrHandler
        If Not sh Is Nothing Then
            Call ApplyPODraftZeroHiddenFormatting(sh)
            Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 5).End(xlUp).Row
            If lastRow < PO_HDR_ROW Then lastRow = PO_HDR_ROW
            Call SetPODraftPrintArea(sh, lastRow)
            n = n + 1
        End If
    Next si

    Application.ScreenUpdating = True
    MsgBox n & "シートに適用しました。保存して確認してください。", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました: (" & Err.Number & ") " & Err.Description, vbCritical
End Sub

' Firm(1〜4週目)/Forecast(5〜13週目)の発注数量セルに、条件付き書式で色を付ける
' (発注数量が0以外の時だけ発色。0の週は色を付けず、数字も表示しない)。
' 実データの行数より十分広い範囲(500行分)に対して設定することで、AddMaterial等で
' 後から追加される行も自動的に対象に含まれるようにする(行を追加するたびに
' 個別に条件付き書式を設定し直す必要が無い)。既存の条件付き書式があれば一旦削除
' してから設定し直すため、何度実行しても安全。
Private Sub ApplyPODraftZeroHiddenFormatting(sh As Worksheet)
    Dim cfLastRow As Long: cfLastRow = PO_DATA_START_ROW + 500
    Dim firmFirstCol As Long: firmFirstCol = PO_FIRST_WEEK_COL
    Dim firmLastCol As Long: firmLastCol = PO_FIRST_WEEK_COL + 3
    Dim forecastFirstCol As Long: forecastFirstCol = PO_FIRST_WEEK_COL + 4
    Dim forecastLastCol As Long: forecastLastCol = PO_FIRST_WEEK_COL + PO_N_WEEKS - 1

    Dim allWeeksRng As Range
    Set allWeeksRng = sh.Range(sh.Cells(PO_DATA_START_ROW, firmFirstCol), sh.Cells(cfLastRow, forecastLastCol))
    allWeeksRng.NumberFormat = "0;-0;;@"
    ' 以前のバージョン(直接の塗りつぶし)やFixMergedPODraftDataRows等の緊急復旧マクロで
    ' 直接色が付いてしまっている可能性があるセルをリセットする。条件付き書式は
    ' 条件に一致しない時は元の(直接指定の)書式にフォールバックするため、直接色が
    ' 残ったままだと0の週でも色が消えない。
    allWeeksRng.Interior.ColorIndex = xlNone
    allWeeksRng.Font.ColorIndex = xlAutomatic

    Dim firmRng As Range: Set firmRng = sh.Range(sh.Cells(PO_DATA_START_ROW, firmFirstCol), sh.Cells(cfLastRow, firmLastCol))
    firmRng.FormatConditions.Delete
    Dim firmAnchor As String: firmAnchor = ColLetter(firmFirstCol) & PO_DATA_START_ROW
    Dim fc1 As FormatCondition
    Set fc1 = firmRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(" & firmAnchor & "<>0," & firmAnchor & "<>"""")")
    fc1.Interior.Color = RGB(255, 193, 193)  ' Firm: FFC1C1
    fc1.Font.Color = RGB(192, 0, 0)          ' 濃い赤: C00000

    Dim forecastRng As Range: Set forecastRng = sh.Range(sh.Cells(PO_DATA_START_ROW, forecastFirstCol), sh.Cells(cfLastRow, forecastLastCol))
    forecastRng.FormatConditions.Delete
    Dim forecastAnchor As String: forecastAnchor = ColLetter(forecastFirstCol) & PO_DATA_START_ROW
    Dim fc2 As FormatCondition
    Set fc2 = forecastRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(" & forecastAnchor & "<>0," & forecastAnchor & "<>"""")")
    fc2.Interior.Color = RGB(235, 241, 222)  ' Forecast: EBF1DE
    fc2.Font.Color = RGB(0, 97, 0)           ' 濃い緑: 006100
End Sub

' PO_Draft_*シートの印刷範囲を設定する。SafetyStock/CurrentStock(F/G列)に加え、
' 基準週(WeekIndex)の入力欄(PO_BASEWEEK_ROW行目のN/P列)も、発注書を印刷・発行する
' 際には不要な内部操作用のセルのため、H:U列の印刷範囲をその行の前後で分割して除外する。
Private Sub SetPODraftPrintArea(sh As Worksheet, lastRow As Long)
    sh.PageSetup.PrintArea = _
        "$A$1:$E$" & lastRow & "," & _
        "$H$1:$U$" & (PO_BASEWEEK_ROW - 1) & "," & _
        "$H$" & (PO_BASEWEEK_ROW + 1) & ":$U$" & lastRow
End Sub
