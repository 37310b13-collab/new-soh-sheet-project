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
Private Const PO_MONTHYEAR_ROW As Long = 20  ' build_soh.pyのPO_HDR_MONTHYEAR_ROWと対応
Private Const PO_FIRST_WEEK_COL As Long = 8  ' H列
Private Const PO_N_WEEKS As Long = 13
Private Const PO_BASEWEEK_ADDR As String = "$P$13"
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

Private Function HasLocalName(sh As Worksheet, nm As String) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = sh.Names(nm)
    On Error GoTo 0
    HasLocalName = Not (n Is Nothing)
End Function

' PO_Draft_Hazardous自身の不具合を修正する。既に修正済み(名前付き範囲"BaseWeek"が
' 既にある)場合は何もせずFalseを返す。
Private Function FixHazardousPODraftLayout(thisWb As Workbook) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = thisWb.Sheets("PO_Draft_Hazardous")
    On Error GoTo 0
    If sh Is Nothing Then
        FixHazardousPODraftLayout = False
        Exit Function
    End If
    If HasLocalName(sh, "BaseWeek") Then
        FixHazardousPODraftLayout = False
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ---- 名前付き範囲を追加(シート固有のローカルスコープ) ----
    sh.Names.Add Name:="BaseWeek", RefersTo:="=" & PO_BASEWEEK_ADDR
    sh.Names.Add Name:="PORevision", RefersTo:="=" & PO_REVISION_ADDR

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

    ' ---- Firm(1〜4週目)/Forecast(5〜13週目)の発注数量セルに色を付ける ----
    If lastRow >= PO_DATA_START_ROW Then
        For r = PO_DATA_START_ROW To lastRow
            For c = PO_FIRST_WEEK_COL To PO_FIRST_WEEK_COL + PO_N_WEEKS - 1
                With sh.Cells(r, c)
                    If c < PO_FIRST_WEEK_COL + 4 Then
                        .Interior.Color = RGB(255, 193, 193)
                        .Font.Color = RGB(192, 0, 0)   ' 濃い赤: C00000
                    Else
                        .Interior.Color = RGB(235, 241, 222)
                        .Font.Color = RGB(0, 97, 0)    ' 濃い緑: 006100
                    End If
                End With
            Next c
        Next r
    End If

    ' ---- SafetyStock/CurrentStock(F/G列)の非表示を解除し、印刷範囲から除外する ----
    sh.Columns("F:G").Hidden = False
    sh.Columns("F").ColumnWidth = 12
    sh.Columns("G").ColumnWidth = 12
    Dim printLastRow As Long
    printLastRow = IIf(lastRow >= PO_DATA_START_ROW, lastRow, PO_HDR_ROW)
    sh.PageSetup.PrintArea = "$A$1:$E$" & printLastRow & ",$H$1:$U$" & printLastRow

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    FixHazardousPODraftLayout = True
End Function

' PO_Draft_Hazardous(修正済み)の構造を複製して、targetSheetNameのシートを新レイアウトに
' 作り直す。既に新レイアウト(名前付き範囲"BaseWeek"が既にある)ならFalseを返して何もしない。
Private Function ClonePODraftLetterheadIfNeeded(thisWb As Workbook, targetSheetName As String, titleText As String) As Boolean
    Dim oldSh As Worksheet
    On Error Resume Next
    Set oldSh = thisWb.Sheets(targetSheetName)
    On Error GoTo 0
    If oldSh Is Nothing Then
        ClonePODraftLetterheadIfNeeded = False
        Exit Function
    End If
    If HasLocalName(oldSh, "BaseWeek") Then
        ClonePODraftLetterheadIfNeeded = False
        Exit Function
    End If

    Dim hazSh As Worksheet
    On Error Resume Next
    Set hazSh = thisWb.Sheets("PO_Draft_Hazardous")
    On Error GoTo 0
    If hazSh Is Nothing Or Not HasLocalName(hazSh, "BaseWeek") Then
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
    If Not IsNumeric(oldBaseWeek) Then oldBaseWeek = hazSh.Range("BaseWeek").Value

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

    ' ---- レターヘッド: タイトル・TO/FROM/CCは仮の文字列に戻す(Hazardousの実際の宛先を
    ' そのまま複製しない。カテゴリによって担当者・取引先が異なる可能性があるため) ----
    newSh.Range("B8").Value = "TO：（サプライヤー／TTAF担当者名を入力）"
    newSh.Range("B9").Value = "　　　　　（会社名）"
    newSh.Range("B11").Value = "CC：（必要であれば入力）"
    newSh.Range("B13").Value = "FROM：（発行者名）"
    newSh.Range("B14").Value = "　　　　　（自社名）"
    newSh.Range("B17").Value = titleText
    newSh.Range("BaseWeek").Value = oldBaseWeek
    newSh.Range("PORevision").Value = oldRevision

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
    newSh.PageSetup.PrintArea = "$A$1:$E$" & finalLastRow & ",$H$1:$U$" & finalLastRow

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ClonePODraftLetterheadIfNeeded = True
End Function
