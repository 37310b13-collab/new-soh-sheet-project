Attribute VB_Name = "RefreshData_BOM"
Option Explicit

' ============================================================================
' RefreshData_BOM モジュール
'
'   RefreshBOM : 「Raw Material - Look Up」を選択すると、M_BOM（原単位）を更新する。
'                対象は4シート:
'                  ・Slurry Data Base / Powder Data Base / Solution … 行そのまま転記。
'                    A列=Intermediate、D列=RM Code、M列=1バッチあたり使用量。
'                    RM CodeにはM_RawMaterials上のRM_Code(例:CHEM-1030)だけでなく、
'                    他の中間体コード(SOL-SCH・TPP-103等、Slurry/PowderがSolutionや
'                    別のPowderを材料として使うケース)がそのまま入ることがあるが、
'                    そのままM_BOMの行として書き込む(Grid_Requirement側で無視される
'                    だけで実害は無い)。
'                  ・Catalyst Data Base … Substrate行(H列"Substrate SC"が埋まっている行)
'                    だけを使う(化学品・スラリー参照行はSlurry側で既に原単位計算済み
'                    のため、ここで拾うと二重計上になる)。Corning供給のSubstrate
'                    (E列Descriptionに"CORNING"を含む)は対象外。中間体名はA列の
'                    Catalyst名(例:"18461-0Q110-1st COAT")から短縮製品コード
'                    (例:"0Q110")を抽出したもの、RM_CodeはH列、数量はF列
'                    (catalyst1個あたりの使用量。PP_Grid側でcatalystは個数管理のため)。
'                新しい材料×中間体の組み合わせが見つかった場合、Material_Detailの
'                該当材料ブロックにも中間体の内訳行（No. of batches／使用量(kg)）を
'                自動的に追加する（SyncMaterialDetailIntermediates）。これにより、
'                AddMaterialで追加したばかりの材料(BOM未登録のため内訳行が無いミニ
'                ブロック)も、RefreshBOMを1回実行するだけで既存の材料と同じ見た目になる
'                (AddMaterialを2回実行する必要はない)。既存材料が新しい中間体で
'                使われ始めた場合も同様に自動反映される。
'
'   FixPassthroughCircularRefs : PP_Gridのパススルー数式(独自の生産計画を持たず、他の
'                中間体の原料としてのみ使われる中間体の行に入っている、SUMPRODUCT+
'                M_BOM[PPGridRow]による逆算数式)には、その中間体自身のレシピ行についても
'                INDEX(PP_Grid[#Data],M_BOM[PPGridRow],...)が評価されてしまい、最終的に
'                0倍されて計算結果には影響しなくても「そのセル自身を参照する経路」が
'                実際に発生してしまうという構造的な欠陥があった(LibreOfficeでの検証では
'                黙って0扱いされるため気づけなかったが、本物のExcelでは循環参照として
'                検出される)。このマクロは、PP_Grid内の既存のパススルー数式を1件ずつ
'                走査し、この欠陥のある数式(IF(...,NA(),...)によるガードがまだ無いもの)
'                だけを安全な形に書き換える。どの中間体をパススルーにするか自体は変更しない
'                (既に数式が入っているセルの中身だけを直す)。
'
' 全体の設計方針(パフォーマンス・DataBodyRange・DisplayAlerts等)はRefreshData_Utilities
' モジュール冒頭のコメントを参照してください。
' ============================================================================

Sub RefreshBOM()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel ファイル (*.xlsx),*.xlsx", , _
        "Raw Material - Look Up を選択してください")
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
    Dim ppGrid As ListObject: Set ppGrid = thisWb.Sheets("PP_Grid").ListObjects("PP_Grid")

    ' RM_Code(Part Name)の存在チェック用と、Catalyst短縮コード抽出時の
    ' 「既知の中間体名かどうか」の判定(末尾"s"表記ゆれの解決)用に、それぞれの
    ' 名前の集合を1回だけ作る。
    Dim rmCodeSet As Object: Set rmCodeSet = BuildNameIndex(rmTbl, "Part Name")
    Dim knownIntermediates As Object: Set knownIntermediates = BuildNameIndex(ppGrid, "Intermediate")

    ' (Intermediate|RM_Code) -> M_BOM内の行番号(データ範囲内)
    Dim pairIndex As Object: Set pairIndex = BuildPairIndex(bomTbl)

    Dim updated As Long, added As Long, unresolved As String
    updated = 0: added = 0: unresolved = ""

    Dim flatSheets As Variant: flatSheets = Array("Slurry Data Base", "Powder Data Base", "Solution")
    Dim sIdx As Integer
    For sIdx = LBound(flatSheets) To UBound(flatSheets)
        Dim shName As String: shName = CStr(flatSheets(sIdx))
        If SheetExists(srcWb, shName) Then
            Call ProcessLookupFlatSheet(srcWb.Sheets(shName), bomTbl, pairIndex, rmCodeSet, _
                knownIntermediates, updated, added, unresolved)
        End If
    Next sIdx
    If SheetExists(srcWb, "Catalyst Data Base") Then
        Call ProcessLookupCatalystSheet(srcWb.Sheets("Catalyst Data Base"), bomTbl, pairIndex, _
            rmCodeSet, knownIntermediates, updated, added, unresolved)
    End If

    ' srcWbが既にNothingになっているケース(取込元ファイル側の自動処理等で、開いた
    ' 直後にワークブックが閉じられてしまう場合がある)でも、後始末処理自体が
    ' 「オブジェクト変数が設定されていません」で落ちないようにガードする。
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False

    ' Material_Detailの中間体内訳行(No. of batches／使用量(kg))を、更新後のM_BOMの内容に
    ' 合わせて同期する。M_BOMの更新自体は上で既に確定しているため、この同期処理で
    ' エラーが出てもM_BOM/PP_Grid/Grid_Requirementへの反映は失われない(On Error Resume Next
    ' で個別に捕捉し、失敗してもメッセージで知らせるだけに留める)。
    Dim addedDetailRows As Long: addedDetailRows = 0
    Dim syncErrMsg As String: syncErrMsg = ""
    On Error Resume Next
    Call SyncMaterialDetailIntermediates(bomTbl, addedDetailRows)
    If Err.Number <> 0 Then
        syncErrMsg = Err.Description
        Err.Clear
    End If
    On Error GoTo ErrHandler

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim msg As String
    msg = "M_BOM を更新しました。" & vbCrLf & "更新: " & updated & " 件、新規追加: " & added & " 件"
    If addedDetailRows > 0 Then
        msg = msg & vbCrLf & "Material_Detailに追加した中間体の内訳行: " & addedDetailRows & " 組"
    End If
    If Len(syncErrMsg) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "(注意) Material_Detailの内訳行の自動追加中にエラーが" & vbCrLf & _
              "発生しました: " & syncErrMsg & vbCrLf & "M_BOM自体の更新は正常に完了しています。"
    End If
    If Len(unresolved) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "M_RawMaterials・PP_Gridのどちらにも見つからないコード" & vbCrLf & _
              "(未登録の新しい材料の可能性。AddMaterialでの登録漏れがないか確認してください):" & vbCrLf & unresolved
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

' PP_Gridの既存のパススルー数式(SUMPRODUCT+M_BOM[PPGridRow])を、そのパススルー中間体
' 自身のレシピ行を誤って参照してしまう欠陥から守る形に書き換える(モジュール冒頭のコメント
' 参照)。まだこの対策が入っていない数式(文字列に"NA()"を含まないもの)だけを対象にするため、
' 複数回実行しても安全(2回目以降は対象0件になるだけ)。
Sub FixPassthroughCircularRefs()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim ppGrid As ListObject: Set ppGrid = thisWb.Sheets("PP_Grid").ListObjects("PP_Grid")
    Dim dataRange As Range: Set dataRange = ppGrid.DataBodyRange
    If dataRange Is Nothing Then
        MsgBox "PP_Gridにデータ行がありません。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 数千セル(材料数×週数)を1セルずつ.Cells(r,c).Formulaで読み書きすると、RefreshBOM等
    ' 他のマクロで対策済みのCOM通信の積み重ねと同じ理由で極めて遅くなる(フリーズしたように
    ' 見える)。範囲全体を1回だけ配列として読み込み、書き換えが必要な要素だけを配列上で
    ' 直し、最後に配列ごと1回だけ書き戻す。
    Dim fixedCells As Long: fixedCells = 0
    Dim firstDataRow As Long: firstDataRow = dataRange.Row
    Dim nRows As Long: nRows = dataRange.Rows.Count
    Dim nCols As Long: nCols = dataRange.Columns.Count

    Dim allFormulas As Variant: allFormulas = dataRange.Formula

    Dim r As Long, c As Long
    For r = 1 To nRows
        Dim actualRow As Long: actualRow = firstDataRow + r - 1
        For c = 2 To nCols  ' 1列目はIntermediate名(数式ではない)なのでスキップ
            If VarType(allFormulas(r, c)) = vbString Then
                Dim f As String: f = CStr(allFormulas(r, c))
                If Left$(f, 1) = "=" And InStr(f, "PPGridRow") > 0 And InStr(f, "NA()") = 0 Then
                    allFormulas(r, c) = Replace(f, "M_BOM[PPGridRow]", _
                        "IF(M_BOM[PPGridRow]=" & actualRow & ",NA(),M_BOM[PPGridRow])")
                    fixedCells = fixedCells + 1
                End If
            End If
        Next c
    Next r

    If fixedCells > 0 Then dataRange.Formula = allFormulas

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "パススルー数式の循環参照リスクを修正しました。" & vbCrLf & _
           "修正したセル数: " & fixedCells, vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "修正処理でエラーが発生しました: (" & errNum & ") " & errMsg, vbCritical
End Sub

' Slurry Data Base/Powder Data Base/Solutionシート共通の処理。行そのままA列=Intermediate、
' D列=RM Code、M列=1バッチあたり使用量として取り込む。RM CodeはM_RawMaterialsのPart Name
' そのものの場合も、他の中間体コード(SOL-SCH等)の場合もあるが、区別せずそのままM_BOMに書き込む
' (Grid_Requirement側はM_RawMaterialsに実在するPart Nameだけを見るため、中間体コードの行は
' 実害なく無視される)。
Private Sub ProcessLookupFlatSheet(sh As Worksheet, bomTbl As ListObject, pairIndex As Object, _
        rmCodeSet As Object, knownIntermediates As Object, ByRef updated As Long, ByRef added As Long, _
        ByRef unresolved As String)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If lastRow > 5000 Then lastRow = 5000  ' 異常値対策
    If lastRow < 2 Then Exit Sub

    ' シート全体を1回だけ配列として読み込む（.Cells(r,c).Valueをループ内で毎回呼ぶと遅いため）
    Dim data As Variant
    data = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, 13)).Value  ' M列(13)まで

    Dim r As Long
    For r = 2 To lastRow
        Dim inter As String: inter = Trim(CStr(data(r, 1)))
        Dim rmCode As String: rmCode = Trim(CStr(data(r, 4)))
        Dim v As Variant: v = data(r, 13)
        If Len(inter) > 0 And Len(rmCode) > 0 And IsNumeric(v) Then
            If CDbl(v) <> 0 Then
                Call UpsertBomRow(bomTbl, pairIndex, inter, rmCode, CDbl(v), updated, added)
                If Not rmCodeSet.Exists(rmCode) And Not knownIntermediates.Exists(rmCode) Then
                    If InStr(unresolved, rmCode) = 0 Then unresolved = unresolved & rmCode & "; "
                End If
            End If
        End If
    Next r
End Sub

' Catalyst Data BaseシートのSubstrate行(H列"Substrate SC"が埋まっている行)だけを使う。
' Corning供給のSubstrate(E列Descriptionに"CORNING"を含む)は対象外。中間体名はA列の
' Catalyst名から短縮製品コードを抽出したもの、RM_CodeはH列、数量はF列(catalyst1個あたり)。
Private Sub ProcessLookupCatalystSheet(sh As Worksheet, bomTbl As ListObject, pairIndex As Object, _
        rmCodeSet As Object, knownIntermediates As Object, ByRef updated As Long, ByRef added As Long, _
        ByRef unresolved As String)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If lastRow > 5000 Then lastRow = 5000
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, 8)).Value  ' H列(8)まで

    Dim r As Long
    For r = 2 To lastRow
        Dim subSC As String: subSC = Trim(CStr(data(r, 8)))
        If Len(subSC) = 0 Then GoTo NextRow
        Dim desc As String: desc = UCase(Trim(CStr(data(r, 5))))
        If InStr(desc, "CORNING") > 0 Then GoTo NextRow  ' Corning供給は対象外
        Dim catName As String: catName = Trim(CStr(data(r, 1)))
        Dim v As Variant: v = data(r, 6)
        If Len(catName) = 0 Or Not IsNumeric(v) Then GoTo NextRow
        If CDbl(v) = 0 Then GoTo NextRow
        Dim code As String: code = CatalystShortCode(catName, knownIntermediates)
        Call UpsertBomRow(bomTbl, pairIndex, code, subSC, CDbl(v), updated, added)
        If Not rmCodeSet.Exists(subSC) Then
            If InStr(unresolved, subSC) = 0 Then unresolved = unresolved & subSC & "; "
        End If
NextRow:
    Next r
End Sub

' Catalyst名(例:"18461-0Q110-1st COAT")から短縮製品コード(例:"0Q110")を抽出する。
' 「18461-」プレフィックスを除去し、最初のスペース/ハイフンまでのトークンを取り出し、
' 先頭が"O"(アルファベットのオー)なら"0"(ゼロ)に置換する(元データの表記ゆれ)。
' それでも中間体マスタに見つからず末尾が"s"の場合は、末尾を除去して再試行する
' (例:"0T420s"→"0T420")。
Private Function CatalystShortCode(catName As String, knownIntermediates As Object) As String
    Dim s As String: s = catName
    If Left(s, 6) = "18461-" Then s = Mid(s, 7)
    Dim i As Long, ch As String, code As String
    code = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = " " Or ch = "-" Then Exit For
        code = code & ch
    Next i
    If Left(code, 1) = "O" Then code = "0" & Mid(code, 2)
    If Not knownIntermediates.Exists(code) Then
        Dim tailCh As String: tailCh = Right(code, 1)
        If (tailCh = "s" Or tailCh = "S") And knownIntermediates.Exists(Left(code, Len(code) - 1)) Then
            code = Left(code, Len(code) - 1)
        End If
    End If
    CatalystShortCode = code
End Function

' M_BOMの(Intermediate, RM_Code)行を1件登録/更新する共通処理。
Private Sub UpsertBomRow(bomTbl As ListObject, pairIndex As Object, inter As String, rmCode As String, _
        qty As Double, ByRef updated As Long, ByRef added As Long)
    Dim pk As String: pk = inter & "|" & rmCode
    If pairIndex.Exists(pk) Then
        Dim rowN As Long: rowN = pairIndex(pk)
        bomTbl.ListRows(rowN).Range.Cells(1, 3).Value = qty
        updated = updated + 1
    Else
        Dim newRow As ListRow
        Set newRow = bomTbl.ListRows.Add
        newRow.Range.Cells(1, 1).Value = inter
        newRow.Range.Cells(1, 2).Value = rmCode
        newRow.Range.Cells(1, 3).Value = qty
        pairIndex(pk) = newRow.Index
        added = added + 1
    End If
End Sub

Private Function SheetExists(wb As Workbook, sName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets(sName)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function

' RefreshBOM実行後に呼ばれる。更新後のM_BOMの内容(材料コード×中間体の組み合わせ)を見て、
' Material_Detailの各材料ブロックに、まだ内訳行(No. of batches／使用量(kg))が無い中間体が
' あれば、その2行を「合計使用量(kg)/週」行の直前にペアで挿入する。
Private Sub SyncMaterialDetailIntermediates(bomTbl As ListObject, ByRef addedPairs As Long)
    addedPairs = 0
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    If Not SheetExists(thisWb, "Material_Detail") Then Exit Sub
    If Not SheetExists(thisWb, "Cal_Weeks") Then Exit Sub
    Dim sh As Worksheet: Set sh = thisWb.Sheets("Material_Detail")
    Dim nWeeks As Long: nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count
    If nWeeks <= 0 Then Exit Sub

    ' M_BOMを「材料コード -> その材料を使う中間体名の一覧(登場順、重複除去)」の
    ' Dictionaryにまとめる(1回だけ配列読み込み。既存のBuildPairIndex等と同じ設計方針)。
    Dim byMat As Object: Set byMat = CreateObject("Scripting.Dictionary")
    byMat.CompareMode = vbTextCompare
    Dim bomN As Long: bomN = bomTbl.ListRows.Count
    If bomN = 0 Then Exit Sub
    Dim bomData As Variant
    bomData = bomTbl.ListColumns(1).DataBodyRange.Resize(bomN, 2).Value  ' 1=Intermediate, 2=Part Name
    Dim bi As Long
    For bi = 1 To bomN
        Dim interN As String: interN = Trim(CStr(bomData(bi, 1)))
        Dim partN As String: partN = Trim(CStr(bomData(bi, 2)))
        If Len(interN) > 0 And Len(partN) > 0 Then
            If Not byMat.Exists(partN) Then
                Dim seenDict As Object: Set seenDict = CreateObject("Scripting.Dictionary")
                seenDict.CompareMode = vbTextCompare
                byMat.Add partN, seenDict
            End If
            If Not byMat(partN).Exists(interN) Then byMat(partN).Add interN, True
        End If
    Next bi
    If byMat.Count = 0 Then Exit Sub

    ' 書式コピー用のテンプレート行(シート内で最初に見つかる「No. of batches」行のペア)。
    ' Rangeオブジェクトとして保持するため、以降このシート上のどこで行挿入が起きても
    ' Excel側が自動的に参照先の行位置を追従してくれる(挿入行の直下に別の行を挿入しても
    ' ズレない、標準的なVBAオブジェクト参照の挙動)。
    Dim lastRowScan As Long: lastRowScan = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row
    Dim templateRow As Long: templateRow = 0
    Dim tr As Long
    For tr = MD_HEADER_ROW + 1 To lastRowScan
        If Trim(CStr(sh.Cells(tr, 3).Value)) = "No. of batches" Then
            templateRow = tr
            Exit For
        End If
    Next tr
    Dim templateRows As Range
    If templateRow > 0 Then Set templateRows = sh.Rows(templateRow & ":" & (templateRow + 1))

    Dim r As Long: r = MD_HEADER_ROW + 1
    Do While r <= lastRowScan
        Dim rmCode As String: rmCode = Trim(CStr(sh.Cells(r, 1).Value))
        If Len(rmCode) = 0 Then
            r = r + 1
        Else
            Dim headerRow As Long: headerRow = r
            Dim existing As Object: Set existing = CreateObject("Scripting.Dictionary")
            existing.CompareMode = vbTextCompare
            Dim rr As Long: rr = headerRow + 1
            Dim sumRow As Long: sumRow = 0
            Do While rr <= lastRowScan
                Dim lbl As String: lbl = Trim(CStr(sh.Cells(rr, 2).Value))
                If lbl = "合計使用量(kg)/週" Then
                    sumRow = rr
                    Exit Do
                End If
                If Trim(CStr(sh.Cells(rr, 3).Value)) = "No. of batches" Then existing(lbl) = True
                rr = rr + 1
            Loop
            If sumRow > 0 And byMat.Exists(rmCode) Then
                Dim allInter As Object: Set allInter = byMat(rmCode)
                Dim k As Variant, insertAt As Long: insertAt = sumRow
                For Each k In allInter.Keys
                    If Not existing.Exists(CStr(k)) Then
                        Call InsertIntermediateRowPair(sh, insertAt, headerRow, CStr(k), nWeeks, templateRows)
                        addedPairs = addedPairs + 1
                        insertAt = insertAt + 2
                        lastRowScan = lastRowScan + 2
                    End If
                Next k
                r = insertAt  ' insertAtは元々sumRowだった行(合計使用量行)の新しい位置。
                               ' A列が空なので、外側ループの空白スキップでそのまま次ブロックまで進む
            Else
                If sumRow > 0 Then r = sumRow + 1 Else r = headerRow + 1
            End If
        End If
    Loop
End Sub

' Material_Detailの指定位置(insertAtRow)に、中間体1件分の内訳行2行
' (No. of batches／使用量(kg))を挿入する。書式はtemplateRows(既存の内訳行ペア)から複製する。
Private Sub InsertIntermediateRowPair(sh As Worksheet, insertAtRow As Long, headerRow As Long, _
        interName As String, nWeeks As Long, templateRows As Range)
    sh.Rows(insertAtRow & ":" & (insertAtRow + 1)).Insert Shift:=xlDown
    Dim batchesRow As Long: batchesRow = insertAtRow
    Dim usageRow As Long: usageRow = insertAtRow + 1
    Dim helperCol As Long: helperCol = MD_WEEK_START_COL + nWeeks
    Dim helperColLetter As String: helperColLetter = ColLetter(helperCol)

    sh.Cells(batchesRow, 2).Value = interName
    sh.Cells(batchesRow, 3).Value = "No. of batches"
    sh.Cells(batchesRow, helperCol).Value = "=IFERROR(MATCH($B" & batchesRow & ",PP_Grid[Intermediate],0),99999)"
    sh.Cells(batchesRow, helperCol).Font.Size = 8
    sh.Cells(batchesRow, helperCol).Font.Color = RGB(128, 128, 128)

    sh.Cells(usageRow, 2).Value = "使用量(kg)"
    sh.Cells(usageRow, 3).Value = "=SUMIFS(M_BOM[RM_Qty_Per_Batch],M_BOM[Intermediate],$B" & batchesRow & _
        ",M_BOM[Part Name],$A" & headerRow & ")"

    Dim w As Long, col As Long, wc As String
    For w = 1 To nWeeks
        col = MD_WEEK_START_COL + w - 1
        wc = ColLetter(col)
        sh.Cells(batchesRow, col).Value = _
            "=IFERROR(INDEX(PP_Grid[#Data],$" & helperColLetter & batchesRow & "," & (w + 1) & "),0)"
        sh.Cells(usageRow, col).Value = "=$C" & usageRow & "*" & wc & batchesRow
    Next w

    If Not templateRows Is Nothing Then
        On Error Resume Next
        templateRows.Copy
        sh.Rows(batchesRow & ":" & usageRow).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        On Error GoTo 0
    End If
End Sub
