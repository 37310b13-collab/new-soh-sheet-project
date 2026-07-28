Attribute VB_Name = "RefreshData_BOM"
Option Explicit

' ============================================================================
' RefreshData_BOM モジュール
'
'   RefreshBOM : 「Usage from Production Engineering」を選択すると、
'                M_BOM（化学原料の原単位）を更新する。新しい材料×中間体の組み合わせが
'                見つかった場合、Material_Detailの該当材料ブロックにも中間体の内訳行
'                （No. of batches／使用量(kg)）を自動的に追加する
'                （SyncMaterialDetailIntermediates）。これにより、AddMaterialで追加した
'                ばかりの材料(BOM未登録のため内訳行が無いミニブロック)も、RefreshBOMを
'                1回実行するだけで既存の材料と同じ見た目になる(AddMaterialを2回実行する
'                必要はない)。既存材料が新しい中間体で使われ始めた場合も同様に自動反映される。
'
' 全体の設計方針(パフォーマンス・DataBodyRange・DisplayAlerts等)はRefreshData_Utilities
' モジュール冒頭のコメントを参照してください。
' ============================================================================

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
