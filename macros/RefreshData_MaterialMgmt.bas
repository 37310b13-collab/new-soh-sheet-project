Attribute VB_Name = "RefreshData_MaterialMgmt"
Option Explicit

' ============================================================================
' RefreshData_MaterialMgmt モジュール
'
' AddMaterial / RemoveMaterial / RemoveIntermediate
'
' 材料(原材料)・中間体(生産される製品コード)をPythonを使わず、Excel(VBA)だけで
' 追加・削除するためのマクロです。
'
'   AddMaterial : 新しい材料(TTAF供給品)をシステムに追加する。InputBoxで
'                 Part Name(RM_Code)・Description・Supplier・Category・TTAF_Codeを
'                 順に入力すると、M_RawMaterials・Grid_Requirement・Grid_Incoming・
'                 Grid_Stock・Grid_TheoreticalStock・T_OpeningStock・T_SelfStock・
'                 T_TTAFStock・Dashboard・Material_Detail・対応するPO_Draft_*シートの
'                 一番下に、必要な行をまとめて追加する。追加直後はまだM_BOMに使用実績が
'                 無いため、Material_Detailのブロックは中間体の内訳が無いミニブロック
'                 (合計欄のみ)になる。RefreshBOM実行後(RefreshData_BOMモジュール)、
'                 実際にこの材料を使う中間体が見つかれば内訳行も自動的に追加される。
'                 Dashboardには材料ごとに「理論在庫」「実在庫」の2行を追加する
'                 （5.6章参照）。
'   RemoveMaterial : 使わなくなった材料をシステムから削除する。InputBoxでPart Name
'                 (RM_Code)を入力すると、AddMaterialが追加する全シートから該当行を
'                 削除する。T_Shipments・T_PlannedOrders・T_StockCount・
'                 T_SelfStock_Log/T_TTAFStock_Log・M_BOMのデータは削除しない
'                 (履歴として残すため。再度AddMaterialで同じPart Nameを追加すれば
'                 自動的に再びつながる)。
'   RemoveIntermediate : 生産中止になった中間体(完成品コード)をシステムから削除する。
'                 InputBoxで中間体名を入力すると、PP_Grid・M_BOMの該当行と、
'                 Material_Detailの該当内訳行(No. of batches／使用量(kg)。この中間体を
'                 使っているすべての材料ブロックから)を削除する。原材料側のデータ
'                 (T_Shipments・T_PlannedOrders・T_OpeningStock・T_StockCount・
'                 実績ログ・M_RawMaterials)は削除しない。
'                 (参考) 中間体の"追加"側は既に自動化済みです。RefreshWeeklyBatches/
'                 RefreshBOMが、生産計画・原単位表に新しい中間体を見つけるたびに
'                 PP_Grid・M_BOM・Material_Detailへ自動的に行を追加するため、
'                 専用の"AddIntermediate"マクロは不要です。
'
' 【追加(AddMaterial)の位置】既存行の途中に割り込ませるのではなく、必ず各シートの
' 一番下に追加します。途中に割り込ませると既存行がずれるリスクが大きいためです。
'
' 【安全性についての補足(RemoveIntermediate)】PP_Grid・M_BOMの行を削除しても、他の
' 中間体・他の材料の計算には影響しません。理由は、Grid_Requirement・Material_Detailから
' PP_Grid/M_BOMを参照する数式がすべてMATCH/構造化参照(テーブル名[列名]形式)で組まれており、
' 固定の行番号を直接使っていないためです(MATCHは削除後の新しい行位置を毎回再計算し、
' 構造化参照は削除後のテーブルの行数に自動的に追従します)。
'
' 【重要な注意点】どの操作も元に戻せません(Ctrl+Zでは戻せない場合があります)。実行前に
' ファイルのバックアップ(コピー)を取っておくことを強くおすすめします。
'
' 全体の設計方針はRefreshData_Utilitiesモジュール冒頭のコメントを参照してください。
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

    ' ---- Grid_Requirement / Grid_Incoming / Grid_Stock / Grid_TheoreticalStock / T_OpeningStock ----
    Dim reqTbl As ListObject: Set reqTbl = thisWb.Sheets("Grid_Requirement").ListObjects("Grid_Requirement")
    Dim inTbl As ListObject: Set inTbl = thisWb.Sheets("Grid_Incoming").ListObjects("Grid_Incoming")
    Dim stTbl As ListObject: Set stTbl = thisWb.Sheets("Grid_Stock").ListObjects("Grid_Stock")
    Dim theoTbl As ListObject: Set theoTbl = thisWb.Sheets("Grid_TheoreticalStock").ListObjects("Grid_TheoreticalStock")
    Dim osTbl As ListObject: Set osTbl = thisWb.Sheets("T_OpeningStock").ListObjects("T_OpeningStock")

    Dim reqRow As ListRow: Set reqRow = reqTbl.ListRows.Add
    Dim inRow As ListRow: Set inRow = inTbl.ListRows.Add
    Dim stRow As ListRow: Set stRow = stTbl.ListRows.Add
    Dim theoRow As ListRow: Set theoRow = theoTbl.ListRows.Add
    Dim osRow As ListRow: Set osRow = osTbl.ListRows.Add

    Dim grow As Long: grow = reqRow.Range.Row  ' Grid_Requirement/Incoming/Stock/TheoreticalStockの実シート行番号(4表とも同じ)

    reqRow.Range.Cells(1, 1).Value = rmCode
    inRow.Range.Cells(1, 1).Value = rmCode
    stRow.Range.Cells(1, 1).Value = rmCode
    theoRow.Range.Cells(1, 1).Value = rmCode
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

        ' Grid_TheoreticalStock: T_StockCount・自社/TTAF実績を一切見ない、純粋な「前週+入庫-消費」
        ' のロールフォワードのみ(Grid_Stockの優先順位チェーンとは無関係)。Dashboardの「理論在庫」行
        ' として、実際の値(Grid_Stock=「実在庫」行)との乖離を確認するために参照する。
        Dim theoPriorExpr As String
        If w = 1 Then
            theoPriorExpr = "IFERROR(INDEX(T_OpeningStock[Opening_Qty],MATCH($A" & grow & ",T_OpeningStock[Part Name],0)),0)"
        Else
            theoPriorExpr = ColLetter(col - 1) & grow
        End If
        theoRow.Range.Cells(1, col).Value = _
            "=" & theoPriorExpr & "+'Grid_Incoming'!" & cl & grow & "-'Grid_Requirement'!" & cl & grow
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
    ' 実際に登録されている表記(大文字小文字)に揃える。Trim()を通し忘れると、
    ' M_RawMaterials側のセルに紛れ込んだ余分な空白がそのままrmCodeに乗り移り、
    ' 以降のDeleteMatchingGridRow等(比較先のセルはTrimしているが、rmCode側は
    ' 呼び出し元で整形済みという前提)で一致しなくなり、一部のシートだけ削除が
    ' 失敗する不具合につながる(実際に報告された不具合)。
    rmCode = Trim(CStr(rmTbl.ListRows(rmFoundRow).Range.Cells(1, 1).Value))
    Dim categoryVal As String: categoryVal = CStr(rmTbl.ListRows(rmFoundRow).Range.Cells(1, 4).Value)

    If MsgBox("材料「" & rmCode & "」を削除します。" & vbCrLf & _
              "関係する全シート(M_RawMaterials・Grid_Requirement・Grid_Incoming・Grid_Stock・" & vbCrLf & _
              "Grid_TheoreticalStock・" & vbCrLf & _
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
    Call DeleteMatchingTableRow(thisWb.Sheets("Grid_TheoreticalStock").ListObjects("Grid_TheoreticalStock"), rmCode)
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

Sub RemoveIntermediate()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim ppGrid As ListObject: Set ppGrid = thisWb.Sheets("PP_Grid").ListObjects("PP_Grid")
    Dim bomTbl As ListObject: Set bomTbl = thisWb.Sheets("M_BOM").ListObjects("M_BOM")

    Dim interName As String
    interName = Trim(InputBox("削除する中間体名(Intermediate)を入力してください。" & vbCrLf & _
        "PP_Gridシートの「Intermediate」列に表示されている名称と一致させてください" & vbCrLf & _
        "(大文字小文字は区別しません)。", "中間体の削除"))
    If Len(interName) = 0 Then Exit Sub

    Dim ppFoundRow As Long: ppFoundRow = FindMaterialRow(ppGrid, interName)
    If ppFoundRow = 0 Then
        MsgBox "PP_Gridに「" & interName & "」という中間体が見つかりませんでした。" & vbCrLf & _
               "PP_Gridシートで正確な名称を確認してください。", vbExclamation
        Exit Sub
    End If
    ' 実際に登録されている表記(大文字小文字)に揃える(入力時のゆらぎを吸収するため)。
    ' RemoveMaterialと同じ理由でTrim()を通す(セル側に紛れ込んだ余分な空白が
    ' 以降の削除処理での不一致につながるのを防ぐ)。
    Dim canonicalName As String
    canonicalName = Trim(CStr(ppGrid.ListRows(ppFoundRow).Range.Cells(1, 1).Value))

    Dim bomCount As Long: bomCount = CountMatchingTableRows(bomTbl, canonicalName)

    If MsgBox("中間体「" & canonicalName & "」を削除します。" & vbCrLf & vbCrLf & _
              "削除される内容:" & vbCrLf & _
              "・PP_Grid: 該当行(週次バッチ数)を1行削除" & vbCrLf & _
              "・M_BOM: この中間体を使う原単位の行を " & bomCount & " 件削除" & vbCrLf & _
              "・Material_Detail: この中間体の内訳行(No. of batches／使用量(kg))を、" & vbCrLf & _
              "  この中間体を使っているすべての材料ブロックから削除" & vbCrLf & vbCrLf & _
              "T_Shipments・T_PlannedOrders・T_OpeningStock・T_StockCount・実績ログ・" & vbCrLf & _
              "M_RawMaterialsなど、原材料側のデータは一切削除されません。" & vbCrLf & vbCrLf & _
              "この操作は元に戻せません。よろしいですか？", _
              vbYesNo + vbExclamation, "中間体の削除の確認") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim removedDetailPairs As Long
    removedDetailPairs = DeleteIntermediateFromMaterialDetail(thisWb.Sheets("Material_Detail"), canonicalName)
    Call DeleteMatchingTableRow(bomTbl, canonicalName)
    ' PP_Grid自体は、上のFindMaterialRowの検索キーとして使い終わってから最後に削除する
    Call DeleteMatchingTableRow(ppGrid, canonicalName)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "中間体「" & canonicalName & "」を削除しました。" & vbCrLf & _
           "Material_Detailから削除した内訳行: " & removedDetailPairs & " 組（材料ブロック数）" & vbCrLf & vbCrLf & _
           "（T_Shipments・T_PlannedOrders・T_OpeningStock・T_StockCount・実績ログ・" & vbCrLf & _
           "M_RawMaterialsは削除されていません）", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errMsg As String: errMsg = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "中間体の削除中にエラーが発生しました: (" & errNum & ") " & errMsg & vbCrLf & vbCrLf & _
           "途中まで削除されている可能性があります。シートの状態を確認してください" & vbCrLf & _
           "(心配な場合は、保存せずにファイルを閉じて開き直せば、直前の保存状態に戻せます)。", vbCritical
End Sub

Private Function IsMaterialCodeFree(rmTbl As ListObject, rmCode As String) As Boolean
    IsMaterialCodeFree = True
    Dim n As Long: n = rmTbl.ListRows.Count
    Dim i As Long
    For i = 1 To n
        If UCase(Trim(CStr(rmTbl.ListRows(i).Range.Cells(1, 1).Value))) = UCase(Trim(rmCode)) Then
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
        If UCase(Trim(CStr(rmTbl.ListRows(i).Range.Cells(1, 1).Value))) = UCase(Trim(rmCode)) Then
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
' Dashboardは材料ごとに2行(理論在庫→実在庫の順)。Part Name(A列)は両方の行に同じ値を書くため、
' RemoveMaterialのDeleteMatchingGridRow(A列一致で削除)が2行ともまとめて削除してくれる
' (削除側の特別対応は不要)。列位置はbuild_soh.pyのLEFT_COLS/DASH_ROW_LABEL_COL/DASH_DIFF_COLと
' 対応: 1=Part Name,2=Description,3=Category,4=基準在庫下限,5=基準在庫上限,6=自社在庫(実績),
' 7=TTAF在庫(実績),8=実績週,9=行種別,10=乖離(kg)、週データは11列目(K列)から。
Private Sub AppendDashboardRow(sh As Worksheet, rmCode As String, nWeeks As Long, ssRow As Long, grow As Long)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim theoRow As Long: theoRow = lastRow + 1
    Dim actualRow As Long: actualRow = lastRow + 2
    Dim lastWeekCol As String: lastWeekCol = ColLetter(1 + nWeeks)

    Dim ssSelfRng As String: ssSelfRng = "'T_SelfStock'!$B$" & ssRow & ":$" & lastWeekCol & "$" & ssRow
    Dim ssTTAFRng As String: ssTTAFRng = "'T_TTAFStock'!$B$" & ssRow & ":$" & lastWeekCol & "$" & ssRow
    Dim ssLabelRng As String: ssLabelRng = "'T_SelfStock'!$B$" & SS_TABLE_ROW & ":$" & lastWeekCol & "$" & SS_TABLE_ROW
    ' 乖離(kg) = 実在庫(Grid_Stock) - 理論在庫(Grid_TheoreticalStock)。表示期間の最終週どうしを
    ' 比較するだけでよい(補正が入った週以降、差分はその後ずっと一定のまま変わらないため)。
    Dim gsLastCol As String: gsLastCol = ColLetter(1 + nWeeks)
    Dim diffFormula As String
    diffFormula = "='Grid_Stock'!" & gsLastCol & grow & "-'Grid_TheoreticalStock'!" & gsLastCol & grow

    Dim rowsArr(1 To 2) As Long
    rowsArr(1) = theoRow
    rowsArr(2) = actualRow
    Dim idx As Long, rr As Long
    For idx = 1 To 2
        rr = rowsArr(idx)
        sh.Cells(rr, 1).Value = rmCode
        sh.Cells(rr, 2).Value = "=IFERROR(INDEX(M_RawMaterials[Description],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),"""")"
        sh.Cells(rr, 3).Value = "=IFERROR(INDEX(M_RawMaterials[Category],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),"""")"
        sh.Cells(rr, 4).Value = "=IFERROR(INDEX(M_RawMaterials[基準在庫下限_要入力],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),0)"
        sh.Cells(rr, 5).Value = "=IFERROR(INDEX(M_RawMaterials[基準在庫上限_要入力],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),0)"
        sh.Cells(rr, 6).Value = "=IFERROR(LOOKUP(2,1/(" & ssSelfRng & "<>"""")," & ssSelfRng & "),"""")"
        sh.Cells(rr, 7).Value = "=IFERROR(LOOKUP(2,1/(" & ssTTAFRng & "<>"""")," & ssTTAFRng & "),"""")"
        sh.Cells(rr, 8).Value = "=IFERROR(LOOKUP(2,1/(" & ssSelfRng & "<>"""")," & ssLabelRng & "),"""")"
        sh.Cells(rr, 10).Value = diffFormula
    Next idx

    sh.Cells(theoRow, 9).Value = "理論在庫"
    sh.Cells(theoRow, 9).Font.Italic = True
    sh.Cells(theoRow, 9).Font.Color = RGB(128, 128, 128)
    sh.Cells(actualRow, 9).Value = "実在庫"
    sh.Cells(actualRow, 9).Font.Bold = True

    Dim w As Long, col As Long
    For w = 1 To nWeeks
        col = 10 + w  ' Dashboardの週データ開始列=11(K列)
        sh.Cells(theoRow, col).Value = "='Grid_TheoreticalStock'!" & ColLetter(1 + w) & grow
        sh.Cells(theoRow, col).Font.Italic = True
        sh.Cells(theoRow, col).Font.Color = RGB(128, 128, 128)
        sh.Cells(actualRow, col).Value = "='Grid_Stock'!" & ColLetter(1 + w) & grow
    Next w

    ' 書式コピー: 既存の最後の材料ペア(直前の理論在庫行・実在庫行)からそれぞれ複製する
    On Error Resume Next
    sh.Rows(lastRow - 1).Copy   ' 直前ペアの理論在庫行
    sh.Rows(theoRow).PasteSpecial xlPasteFormats
    sh.Rows(lastRow).Copy       ' 直前ペアの実在庫行
    sh.Rows(actualRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

' Material_Detailの一番下に新しい材料のブロックを追加する。追加直後はM_BOMに未登録のため、
' 中間体の行(No. of batches/使用量)は無く、「合計使用量(0のはず)・TTAF在庫・自社在庫・
' Order・合計在庫・注記」の6行だけのミニブロックになる。RefreshBOM実行後、この材料が
' 実際に使われている中間体が見つかれば、SyncMaterialDetailIntermediates(RefreshData_BOM
' モジュール、RefreshBOMの末尾で自動的に呼ばれる)がこのブロックに中間体の内訳行を
' 自動的に追加する（AddMaterialを再度実行する必要はない）。
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
    Dim key As String: key = UCase(Trim(rmCode))
    For i = n To 1 Step -1
        If UCase(Trim(CStr(tbl.ListRows(i).Range.Cells(1, 1).Value))) = key Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

' テーブル機能を使わない罫線グリッド(T_SelfStock/T_TTAFStock/Dashboard/PO_Draft_*)から、
' 指定した材料コードに一致する行を削除する。nameCol=材料コードが入っている列番号。
Private Sub DeleteMatchingGridRow(sh As Worksheet, rmCode As String, nameCol As Long)
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, nameCol).End(xlUp).Row
    Dim r As Long
    Dim key As String: key = UCase(Trim(rmCode))
    For r = lastRow To 1 Step -1
        If UCase(Trim(CStr(sh.Cells(r, nameCol).Value))) = key Then
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

' テーブルの列1(RemoveIntermediateではPP_Grid/M_BOMの「Intermediate」列)に一致する行数を数える
' (実際の削除前に、確認ダイアログでユーザーに件数を提示するためだけに使う)。
Private Function CountMatchingTableRows(tbl As ListObject, keyVal As String) As Long
    Dim n As Long: n = tbl.ListRows.Count
    Dim i As Long, cnt As Long: cnt = 0
    Dim key As String: key = UCase(Trim(keyVal))
    For i = 1 To n
        If UCase(Trim(CStr(tbl.ListRows(i).Range.Cells(1, 1).Value))) = key Then cnt = cnt + 1
    Next i
    CountMatchingTableRows = cnt
End Function

' Material_Detailの全ブロックを走査し、指定した中間体名の内訳行ペア(No. of batches／使用量(kg))を
' 見つけ次第削除する。1つの中間体が複数の材料ブロックで使われている場合、該当するすべての
' ブロックから削除する(削除した組数を返す)。RefreshData_BOMのSyncMaterialDetailIntermediates
' (追加側)と対になる削除側の処理で、同じブロック走査の考え方(ライブにセルを読みながら、
' 削除で行がずれてもその場で辻褄が合うようにする)を使っている。
Private Function DeleteIntermediateFromMaterialDetail(sh As Worksheet, interName As String) As Long
    Dim removedPairs As Long: removedPairs = 0
    Dim lastRowScan As Long: lastRowScan = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row

    Dim r As Long: r = MD_HEADER_ROW + 1
    Do While r <= lastRowScan
        Dim rmCode As String: rmCode = Trim(CStr(sh.Cells(r, 1).Value))
        If Len(rmCode) = 0 Then
            r = r + 1
        Else
            Dim headerRow As Long: headerRow = r
            Dim rr As Long: rr = headerRow + 1
            Dim sumRow As Long: sumRow = 0
            Do While rr <= lastRowScan
                Dim lbl As String: lbl = Trim(CStr(sh.Cells(rr, 2).Value))
                If lbl = "合計使用量(kg)/週" Then
                    sumRow = rr
                    Exit Do
                End If
                If Trim(CStr(sh.Cells(rr, 3).Value)) = "No. of batches" And StrComp(lbl, interName, vbTextCompare) = 0 Then
                    sh.Rows(rr & ":" & (rr + 1)).Delete
                    removedPairs = removedPairs + 1
                    lastRowScan = lastRowScan - 2
                    ' rrはそのまま(削除により、次の行がこの位置に繰り上がってくるため)
                Else
                    rr = rr + 1
                End If
            Loop
            If sumRow > 0 Then r = sumRow + 1 Else r = headerRow + 1
        End If
    Loop
    DeleteIntermediateFromMaterialDetail = removedPairs
End Function
