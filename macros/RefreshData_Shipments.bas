Attribute VB_Name = "RefreshData_Shipments"
Option Explicit

' ============================================================================
' RefreshData_Shipments モジュール
'
'   RefreshShipments : 「CSA Report」を選択すると、その中の「Shipping Schedule」シートを
'                      丸ごと取り込み、T_Shipments（発注〜着荷の実績・予定）を更新する。
'                      PO No単位ではなく全件まとめて取り込む(発注から着荷まで4〜6ヶ月かかり、
'                      常に複数件のPOが並行して進むため)。
'
' 列: D=CSA Product Code(材料コード)、G=CSA Order firm month(発注月)、I=CSA PO No.、
' K=Vessel、L=Container、N=Confirmed Order Qty、O=Original ETD、
' Q=2 week transit to TTAF(Latest ETA[P列]+14日。実際の入荷予測に使うのはP列そのものでは
' なくこちら)、S=Received At TTAF、T=Status。
' CSA Product CodeはTTAF PART NUMBERと異なり、既にこちらのRM_Codeとほぼ同じ表記のため、
' TTAF_Code/Description経由の照合(RefreshData_StockActuals側のResolveTTAFPart)ではなく、
' RM_Code同士を直接(大文字小文字・前後空白を無視して)照合する。
'
' 【重要】T_Shipmentsの一意キーについて。以前は「材料名＋PO番号」だけで行を一意に
' 管理していたが、実際のCSA Reportでは同じ材料・同じPO番号が複数回に分けて届く分割出荷が
' 頻繁にある(1つのPOで5行に分かれているケース等も普通にある)。材料名＋PO番号だけの
' キーだとこれらが同じ行として扱われ、後から読んだ行が前の行を上書きし、実際には届いて
' いるはずの数量が静かに失われる不具合があった(このプロジェクトの実データで55組も
' 該当箇所が見つかった)。そのため、コンテナ番号(L列)・Original ETD(O列)まで含めた
' 複合キーで行を区別する。それでも完全に同じ組み合わせ(同じコンテナに複数バッチが
' 混載されている等、ごく稀なケース)が複数行ある場合だけ、ファイル内の出現順の連番で
' 最終的に区別する(DateKeyStr/BuildShipmentRowIndex参照)。
' 既存の運用中ブックはT_Shipmentsがまだ9列(Vessel/Container/Original_ETD列が無い)の
' ため、このモジュールを貼り替えた後、一度だけ AddShipmentSplitColumns を実行して列を
' 追加してから RefreshShipments を実行し直すこと。これにより、以前は材料名+PO番号の
' 重複で上書きされて消えていた分割出荷の行が、複合キーで正しく区別されて自動的に
' 追加され直す(過去に失われた数量が復元される)。
'
' T_ShipmentsはGrid_Incoming(材料×週のSUMIFS)から大量に参照される重量級テーブルのため、
' RefreshBOM(M_BOM)・RefreshWeeklyBatches(PP_Grid)と同じ理由で、新規行はまとめて件数を
' 数えてから1回のResizeで追加し、既存行の更新も行単位でまとめて読み書きする(1件ずつ
' ListRows.Addを呼ぶとExcelが応答なしになる恐れがあるため)。
'
' 【発注管理(Material_Detail連携)について】T_Shipmentsを取り込んだ後、
' SyncMaterialDetailOrders を呼び、Material_DetailのOrder行(発注予定,kg)・PO_No行
' (Order行の直下)を、CSA ReportのStatus列に合わせて自動更新する。
'   ・Status="Unconfirmed"でETAが未定(TBC): Order_Month + M_RawMaterials[LeadTime_Weeks_要入力]
'     から仮の週を計算し、その週にOrder/PO_Noセルを移動する(あくまで仮の予測)。
'   ・Status="Unconfirmed"/"In-transit"でETAが判明: T_Shipments[Effective_Week](Q列の
'     日付が反映済み)の週に移動する。ETAが更新されるたびに追従する。
'   ・Status="TTAF Stock": 最後に分かっている週に固定し、PO_Noセルに"[済]"を付けて
'     Grid_Incomingの計算対象から除外する(数字自体は履歴として残す)。
'   ・同じPO番号で出荷が複数行に分かれている場合(分割出荷)、Order/PO_Noセルもその週数分に
'     自動的に分割する。
'   ・セルを動かした/分割した/確定させた場合は、変更内容をセルコメントに残す。
' Material_Detailにブロックが無い材料(BOMで使われない梱包資材等)や、PO_Noが
' Material_Detailのどこにも入力されていない出荷は対象外(Grid_Incoming側でT_Shipmentsを
' 直接見るフォールバックが効く)。
'
' 全体の設計方針(パフォーマンス・DataBodyRange・DisplayAlerts等)はRefreshData_Utilities
' モジュール冒頭のコメントを参照してください。
' ============================================================================

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

    ' 同じCSA Report内のピボットが更新し忘れられたまま送られてくる可能性があるため
    ' (RefreshTTAFStockと同じ理由)、データを読む前に必ず更新しておく。
    srcWb.RefreshAll
    DoEvents

    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim shipTbl As ListObject: Set shipTbl = thisWb.Sheets("T_Shipments").ListObjects("T_Shipments")
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")

    ' 複合キー対応(Vessel/Container/Original_ETD列)の移行がまだの場合、このまま進めると
    ' 列位置がずれて実行時エラーで異常終了する(不可解なエラーになるのを避けるため、
    ' ここで分かりやすいメッセージを出して安全に中断する)。
    If shipTbl.ListColumns.Count < 12 Then
        srcWb.Close SaveChanges:=False
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        MsgBox "T_Shipmentsがまだ新しい列構成(Vessel/Container/Original_ETD)に移行されていません。" & vbCrLf & _
               "先に「AddShipmentSplitColumns」マクロを一度だけ実行してから、" & vbCrLf & _
               "あらためてRefreshShipmentsを実行してください。", vbExclamation
        Exit Sub
    End If

    Dim sh As Worksheet: Set sh = srcWb.Sheets("Shipping Schedule")

    Dim rmCodeIdx As Object: Set rmCodeIdx = CreateObject("Scripting.Dictionary")
    Call BuildRMCodeIndex(rmTbl, rmCodeIdx)
    Dim knownAliasIdx As Object: Set knownAliasIdx = BuildKnownAliasIndex()

    Dim shipIdx As Object: Set shipIdx = BuildShipmentRowIndex(shipTbl)

    ' シートを1セルずつ読むと遅くなるため、余裕を持った範囲を1回だけ配列で読み込んでから走査する。
    Const MAX_ROWS As Long = 3000
    Dim data As Variant
    data = sh.Range(sh.Cells(2, 1), sh.Cells(MAX_ROWS, 20)).Value

    ' 【重要】T_ShipmentsはGrid_Incoming(材料×週のSUMIFS、90材料×104週分)から大量に参照される
    ' テーブル。以前M_BOM(RefreshBOM)・PP_Grid(RefreshWeeklyBatches)で、新規行をListRows.Addで
    ' 1件ずつ追加してExcelが応答なしになった不具合と同じ構造(重量級テーブル×大量の1件ずつAdd)
    ' のため、ここでも新規行はまず件数を数えてから1回のResizeでまとめて追加する。既存行の
    ' 更新も、行ごとに複数セルを個別に書き込むのではなく行単位でまとめて読み書きする。
    Dim r As Long, added As Long, updated As Long, unresolved As String
    added = 0: updated = 0: unresolved = ""
    Dim updateVals As Object: Set updateVals = CreateObject("Scripting.Dictionary")  ' 行番号 -> Array(qty,eta,receivedDate,status,orderMonth,vessel,container,origEtd)
    Dim newRecords As Object: Set newRecords = CreateObject("Scripting.Dictionary")  ' 複合キー -> Array(partName,poNo,qty,eta,receivedDate,status,orderMonth,vessel,container,origEtd)
    ' 複合キーの元(材料+PO番号+コンテナ+OriginalETD)ごとの、このファイル内での出現回数
    Dim seqCounter As Object: Set seqCounter = CreateObject("Scripting.Dictionary")
    seqCounter.CompareMode = vbTextCompare

    For r = 1 To (MAX_ROWS - 2 + 1)
        Dim rmCodeRaw As String: rmCodeRaw = Trim(CStr(data(r, 4)))
        Dim poNo As String: poNo = Trim(CStr(data(r, 9)))
        If Len(rmCodeRaw) = 0 Or Len(poNo) = 0 Then GoTo NextRow

        Dim kRow As String: kRow = NormalizeText(rmCodeRaw)
        Dim matchedPart As String: matchedPart = ""
        If rmCodeIdx.Exists(kRow) Then
            matchedPart = rmCodeIdx(kRow)
        Else
            ' "0"(ゼロ)/"O"(オー)表記ゆれの読み替えを試す(RefreshData_StockActualsの
            ' ResolveTTAFPartと同じ理由。TTAF側の元データで度々見つかる表記ゆれ。
            ' 例: CSA ReportのCSA Product Codeが"0JN"、M_RawMaterials側は正式に"OJN")。
            Dim kZeroToO As String: kZeroToO = Replace(kRow, "0", "O")
            If kZeroToO <> kRow And rmCodeIdx.Exists(kZeroToO) Then
                matchedPart = rmCodeIdx(kZeroToO)
            Else
                Dim kOToZero As String: kOToZero = Replace(kRow, "O", "0")
                If kOToZero <> kRow And rmCodeIdx.Exists(kOToZero) Then matchedPart = rmCodeIdx(kOToZero)
            End If
            ' それでも見つからなければ、記号ゆれでは吸収しきれない既知の別名(BuildKnownAliasIndex
            ' 参照)を試す。
            If Len(matchedPart) = 0 And knownAliasIdx.Exists(kRow) Then
                Dim aliasKey As String: aliasKey = NormalizeText(CStr(knownAliasIdx(kRow)))
                If rmCodeIdx.Exists(aliasKey) Then matchedPart = rmCodeIdx(aliasKey)
            End If
        End If

        If Len(matchedPart) = 0 Then
            If InStr(unresolved, rmCodeRaw) = 0 Then unresolved = unresolved & rmCodeRaw & "; "
            GoTo NextRow
        End If

        Dim qty As Variant: qty = data(r, 14)
        If Not IsNumeric(qty) Then qty = 0

        Dim eta As Variant: eta = data(r, 17)  ' Q列(2 week transit to TTAF)。P列そのものは使わない
        If Not IsDate(eta) Then eta = Empty

        Dim receivedDate As Variant: receivedDate = data(r, 19)
        If Not IsDate(receivedDate) Then receivedDate = Empty

        Dim statusVal As String: statusVal = Trim(CStr(data(r, 20)))

        Dim orderMonth As Variant: orderMonth = data(r, 7)
        If Not IsDate(orderMonth) Then orderMonth = Empty

        Dim vessel As String: vessel = Trim(CStr(data(r, 11)))
        Dim container As String: container = Trim(CStr(data(r, 12)))
        Dim origEtd As Variant: origEtd = data(r, 15)
        If Not IsDate(origEtd) Then origEtd = Empty

        Dim baseKey As String
        baseKey = matchedPart & "|" & poNo & "|" & container & "|" & DateKeyStr(origEtd)
        Dim seq As Long
        If seqCounter.Exists(baseKey) Then
            seq = seqCounter(baseKey) + 1
        Else
            seq = 1
        End If
        seqCounter(baseKey) = seq
        Dim key As String: key = baseKey & "|" & seq

        If shipIdx.Exists(key) Then
            updateVals(shipIdx(key)) = Array(CDbl(qty), eta, receivedDate, statusVal, orderMonth, vessel, container, origEtd)
            updated = updated + 1
        Else
            If Not newRecords.Exists(key) Then added = added + 1
            newRecords(key) = Array(matchedPart, poNo, CDbl(qty), eta, receivedDate, statusVal, orderMonth, vessel, container, origEtd)
        End If
NextRow:
    Next r

    ' ---- 既存行の更新をまとめて反映(行ごとに読み込み→書き換え→書き戻しを1回ずつ) ----
    ' 【重要・過去からの不具合を今回発見して修正】8列目(Effective_Week)は数式セルのため、
    ' 行全体を1回の.Range.Value=配列で書き戻すと、.Valueで読み込んだ時点の「計算済みの
    ' 値」がそのまま書き込まれてしまい、数式そのものが壊れて固定値になってしまう
    ' (以前はこのバグにより、一度でも更新された行のEffective_Weekがその時点の値のまま
    ' 永久に凍結され、その後Latest_ETA/Received_Dateが変わっても週がまったく
    ' 追従しなくなっていた)。8列目を挟んで前半(1〜7列)・後半(9〜12列)の2つに
    ' 分けて書き戻すことで、8列目の数式には一切触れないようにする。
    Dim rowKey As Variant
    For Each rowKey In updateVals.Keys
        Dim rowN As Long: rowN = CLng(rowKey)
        Dim uVals As Variant: uVals = updateVals(rowKey)
        Dim rowRng As Range: Set rowRng = shipTbl.ListRows(rowN).Range
        Dim rowArr As Variant: rowArr = rowRng.Value
        rowArr(1, 4) = uVals(0)
        If Not IsEmpty(uVals(1)) Then rowArr(1, 5) = uVals(1)
        If Not IsEmpty(uVals(2)) Then rowArr(1, 6) = uVals(2)
        rowArr(1, 7) = uVals(3)
        If Not IsEmpty(uVals(4)) Then rowArr(1, 9) = uVals(4)
        rowArr(1, 10) = uVals(5)
        rowArr(1, 11) = uVals(6)
        If Not IsEmpty(uVals(7)) Then rowArr(1, 12) = uVals(7)

        Dim leftPart As Variant
        ReDim leftPart(1 To 1, 1 To 7)
        Dim ci As Long
        For ci = 1 To 7: leftPart(1, ci) = rowArr(1, ci): Next ci
        rowRng.Resize(1, 7).Value = leftPart

        Dim rightPart As Variant
        ReDim rightPart(1 To 1, 1 To 4)
        For ci = 9 To 12: rightPart(1, ci - 8) = rowArr(1, ci): Next ci
        rowRng.Cells(1, 9).Resize(1, 4).Value = rightPart
    Next rowKey

    ' ---- 新規行をまとめて追加(1回のResize+配列書き込み。Effective_Week列は計算列の
    ' 自動複製が効かないため、build_soh.pyのweek_index_formula_clampedと同じ式を明示的に書く) ----
    If newRecords.Count > 0 Then
        Dim oldRowCount As Long: oldRowCount = shipTbl.ListRows.Count
        Dim nWeeksCal As Long: nWeeksCal = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count
        shipTbl.Resize shipTbl.Range.Resize(shipTbl.Range.Rows.Count + newRecords.Count, shipTbl.Range.Columns.Count)
        Dim nNew As Long: nNew = newRecords.Count
        Dim outArr() As Variant
        ReDim outArr(1 To nNew, 1 To 12)
        Dim ni As Long: ni = 0
        Dim recKey As Variant
        For Each recKey In newRecords.Keys
            ni = ni + 1
            Dim rec As Variant: rec = newRecords(recKey)
            Dim absRow As Long: absRow = oldRowCount + 1 + ni  ' このテーブルの新規行の実シート行番号(ヘッダー1行分+既存行+ni)
            outArr(ni, 1) = rec(0)   ' Part Name
            outArr(ni, 2) = rec(1)   ' PO_No
            outArr(ni, 3) = Empty    ' Order_Date(手入力欄のため触れない)
            outArr(ni, 4) = rec(2)   ' Confirmed_Qty
            outArr(ni, 5) = rec(3)   ' Latest_ETA
            outArr(ni, 6) = rec(4)   ' Received_Date
            outArr(ni, 7) = rec(5)   ' Status
            outArr(ni, 8) = "=IFERROR(MAX(1,MIN(" & nWeeksCal & ",INT((IF(F" & absRow & "="""",E" & absRow & ",F" & absRow & _
                ")-(DATE(Cal_Weeks!$B$1,1,1)-WEEKDAY(DATE(Cal_Weeks!$B$1,1,1),3)))/7)+1)),"""")"
            outArr(ni, 9) = rec(6)   ' Order_Month
            outArr(ni, 10) = rec(7)  ' Vessel
            outArr(ni, 11) = rec(8)  ' Container
            outArr(ni, 12) = rec(9)  ' Original_ETD
        Next recKey
        shipTbl.ListRows(oldRowCount + 1).Range.Resize(nNew, 12).Formula = outArr
    End If

    ' 【重要】Application.Calculation=xlCalculationManualのままなので、既存行のLatest_ETA/
    ' Received_Date更新も新規行のEffective_Week数式も、このままではまだ再計算されていない
    ' (SyncMaterialDetailOrdersがDataBodyRange.Valueで読むのは古いキャッシュ値のままになる
    ' 恐れがある)。T_Shipmentsの範囲だけを明示的に再計算する(ブック全体の再計算は避ける)。
    shipTbl.Range.Calculate

    ' T_Shipmentsの取り込みが終わった後、Material_DetailのOrder/PO_No行をCSA Reportの
    ' 最新のStatus/ETAに合わせて同期する(モジュール冒頭コメント参照)。
    Dim mdChanged As Long, mdFrozen As Long
    Call SyncMaterialDetailOrders(thisWb, rmTbl, shipTbl, mdChanged, mdFrozen)

    ' srcWbが既にNothingになっているケース(取込元ファイル側の自動処理等で、開いた
    ' 直後にワークブックが閉じられてしまう場合がある)でも、後始末処理自体が
    ' 「オブジェクト変数が設定されていません」で落ちないようにガードする。
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim msg As String
    msg = "T_Shipments を更新しました。" & vbCrLf & "追加: " & added & " 件、更新: " & updated & " 件" & vbCrLf & _
          "（PO No＋材料の組み合わせが同じ行は上書きされます。Order_Date欄は手入力のため上書きしません）" & vbCrLf & vbCrLf & _
          "Material_DetailのOrder/PO_No自動更新: " & mdChanged & " 件(うちTTAF在庫として確定・計算対象から除外: " & mdFrozen & " 件)" & vbCrLf & _
          "（変更箇所にはセルコメントを付けています）"
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

' T_Shipmentsに、分割出荷を区別するためのVessel/Container/Original_ETD列(10〜12列目)を
' 追加する(一度だけ実行する移行用マクロ)。build_soh.pyで新規に作られるブックは最初から
' この列があるが、既存のライブブックには無い。既存行はこれらの列が空欄のまま追加されるが、
' 次にRefreshShipmentsを実行すると、以前は材料名+PO番号だけの一意キーで上書きされてしまい
' 失われていた分割出荷の行が、複合キーで正しく区別されて自動的に追加され直す
' (モジュール冒頭コメント参照)。列が既にある場合は列追加をスキップするだけで、
' 下記のEffective_Week修復は毎回必ず行う(誤って複数回実行しても安全)。
'
' 【重要・今回の再検証で新たに発見した別の不具合の修復も兼ねる】以前のRefreshShipmentsには、
' 既存行を更新する際にEffective_Week(8列目、着荷予定週を計算する数式)を、その時点の
' 計算結果の値でまるごと上書きしてしまい、数式自体を破壊する不具合があった(コード側は
' 今回のモジュール貼り替えで修正済みだが、過去に実行された分は直らないまま残っている)。
' これにより、一度でも更新されたことがある行は、ETAがその後どれだけ変わってもEffective_Week
' が更新時点の値のまま凍結され、Material_Detailへの反映が追従しなくなっていた。
' そのため、このマクロで全行のEffective_Week数式を一括で正しい状態に復元する。
Sub AddShipmentSplitColumns()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim shipTbl As ListObject: Set shipTbl = thisWb.Sheets("T_Shipments").ListObjects("T_Shipments")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim alreadyHadColumns As Boolean: alreadyHadColumns = (shipTbl.ListColumns.Count >= 12)
    If Not alreadyHadColumns Then
        shipTbl.Resize shipTbl.Range.Resize(shipTbl.Range.Rows.Count, 12)
        shipTbl.HeaderRowRange.Cells(1, 10).Value = "Vessel"
        shipTbl.HeaderRowRange.Cells(1, 11).Value = "Container"
        shipTbl.HeaderRowRange.Cells(1, 12).Value = "Original_ETD"
    End If

    Dim repaired As Long: repaired = 0
    Dim n As Long: n = shipTbl.ListRows.Count
    If n > 0 Then
        Dim nWeeksCal As Long: nWeeksCal = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count
        Dim fArr() As Variant
        ReDim fArr(1 To n, 1 To 1)
        Dim i As Long
        For i = 1 To n
            Dim absRow As Long: absRow = i + 1  ' ヘッダー1行分のオフセット
            fArr(i, 1) = "=IFERROR(MAX(1,MIN(" & nWeeksCal & ",INT((IF(F" & absRow & "="""",E" & absRow & ",F" & absRow & _
                ")-(DATE(Cal_Weeks!$B$1,1,1)-WEEKDAY(DATE(Cal_Weeks!$B$1,1,1),3)))/7)+1)),"""")"
        Next i
        shipTbl.ListColumns(8).DataBodyRange.Formula = fArr
        repaired = n
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim msg As String
    If alreadyHadColumns Then
        msg = "T_Shipmentsは列構成(Vessel/Container/Original_ETD)は既に最新でした。"
    Else
        msg = "T_ShipmentsにVessel/Container/Original_ETD列を追加しました。" & vbCrLf & vbCrLf & _
              "既存の行はこの3列が空欄のままです。次にRefreshShipmentsを実行すると、" & vbCrLf & _
              "以前は同じ材料+PO番号で上書きされて消えていた分割出荷の行が、" & vbCrLf & _
              "自動的に正しく追加され直します。" & vbCrLf & vbCrLf
    End If
    msg = msg & "あわせて、Effective_Week(着荷予定週)の数式を全" & repaired & "行分、正しい状態に" & vbCrLf & _
          "復元しました(以前のRefreshShipmentsには、既存行を更新するたびにこの数式を計算済みの" & vbCrLf & _
          "値で上書きして壊してしまう不具合があり、その影響を受けていた行を修復しています)。"
    MsgBox msg, vbInformation
    Exit Sub

ErrHandler:
    Dim errNum3 As Long: errNum3 = Err.Number
    Dim errMsg3 As String: errMsg3 = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "更新処理でエラーが発生しました: (" & errNum3 & ") " & errMsg3, vbCritical
End Sub

' 【一度だけ実行する移行用マクロ】AddShipmentSplitColumns実行後、最初のRefreshShipments実行で
' 新形式(Vessel/Container/Original_ETDが入った)行が追加されると、移行前から存在した旧形式の行
' (Container・Original_ETDが両方とも空欄)は、複合キーが一致しないため以後RefreshShipmentsで
' 二度と更新されずそのまま残り続ける。もしその旧形式行のPO番号がMaterial_Detail側でまだ
' 確定([済])していなければ、SyncMaterialDetailOrdersは同じPO番号の旧行・新行の両方を合算して
' しまい、実際より多い数量を発注予定として計上してしまう(二重計上)。
' そのため、旧形式の行のうち、同じ材料+PO番号の新形式の行が既に存在するものだけを削除する
' (新形式の行の方が正しい最新データなので、旧形式は完全に不要になっている。まだ新形式の行が
' 無いPO番号の旧行はそのまま残す=削除しない。既に[済]で確定済みのPO番号は、そもそも
' SyncMaterialDetailOrdersの集計対象から外れるため元々問題にならないが、対象PO番号の判定は
' 単純にするため、確定済みかどうかは見ずに「新形式の行が存在するかどうか」だけで判定する)。
' 【重要】実データ477件で確認した限り、コンテナ番号・Original ETDが両方空欄になることは
' 無いため、「両方空欄」を旧形式行の判定条件として安全に使える。
Sub CleanupOrphanedPreSplitShipmentRows()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Dim shipTbl As ListObject: Set shipTbl = thisWb.Sheets("T_Shipments").ListObjects("T_Shipments")

    If shipTbl.ListColumns.Count < 12 Then
        MsgBox "先に AddShipmentSplitColumns を実行してください。", vbExclamation
        Exit Sub
    End If

    Dim n As Long: n = shipTbl.ListRows.Count
    If n = 0 Then
        MsgBox "T_Shipmentsにデータがありません。", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim data As Variant: data = shipTbl.DataBodyRange.Value

    ' 新形式の行(コンテナ・Original ETDのどちらかが入力済み)がある「材料+PO番号」を集める
    Dim hasNewFormat As Object: Set hasNewFormat = CreateObject("Scripting.Dictionary")
    hasNewFormat.CompareMode = vbTextCompare
    Dim i As Long
    For i = 1 To n
        Dim containerV As String: containerV = Trim(CStr(data(i, 11)))
        Dim origEtdV As Variant: origEtdV = data(i, 12)
        If Len(containerV) > 0 Or IsDate(origEtdV) Then
            Dim mpKey As String: mpKey = Trim(CStr(data(i, 1))) & "|" & Trim(CStr(data(i, 2)))
            If Not hasNewFormat.Exists(mpKey) Then hasNewFormat.Add mpKey, True
        End If
    Next i

    ' 旧形式(コンテナ・Original ETDが両方空欄)の行のうち、新形式の行が既にある
    ' 材料+PO番号のものだけを削除対象にする
    Dim rowsToDelete As Object: Set rowsToDelete = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        Dim containerV2 As String: containerV2 = Trim(CStr(data(i, 11)))
        Dim origEtdV2 As Variant: origEtdV2 = data(i, 12)
        If Len(containerV2) = 0 And Not IsDate(origEtdV2) Then
            Dim poV As String: poV = Trim(CStr(data(i, 2)))
            Dim mpKey2 As String: mpKey2 = Trim(CStr(data(i, 1))) & "|" & poV
            If Len(poV) > 0 And hasNewFormat.Exists(mpKey2) Then rowsToDelete(i) = True
        End If
    Next i

    Dim delCount As Long: delCount = rowsToDelete.Count
    If delCount = 0 Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "削除対象の旧形式行はありませんでした(既にクリーンな状態です)。", vbInformation
        Exit Sub
    End If

    ' 行番号の大きい方から順に削除する(小さい番号から消すとインデックスがずれるため)
    Dim delRows() As Long
    ReDim delRows(1 To delCount)
    Dim delRow As Variant, di As Long: di = 0
    For Each delRow In rowsToDelete.Keys
        di = di + 1
        delRows(di) = CLng(delRow)
    Next delRow
    Dim a As Long, b As Long, tmp As Long
    For a = 1 To delCount - 1
        For b = a + 1 To delCount
            If delRows(b) > delRows(a) Then
                tmp = delRows(a): delRows(a) = delRows(b): delRows(b) = tmp
            End If
        Next b
    Next a
    For a = 1 To delCount
        shipTbl.ListRows(delRows(a)).Delete
    Next a

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "二重計上の原因になり得る旧形式の行を " & delCount & " 件削除しました。" & vbCrLf & vbCrLf & _
           "(移行前から存在し、同じ材料+PO番号の新形式の行に置き換わった行のみを削除しています。" & vbCrLf & _
           "まだ新形式の行が無いPO番号の行には触れていません。削除後、あらためて" & vbCrLf & _
           "RefreshShipmentsを実行してMaterial_Detailを最新状態に同期してください)", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum4 As Long: errNum4 = Err.Number
    Dim errMsg4 As String: errMsg4 = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました: (" & errNum4 & ") " & errMsg4, vbCritical
End Sub

' M_RawMaterialsのPart Name(=RM_Code)自体を正規化テキストでインデックス化する。
' TTAF_Code/Descriptionでの照合(RefreshData_StockActuals側のBuildTTAFCodeAndDescIndex)とは
' 別物: CSA Product Codeは既にRM_Codeとほぼ同じ表記のため、RM_Code同士を直接照合する方が確実。
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

' CSA Report側の品名表記が、M_RawMaterials側の正式なPart Nameと大きく異なる既知のケースを
' 個別に対応する(単純な記号ゆれ・0/O表記ゆれでは吸収しきれないもの)。
' キーはCSA Product Code側の表記をNormalizeTextしたもの、値はM_RawMaterials側の正式な
' Part Name(こちらもRefreshShipments側でNormalizeTextしてrmCodeIdxと突き合わせる)。
' 新しいケースが見つかったら、ここに1行追加するだけで対応できる。
Private Function BuildKnownAliasIndex() As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbTextCompare
    ' CSA Reportの"PET FILM(900*1000)"は、M_RawMaterials上は"Ester Film"として登録されている
    ' (当初"PP Film"だと聞いていたが、確認の結果こちらが正しいと判明した)。
    idx(NormalizeText("PET FILM(900*1000)")) = "Ester Film"
    Set BuildKnownAliasIndex = idx
End Function

' T_Shipmentsの複合キー(材料名+PO番号+コンテナ+OriginalETD+出現順連番)->行番号の
' インデックスを1回だけ作る。RefreshShipments側の複合キー生成と全く同じロジックで、
' T_Shipments自身の現在の並び順から連番を振り直す(モジュール冒頭コメント参照)。
Private Function BuildShipmentRowIndex(tbl As ListObject) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim n As Long: n = tbl.ListRows.Count
    If n > 0 Then
        Dim data As Variant
        data = tbl.DataBodyRange.Value  ' 全12列(Part Name...Original_ETD)をまとめて読む
        Dim seqCounter As Object: Set seqCounter = CreateObject("Scripting.Dictionary")
        seqCounter.CompareMode = vbTextCompare
        Dim i As Long
        For i = 1 To n
            Dim partNameV As String: partNameV = Trim(CStr(data(i, 1)))
            Dim poNoV As String: poNoV = Trim(CStr(data(i, 2)))
            Dim containerV As String: containerV = Trim(CStr(data(i, 11)))
            Dim origEtdV As Variant: origEtdV = data(i, 12)
            Dim baseKey As String
            baseKey = partNameV & "|" & poNoV & "|" & containerV & "|" & DateKeyStr(origEtdV)
            Dim seq As Long
            If seqCounter.Exists(baseKey) Then
                seq = seqCounter(baseKey) + 1
            Else
                seq = 1
            End If
            seqCounter(baseKey) = seq
            idx(baseKey & "|" & seq) = i
        Next i
    End If
    Set BuildShipmentRowIndex = idx
End Function

' 日付(またはEmpty)を、複合キーに使うための安定した文字列に変換する
' (地域設定による日付表示の違いに影響されないよう、シリアル値の整数部分をそのまま使う)。
Private Function DateKeyStr(v As Variant) As String
    If IsEmpty(v) Then
        DateKeyStr = ""
    ElseIf IsDate(v) Then
        DateKeyStr = CStr(CLng(CDate(v)))
    Else
        DateKeyStr = ""
    End If
End Function

' 列: Part Name(1), PO_No(2), Order_Date_発注日(3, 手入力のため触れない), Confirmed_Qty(4),
' Latest_ETA(5), Received_Date(6), Status(7), Effective_Week(8, 数式), Order_Month(9),
' Vessel(10), Container(11), Original_ETD(12)。
' Effective_Week(8列目)の数式は、RefreshShipments側で新規行をまとめて書き込む際に
' build_soh.pyのweek_index_formula_clampedと同じ式を明示的に生成している。
' Vessel/Container/Original_ETDは、分割出荷の各行を一意に区別するための複合キーに
' 使われる(BuildShipmentRowIndex/DateKeyStr参照)。

' Material_DetailのOrder行(発注予定,kg)・PO_No行(Order行の直下)を、T_Shipmentsの最新の
' Status/Effective_Weekに合わせて自動更新する(モジュール冒頭コメント参照)。
' 材料1件につきOrder行・PO_No行の2行だけをまとめて読み込み→メモリ上で書き換え→
' まとめて書き戻す(セル単位の読み書きはしない。重量級テーブルではないため件数的に
' 大きな問題にはなりにくいが、既存の設計方針に合わせておく)。
Private Sub SyncMaterialDetailOrders(thisWb As Workbook, rmTbl As ListObject, shipTbl As ListObject, _
        ByRef changedCells As Long, ByRef frozenCells As Long)
    changedCells = 0
    frozenCells = 0
    Dim mdSheet As Worksheet
    On Error Resume Next
    Set mdSheet = thisWb.Sheets("Material_Detail")
    On Error GoTo 0
    If mdSheet Is Nothing Then Exit Sub

    Dim shipN As Long: shipN = shipTbl.ListRows.Count
    If shipN = 0 Then Exit Sub
    Dim shipData As Variant: shipData = shipTbl.DataBodyRange.Value
    ' shipData列: 1=Part Name,2=PO_No,3=Order_Date,4=Confirmed_Qty,5=Latest_ETA,
    '             6=Received_Date,7=Status,8=Effective_Week,9=Order_Month

    ' "材料名|PO番号" -> その組み合わせのshipData行番号一覧(分割出荷は複数行になる)
    Dim byMatPo As Object: Set byMatPo = CreateObject("Scripting.Dictionary")
    byMatPo.CompareMode = vbTextCompare
    Dim i As Long
    For i = 1 To shipN
        Dim sPoNo As String: sPoNo = Trim(CStr(shipData(i, 2)))
        If Len(sPoNo) = 0 Then GoTo NextShipRow
        Dim mpKey As String: mpKey = Trim(CStr(shipData(i, 1))) & "|" & sPoNo
        Dim lst As Object
        If byMatPo.Exists(mpKey) Then
            Set lst = byMatPo(mpKey)
        Else
            Set lst = CreateObject("Scripting.Dictionary")
            byMatPo.Add mpKey, lst
        End If
        lst.Add lst.Count, i
NextShipRow:
    Next i
    If byMatPo.Count = 0 Then Exit Sub

    ' M_RawMaterialsのLeadTime_Weeks_要入力を材料名で引けるように1回だけ索引化
    ' (ETA未定(TBC)の仮予測に使う)。
    Dim ltIdx As Object: Set ltIdx = CreateObject("Scripting.Dictionary")
    ltIdx.CompareMode = vbTextCompare
    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    If rmN > 0 Then
        Dim rmNameLt As Variant: rmNameLt = rmTbl.ListColumns(1).DataBodyRange.Resize(rmN, 1).Value
        Dim rmLtCol As Variant: rmLtCol = rmTbl.ListColumns("LeadTime_Weeks_要入力").DataBodyRange.Value
        Dim li As Long
        For li = 1 To rmN
            Dim ltk As String: ltk = Trim(CStr(rmNameLt(li, 1)))
            If Len(ltk) > 0 And Not ltIdx.Exists(ltk) Then
                Dim ltv As Double: ltv = 0
                If IsNumeric(rmLtCol(li, 1)) Then ltv = CDbl(rmLtCol(li, 1))
                ltIdx(ltk) = ltv
            End If
        Next li
    End If

    ' Material_Detailの各ブロックの「Order」行位置を1回だけ収集する(PO_No行はその1つ下)。
    Dim lastRow As Long: lastRow = mdSheet.Cells(mdSheet.Rows.Count, 2).End(xlUp).Row
    Dim orderRowByMat As Object: Set orderRowByMat = CreateObject("Scripting.Dictionary")
    orderRowByMat.CompareMode = vbTextCompare
    Dim curMatCode As String: curMatCode = ""
    Dim r As Long
    For r = MD_HEADER_ROW + 1 To lastRow
        Dim colAVal As String: colAVal = Trim(CStr(mdSheet.Cells(r, 1).Value))
        If Len(colAVal) > 0 Then curMatCode = colAVal
        If Len(curMatCode) > 0 And Trim(CStr(mdSheet.Cells(r, 2).Value)) = "Order(発注予定,kg)" Then
            If Not orderRowByMat.Exists(curMatCode) Then orderRowByMat.Add curMatCode, r
        End If
    Next r
    If orderRowByMat.Count = 0 Then Exit Sub

    Dim nWeeks As Long: nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count
    Dim lastWeekCol As Long: lastWeekCol = MD_WEEK_START_COL + nWeeks - 1

    ' 出荷情報の"材料名"側だけを抜き出し、重複を除いて1材料につき1回だけ処理する。
    Dim matNamesSeen As Object: Set matNamesSeen = CreateObject("Scripting.Dictionary")
    matNamesSeen.CompareMode = vbTextCompare
    Dim mpKeyOuter As Variant
    For Each mpKeyOuter In byMatPo.Keys
        Dim sepPos As Long: sepPos = InStr(CStr(mpKeyOuter), "|")
        Dim matName As String: matName = Left(CStr(mpKeyOuter), sepPos - 1)
        If matNamesSeen.Exists(matName) Then GoTo NextMatKey
        matNamesSeen.Add matName, True
        If Not orderRowByMat.Exists(matName) Then GoTo NextMatKey  ' Material_Detailにブロックが無い材料

        Dim orderRow As Long: orderRow = orderRowByMat(matName)
        Dim poRow As Long: poRow = orderRow + 1
        Dim blockArr As Variant
        blockArr = mdSheet.Range(mdSheet.Cells(orderRow, MD_WEEK_START_COL), mdSheet.Cells(poRow, lastWeekCol)).Value
        ' blockArr(1,w)=Order行の値, blockArr(2,w)=PO_No行の値 (w=1..nWeeks)

        ' この材料のPO_No行に現れる、まだ確定していない("[済]"が付いていない)PO番号を
        ' 週ごとに集め、PO番号ごとにグループ化する。
        Dim activeByPo As Object: Set activeByPo = CreateObject("Scripting.Dictionary")
        activeByPo.CompareMode = vbTextCompare
        Dim w As Long
        For w = 1 To nWeeks
            Dim poCellVal As String: poCellVal = Trim(CStr(blockArr(2, w)))
            If Len(poCellVal) > 0 And InStr(poCellVal, "[済]") = 0 Then
                Dim poLst As Object
                If activeByPo.Exists(poCellVal) Then
                    Set poLst = activeByPo(poCellVal)
                Else
                    Set poLst = CreateObject("Scripting.Dictionary")
                    activeByPo.Add poCellVal, poLst
                End If
                poLst.Add poLst.Count, w
            End If
        Next w
        If activeByPo.Count = 0 Then GoTo NextMatKey

        Dim blockChanged As Boolean: blockChanged = False
        Dim commentsToAdd As Object: Set commentsToAdd = CreateObject("Scripting.Dictionary")  ' 週 -> コメント文字列

        Dim poKey As Variant
        For Each poKey In activeByPo.Keys
            Dim thisMpKey As String: thisMpKey = matName & "|" & CStr(poKey)
            If Not byMatPo.Exists(thisMpKey) Then GoTo NextPoKey  ' このPO番号の出荷情報はまだCSA Reportに無い

            Dim shipRowsForPo As Object: Set shipRowsForPo = byMatPo(thisMpKey)
            Dim existingWeeks As Object: Set existingWeeks = activeByPo(poKey)

            ' --- CSA Reportの最新情報から、このPO番号の「あるべき状態」を組み立てる ---
            Dim newWeeks As Object: Set newWeeks = CreateObject("Scripting.Dictionary")  ' 週 -> Array(qty, frozen)
            Dim siKey As Variant
            For Each siKey In shipRowsForPo.Keys
                Dim shipRow As Long: shipRow = shipRowsForPo(siKey)
                Dim qty As Double: qty = 0
                If IsNumeric(shipData(shipRow, 4)) Then qty = CDbl(shipData(shipRow, 4))
                Dim statusText As String: statusText = Trim(CStr(shipData(shipRow, 7)))
                Dim targetWeek As Long: targetWeek = 0
                If IsNumeric(shipData(shipRow, 8)) Then
                    targetWeek = CLng(shipData(shipRow, 8))
                Else
                    ' ETAが未定(TBC)。Order_Month + LeadTime_Weeksから仮の週を計算する
                    ' (月の中央=15日を起点にすることで、月初/月末どちらかに偏らないようにする)。
                    Dim orderMonthVal As Variant: orderMonthVal = shipData(shipRow, 9)
                    If IsDate(orderMonthVal) Then
                        Dim ltWeeks As Double: ltWeeks = 0
                        If ltIdx.Exists(matName) Then ltWeeks = ltIdx(matName)
                        Dim provDate As Date: provDate = DateSerial(Year(CDate(orderMonthVal)), Month(CDate(orderMonthVal)), 15) + ltWeeks * 7
                        targetWeek = WeekIndexForDate(thisWb, provDate)
                    End If
                End If
                If targetWeek > 0 Then
                    Dim frozenFlag As Boolean: frozenFlag = (statusText = "TTAF Stock")
                    ' 同じPO番号の複数の出荷行(分割出荷)が同じ週に重なることもあるため、
                    ' 上書きせず合算する。frozen(確定扱い)は、その週に合算された行の
                    ' 全部がTTAF Stockになって初めて真にする(1件でもまだ輸送中なら、
                    ' その週はまだGrid_Incomingの計算対象から外さない)。
                    If newWeeks.Exists(targetWeek) Then
                        Dim existingNW As Variant: existingNW = newWeeks(targetWeek)
                        newWeeks(targetWeek) = Array(CDbl(existingNW(0)) + qty, CBool(existingNW(1)) And frozenFlag)
                    Else
                        newWeeks(targetWeek) = Array(qty, frozenFlag)
                    End If
                End If
            Next siKey
            If newWeeks.Count = 0 Then GoTo NextPoKey

            ' --- 「あるべき状態」と今の状態を、週ごとに個別に比較する ---
            ' (以前の「PO全体をまとめてクリア→まとめて書き直す」方式だと、分割出荷の一部だけ
            ' 既に確定[済]で残りがまだ輸送中、というケースで、既に確定済みで何も変わっていない
            ' 週まで毎回クリア→再書き込みされてしまい、コメントが無関係な週の移動として
            ' 表示されたり、frozenCellsが同じ確定を何度も数えてしまう不具合があった。
            ' そのため、週ごとに「今の内容と違うものだけ」を書き換える。)
            ' weeksToClear: 今アクティブ(未確定)な週のうち、新しい状態には存在しない(=移動元)週
            Dim weeksToClear As Object: Set weeksToClear = CreateObject("Scripting.Dictionary")
            Dim ewk As Variant
            For Each ewk In existingWeeks.Items
                Dim ewkL As Long: ewkL = CLng(ewk)
                If Not newWeeks.Exists(ewkL) Then weeksToClear(ewkL) = True
            Next ewk
            ' weeksToWrite: 新しい状態のうち、今のセルの内容(数量・確定マーク)と実際に違うものだけ
            Dim weeksToWrite As Object: Set weeksToWrite = CreateObject("Scripting.Dictionary")  ' 週 -> Array(qty,frozen,poText)
            Dim wk As Variant
            For Each wk In newWeeks.Keys
                Dim nwk As Long: nwk = CLng(wk)
                Dim info As Variant: info = newWeeks(wk)
                Dim wantQty As Double: wantQty = CDbl(info(0))
                Dim wantFrozen As Boolean: wantFrozen = CBool(info(1))
                Dim wantPoText As String: wantPoText = CStr(poKey)
                If wantFrozen Then wantPoText = wantPoText & " [済]"

                Dim curQty As Double: curQty = 0
                If IsNumeric(blockArr(1, nwk)) Then curQty = CDbl(blockArr(1, nwk))
                Dim curPoText As String: curPoText = Trim(CStr(blockArr(2, nwk)))

                If Abs(curQty - wantQty) > 0.0001 Or curPoText <> wantPoText Then
                    weeksToWrite(nwk) = Array(wantQty, wantFrozen, wantPoText)
                End If
            Next wk
            If weeksToClear.Count = 0 And weeksToWrite.Count = 0 Then GoTo NextPoKey

            ' --- 移動元(weeksToClear)をクリアする(材料ブロックのメモリ上の配列だけを操作) ---
            Dim oldWeeksSummary As String: oldWeeksSummary = ""
            Dim cwk2 As Variant
            For Each cwk2 In weeksToClear.Keys
                Dim ow As Long: ow = CLng(cwk2)
                If Len(oldWeeksSummary) > 0 Then oldWeeksSummary = oldWeeksSummary & "、"
                oldWeeksSummary = oldWeeksSummary & "週" & ow & "(" & Format(blockArr(1, ow), "0") & "kg)"
                blockArr(1, ow) = Empty
                blockArr(2, ow) = Empty
            Next cwk2

            Dim changeNote As String
            If Len(oldWeeksSummary) > 0 Then
                changeNote = "PO#" & poKey & ": " & oldWeeksSummary & " から自動更新されました(" & Format(Date, "yyyy-mm-dd") & ")"
            Else
                changeNote = "PO#" & poKey & ": 自動更新されました(" & Format(Date, "yyyy-mm-dd") & ")"
            End If
            Dim wwk As Variant
            For Each wwk In weeksToWrite.Keys
                Dim nw As Long: nw = CLng(wwk)
                Dim winfo As Variant: winfo = weeksToWrite(wwk)
                blockArr(1, nw) = CDbl(winfo(0))
                blockArr(2, nw) = CStr(winfo(2))
                If CBool(winfo(1)) Then frozenCells = frozenCells + 1
                commentsToAdd(nw) = changeNote
            Next wwk

            blockChanged = True
            changedCells = changedCells + 1
NextPoKey:
        Next poKey

        If blockChanged Then
            mdSheet.Range(mdSheet.Cells(orderRow, MD_WEEK_START_COL), mdSheet.Cells(poRow, lastWeekCol)).Value = blockArr
            Dim cwKey As Variant
            For Each cwKey In commentsToAdd.Keys
                Dim commentCol As Long: commentCol = MD_WEEK_START_COL + CLng(cwKey) - 1
                Dim targetCell As Range: Set targetCell = mdSheet.Cells(orderRow, commentCol)
                On Error Resume Next
                targetCell.Comment.Delete
                targetCell.AddComment CStr(commentsToAdd(cwKey))
                targetCell.Comment.Shape.TextFrame.AutoSize = True
                On Error GoTo 0
            Next cwKey
        End If
NextMatKey:
    Next mpKeyOuter
End Sub
