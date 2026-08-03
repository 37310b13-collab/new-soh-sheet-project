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
' N=Confirmed Order Qty、P=Latest ETA、S=Received At TTAF、T=Status。
' CSA Product CodeはTTAF PART NUMBERと異なり、既にこちらのRM_Codeとほぼ同じ表記のため、
' TTAF_Code/Description経由の照合(RefreshData_StockActuals側のResolveTTAFPart)ではなく、
' RM_Code同士を直接(大文字小文字・前後空白を無視して)照合する。
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
