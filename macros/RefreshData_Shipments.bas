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
' T_ShipmentsはGrid_Incoming(材料×週のSUMIFS)から大量に参照される重量級テーブルのため、
' RefreshBOM(M_BOM)・RefreshWeeklyBatches(PP_Grid)と同じ理由で、新規行はまとめて件数を
' 数えてから1回のResizeで追加し、既存行の更新も行単位でまとめて読み書きする(1件ずつ
' ListRows.Addを呼ぶとExcelが応答なしになる恐れがあるため)。
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

    ' 【重要】T_ShipmentsはGrid_Incoming(材料×週のSUMIFS、90材料×104週分)から大量に参照される
    ' テーブル。以前M_BOM(RefreshBOM)・PP_Grid(RefreshWeeklyBatches)で、新規行をListRows.Addで
    ' 1件ずつ追加してExcelが応答なしになった不具合と同じ構造(重量級テーブル×大量の1件ずつAdd)
    ' のため、ここでも新規行はまず件数を数えてから1回のResizeでまとめて追加する。既存行の
    ' 更新も、行ごとに複数セルを個別に書き込むのではなく行単位でまとめて読み書きする。
    Dim r As Long, added As Long, updated As Long, unresolved As String
    added = 0: updated = 0: unresolved = ""
    Dim updateVals As Object: Set updateVals = CreateObject("Scripting.Dictionary")  ' 行番号 -> Array(qty,eta,receivedDate,status,orderMonth)
    Dim newRecords As Object: Set newRecords = CreateObject("Scripting.Dictionary")  ' "PartName|PO_No" -> Array(partName,poNo,qty,eta,receivedDate,status,orderMonth)

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

        Dim key As String: key = matchedPart & "|" & poNo
        If shipIdx.Exists(key) Then
            updateVals(shipIdx(key)) = Array(CDbl(qty), eta, receivedDate, statusVal, orderMonth)
            updated = updated + 1
        Else
            If Not newRecords.Exists(key) Then added = added + 1
            newRecords(key) = Array(matchedPart, poNo, CDbl(qty), eta, receivedDate, statusVal, orderMonth)
        End If
NextRow:
    Next r

    ' ---- 既存行の更新をまとめて反映(行ごとに読み込み→書き換え→書き戻しを1回ずつ) ----
    Dim rowKey As Variant
    For Each rowKey In updateVals.Keys
        Dim rowN As Long: rowN = CLng(rowKey)
        Dim uVals As Variant: uVals = updateVals(rowKey)
        Dim rowArr As Variant: rowArr = shipTbl.ListRows(rowN).Range.Value
        rowArr(1, 4) = uVals(0)
        If Not IsEmpty(uVals(1)) Then rowArr(1, 5) = uVals(1)
        If Not IsEmpty(uVals(2)) Then rowArr(1, 6) = uVals(2)
        rowArr(1, 7) = uVals(3)
        If Not IsEmpty(uVals(4)) Then rowArr(1, 9) = uVals(4)
        shipTbl.ListRows(rowN).Range.Value = rowArr
    Next rowKey

    ' ---- 新規行をまとめて追加(1回のResize+配列書き込み。Effective_Week列は計算列の
    ' 自動複製が効かないため、build_soh.pyのweek_index_formula_clampedと同じ式を明示的に書く) ----
    If newRecords.Count > 0 Then
        Dim oldRowCount As Long: oldRowCount = shipTbl.ListRows.Count
        Dim nWeeksCal As Long: nWeeksCal = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count
        shipTbl.Resize shipTbl.Range.Resize(shipTbl.Range.Rows.Count + newRecords.Count, shipTbl.Range.Columns.Count)
        Dim nNew As Long: nNew = newRecords.Count
        Dim outArr() As Variant
        ReDim outArr(1 To nNew, 1 To 9)
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
        Next recKey
        shipTbl.ListRows(oldRowCount + 1).Range.Resize(nNew, 9).Formula = outArr
    End If

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
' Effective_Week(8列目)の数式は、RefreshShipments側で新規行をまとめて書き込む際に
' build_soh.pyのweek_index_formula_clampedと同じ式を明示的に生成している。
