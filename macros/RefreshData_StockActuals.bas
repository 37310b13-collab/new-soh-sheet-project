Attribute VB_Name = "RefreshData_StockActuals"
Option Explicit

' ============================================================================
' RefreshData_StockActuals モジュール
'
'   RefreshSelfStock : 「Raw materials daily check」(自社倉庫の現物確認シート、ファイル名に
'                      DD.MM.YYYY形式の日付を含む)を選択すると、T_SelfStockにその週の実績を
'                      追加/更新する。毎週月曜の朝に確認する運用のため、ファイル名の日付は
'                      「その週の月曜(祝日の場合は翌営業日)」だが、これは前週末時点の在庫を
'                      表す。そのため7日引いてから対象週を判定する(RefreshTTAFStockと同じ考え方。
'                      詳細は下のRefreshTTAFStockの説明を参照)。
'   RefreshTTAFStock : 「CSA Report」を選択すると、その中の「Stock invoiced to CSA」シート
'                      (A列=TTAF PART NUMBER、D列=Description、F列=在庫数量。ヘッダーは4行目、
'                      データは5行目から)からT_TTAFStockにその週の実績を追加/更新する。手入力の
'                      生データを直接読むため、ピボット(旧「 COUNT SHEET SOH」「PIVOT SOH TTAF」)の
'                      更新忘れに左右されない。対象週はF4セルの日付で判定する(ファイル名には
'                      依存しない)。材料の照合はTTAF_Codeを優先し、見つからなければ
'                      Descriptionの正規化テキストで照合する。
'
' どちらも、既存の(原材料, 週)の組み合わせがあれば値を上書き、無ければ新しい行として追加する
' (同じ週内に複数回取り込んでも1行にまとまる。取り込む順序は問わない)。
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
' 全体の設計方針(パフォーマンス・DataBodyRange・DisplayAlerts等)はRefreshData_Utilities
' モジュール冒頭のコメントを参照してください。
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
    ' ファイル名の日付は「確認した月曜日(祝日の場合は翌営業日)」だが、その数値は前週末
    ' 時点の在庫を表す。そのため7日引いてから週Noを判定する(祝日で月曜以外の日になっていても、
    ' ちょうど1週間前にずらすだけなので、前週の範囲内に正しく収まる。RefreshTTAFStockと同じ考え方)。
    Dim reportDate As Date: reportDate = ExtractDateFromName(CStr(srcPath)) - 7

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

    ' 「Stock invoiced to CSA」シートを使う。A列=TTAF PART NUMBER、D列=Description、
    ' F列=在庫数量。ヘッダーは4行目、データは5行目から。手入力の生データなので、
    ' 「 COUNT SHEET SOH」/「PIVOT SOH TTAF」のようなピボット更新忘れの心配が無い。
    Dim sh As Worksheet: Set sh = srcWb.Sheets("Stock invoiced to CSA")
    ' F4の日付は「レポートが届いた月曜日(祝日の場合は翌営業日)」だが、その数値は前週金曜
    ' 営業終了後の在庫を表す。そのため7日引いてから週Noを判定する(月曜も金曜もExcel上は
    ' 同じ月〜日の週に属するため、-7でも-3でも週Noの判定結果は変わらない。日付自体は
    ' 週の起点であるMondayに揃えておいた方が他の実績(T_SelfStock等)と一貫するため-7を使う)。
    Dim reportDate As Date: reportDate = ExtractDDMMYYYYFromText(CStr(sh.Cells(4, 6).Value)) - 7
    Dim wIdx As Long: wIdx = WeekIndexForDate(thisWb, reportDate)
    Dim ttafIdx As Object: Set ttafIdx = BuildStockRowIndex(ttafTbl)

    Dim ttafCodeIdx As Object: Set ttafCodeIdx = CreateObject("Scripting.Dictionary")
    Dim descIdx As Object: Set descIdx = CreateObject("Scripting.Dictionary")
    Call BuildTTAFCodeAndDescIndex(rmTbl, ttafCodeIdx, descIdx)

    ' シートを1セルずつ読むと遅くなるため、余裕を持った範囲を1回だけ配列で読み込んでから走査する
    ' (A列=TTAF PART NUMBER, D列=Description, F列=在庫数量)。
    Const MAX_ROWS As Long = 2000
    Dim data As Variant
    data = sh.Range(sh.Cells(5, 1), sh.Cells(MAX_ROWS, 6)).Value

    Dim r As Long, added As Long, updated As Long, unresolved As String
    added = 0: updated = 0: unresolved = ""
    For r = 1 To (MAX_ROWS - 5 + 1)
        Dim ttafCodeRaw As String: ttafCodeRaw = Trim(CStr(data(r, 1)))
        If Len(ttafCodeRaw) = 0 Then GoTo NextRow
        Dim v As Variant: v = data(r, 6)
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
