Attribute VB_Name = "RefreshData_StockActuals"
Option Explicit

' ============================================================================
' RefreshData_StockActuals モジュール
'
'   RefreshSelfStock : 「Raw materials daily check」(自社倉庫の現物確認シート、ファイル名に
'                      DD.MM.YYYY形式の日付を含む)を選択すると、T_SelfStockにその週の実績を
'                      追加/更新する。「CHEMICAL SOH」シート(B列=CSA/コード、C列・D列=在庫
'                      〈化学品セクションはC=WH・D=Floor、Substrate/Consumablesセクションは
'                      C=Floor・D=WHと列の意味が入れ替わるが、在庫合計は単純にC+D)を読む。
'                      毎週月曜の朝に確認する運用のため、ファイル名の日付は
'                      「その週の月曜(祝日の場合は翌営業日)」だが、これは前週末時点の在庫を
'                      表す。そのため7日引いてから対象週を判定する(RefreshTTAFStockと同じ考え方。
'                      詳細は下のRefreshTTAFStockの説明を参照)。
'   RefreshTTAFStock : 「CSA Report」を選択すると、その中の「Stock invoiced to CSA」シート
'                      (A列=TTAF PART NUMBER、C列=Part No、D列=Description、F列=在庫数量。
'                      ヘッダーは4行目、データは5行目から)からT_TTAFStockにその週の実績を
'                      追加/更新する。手入力の生データを直接読むため、ピボット(旧「 COUNT
'                      SHEET SOH」「PIVOT SOH TTAF」)の更新忘れに左右されない。対象週は
'                      F4セルの日付で判定する(ファイル名には依存しない)。材料の照合は
'                      TTAF_Codeを優先し、見つからなければPart No(M_RawMaterialsのPart Name
'                      がそのまま入っていることが多く、Descriptionより確実。TTAF側データの
'                      "0"(ゼロ)/"O"(オー)表記ゆれ〈例:CSA ReportのPart No列の"0JN"、
'                      M_RawMaterials側は正式に"OJN"〉も読み替えて試す)で照合し、それでも
'                      見つからない場合だけDescriptionの正規化テキストで照合する。
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
    ' C列(CSA code)がM_RawMaterialsに実在するPart Nameかどうかで対象行を判定する
    ' (以前はLeft(code,4)="CHEM"で化学品のみに限定していたため、Substrate行
    ' 〈OJN・7EP等の短いコード〉が一件も反映されない不具合があった)。
    Dim rmTbl As ListObject: Set rmTbl = thisWb.Sheets("M_RawMaterials").ListObjects("M_RawMaterials")
    Dim rmCodeSet As Object: Set rmCodeSet = BuildNameIndex(rmTbl, "Part Name")

    ' 「Stock」シートのJ列(Total)ではなく「CHEMICAL SOH」シートを直接使う。「Stock」側は
    ' 化学品がK列+L列の数式、Substrateは別シートを参照するVLOOKUPと、材料の種類によって
    ' 仕組みがバラバラな上に、どちらもWorkbooks.Open直後は再計算されず取込元ファイルの
    ' 保存時点のキャッシュ値のまま(古いまま)読み込まれてしまう問題があった。「CHEMICAL SOH」
    ' シートは化学品・Substrate・Ester Film/Original Towel/PP Film等の実測値が全て
    ' B列(CSA/コード)・C列・D列という共通の生データとして入っている(化学品セクションは
    ' C=WH,D=Floor、それ以降のセクションはC=Floor,D=WHと列の意味が入れ替わるが、どちらに
    ' せよ在庫合計はC+D列の単純合計になるため、セクションを区別する必要はない)。
    ' なお、MATコード(18456-xxxxx)はこのシートに載っておらず「Stock」シート側にしかない。
    ' MATは自社在庫の管理対象外のため未対応のままにしている。
    Dim sh As Worksheet: Set sh = srcWb.Sheets("CHEMICAL SOH")
    ' シートを1セルずつ読むと遅くなるため、対象範囲(5〜300行, A〜D列)を1回だけ配列で読み込む。
    Dim data As Variant
    data = sh.Range(sh.Cells(5, 1), sh.Cells(300, 4)).Value
    Dim r As Long, added As Long, updated As Long
    added = 0: updated = 0
    For r = 1 To (300 - 5 + 1)
        Dim code As String
        code = Trim(CStr(data(r, 2)))
        If Len(code) > 0 Then
            Dim matchedCode As String: matchedCode = ResolveSelfStockCode(rmTbl, rmCodeSet, code)
            If Len(matchedCode) > 0 Then
                Dim vC As Double: vC = 0
                Dim vD As Double: vD = 0
                If IsNumeric(data(r, 3)) Then vC = CDbl(data(r, 3))
                If IsNumeric(data(r, 4)) Then vD = CDbl(data(r, 4))
                Call UpsertStockRowIndexed(selfTbl, selfIdx, matchedCode, reportDate, vC + vD, added, updated)
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

    ' 「Stock invoiced to CSA」シートを使う。A列=TTAF PART NUMBER、C列=Part No、
    ' D列=Description、F列=在庫数量。ヘッダーは4行目、データは5行目から。手入力の
    ' 生データなので、「 COUNT SHEET SOH」/「PIVOT SOH TTAF」のようなピボット更新忘れの
    ' 心配が無い。
    Dim sh As Worksheet: Set sh = srcWb.Sheets("Stock invoiced to CSA")
    ' F4の日付は「レポートが届いた月曜日(祝日の場合は翌営業日)」だが、その数値は前週金曜
    ' 営業終了後の在庫を表す。そのため7日引いてから週Noを判定する(月曜も金曜もExcel上は
    ' 同じ月〜日の週に属するため、-7でも-3でも週Noの判定結果は変わらない。日付自体は
    ' 週の起点であるMondayに揃えておいた方が他の実績(T_SelfStock等)と一貫するため-7を使う)。
    ' 「TTAF count(dd.mm.yyyy)」の見出しはF3:F4等で結合されていることがあり、結合セルは
    ' 左上(アンカー)のセルにしか値を持たない。.Cells(4,6)がアンカーでない場合に空を
    ' 読んでしまわないよう、.MergeArea.Cells(1,1)で必ずアンカー側の値を取得する。
    Dim reportDate As Date: reportDate = ExtractDDMMYYYYFromText(sh.Cells(4, 6).MergeArea.Cells(1, 1).Value) - 7
    Dim wIdx As Long: wIdx = WeekIndexForDate(thisWb, reportDate)
    Dim ttafIdx As Object: Set ttafIdx = BuildStockRowIndex(ttafTbl)

    Dim ttafCodeIdx As Object: Set ttafCodeIdx = CreateObject("Scripting.Dictionary")
    Dim descIdx As Object: Set descIdx = CreateObject("Scripting.Dictionary")
    Dim rmNameIdx As Object: Set rmNameIdx = CreateObject("Scripting.Dictionary")
    Call BuildTTAFCodeAndDescIndex(rmTbl, ttafCodeIdx, descIdx, rmNameIdx)

    ' シートを1セルずつ読むと遅くなるため、余裕を持った範囲を1回だけ配列で読み込んでから走査する
    ' (A列=TTAF PART NUMBER, C列=Part No, D列=Description, F列=在庫数量)。
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

        Dim partNoRaw As String: partNoRaw = Trim(CStr(data(r, 3)))
        Dim descRaw As String: descRaw = Trim(CStr(data(r, 4)))
        Dim matchedPart As String
        matchedPart = ResolveTTAFPart(ttafCodeIdx, rmNameIdx, descIdx, ttafCodeRaw, partNoRaw, descRaw)

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

' M_RawMaterialsから、TTAF_Code(正規化済み)->Part Name、Description(正規化済み)->Part Name、
' Part Name(大文字小文字・前後空白を無視)->Part Nameのインデックスを1回だけ作る。
' RefreshTTAFStockが使う共通処理。Dictionaryはオブジェクト(参照渡し)なので、ByRefを
' 明示しなくても呼び出し元のttafCodeIdx/descIdx/rmNameIdxにそのまま反映される。
Private Sub BuildTTAFCodeAndDescIndex(rmTbl As ListObject, ttafCodeIdx As Object, descIdx As Object, rmNameIdx As Object)
    rmNameIdx.CompareMode = vbTextCompare
    Dim rmN As Long: rmN = rmTbl.ListRows.Count
    If rmN > 0 Then
        Dim rmNameDesc As Variant
        rmNameDesc = rmTbl.ListColumns(1).DataBodyRange.Resize(rmN, 2).Value  ' Part Name, Description
        Dim rmTtafCode As Variant
        rmTtafCode = rmTbl.ListColumns(9).DataBodyRange.Value                 ' TTAF_Code
        Dim i As Long
        For i = 1 To rmN
            Dim partName As String: partName = Trim(CStr(rmNameDesc(i, 1)))
            Dim tKeyBuild As String: tKeyBuild = NormalizeText(CStr(rmTtafCode(i, 1)))
            If Len(tKeyBuild) > 0 And Not ttafCodeIdx.Exists(tKeyBuild) Then ttafCodeIdx(tKeyBuild) = partName
            Dim dKeyBuild As String: dKeyBuild = NormalizeText(CStr(rmNameDesc(i, 2)))
            If Len(dKeyBuild) > 0 And Not descIdx.Exists(dKeyBuild) Then descIdx(dKeyBuild) = partName
            If Len(partName) > 0 And Not rmNameIdx.Exists(partName) Then rmNameIdx(partName) = partName
        Next i
    End If

    ' 既知の表記ゆれの手動エイリアス(RefreshData_ShipmentsのBuildKnownAliasIndexと同じ考え方)。
    ' 「Stock invoiced to CSA」シートのTTAF PART NUMBER列(A列)は、ND TAC(CHEM-1280)だけ
    ' 他の材料と違い数値のTTAF_Code(83988002202、CSA ReportのShipping Scheduleと一致)ではなく
    ' 文字列"NDTAC"になっている(TTAF側の入力ゆれ)。現状はDescription("ND TAC")側の一致で
    ' 拾えているため実害は出ていないが、将来Description表記が変わった場合に備え、
    ' TTAF_Code側でも直接一致するようにしておく。
    Dim aliasKey As String: aliasKey = NormalizeText("NDTAC")
    If Not ttafCodeIdx.Exists(aliasKey) Then ttafCodeIdx(aliasKey) = "CHEM-1280"
End Sub

' RefreshSelfStockが使う。「CHEMICAL SOH」シートのB列(CSA/コード)がM_RawMaterialsの
' Part Nameと完全一致すればそのまま返す。見つからなければ"0"(ゼロ)/"O"(オー)の表記ゆれ
' 〈例:"0JN" vs "OJN"。daily check・CSA Report等TTAF側の複数ファイルで繰り返し見つかって
' いる〉を両方向で読み替えて試す。見つかった場合はM_RawMaterials側の正式な表記(Part Name)
' を返す(取込元ファイルの表記ゆれのまま書き込むと、既存のT_SelfStock_Log行や他シートの
' 正式表記と食い違い、グリッド表示側で該当材料として認識されなくなるため)。
' どれでも見つからなければ空文字を返す。
Private Function ResolveSelfStockCode(rmTbl As ListObject, rmCodeSet As Object, codeRaw As String) As String
    If rmCodeSet.Exists(codeRaw) Then
        ResolveSelfStockCode = Trim(CStr(rmTbl.ListRows(rmCodeSet(codeRaw)).Range.Cells(1, 1).Value))
        Exit Function
    End If
    Dim zeroToO As String: zeroToO = Replace(codeRaw, "0", "O")
    If zeroToO <> codeRaw And rmCodeSet.Exists(zeroToO) Then
        ResolveSelfStockCode = Trim(CStr(rmTbl.ListRows(rmCodeSet(zeroToO)).Range.Cells(1, 1).Value))
        Exit Function
    End If
    Dim oToZero As String: oToZero = Replace(codeRaw, "O", "0")
    If oToZero <> codeRaw And rmCodeSet.Exists(oToZero) Then
        ResolveSelfStockCode = Trim(CStr(rmTbl.ListRows(rmCodeSet(oToZero)).Range.Cells(1, 1).Value))
        Exit Function
    End If
    ResolveSelfStockCode = ""
End Function

' TTAF_Codeでの照合を優先し、見つからなければPart No(C列。M_RawMaterialsのPart Nameが
' そのまま入っていることが多く、Descriptionより確実)で照合し、それでも見つからない場合だけ
' Description(材料名)の正規化テキストで照合する。どれでも見つからなければ空文字を返す。
Private Function ResolveTTAFPart(ttafCodeIdx As Object, rmNameIdx As Object, descIdx As Object, _
        ttafCodeRaw As String, partNoRaw As String, descRaw As String) As String
    Dim tKey As String: tKey = NormalizeText(ttafCodeRaw)
    If Len(tKey) > 0 And ttafCodeIdx.Exists(tKey) Then
        ResolveTTAFPart = ttafCodeIdx(tKey)
        Exit Function
    End If
    If Len(partNoRaw) > 0 And rmNameIdx.Exists(partNoRaw) Then
        ResolveTTAFPart = rmNameIdx(partNoRaw)
        Exit Function
    End If
    ' TTAF側の元データで"0"(ゼロ)と"O"(アルファベットのオー)の表記ゆれが度々見つかっている
    ' (例: CSA ReportのPart No列で"0JN"、M_RawMaterials側は正式に"OJN")。完全一致で
    ' 見つからない場合、どちらの表記に読み替えても一致するか試す。
    If Len(partNoRaw) > 0 Then
        Dim zeroToO As String: zeroToO = Replace(partNoRaw, "0", "O")
        If zeroToO <> partNoRaw And rmNameIdx.Exists(zeroToO) Then
            ResolveTTAFPart = rmNameIdx(zeroToO)
            Exit Function
        End If
        Dim oToZero As String: oToZero = Replace(partNoRaw, "O", "0")
        If oToZero <> partNoRaw And rmNameIdx.Exists(oToZero) Then
            ResolveTTAFPart = rmNameIdx(oToZero)
            Exit Function
        End If
    End If
    Dim dKey As String: dKey = NormalizeText(descRaw)
    If descIdx.Exists(dKey) Then
        ResolveTTAFPart = descIdx(dKey)
        Exit Function
    End If
    ResolveTTAFPart = ""
End Function

' F4セルは「TTAF count(dd.mm.yyyy)」という見出し文字列だが、これがセルの表示形式
' (カスタム書式)による見た目だけの場合、実際の.Valueは本物の日付シリアル値になっている
' ことがある。その場合CStr()で文字列化すると地域設定依存の別形式(例: 6/29/2026)に
' なってしまい、DD.MM.YYYYの正規表現に一致しない。そのため、まずIsDate()で本物の日付値
' かどうかを確認し、日付値ならそのまま使う(文字列化を経由しない)。
' 文字列として読む場合も、全角数字(２９.０６.２０２６)で入力されているケースに備えて
' StrConv(..., vbNarrow)で半角に正規化してからマッチさせる。
Private Function ExtractDDMMYYYYFromText(cellValue As Variant) As Date
    If IsDate(cellValue) Then
        ExtractDDMMYYYYFromText = CDate(cellValue)
        Exit Function
    End If
    Dim text As String: text = StrConv(CStr(cellValue), vbNarrow)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d{1,2})\.(\d{1,2})\.(\d{4})"
    Dim m As Object
    Set m = re.Execute(text)
    If m.Count = 0 Then
        Err.Raise vbObjectError + 1, , "TTAF countの日付(DD.MM.YYYY)を読み取れませんでした: " & CStr(cellValue)
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
