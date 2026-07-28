Attribute VB_Name = "RefreshData_Display"
Option Explicit

' ============================================================================
' RefreshData_Display モジュール
'
' HideInactiveIntermediates / ShowAllIntermediates / JumpToSelectedWeek
'
' いずれもデータを更新するマクロではなく、見た目(行の表示/非表示、ウィンドウの
' スクロール位置)だけを操作するものです。
'
' Material_Detailシートは、材料(RM_Code)ごとに「中間体名／No. of batches」
' 「使用量(kg)」の行が縦に並ぶブロック構成になっている。将来しばらく(半年・1年など)
' 生産予定の無い中間体の行を、ユーザー指定期間に応じてボタン一つで折りたたみ、
' 見た目のスクロール量を減らせるようにする。材料の合計使用量・TTAF在庫実績・
' 自社在庫実績・合計在庫(週末時点)の行は、非表示の対象にせず常に見える状態を保つ
' （在庫量そのものは常に把握できるようにするため）。
'
' 判定方法: 「今週」から指定月数分の週について、その中間体の「No. of batches」行が
' 全て0(または空白)であれば、その中間体の2行(No. of batches行＋使用量(kg)行)を隠す。
' 1つでも0以外の週があれば表示したままにする。実行のたびにまず全行を再表示してから
' 判定し直すため、期間を変えて何度実行しても結果は指定した条件どおりになる。
'
' 行位置はラベル文字列(列C="No. of batches")で判定しており、build_soh.py側で
' 行数が増減しても自動的に追従する。ただし週データの開始列(MD_WEEK_START_COL=E列)
' とヘッダー行(MD_HEADER_ROW=6行目、RefreshData_Utilitiesで定義)は、build_soh.pyの
' WEEK_START_COL/MD_TABLE_ROWと値を合わせる必要がある(シート構成を変更した場合は
' ここも合わせて変更すること)。
'
' 【ボタンの割り当て方(手動での一度だけの作業。openpyxlではボタンを自動作成できないため)】
'   1. Material_Detailシートを開く
'   2. 「挿入」タブ →「図形」等で好きな形の図形を1~2個描く(例:「非表示にする」「全部表示」)
'   3. 図形を右クリック →「マクロの登録」→ HideInactiveIntermediates (もう1つには
'      ShowAllIntermediates)を選択
'   4. お好みでシート上部の空いている場所(A1付近など)に配置する
'
' 【JumpToSelectedWeekの導入方法(任意)】標準モジュールへのインポートだけでは動きません。
' Dashboard/Material_Detail/T_SelfStock/T_TTAFStock各シート自身のコードモジュールに
' Worksheet_Changeを設置する必要があります。詳細はRefreshData_Utilitiesモジュール冒頭の
' コメントを参照してください。
' ============================================================================

Sub HideInactiveIntermediates()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets("Material_Detail")
    On Error GoTo 0
    If sh Is Nothing Then
        MsgBox "Material_Detailシートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim monthsStr As String
    monthsStr = InputBox("今週から何ヶ月間、全週バッチ数0の中間体を非表示にしますか？" & _
        vbCrLf & "（例: 6 → 半年間生産予定の無い中間体を非表示、12 → 1年間）" & _
        vbCrLf & "半角数字で入力してください。", "非表示にする期間(月数)", "6")
    If monthsStr = "" Then Exit Sub  ' キャンセル
    If Not IsNumeric(monthsStr) Then
        MsgBox "数値を入力してください。", vbExclamation
        Exit Sub
    End If
    Dim months As Double: months = CDbl(monthsStr)
    If months <= 0 Then
        MsgBox "1以上の数値を入力してください。", vbExclamation
        Exit Sub
    End If

    Dim nWeeks As Long
    nWeeks = wb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count

    Dim curWeek As Long: curWeek = WeekIndexForDate(wb, Date)
    Dim thresholdWeeks As Long
    thresholdWeeks = CLng(Application.WorksheetFunction.RoundUp(months * 52 / 12, 0))

    Dim endWeek As Long: endWeek = curWeek + thresholdWeeks - 1
    If endWeek > nWeeks Then endWeek = nWeeks
    Dim startCol As Long: startCol = MD_WEEK_START_COL + curWeek - 1
    Dim endCol As Long: endCol = MD_WEEK_START_COL + endWeek - 1
    If endCol < startCol Then
        MsgBox "対象週が見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row

    ' まず全行を再表示してから判定する(期間を変えて再実行しても常に指定条件どおりになるように)
    If lastRow >= MD_HEADER_ROW Then sh.Rows(MD_HEADER_ROW & ":" & lastRow).Hidden = False

    Dim r As Long, hiddenCount As Long, shownCount As Long
    r = MD_HEADER_ROW + 1
    Do While r <= lastRow
        If sh.Cells(r, 3).Value = "No. of batches" Then
            Dim vals As Variant
            vals = sh.Range(sh.Cells(r, startCol), sh.Cells(r, endCol)).Value
            Dim allZero As Boolean: allZero = True
            Dim c As Long
            For c = 1 To UBound(vals, 2)
                If IsNumeric(vals(1, c)) Then
                    If CDbl(vals(1, c)) <> 0 Then
                        allZero = False
                        Exit For
                    End If
                End If
            Next c
            If allZero Then
                sh.Rows(r).Hidden = True
                sh.Rows(r + 1).Hidden = True   ' 対応する「使用量(kg)」行も一緒に隠す
                hiddenCount = hiddenCount + 1
            Else
                shownCount = shownCount + 1
            End If
            r = r + 2   ' No. of batches行＋使用量(kg)行のペア分をスキップ
        Else
            r = r + 1
        End If
    Loop

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "完了しました。" & vbCrLf & _
           "今週(week " & curWeek & ")から" & thresholdWeeks & "週間、全週バッチ数0の中間体を非表示にしました。" & vbCrLf & _
           "非表示: " & hiddenCount & " 件 / 表示中: " & shownCount & " 件" & vbCrLf & _
           "（材料名の行・在庫関連の行は常に表示されます）", vbInformation
End Sub

' Dashboard/Material_Detail/T_SelfStock/T_TTAFStockのC1(選択週)を入力したときに呼ばれる
' 想定の共通処理。選択週の値を別セルに複製する(ピン留め列)方式は廃止し、代わりに
' 「本物の週データ列」が常にラベル列のすぐ右(固定ペインの直後)に見えるよう、ウィンドウを
' 横スクロールするだけにしている。値の複製が一切ないため、各シート・Grid_Stock等の間で
' 数字が食い違う余地がない。
' 呼び出し元のシート自身のWorksheet_Changeから、対象シート・週No解決済みセル(F1)・
' 週データ開始列(Dashboardは9=I列、Material_Detailは4=D列、T_SelfStock/T_TTAFStockは
' 2=B列)を渡して呼び出す。
Public Sub JumpToSelectedWeek(sh As Worksheet, weekIndexCell As String, weekStartCol As Long)
    Dim wIdx As Variant
    wIdx = sh.Range(weekIndexCell).Value
    If Not IsNumeric(wIdx) Then Exit Sub
    Dim targetCol As Long
    targetCol = weekStartCol + CLng(wIdx) - 1
    On Error Resume Next
    ActiveWindow.ScrollColumn = targetCol
    On Error GoTo 0
End Sub

Sub ShowAllIntermediates()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets("Material_Detail")
    On Error GoTo 0
    If sh Is Nothing Then
        MsgBox "Material_Detailシートが見つかりません。", vbExclamation
        Exit Sub
    End If
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row
    If lastRow >= MD_HEADER_ROW Then sh.Rows(MD_HEADER_ROW & ":" & lastRow).Hidden = False
    MsgBox "すべての中間体行を再表示しました。", vbInformation
End Sub
