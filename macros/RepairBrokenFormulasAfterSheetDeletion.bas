Attribute VB_Name = "RepairBrokenFormulasAfterSheetDeletion"
Option Explicit

' ============================================================================
' RepairBrokenFormulasAfterSheetDeletion モジュール(一度きりの緊急復旧用)
'
' 背景: Grid_TheoreticalStock・T_SelfStock_Log・T_TTAFStock_Logの3シートが誤って
' 削除され、別ブック(バックアップ)からSheets.Copyで複製し直して復元した際に、
' 以下の問題が生じた。
'   ①複製されたシート内の「同じブック内の別シートへの参照」が、Excelの仕様により
'     「外部ブックへの参照([1]!...)」に変換されてしまっていた
'     (Grid_TheoreticalStock・T_SelfStock_Log・T_TTAFStock_Log自身の数式)
'   ②Dashboard・T_SelfStock・T_TTAFStockの数式は、3シートが削除された瞬間に
'     永久に#REF!へ変換されており、参照先シートを復活させただけでは直らない
'     (Excelは、消えたシートへの参照が含まれる数式を、消えた時点でテキストごと
'     #REF!に書き換えてしまうため。後から同じ名前のシートが復活しても、既に
'     #REF!化した数式は自動的には元に戻らない)
'   ③本件とは別に、以前から存在していた不具合(T_OpeningStockの列名参照)により、
'     Grid_Stockの週1列の一部が#REF!のままだった(FixOpeningStockColumnReferenceは
'     文字列置換のため、既に#REF!まで壊れてしまった箇所までは直せなかった)
'
' これらのシート(Grid_TheoreticalStock・T_SelfStock・T_TTAFStock・Dashboard・
' Grid_Stock)はすべて数式のみで構成されており、手入力の実データは一切含まれない
' (実データはT_Shipments・T_OpeningStock・T_StockCount・T_SelfStock_Log・
' T_TTAFStock_Log・M_RawMaterials・M_BOM等にあり、今回一切変更しない)。そのため、
' 数式を書き直しても実データが失われることはない。
'
' 【使い方】標準モジュールとして貼り付け、RepairBrokenFormulasAfterSheetDeletionを
' 一度だけ実行する。既に正しい数式のセルは同じ内容で上書きされるだけなので、
' 誤って複数回実行しても安全(ただし通常は1回で完了する)。
' ============================================================================

Sub RepairBrokenFormulasAfterSheetDeletion()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim nWeeks As Long: nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count

    Dim fixed1 As Long, fixed2 As Long, fixed3 As Long
    fixed1 = StripExternalLinkMarker(thisWb.Sheets("Grid_TheoreticalStock"))
    fixed2 = StripExternalLinkMarker(thisWb.Sheets("T_SelfStock_Log"))
    fixed3 = StripExternalLinkMarker(thisWb.Sheets("T_TTAFStock_Log"))

    Dim regenSS As Long, regenTTAF As Long, regenDash As Long
    regenSS = RegenerateStockLogGrid(thisWb.Sheets("T_SelfStock"), nWeeks, "T_SelfStock_Log", "Self_Qty")
    regenTTAF = RegenerateStockLogGrid(thisWb.Sheets("T_TTAFStock"), nWeeks, "T_TTAFStock_Log", "TTAF_Qty")
    regenDash = RegenerateDashboardFormulas(thisWb, nWeeks)

    Dim fixed4 As Long, fixed5 As Long
    fixed4 = FixWeek1OpeningStockRef(thisWb.Sheets("Grid_Stock"))
    fixed5 = FixWeek1OpeningStockRef(thisWb.Sheets("Grid_TheoreticalStock"))

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.CalculateFull

    MsgBox "修復が完了しました。" & vbCrLf & vbCrLf & _
           "①外部参照を修正したセル数:" & vbCrLf & _
           "  Grid_TheoreticalStock " & fixed1 & " / T_SelfStock_Log " & fixed2 & " / T_TTAFStock_Log " & fixed3 & vbCrLf & vbCrLf & _
           "②数式を再生成した行数:" & vbCrLf & _
           "  T_SelfStock " & regenSS & " / T_TTAFStock " & regenTTAF & " / Dashboard(材料ペア数) " & regenDash & vbCrLf & vbCrLf & _
           "③T_OpeningStock参照を修正したセル数:" & vbCrLf & _
           "  Grid_Stock " & fixed4 & " / Grid_TheoreticalStock " & fixed5 & vbCrLf & vbCrLf & _
           "保存する前に、Dashboard・T_SelfStock・T_TTAFStock・Grid_Stockに#REF!が" & vbCrLf & _
           "残っていないか確認してください。", vbInformation
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "修復中にエラーが発生しました: (" & Err.Number & ") " & Err.Description & vbCrLf & vbCrLf & _
           "途中まで反映されている可能性があります。保存せずに閉じて開き直せば、実行前の状態に戻せます。" & vbCrLf & _
           "エラー内容をそのまま報告してください。", vbCritical
End Sub

' シート内の全セルを走査し、数式に"[1]"(外部ブック参照マーカー)が含まれていれば、
' 同じブック内の参照に戻す(修正したセル数を返す)。
'   "[1]!TableName[Column]" 形式(構造化参照) -> "TableName[Column]"
'   "[1]SheetName!Cell"     形式(シート直接参照) -> "SheetName!Cell"
' の両方に対応するため、"[1]!"を先に、次に残った"[1]"を置換する(この順序が重要)。
Private Function StripExternalLinkMarker(sh As Worksheet) As Long
    Dim n As Long: n = 0
    Dim usedRng As Range: Set usedRng = sh.UsedRange
    Dim cell As Range
    For Each cell In usedRng
        If cell.HasFormula Then
            Dim f As String: f = cell.Formula
            If InStr(f, "[1]") > 0 Then
                f = Replace(f, "[1]!", "")
                f = Replace(f, "[1]", "")
                cell.Formula = f
                n = n + 1
            End If
        End If
    Next cell
    StripExternalLinkMarker = n
End Function

' T_SelfStock/T_TTAFStockの全データ行・全週列の数式を、RefreshData_MaterialMgmt.bas
' (InsertOrAppendStockGridRow)と同じパターンで再生成する(処理した行数を返す)。
Private Function RegenerateStockLogGrid(sh As Worksheet, nWeeks As Long, logTableName As String, qtyColName As String) As Long
    Const SS_TABLE_ROW As Long = 5
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, w As Long, col As Long, n As Long: n = 0
    For r = SS_TABLE_ROW + 1 To lastRow
        If Len(Trim(CStr(sh.Cells(r, 1).Value))) > 0 Then
            For w = 1 To nWeeks
                col = 1 + w
                sh.Cells(r, col).Value = _
                    "=IF(COUNTIFS(" & logTableName & "[Part Name],$A" & r & "," & logTableName & "[WeekIndex]," & w & ")=0,"""",SUMIFS(" & _
                    logTableName & "[" & qtyColName & "]," & logTableName & "[Part Name],$A" & r & "," & logTableName & "[WeekIndex]," & w & "))"
            Next w
            n = n + 1
        End If
    Next r
    RegenerateStockLogGrid = n
End Function

' Dashboardの全材料ペア(理論在庫/実在庫)の数式を、RefreshData_MaterialMgmt.bas
' (AppendDashboardRow)と同じパターンで再生成する(処理したペア数を返す)。
' M_RawMaterials・Grid_Stock・T_SelfStockは行の並び順が完全に一致している設計のため、
' 算術計算だけで対応行を求める。念のため、実際にその行の材料コードが一致しているか
' 毎回確認し、想定外の構造だった場合は処理を中断してエラーを報告する
' (誤った行に書き込んでしまうことを防ぐため)。
Private Function RegenerateDashboardFormulas(thisWb As Workbook, nWeeks As Long) As Long
    Const DASH_DATA_START_ROW As Long = 7
    Const SS_FIRST_DATA_ROW As Long = 6
    Dim dashSh As Worksheet: Set dashSh = thisWb.Sheets("Dashboard")
    Dim gsSh As Worksheet: Set gsSh = thisWb.Sheets("Grid_Stock")
    Dim ssSh As Worksheet: Set ssSh = thisWb.Sheets("T_SelfStock")
    Dim gsTbl As ListObject: Set gsTbl = gsSh.ListObjects("Grid_Stock")

    Dim lastRow As Long: lastRow = dashSh.Cells(dashSh.Rows.Count, 1).End(xlUp).Row
    Dim lastWeekColLetter As String: lastWeekColLetter = ColLetterLocal(1 + nWeeks)
    Dim gsFirstDataRow As Long: gsFirstDataRow = gsTbl.DataBodyRange.Row

    Dim pairCount As Long: pairCount = 0
    Dim r As Long: r = DASH_DATA_START_ROW
    Do While r <= lastRow
        Dim rmCode As String: rmCode = Trim(CStr(dashSh.Cells(r, 1).Value))
        If Len(rmCode) = 0 Then
            r = r + 1
        Else
            Dim theoRow As Long: theoRow = r
            Dim actualRow As Long: actualRow = r + 1
            Dim actualCode As String: actualCode = Trim(CStr(dashSh.Cells(actualRow, 1).Value))

            If actualCode <> rmCode Then
                Err.Raise vbObjectError + 1, , "Dashboard " & theoRow & "行目(理論在庫=" & rmCode & _
                    ")と" & actualRow & "行目(実在庫=" & actualCode & ")の材料コードが一致しません。" & _
                    "想定外の構造のため処理を中断しました。手動で確認してください。"
            End If

            Dim rowIdx As Long: rowIdx = (theoRow - DASH_DATA_START_ROW) \ 2
            Dim ssRow As Long: ssRow = SS_FIRST_DATA_ROW + rowIdx
            Dim grow As Long: grow = gsFirstDataRow + rowIdx

            Dim gsCode As String: gsCode = Trim(CStr(gsSh.Cells(grow, 1).Value))
            Dim ssCode As String: ssCode = Trim(CStr(ssSh.Cells(ssRow, 1).Value))
            If gsCode <> rmCode Or ssCode <> rmCode Then
                Err.Raise vbObjectError + 2, , "Dashboard " & theoRow & "行目(" & rmCode & ")に対応するはずの" & _
                    "Grid_Stock" & grow & "行目(" & gsCode & ")・T_SelfStock" & ssRow & "行目(" & ssCode & ")が" & _
                    "一致しません。想定外の構造のため処理を中断しました。手動で確認してください。"
            End If

            Dim ssSelfRng As String: ssSelfRng = "'T_SelfStock'!$B$" & ssRow & ":$" & lastWeekColLetter & "$" & ssRow
            Dim ssTTAFRng As String: ssTTAFRng = "'T_TTAFStock'!$B$" & ssRow & ":$" & lastWeekColLetter & "$" & ssRow
            Dim ssLabelRng As String: ssLabelRng = "'T_SelfStock'!$B$5:$" & lastWeekColLetter & "$5"
            Dim lastActualColIdx As String
            lastActualColIdx = "LOOKUP(2,1/(" & ssSelfRng & "<>""""),COLUMN(" & ssSelfRng & "))"
            Dim diffFormula As String
            diffFormula = "=IFERROR(INDEX('Grid_Stock'!$A" & grow & ":$" & lastWeekColLetter & grow & ",1," & lastActualColIdx & ")" & _
                "-INDEX('Grid_TheoreticalStock'!$A" & grow & ":$" & lastWeekColLetter & grow & ",1," & lastActualColIdx & "),0)"

            Dim idx As Long, rr As Long, rowsArr(1 To 2) As Long
            rowsArr(1) = theoRow: rowsArr(2) = actualRow
            For idx = 1 To 2
                rr = rowsArr(idx)
                dashSh.Cells(rr, 2).Value = "=IFERROR(INDEX(M_RawMaterials[Description],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),"""")"
                dashSh.Cells(rr, 3).Value = "=IFERROR(INDEX(M_RawMaterials[Category],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),"""")"
                dashSh.Cells(rr, 4).Value = "=IFERROR(INDEX(M_RawMaterials[基準在庫下限_要入力],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),0)"
                dashSh.Cells(rr, 5).Value = "=IFERROR(INDEX(M_RawMaterials[基準在庫上限_要入力],MATCH($A" & rr & ",M_RawMaterials[Part Name],0)),0)"
                dashSh.Cells(rr, 6).Value = "=IFERROR(LOOKUP(2,1/(" & ssSelfRng & "<>"""")," & ssSelfRng & "),"""")"
                dashSh.Cells(rr, 7).Value = "=IFERROR(LOOKUP(2,1/(" & ssTTAFRng & "<>"""")," & ssTTAFRng & "),"""")"
                dashSh.Cells(rr, 8).Value = "=IFERROR(LOOKUP(2,1/(" & ssSelfRng & "<>"""")," & ssLabelRng & "),"""")"
                dashSh.Cells(rr, 10).Value = diffFormula
            Next idx

            Dim w As Long, col As Long
            For w = 1 To nWeeks
                col = 10 + w
                dashSh.Cells(theoRow, col).Value = "='Grid_TheoreticalStock'!" & ColLetterLocal(1 + w) & grow
                dashSh.Cells(actualRow, col).Value = "='Grid_Stock'!" & ColLetterLocal(1 + w) & grow
            Next w

            pairCount = pairCount + 1
            r = r + 2
        End If
    Loop
    RegenerateDashboardFormulas = pairCount
End Function

' Grid_Stock/Grid_TheoreticalStockの週1列(2列目)にある、以前からの既知の不具合
' (T_OpeningStockの列名参照が完全に壊れて#REF!化したもの)を修正する。
' 週1列以外は対象にしない(他の週は前週セルを参照するだけの式で、この不具合の対象外)。
Private Function FixWeek1OpeningStockRef(sh As Worksheet) As Long
    Dim tbl As ListObject: Set tbl = sh.ListObjects(1)
    Dim dataRange As Range: Set dataRange = tbl.DataBodyRange
    If dataRange Is Nothing Then
        FixWeek1OpeningStockRef = 0
        Exit Function
    End If
    Dim col1 As Range: Set col1 = dataRange.Columns(2)
    Dim n As Long: n = 0
    Dim r As Long
    For r = 1 To col1.Rows.Count
        Dim f As String: f = col1.Cells(r, 1).Formula
        If InStr(f, "#REF!") > 0 Then
            f = Replace(f, "INDEX(#REF!,", "INDEX(T_OpeningStock[Opening_Qty_要入力],")
            col1.Cells(r, 1).Formula = f
            n = n + 1
        End If
    Next r
    FixWeek1OpeningStockRef = n
End Function

Private Function ColLetterLocal(colNum As Long) As String
    Dim s As String, n As Long, r As Long
    n = colNum
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr(65 + r) & s
        n = (n - r - 1) \ 26
    Loop
    ColLetterLocal = s
End Function
