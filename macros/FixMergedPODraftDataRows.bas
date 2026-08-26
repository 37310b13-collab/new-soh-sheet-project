Attribute VB_Name = "FixMergedPODraftDataRows"
Option Explicit

' ============================================================================
' FixMergedPODraftDataRows モジュール(一度きりの緊急復旧用)
'
' 背景: AppendPODraftRow(RefreshData_MaterialMgmt.bas)の不具合により、空の
' PO_Draft_*シートに1件目の材料行を追加する際、書式コピー元がヘッダー行(26行目。
' H26:K26/L26:T26が"Firm"/"Forecast"ラベル用に結合されている)になってしまい、
' その結合状態までコピーされていた。結合により、週ごとに個別で書き込んだはずの
' 発注数量の数式が「最後の1つだけ残して上書き」される形で失われ、さらに2件目以降も
' 直前行(=結合された1件目)から書式をコピーするため、結合が全行へ伝播していた。
' (この不具合の原因自体はRefreshData_MaterialMgmt.basのAppendPODraftRowで修正済み。
' このモジュールは、修正前の状態で既に結合・破損してしまった既存行を直すためのもの)
'
' 【使い方】標準モジュールとして貼り付け、FixMergedPODraftDataRowsを一度だけ実行する。
' PO_Draft_Chemical・PO_Draft_Hazardous・PO_Draft_Substrate_JPN_CHN・
' PO_Draft_Substrate_Polandの全データ行を対象に、結合を解除した上で、
' 発注数量の数式(Material_DetailのOrder行を週ごとに参照する式)とFirm/Forecastの
' 色分けを、AppendPODraftRowと同じパターンで書き直す。既に正常な行も同じ内容で
' 上書きされるだけなので、誤って複数回実行しても安全。
' ============================================================================

Sub FixMergedPODraftDataRows()
    On Error GoTo ErrHandler
    Dim thisWb As Workbook: Set thisWb = ThisWorkbook
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim nWeeks As Long: nWeeks = thisWb.Sheets("Cal_Weeks").ListObjects("Cal_Weeks").ListRows.Count

    Const MD_WEEK_START_COL As Long = 4  ' Material_Detail: 週データ開始列(D列)
    Const PO_FIRST_WEEK_COL As Long = 8
    Const PO_N_WEEKS As Long = 13
    Const PO_DATA_START_ROW As Long = 27

    Dim mdOrderHelperCol As Long: mdOrderHelperCol = MD_WEEK_START_COL + nWeeks + 1
    Dim mdOrderHelperColLetter As String: mdOrderHelperColLetter = ColLetterLocal2(mdOrderHelperCol)
    Dim mdWeekFirstColLetter As String: mdWeekFirstColLetter = ColLetterLocal2(MD_WEEK_START_COL)
    Dim mdWeekLastColLetter As String: mdWeekLastColLetter = ColLetterLocal2(MD_WEEK_START_COL + nWeeks - 1)

    Dim sheetNames As Variant
    sheetNames = Array("PO_Draft_Chemical", "PO_Draft_Hazardous", "PO_Draft_Substrate_JPN_CHN", "PO_Draft_Substrate_Poland")

    Dim si As Long, totalRows As Long, totalUnmerged As Long
    totalRows = 0: totalUnmerged = 0
    For si = LBound(sheetNames) To UBound(sheetNames)
        Dim sh As Worksheet
        On Error Resume Next
        Set sh = Nothing
        Set sh = thisWb.Sheets(CStr(sheetNames(si)))
        On Error GoTo ErrHandler
        If Not sh Is Nothing Then
            Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.Count, 5).End(xlUp).Row
            Dim bwRefExpr As String: bwRefExpr = BaseWeekRefLocal2(sh)
            Dim r As Long
            For r = PO_DATA_START_ROW To lastRow
                Dim rmCode As String: rmCode = Trim(CStr(sh.Cells(r, 4).Value))
                If Len(rmCode) > 0 Then
                    Dim wkRng As Range
                    Set wkRng = sh.Range(sh.Cells(r, PO_FIRST_WEEK_COL), sh.Cells(r, PO_FIRST_WEEK_COL + PO_N_WEEKS - 1))
                    If wkRng.MergeCells Then
                        wkRng.UnMerge
                        totalUnmerged = totalUnmerged + 1
                    End If

                    Dim mdOrderMatch As String
                    mdOrderMatch = "MATCH($D" & r & ",Material_Detail!$" & mdOrderHelperColLetter & ":$" & mdOrderHelperColLetter & ",0)"
                    Dim w As Long, col As Long
                    For w = 1 To PO_N_WEEKS
                        col = PO_FIRST_WEEK_COL + w - 1
                        sh.Cells(r, col).Value = "=IFERROR(INDEX(Material_Detail!$" & mdWeekFirstColLetter & ":$" & mdWeekLastColLetter & _
                            "," & mdOrderMatch & "," & bwRefExpr & "+" & (w - 1) & "),0)"
                        With sh.Cells(r, col)
                            If w <= 4 Then
                                .Interior.Color = RGB(255, 193, 193)  ' Firm: FFC1C1
                                .Font.Color = RGB(192, 0, 0)          ' 濃い赤: C00000
                            Else
                                .Interior.Color = RGB(235, 241, 222)  ' Forecast: EBF1DE
                                .Font.Color = RGB(0, 97, 0)           ' 濃い緑: 006100
                            End If
                        End With
                    Next w
                    totalRows = totalRows + 1
                End If
            Next r
        End If
    Next si

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Merged blocks unmerged: " & totalUnmerged & vbCrLf & _
           "Rows regenerated: " & totalRows & vbCrLf & vbCrLf & _
           "Please check that week columns show independent values (not merged), then save.", vbInformation
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: (" & Err.Number & ") " & Err.Description & vbCrLf & _
           "Close without saving to revert if unsure, and report the error text.", vbCritical
End Sub

' 名前付き範囲"BaseWeek"があればその名前を、無ければ"$P$7"を返す
' (RefreshData_UtilitiesのBaseWeekRefと同じロジック。このモジュール単体でも
' 動作するよう複製している)。
Private Function BaseWeekRefLocal2(sh As Worksheet) As String
    On Error Resume Next
    Dim nm As Name: Set nm = sh.Names("BaseWeek")
    On Error GoTo 0
    If Not nm Is Nothing Then
        BaseWeekRefLocal2 = "BaseWeek"
    Else
        BaseWeekRefLocal2 = "$P$7"
    End If
End Function

Private Function ColLetterLocal2(colNum As Long) As String
    Dim s As String, n As Long, r As Long
    n = colNum
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr(65 + r) & s
        n = (n - r - 1) \ 26
    Loop
    ColLetterLocal2 = s
End Function
