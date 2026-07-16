Attribute VB_Name = "JumpToWeek"
Option Explicit

' ============================================================================
' JumpToWeek モジュール（任意）
'
' Dashboard の C1 に週(例: "2026-W23")を入力した状態でこのマクロを実行すると、
' 該当する週の列が画面に見える位置までスクロールします。
' 条件付き書式による黄色ハイライトと合わせて使うと、目的の週を見つけやすくなります。
'
' 【注意】この環境ではVBAの実行検証ができないため未検証です。
'
' 【使い方】
'   Alt+F8 → GoToWeek を選択して実行、または図形にマクロ登録してボタン化
' ============================================================================

Sub GoToWeek()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    Dim target As String
    target = Trim(ws.Range("C1").Value)
    If Len(target) = 0 Then
        MsgBox "Dashboard の C1 に、ジャンプ先の週(例: 2026-W23)を入力してください。", vbExclamation
        Exit Sub
    End If

    Dim headerRow As Long: headerRow = 6  ' HDR_TABLE_ROW
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim foundCell As Range
    Set foundCell = ws.Range(ws.Cells(headerRow, 6), ws.Cells(headerRow, lastCol)).Find(What:=target, LookAt:=xlWhole)

    If foundCell Is Nothing Then
        MsgBox "'" & target & "' が見つかりませんでした。'2026-W23' のような形式で入力してください。", vbExclamation
        Exit Sub
    End If

    ws.Activate
    Application.Goto ws.Cells(7, foundCell.Column), True
    ActiveWindow.ScrollColumn = Application.Max(1, foundCell.Column - 1)
End Sub
