Attribute VB_Name = "PO_Export"
Option Explicit

' ============================================================================
' PO_Export モジュール
'
' 【導入方法】
'   1. SOH_Master.xlsm を「マクロを有効にする」で開く
'   2. Alt+F11 でVBEを開く → 「ファイル」→「ファイルのインポート」→ このファイル(PO_Export.bas)を選択
'      （コピー＆貼り付けで導入する場合は、1行目の Attribute VB_Name = "..." を必ず削除してから
'       貼り付けてください。この行は貼り付けでは使えず、含めるとコンパイルエラーになります）
'   3. Dashboardシート等にボタンを配置し、下記のExportChemical / ExportHazardous /
'      ExportSubstrate をボタンのマクロ登録先として割り当てる
'
' 【動作】
'   PO_Draft_* シートを別ブックとして複製し、数式を値に変換（発行時点でスナップ
'   ショット化）した上で、日付・Revision付きのファイル名で保存します。
'   保存後、元シートのRevisionセルを自動で+1します。
' ============================================================================

Sub ExportChemical()
    ExportPODraft "PO_Draft_Chemical", "Chemical_Release"
End Sub

Sub ExportHazardous()
    ExportPODraft "PO_Draft_Hazardous", "Hazardous_Chemical_Release"
End Sub

Sub ExportSubstrate()
    ExportPODraft "PO_Draft_Substrate", "Substrate_Release"
End Sub

Private Sub ExportPODraft(ByVal sheetName As String, ByVal fileLabel As String)
    Dim srcWs As Worksheet
    Dim newWb As Workbook
    Dim revCell As Range
    Dim revNo As Long
    Dim saveFolder As String
    Dim fileName As String
    Dim fso As Object

    On Error GoTo ErrHandler

    Set srcWs = ThisWorkbook.Worksheets(sheetName)

    ' --- Revision番号を読み取り、発行後にインクリメントする ---
    Set revCell = srcWs.Range("P5")
    If IsNumeric(revCell.Value) Then
        revNo = CLng(revCell.Value)
    Else
        revNo = 0
    End If

    ' --- シートを新規ブックへ複製 ---
    srcWs.Copy
    Set newWb = ActiveWorkbook

    ' --- 数式を値に変換（発行時点の内容を固定） ---
    With newWb.Worksheets(1).UsedRange
        .Value = .Value
    End With

    ' --- 保存先・ファイル名 ---
    saveFolder = ThisWorkbook.Path & Application.PathSeparator & "PO_Issued"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(saveFolder) Then
        fso.CreateFolder saveFolder
    End If

    fileName = fileLabel & "_" & Format(Date, "yyyymmdd") & "_Rev" & Format(revNo, "00") & ".xlsx"
    newWb.SaveAs fileName:=saveFolder & Application.PathSeparator & fileName, _
                 FileFormat:=xlOpenXMLWorkbook

    newWb.Close SaveChanges:=False

    ' --- 元シートのRevisionを更新 ---
    revCell.Value = revNo + 1
    ThisWorkbook.Save

    MsgBox "発注書を発行しました:" & vbCrLf & saveFolder & Application.PathSeparator & fileName, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "発行処理でエラーが発生しました: " & Err.Description, vbCritical
End Sub
