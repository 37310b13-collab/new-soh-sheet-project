Attribute VB_Name = "PO_Export"
Option Explicit

' ============================================================================
' PO_Export module
'
' [Caution] The 4 entry points below (ExportChemical / ExportHazardous /
' ExportSubstrateJPNCHN / ExportSubstratePoland) are meant to be run from
' buttons placed on the Dashboard sheet, not typed macro names - assign
' each one to a button via "Assign Macro".
'
' [Behavior]
'   Duplicates the PO_Draft_* sheet into a separate workbook, converts formulas to
'   values (snapshotting the content at the time of issue), then saves it with a
'   filename that includes the date and revision number.
'   After saving, the Revision cell on the original sheet is automatically incremented by 1.
' ============================================================================

Sub ExportChemical()
    ExportPODraft "PO_Draft_Chemical", "Chemical_Release"
End Sub

Sub ExportHazardous()
    ExportPODraft "PO_Draft_Hazardous", "Hazardous_Chemical_Release"
End Sub

Sub ExportSubstrateJPNCHN()
    ExportPODraft "PO_Draft_Substrate_JPN_CHN", "Substrate_JPN_CHN_Release"
End Sub

Sub ExportSubstratePoland()
    ExportPODraft "PO_Draft_Substrate_Poland", "Substrate_Poland_Release"
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

    ' --- Read the revision number and increment it after issuing ---
    ' Use the named range "PORevision" if it exists (the current layout has the
    ' Revision cell at N11). If it doesn't exist (an old, unmigrated layout),
    ' fall back to P5 as before. Without this fallback, migrating to the new
    ' layout would have kept reading the always-blank P5 (treating it as
    ' revNo=0) and writing "1" into P5 on every issue, silently corrupting a
    ' blank letterhead cell.
    On Error Resume Next
    Set revCell = srcWs.Range("PORevision")
    On Error GoTo ErrHandler
    If revCell Is Nothing Then Set revCell = srcWs.Range("P5")
    If IsNumeric(revCell.Value) Then
        revNo = CLng(revCell.Value)
    Else
        revNo = 0
    End If

    ' --- Duplicate the sheet into a new workbook ---
    srcWs.Copy
    Set newWb = ActiveWorkbook

    ' --- Convert formulas to values (freeze the content as of the issue date) ---
    With newWb.Worksheets(1).UsedRange
        .Value = .Value
    End With

    ' --- Save location / filename ---
    saveFolder = ThisWorkbook.Path & Application.PathSeparator & "PO_Issued"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(saveFolder) Then
        fso.CreateFolder saveFolder
    End If

    fileName = fileLabel & "_" & Format(Date, "yyyymmdd") & "_Rev" & Format(revNo, "00") & ".xlsx"
    newWb.SaveAs fileName:=saveFolder & Application.PathSeparator & fileName, _
                 FileFormat:=xlOpenXMLWorkbook

    newWb.Close SaveChanges:=False

    ' --- Update the Revision on the original sheet ---
    revCell.Value = revNo + 1
    ThisWorkbook.Save

    MsgBox "Order form issued:" & vbCrLf & saveFolder & Application.PathSeparator & fileName, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred while issuing the order form: " & Err.Description, vbCritical
End Sub
