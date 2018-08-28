VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queVAappraisal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click


DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub



Private Sub cmdUpload_Click()
Dim Filespec As String, fileextension As String, Path As String, FileName As String, newfilename As String, i As Integer, Prompt As String
Dim GroupCode As String, DocType As String, rstDoc As Recordset, DocDateInput As String, DocDate As Date

On Error GoTo Err_cmdAddDoc_Click

DoCmd.OpenForm "VAappraisal Search Info"
Forms![vaappraisal Search Info]!FileNumber = lstFiles

OpenCase lstFiles
Dim stDocName As String
Dim stLinkCriteria As String
Dim Details As Recordset

stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & Me![lstFiles] & " AND Current = True"
DoCmd.OpenForm stDocName, , , stLinkCriteria
Call VAnames(lstFiles)
Call VAappraisalCallFromQueue(lstFiles)
Forms![vaappraisal Search Info].SetFocus
Forms![Case List]!Page97.SetFocus



Exit_cmdAddDoc_Click:
    Exit Sub

Err_cmdAddDoc_Click:
    If Err.Number = 76 Then     ' path not found
        MkDir DocLocation & DocBucket(lstFiles) & "\" & lstFiles & "\"
        Resume
    Else
        MsgBox Err.Description
        Resume Exit_cmdAddDoc_Click
    End If
End Sub

Private Sub Form_Current()
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueVAappraisal", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
