VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queSCRA9"
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



Private Sub cmdUploadSCRA_Click()

    


Dim Filespec As String, fileextension As String, Path As String, FileName As String, newfilename As String, i As Integer, Prompt As String
Dim GroupCode As String, DocType As String, rstDoc As Recordset, DocDateInput As String, DocDate As Date, rstqueue As Recordset, Reason As String

On Error GoTo Err_cmdAddDoc_Click


DoCmd.OpenForm "SCRA Search Info"
Forms![SCRA Search Info]!FileNumber = lstFiles
Call SCRAnames(lstFiles)

OpenCase lstFiles
Set rstqueue = CurrentDb.OpenRecordset("select * from qryqueueSCRA9 where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
Reason = rstqueue!Reason
Select Case Reason
Case "Referral"
Forms![Case List]!SCRAID = 10
Case "Hearing"
Forms![Case List]!SCRAID = 11
Case "Lockout"
Forms![Case List]!SCRAID = 12
End Select
Set rstqueue = Nothing

Forms![Case List]!Page97.SetFocus
Forms![SCRA Search Info].SetFocus


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
Private Sub cmdUploadPACER_Click()
Dim Filespec As String, fileextension As String, Path As String, FileName As String, newfilename As String, i As Integer, Prompt As String
Dim GroupCode As String, DocType As String, rstDoc As Recordset, DocDateInput As String, DocDate As Date, rstqueue As Recordset, Reason As String

On Error GoTo Err_cmdAddDoc_Click


DoCmd.OpenForm "SCRA Search Info"
Forms![SCRA Search Info]!FileNumber = lstFiles
Call SCRAnames(lstFiles)

OpenCase lstFiles
Set rstqueue = CurrentDb.OpenRecordset("select * from qryqueueSCRA9 where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
Reason = rstqueue!Reason
Select Case Reason
Case "Referral"
Forms![Case List]!SCRAID = 10
Case "Hearing"
Forms![Case List]!SCRAID = 11
Case "Lockout"
Forms![Case List]!SCRAID = 12
End Select
Set rstqueue = Nothing

Forms![Case List]!Page97.SetFocus
Forms![SCRA Search Info].SetFocus


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

Private Sub cmdWaiting_Click()
Dim rstqueue As Recordset
DoCmd.OpenForm "EnterEVSCRAReason"
Forms!enterevscrareason!FileNumber = lstFiles



End Sub

Private Sub Form_Current()
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueSCRA9", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
