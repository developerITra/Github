VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queSCRABK"
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
Dim GroupCode As String, DocType As String, rstDoc As Recordset, DocDateInput As String, DocDate As Date, SCRAID As String

On Error GoTo Err_cmdAddDoc_Click

AddToList (lstFiles)
DoCmd.OpenForm "SCRA Search Info"
Forms![SCRA Search Info]!FileNumber = lstFiles

'added 2/9/15
            
strStage = Trim(Me.lstFiles.Column(3))
Call SCRAnames(lstFiles)

SCRAID = DLookup("SCRAstageID", "qryqueueSCRAbk", "file=" & lstFiles)
OpenCase lstFiles
Select Case SCRAID
Case 100
SCRAID = 14
Case 110
SCRAID = 15
Case 120
SCRAID = 16
End Select

Forms![Case List]!SCRAID = SCRAID
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

Private Sub Form_Current()
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueSCRABK", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
