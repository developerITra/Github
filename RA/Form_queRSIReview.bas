VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queRSIReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
Dim rstqueue As Recordset
On Error GoTo Err_cmdOK_Click
        
Restart1CallFromQueue lstFiles

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click


DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdRefresh_Click()
Me!lstFiles.Requery
Me.Requery
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueRSIreview", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
End Sub

Private Sub ComConf_Click()

On Error GoTo Err_cmdOK_Click
If Forms!queRSIReview.lstFiles.Column(6) = "Potential Employee Conflict" Then
    DoCmd.OpenForm "Staffconflict", , , "fileNumber= " & Forms!queRSIReview.lstFiles.Column(0)

Else
    MsgBox ("This is not conflict file")
    Exit Sub
End If



Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox "Select File" 'Err.Description"
    Resume Exit_cmdOK_Click

End Sub

Private Sub Form_Current()
If conflicts Then ComConf.Visible = True

End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)
AddToList (lstFiles)

If Forms!queRSIReview.lstFiles.Column(6) = "Potential Employee Conflict" Then
    If conflicts Then
    EditFormRSI = True
    Restart1CallFromReviewMgrQueue lstFiles
    EditFormRSI = False
    Else
    Exit Sub
    End If
Else


EditFormRSI = True
Restart1CallFromReviewMgrQueue lstFiles
EditFormRSI = False

End If


End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueRSIreview", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
