VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queVASaleSettingWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdExcel_Click()
Call DoReport("qryqueuevasalesettingwaiting", -3)
End Sub

Private Sub cmdOK_Click()
Dim rstqueue As Recordset
On Error GoTo Err_cmdOK_Click
        
VAsalesettingCallFromQueue lstFiles
Forms!foreclosuredetails!cmdWaiting1.Visible = True

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!VASaleSettingWaitingUser = StaffID
!VASaleSettingWaitinglastedited = Date
.Update
End With

Set rstqueue = Nothing

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
Me!lstfilesRed.Requery
Me.Requery
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuevasalesettingwaitinglst", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
End Sub

Private Sub cmdViewItems_Click()
DoCmd.OpenForm "MissingDocsListvasalesettingviewonly"
Forms!MissingDocsListvasalesettingviewonly!FileNbr = lstFiles
End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)

Dim rstqueue As Recordset
AddToList (lstFiles)
VAsalesettingCallFromQueue lstFiles
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!VASaleSettingWaitingUser = StaffID
!VASaleSettingWaitinglastedited = Date
.Update
End With
Set rstqueue = Nothing
End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuevasalesettingwaitinglst", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub

Private Sub lstFiles_GotFocus()
cmdViewItems.Enabled = True
End Sub


Private Sub lstFilesred_DblClick(Cancel As Integer)
Dim rstqueue As Recordset

AddToList (lstfilesRed)
VAsalesettingCallFromQueue lstfilesRed
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstfilesRed & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!VASaleSettingWaitingUser = StaffID
!VASaleSettingWaitinglastedited = Date
.Update
End With
Set rstqueue = Nothing
End Sub
