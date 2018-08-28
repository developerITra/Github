VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queAttyMilestone1_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCorrections_Click()
Dim rstqueue As Recordset, rstJnl As Recordset

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!AttyMilestone1_5Reject = False
!AttyMilestone1_5 = Now
!AttyMilestone1_5Approver = GetStaffID
!DateSentAttyNOI = Null
.Update
End With

AddStatus lstFiles, Date, "NOI Review by Attorney Completed"

Me.Refresh
Set rstqueue = Nothing
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttyMilestone1_5", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)

    DoCmd.SetWarnings False
    strinfo = "NOI Review- Approved with the following corrections made:  " & InputBox("Enter Description of Corrections Made", "Attorney Review")
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone1_5!lstFiles,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True


'



End Sub

Private Sub cmdReject_Click()
Dim rstqueue As Recordset, rstJnl As Recordset, rstFCdetails As Recordset
On Error GoTo Err_cmdOK_Click


'Set rstFCdetails = CurrentDb.OpenRecordset("SELECT NOI FROM FCDetails WHERE FileNumber = " & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)
'With rstFCdetails
'.Edit
'!NOI = Null
'.Update
'End With

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!AttyMilestone1_5Reject = True
!AttyMilestoneMgr1_5 = Now
!AttyMilestone1_5Approver = GetStaffID
!DateSentAttyNOI = Null
!DateInQueueNOI = Null
!NOIBulkUpload = False
.Update
End With

Me.Refresh
Set rstqueue = Nothing
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttyMilestone1_5", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing




'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
    DoCmd.SetWarnings False
    strinfo = "NOI Review- Rejected because " & InputBox("Enter Reason for Rejection", "Attorney Review")
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone1_5!lstFiles,Now,GetFullName(),'" & strinfo & "',4 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = lstFiles
'!JournalDate = Now
'!Who = GetFullName
'!Info = "NOI Review- Rejected because " & InputBox("Enter Reason for Rejection", "Attorney Review")
'' 2012.02.28 DaveW - Color to red
'' 2012.04.16 PJF - Color to  blue
'!Color = 4
'.Update
'End With
'Set rstJnl = Nothing



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

Private Sub cmdApprove_Click()

Dim rstqueue As Recordset, rstJnl As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!AttyMilestone1_5Reject = False
!AttyMilestone1_5 = Now
!AttyMilestone1_5Approver = GetStaffID
!DateSentAttyNOI = Null
.Update
End With

AddStatus lstFiles, Date, "NOI Review by Attorney Completed"

'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)

    DoCmd.SetWarnings False
    strinfo = "NOI Review by Attorney Completed"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone1_5!lstFiles,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True




Set rstqueue = Nothing
Me.Refresh
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttyMilestone1_5", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)

    OpenCase lstFiles
    
End Sub
Private Sub Form_Current()
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttyMilestone1_5", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
