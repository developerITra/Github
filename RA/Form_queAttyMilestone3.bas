VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queAttyMilestone3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCorrections_Click()
Dim rstqueue As Recordset, rstJnl As Recordset, jnltxt As String

jnltxt = InputBox("Enter Brief Reason Approval is Pending (30 character max)", "Attorney Review")
If jnltxt = "" Then
MsgBox "Please enter a description", vbCritical
Exit Sub
End If

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!AttyMilestone3Reject = False
!AttyMilestone3 = Now
!AttyMilestone3Approver = GetStaffID
!VASaleSettingAttyReview = Null
!VASaleSettingReason = 2
!VASaleSettingComment = jnltxt
.Update
End With

AddStatus lstFiles, Date, "Sale Setting Review by Attorney Completed"

Me.Refresh
Set rstqueue = Nothing
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttymilestone3", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing


    DoCmd.SetWarnings False
    strinfo = "VA Sale Setting Review- Approved with exceptions.  Reason is:  " & jnltxt
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone3!lstFiles,Now,GetFullName(),'" & strinfo & "',2)"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    
Forms!Journal.Requery
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
''Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = lstFiles
'!JournalDate = Now
'!Who = GetFullName
'!Info = "VA Sale Setting Review- Approved with exceptions.  Reason is:  " & jnltxt
'' 2012.02.28 DaveW - Color to red
'!Color = 2
'.Update
'End With
'Set rstJnl = Nothing




End Sub

Private Sub cmdReject_Click()
Dim rstqueue As Recordset, rstJnl As Recordset, cntr As Integer
On Error GoTo Err_cmdOK_Click


Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!AttyMilestone3Reject = True
!AttyMilestone3 = Now
!AttyMilestone3Approver = GetStaffID
!VASaleSettingAttyReview = Null
!VASaleSettingReason = 5

End With


'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = lstFiles
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Sale Setting Review- Rejected because " & InputBox("Enter Reason for Rejection", "Attorney Review")
'' 2012.04.16 PJF - Color to  blue
'!Color = 4
'.Update
'End With
'Set rstJnl = Nothing

DoCmd.OpenForm "EnterVASalesettingDocs"
Forms!enterVAsalesettingdocs!FileNumber = lstFiles

Dim rstdocs As Recordset
Set rstdocs = CurrentDb.OpenRecordset("select * from VAsalesettingdocsneeded where docreceived is null AND filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
With rstdocs
Do Until .EOF
If Not IsNull(!DocName) Then
Select Case !DocName
Case "Note"
Forms!enterVAsalesettingdocs!btn9 = True
Case "SOT"
Forms!enterVAsalesettingdocs!btn10 = True
Case "FD"
Forms!enterVAsalesettingdocs!btn11 = True
Case "Demand"
Forms!enterVAsalesettingdocs!btn12 = True
Case "LNA/Notice"
Forms!enterVAsalesettingdocs!btn13 = True
Case "Assignment"
Forms!enterVAsalesettingdocs!btn14 = True
Case "Other"
Forms!enterVAsalesettingdocs!btn16 = True
End Select
End If
.MoveNext
Loop
End With

rstqueue!VASaleSettingComment = InputBox("Enter Brief Reason for Rejection (30 character max)", "Attorney Review")
rstqueue.Update

Me.Refresh
Set rstqueue = Nothing
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttymilestone3", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
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

Private Sub cmdApprove_Click()

Dim rstqueue As Recordset, rstJnl As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!AttyMilestone3Reject = False
!AttyMilestone3 = Now
!AttyMilestone3Approver = GetStaffID
!VASaleSettingAttyReview = Null
!VASaleSettingReason = 1
.Update
End With
Set rstqueue = Nothing
AddStatus lstFiles, Date, "VA Sale Setting Review by Attorney Completed"


    DoCmd.SetWarnings False
    strinfo = "VA Sale Setting Review by Attorney Completed"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone3!lstFiles,Now,GetFullName(),'" & strinfo & "',1)"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    
Forms!Journal.Requery


'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = lstFiles
'!JournalDate = Now
'!Who = GetFullName
'!Info = "VA Sale Setting Review by Attorney Completed"
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing

Me.Refresh
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttymilestone3", dbOpenDynaset, dbSeeChanges)
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
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttymilestone3", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
