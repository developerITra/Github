VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queAttyMilestone2Mgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCorrections_Click()
Dim rstqueue As Recordset, rstJnl As Recordset
Dim cntr As Integer
Dim mgrInput As String

If Me.lstFiles.Column(7) <> "R " Then

            
            ' 2012.01.31:  Do not allow the input to be left empty.
           
            mgrInput = InputBox("Enter Description of Corrections Made", "Manager Review")
            If 0 = Len(mgrInput) _
            Then
                MsgBox ("Change will not be processed unless Description of Corrections is supplied.")
                Exit Sub
            End If
            
            Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)
            
            With rstqueue
            .Edit
            !AttyMilestone2Reject = False
            !AttyMilestone2 = Null
            .Update
            End With
            
            Me.Refresh
            Set rstqueue = Nothing
            Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttyMilestone2Mgr", dbOpenDynaset, dbSeeChanges)
            Do Until rstqueue.EOF
            cntr = cntr + 1
            rstqueue.MoveNext
            Loop
            QueueCount = cntr
            Set rstqueue = Nothing
            
            'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
            
            '2/11/14
            'lisa
                DoCmd.SetWarnings False
                strinfo = "Intake Review- Manager corrections made:  " & mgrInput
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone2Mgr!lstFiles,Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True

Else
           ' 2012.01.31:  Do not allow the input to be left empty.
            
            mgrInput = InputBox("Enter Description of Corrections Made", "Manager Review")
            If 0 = Len(mgrInput) _
            Then
                MsgBox ("Change will not be processed unless Description of Corrections is supplied.")
                Exit Sub
            End If
            
            Set rstqueue = CurrentDb.OpenRecordset("Select * FROM WizardSupportTwo where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)
            
            With rstqueue
            .Edit
            !AttyMilestonerestartReject = False
            !AttyMilestoneRestart = Null
            .Update
            End With
            
            Me.Refresh
            Set rstqueue = Nothing
            Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttyMilestone2Mgr", dbOpenDynaset, dbSeeChanges)
            Do Until rstqueue.EOF
            cntr = cntr + 1
            rstqueue.MoveNext
            Loop
            QueueCount = cntr
            Set rstqueue = Nothing
            
            'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
            
            '2/11/14
            'lisa
                DoCmd.SetWarnings False
                strinfo = "Restart Review- Manager corrections made:  " & mgrInput
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone2Mgr!lstFiles,Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                
End If



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


Private Sub lstFiles_DblClick(Cancel As Integer)

    OpenCase lstFiles
    Dim rstqueue As Recordset, rstJnl As Recordset

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!AttyMilestoneMgr2 = Date
!AttyMilestone2Approver = GetStaffID
.Update
End With

Me.Refresh
Set rstqueue = Nothing

End Sub
Private Sub Form_Current()
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttyMilestone2Mgr", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
