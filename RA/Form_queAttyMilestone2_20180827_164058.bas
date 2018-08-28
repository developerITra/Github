VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queAttyMilestone2"
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
!AttyMilestone2Reject = False
!AttyMilestone2 = Now
!AttyMilestone2Approver = GetStaffID
.Update
End With

AddStatus lstFiles, Date, "Intake Review by Attorney Completed"

Me.Refresh
Set rstqueue = Nothing
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttymilestone2", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)

    DoCmd.SetWarnings False
    strinfo = "Intake Review- Approved with the following corrections made:  " & InputBox("Enter Description of Corrections Made", "Attorney Review")
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone2!lstFiles,Now,GetFullName(),'" & strinfo & "',2)"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    

'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = lstFiles
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Intake Review- Approved with the following corrections made:  " & InputBox("Enter Description of Corrections Made", "Attorney Review")
'' 2012.02.28 DaveW - Color to red
'!Color = 2
'.Update
'End With
'Set rstJnl = Nothing




End Sub

Private Sub cmdReject_Click()
Dim rstqueue As Recordset, rstJnl As Recordset
Dim InterJour As String
Dim cntr As Integer
Dim rstvalumeintake As Recordset

If Me.lstFiles.Column(7) <> "R " Then


            InterJour = InputBox("Enter Reason for Rejection", "Attorney Review")
            If InterJour = "" Then
            Exit Sub
            Else
            
             
                
                
                Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)
                
                With rstqueue
                .Edit
                !AttyMilestone2Reject = True
                !AttyMilestone2 = Now
                !AttyMilestone2Approver = GetStaffID
                .Update
                End With
                
                Me.Refresh
                Set rstqueue = Nothing
                
                Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttymilestone2", dbOpenDynaset, dbSeeChanges)
                Do Until rstqueue.EOF
                cntr = cntr + 1
                rstqueue.MoveNext
                Loop
                QueueCount = cntr
                Set rstqueue = Nothing
                
                'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
                
                 
                   Set rstvalumeintake = CurrentDb.OpenRecordset("Select * from ValumeIntake", dbOpenDynaset, dbSeeChanges)
                   With rstvalumeintake
                   .AddNew
                   !CaseFile = Forms!queAttyMilestone2!lstFiles
                   !Client = Forms!queAttyMilestone2!lstFiles.Column(2) 'DLookup("ShortClientName", "ClientList", "ClientID = " & Forms![Case List]!ClientID)
                   !AttyRej = Now
                   !AttyRejC = 1
                   !Name = GetFullName()
                   .Update
                   End With
                   Set rstvalumeintake = Nothing
                    
                    
                    
                    
                    DoCmd.SetWarnings False
                    strinfo = "Intake Review- Rejected because " & InterJour 'InputBox("Enter Reason for Rejection", "Attorney Review")
                    strinfo = Replace(strinfo, "'", "''")
                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone2!lstFiles,Now,GetFullName(),'" & strinfo & "',4 )"
                    DoCmd.RunSQL strSQLJournal
                    DoCmd.SetWarnings True
            
            End If
            
    Else
    
                InterJour = InputBox("Enter Reason for Rejection", "Attorney Review")
            If InterJour = "" Then
            Exit Sub
            Else
            
             
                
                
                Set rstqueue = CurrentDb.OpenRecordset("Select * FROM WizardSupportTwo where filenumber=" & Me.lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)
                
                With rstqueue
                .Edit
                !AttyMilestonerestartReject = True
                !AttyMilestoneRestart = Now
                !AttyMilestonerestartApprover = GetStaffID
                .Update
                End With
                
                Me.Refresh
                Set rstqueue = Nothing
                
                Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttymilestone2", dbOpenDynaset, dbSeeChanges)
                Do Until rstqueue.EOF
                cntr = cntr + 1
                rstqueue.MoveNext
                Loop
                QueueCount = cntr
                Set rstqueue = Nothing
                
                'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
                
                    
                       Set rstvalumeintake = CurrentDb.OpenRecordset("Select * from ValumeRestart", dbOpenDynaset, dbSeeChanges)
                       With rstvalumeintake
                       .AddNew
                       !CaseFile = Forms!queAttyMilestone2!lstFiles
                       !Client = Forms!queAttyMilestone2!lstFiles.Column(2) 'DLookup("ShortClientName", "ClientList", "ClientID = " & Forms![Case List]!ClientID)
                       !AttyRej = Now
                       !AttyRejC = 1
                       !Name = GetFullName()
                       .Update
                       End With
                       Set rstvalumeintake = Nothing
                    
                    
                    
                    
                    DoCmd.SetWarnings False
                    strinfo = "Restart Review- Rejected because " & InterJour 'InputBox("Enter Reason for Rejection", "Attorney Review")
                    strinfo = Replace(strinfo, "'", "''")
                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone2!lstFiles,Now,GetFullName(),'" & strinfo & "',4 )"
                    DoCmd.RunSQL strSQLJournal
                    DoCmd.SetWarnings True
                
                
                
                   
            
            End If
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

Private Sub cmdApprove_Click()
Dim cntr As Integer
Dim rstqueue As Recordset, rstJnl As Recordset

If Me.lstFiles.Column(7) <> "R " Then
'
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)
        
        With rstqueue
        .Edit
        !AttyMilestone2Reject = False
        !AttyMilestone2 = Now
        !AttyMilestone2Approver = GetStaffID
        .Update
        End With
        
        AddStatus lstFiles, Date, "Intake Review by Attorney Completed"
        
        'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
        
            DoCmd.SetWarnings False
            strinfo = "Intake Review by Attorney Completed"
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone2!lstFiles,Now,GetFullName(),'" & strinfo & "',1)"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            
        
        
        
        Set rstqueue = Nothing
        Me.Refresh
        
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttymilestone2", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        QueueCount = cntr
        Set rstqueue = Nothing
        
Else


     Set rstqueue = CurrentDb.OpenRecordset("Select * FROM WizardSupportTwo where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)
        
        With rstqueue
        .Edit
        rstqueue!DateSentAttyrestart = Null
        rstqueue!AttyMilestonerestartReject = False
        rstqueue!AttyMilestoneRestart = Now
        rstqueue!AttyMilestonerestartApprover = GetStaffID
        .Update
        End With
        
        AddStatus lstFiles, Date, "Restart Review by Attorney Completed"
        
        'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
        
            DoCmd.SetWarnings False
            strinfo = "Restart Review by Attorney Completed"
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!queAttyMilestone2!lstFiles,Now,GetFullName(),'" & strinfo & "',1)"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            
        
        
        
        Set rstqueue = Nothing
        Me.Refresh
        
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttymilestone2", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        QueueCount = cntr
        Set rstqueue = Nothing
        
End If


End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)

        Dim Fnumber As Long
        Dim Value As String
        Dim blnFound As Boolean
        Dim K As Integer
        Dim J As Integer
        Dim A As Integer
   
If Me.lstFiles.Column(7) <> "R " Then
   
        OpenCase lstFiles
       ' Forms![Case list]!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name],DocTitleID FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & lstFiles & " ' AND Filespec IS NOT NULL and DeleteDate is null ORDER BY Datestamp Desc"
    
      
        
        K = Forms![Case List]!lstDocs.ListCount - 1
        Fnumber = 1579
                    blnFound = False
                   
                    For J = K To 0 Step -1
                    'For J = 0 To Forms![Case list]!lstDocs.ListCount - 1
                    
                     
                       Value = Forms![Case List]!lstDocs.Column(4, J)
                       If InStr(Value, Fnumber) Then
                            blnFound = True
                             A = J
                            Forms![Case List].lstDocs.Selected(A) = True
                        Exit For
                        End If
                    Next J
                    
                  '  If Not blnFound Then MsgBox ("Draft Intake not in the document list.")
                    Forms![Case List]!lstDocs.SetFocus
                
Else


      OpenCase lstFiles
       ' Forms![Case list]!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name],DocTitleID FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & lstFiles & " ' AND Filespec IS NOT NULL and DeleteDate is null ORDER BY Datestamp Desc"
    
      
       
        K = Forms![Case List]!lstDocs.ListCount - 1
        Fnumber = 1585
                    blnFound = False
                    
                    For J = K To 0 Step -1
                    'For J = 0 To Forms![Case list]!lstDocs.ListCount - 1
                    
                     
                       Value = Forms![Case List]!lstDocs.Column(4, J)
                       If InStr(Value, Fnumber) Then
                            blnFound = True
                             A = J
                            Forms![Case List].lstDocs.Selected(A) = True
                        Exit For
                        End If
                    Next J
                    
                  '  If Not blnFound Then MsgBox ("Draft Intake not in the document list.")
                    Forms![Case List]!lstDocs.SetFocus
End If

    
End Sub
Private Sub Form_Current()
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttymilestone2", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
