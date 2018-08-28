VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queDecea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCorrections_Click()
Dim rstqueue As Recordset, rstJnl As Recordset

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!AttyMilestone1Reject = False
!AttyMilestone1 = Now
!AttyMilestone1Approver = GetStaffID
.Update
End With

AddStatus lstFiles, Date, "Intake Review by Attorney Completed"

Me.Refresh
Set rstqueue = Nothing
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueAttyMilestone1", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)02/05/14
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = lstFiles
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Intake Review- Approved with the following corrections made:  " & InputBox("Enter Description of Corrections Made", "Attorney Review")
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing

DoCmd.SetWarnings False
strinfo = "Intake Review- Approved with the following corrections made:  " & InputBox("Enter Description of Corrections Made", "Attorney Review")
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & lstFiles & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

End Sub

Private Sub cmdAccept_Click()
Dim rstqueue As Recordset, rstJnl As Recordset, rstDE As Recordset, rstUpdateName As Recordset, cntr As Integer, rstQueueb As Recordset, rstJnlB As Recordset, rstUpdateNameB As Recordset
On Error GoTo Err_cmdOK_Click

'If lstFiles <> Null Then
''If lstFiles.Column(7) <> "Locked" Then
'
'        Set rstDE = CurrentDb.OpenRecordset("Select * FROM qryQueueDeceaALLINFO where File=" & lstFiles, dbOpenDynaset, dbSeeChanges)
'
'        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea where NameID=" & lstFiles.Column(1), dbOpenDynaset, dbSeeChanges)
'
'        Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
'


If Not IsNull(lstFiles) Then
If lstFiles.Column(7) Like "Available" Then

    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea where NameID=" & lstFiles.Column(1), dbOpenDynaset, dbSeeChanges)
    
    'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)02/5/14
'    Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
    
        Set rstUpdateName = CurrentDb.OpenRecordset("Select * FROM Names Where ID=" & lstFiles.Column(1), dbOpenDynaset, dbSeeChanges)

        
        With rstqueue
        .Edit
        !DataDisiminated = True
        !WhenUpdated = Now
        !WhoUpdatedIt = GetStaffID
        .Update
        End With
        
        With rstUpdateName
        .Edit
        !Deceased = True
        .Update
        End With
        
'        With rstJnl
'        .AddNew
'        !FileNumber = lstFiles
'        !JournalDate = Now
'        !Who = GetFullName
'        !Info = " DECEASED FOUND IN LEXIS/NEXIS MONITORING:  Firt Name: " & lstFiles.Column(3) & " Last Name : " & lstFiles.Column(4) & " State Of Last Residence  " & lstFiles.Column(5) & " Deceased on: " & lstFiles.Column(6) & " SSN is exact match"
'        !Color = 1
'        .Update
'        End With
'        Set rstJnl = Nothing
        
DoCmd.SetWarnings False
strinfo = " DECEASED FOUND IN LEXIS/NEXIS MONITORING:  Firt Name: " & lstFiles.Column(3) & " Last Name : " & lstFiles.Column(4) & " State Of Last Residence  " & lstFiles.Column(5) & " Deceased on: " & lstFiles.Column(6) & " SSN is exact match"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & lstFiles & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

        Set rstUpdateName = Nothing
        
        AddStatus lstFiles, Date, "Lexis/Nexis Deceased Hit Identified and Noted in File. SSN exact match."
        
        Set rstqueue = Nothing
        
       
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        QueueCount = cntr
        
        DoCmd.SetWarnings False
        DoCmd.OpenQuery "DecInputFromLexisNexisDismissedSSNUpdarte0"
        DoCmd.Close acQuery, "DecInputFromLexisNexisDismissedSSNUpdarte0"
        DoCmd.SetWarnings True
        
        Set rstqueue = Nothing
        
        OpenCaseDONTCloseForms_S2 lstFiles
     Else
    OpenCaseDONTCloseForms_S2 lstFiles
    End If
 
Else
    If ListS.Column(7) Like "Available" Then
'Set rstDE = CurrentDb.OpenRecordset("Select * FROM qryQueueDeceaALLINFO where File=" & ListS, dbOpenDynaset, dbSeeChanges)
        
        Set rstQueueb = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea where NameID=" & ListS.Column(1), dbOpenDynaset, dbSeeChanges)
        
        'Set rstJnlB = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & ListS, dbOpenDynaset, dbSeeChanges)
'        Set rstJnlB = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
        
        Set rstUpdateNameB = CurrentDb.OpenRecordset("Select * FROM Names Where ID=" & ListS.Column(1), dbOpenDynaset, dbSeeChanges)
        
        
        With rstQueueb
        .Edit
        !DataDisiminated = True
        !WhenUpdated = Now
        !WhoUpdatedIt = GetStaffID
        .Update
        End With
        
        With rstUpdateNameB
        .Edit
        !Deceased = True
        .Update
        End With
        
'        With rstJnlB
'        .AddNew
'        !FileNumber = ListS
'        !JournalDate = Now
'        !Who = GetFullName
'        !Info = " DECEASED FOUND IN LEXIS/NEXIS MONITORING:  Firt Name: " & ListS.Column(3) & " Last Name : " & ListS.Column(4) & " State Of Last Residence  " & ListS.Column(5) & " Deceased on: " & ListS.Column(6) & " SSN is exact match"
'        !Color = 1
'        .Update
'        End With
'        Set rstJnlB = Nothing
        
DoCmd.SetWarnings False
strinfo = " DECEASED FOUND IN LEXIS/NEXIS MONITORING:  Firt Name: " & ListS.Column(3) & " Last Name : " & ListS.Column(4) & " State Of Last Residence  " & ListS.Column(5) & " Deceased on: " & ListS.Column(6) & " SSN is exact match"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(ListS,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
        Set rstUpdateNameB = Nothing
        
        AddStatus ListS, Date, "Lexis/Nexis Deceased Hit Identified and Noted in File. SSN exact match."
        
        Set rstQueueb = Nothing
        
    '    Dim cntr As Integer
        Set rstQueueb = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea", dbOpenDynaset, dbSeeChanges)
        Do Until rstQueueb.EOF
        cntr = cntr + 1
        rstQueueb.MoveNext
        Loop
        QueueCount = cntr
        Set rstQueueb = Nothing
        
        OpenCaseDONTCloseForms ListS
   Else

    OpenCaseDONTCloseForms ListS
    End If

End If



        
Exit_cmdOK_Click:
            Exit Sub
        
Err_cmdOK_Click:
            MsgBox ("You Must First Select A File")
            ' MsgBox Err.Description
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
Me.Refresh
End Sub

Private Sub cmdRemove_Click()

Dim rstqueue As Recordset, rstJnl As Recordset, rstDE As Recordset, rstUpdateName As Recordset, cntr As Integer

On Error GoTo Err_cmdRemove_Click

If Not IsNull(lstFiles) Then

    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea where NameID=" & lstFiles.Column(1), dbOpenDynaset, dbSeeChanges)
    
   
    
    
    
    With rstqueue
    .Edit
    !DataDisiminated = True
    !WhenUpdated = Now
    !WhoUpdatedIt = GetStaffID
    .Update
    End With
    
    'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
'    Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'    With rstJnl
'    .AddNew
'    !FileNumber = lstFiles
'    !JournalDate = Now
'    !Who = GetFullName
'    !Info = "False Info Found by Lexis/Nexis by" & GetFullName & " PLEASE NOTE: No Action will take to update the File."
'    !Color = 1
'    .Update
'    End With
'    Set rstJnl = Nothing
    
    DoCmd.SetWarnings False
    strinfo = "False Info Found by Lexis/Nexis by" & GetFullName & " PLEASE NOTE: No Action will take to update the File."
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & lstFiles & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    
    'AddStatus lstFiles, Date, "Lexis/Nexis BK Hit Not Germaine. Reference Deleted."
    
    Set rstqueue = Nothing
    Me.Refresh
    'Dim cntr As Integer
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea", dbOpenDynaset, dbSeeChanges)
    Do Until rstqueue.EOF
    cntr = cntr + 1
    rstqueue.MoveNext
    Loop
    QueueCount = cntr
    Set rstqueue = Nothing
Else

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea where NameID=" & ListS.Column(1), dbOpenDynaset, dbSeeChanges)
    
   
    
    
    
    With rstqueue
    .Edit
    !DataDisiminated = True
    !WhenUpdated = Now
    !WhoUpdatedIt = GetStaffID
    .Update
    End With
'    Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'    'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & ListS, dbOpenDynaset, dbSeeChanges)
'    With rstJnl
'    .AddNew
'    !FileNumber = lstFiles
'    !JournalDate = Now
'    !Who = GetFullName
'    !Info = "False Info Found by Lexis/Nexis by" & GetFullName & " PLEASE NOTE: No Action will take to update the File."
'    !Color = 1
'    .Update
'    End With
'    Set rstJnl = Nothing
    
    DoCmd.SetWarnings False
strinfo = "False Info Found by Lexis/Nexis by" & GetFullName & " PLEASE NOTE: No Action needs to be taken to update the File."
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & lstFiles & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
    
    'AddStatus lstFiles, Date, "Lexis/Nexis BK Hit Not Germaine. Reference Deleted."
    
    Set rstqueue = Nothing
    Me.Refresh
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea", dbOpenDynaset, dbSeeChanges)
    Do Until rstqueue.EOF
    cntr = cntr + 1
    rstqueue.MoveNext
    Loop
    QueueCount = cntr
    Set rstqueue = Nothing
End If



Exit_cmdRemove_Click:
    Exit Sub

Err_cmdRemove_Click:
    MsgBox ("You Must First Select A File")
    ' MsgBox Err.Description
    Resume Exit_cmdRemove_Click

End Sub

Private Sub cmdReview_Click()

On Error GoTo Err_cmdReview_Click
If Not IsNull(lstFiles) Then

OpenCaseDONTCloseForms_S2 lstFiles

Else
OpenCaseDONTCloseForms_S2 ListS
End If


Exit_cmdReview_Click:
    Exit Sub

Err_cmdReview_Click:
    MsgBox ("You Must First Select A File")
    ' MsgBox Err.Description
    Resume Exit_cmdReview_Click
    
End Sub

Private Sub List100_Click()


End Sub

Private Sub ListS_Click()
ListS.SetFocus
lstFiles = Null
End Sub

Private Sub lstFiles_Click()
lstFiles.SetFocus
ListS = Null


End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)

    OpenCaseDONTCloseForms_S2 lstFiles
    
End Sub
Private Sub Form_Open(Cancel As Integer)

Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
