VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queBankoNoSS"
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

'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
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
Dim rstqueue As Recordset, rstJnl As Recordset, rstbk As Recordset
On Error GoTo Err_cmdOK_Click

Set rstbk = CurrentDb.OpenRecordset("Select * FROM qryQueueBankoALLINFO where File=" & lstFiles, dbOpenDynaset, dbSeeChanges)

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueBankoNoSSN where File=" & lstFiles, dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!DataDisiminated = True
!WhenUpdated = Now
!WhoUpdatedIt = GetStaffID
.Update
End With

'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = lstFiles
'!JournalDate = Now
'!Who = GetFullName
'!Info = "BANKRUPTCY FOUND IN LEXIS/NEXIS MONITORING NOT MATCH SSN : Case Number:" & rstbk.Fields("CaseNumber") & " Chapter:" & rstbk.Fields("Chapter") & " File Date:" & rstbk.Fields("FileDate") & " Disposition:" & rstbk.Fields("DispositionCode") & " City Filed:" & rstbk.Fields("CityFiled") & " StateFiled:" & rstbk.Fields("StateFiled") & " County:" & rstbk.Fields("County") & " FirstName:" & rstbk.Fields("FirstName") & " LastName:" & rstbk.Fields("LastName") & " Debtors City:" & rstbk.Fields("DebtorsCity") & " Debtors State:" & rstbk.Fields("DebtorsState") _
'& " Law Firm:" & rstbk.Fields("LawFirm") & " Attorney Name:" & rstbk.Fields("AttorneyName") & " Attorney Address:" & rstbk.Fields("AttorneyAddress") & " Attorney City:" & rstbk.Fields("AttorneyCity") & " Attorney State:" & rstbk.Fields("AttorneyState") & " Attorney Zip:" & rstbk.Fields("AttorneyZip") & " AttorneyPhone:" & rstbk.Fields("AttorneyPhone") & " 341 Date:" & rstbk.Fields("341Date") & " 341 Time:" & rstbk.Fields("341Time") & " 341 Location:" & rstbk.Fields("341Location") & " Trustee:" & rstbk.Fields("Trustee") & " Trustee Address:" & rstbk.Fields("TrusteeAddress") _
' & " Trustee City:" & rstbk.Fields("TrusteeCity") & " Trustee State:" & rstbk.Fields("TrusteeState") & " Trustee Zip:" & rstbk.Fields("TrusteeZip") & " Trustee Phone:" & rstbk.Fields("TrusteePhone") & " Judges Initials:" & rstbk.Fields("JudgesInitials") & " Court District:" & rstbk.Fields("CourtDistrict") & " Court Phone:" & rstbk.Fields("CourtPhone") & " Debtor Phone:" & rstbk.Fields("DebtorPhone") & " Voluntary or Involuntary Dismissal:" & rstbk.Fields("VoluntaryInvoluntaryDismissal") & " Proof Of Claim Date:" & rstbk.Fields("ProofOfClaimDate")
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing

DoCmd.SetWarnings False
strinfo = "BANKRUPTCY FOUND IN LEXIS/NEXIS MONITORING NOT MATCH SSN : Case Number:" & rstbk.Fields("CaseNumber") & " Chapter:" & rstbk.Fields("Chapter") & " File Date:" & rstbk.Fields("FileDate") & " Disposition:" & rstbk.Fields("DispositionCode") & " City Filed:" & rstbk.Fields("CityFiled") & " StateFiled:" & rstbk.Fields("StateFiled") & " County:" & rstbk.Fields("County") & " FirstName:" & rstbk.Fields("FirstName") & " LastName:" & rstbk.Fields("LastName") & " Debtors City:" & rstbk.Fields("DebtorsCity") & " Debtors State:" & rstbk.Fields("DebtorsState") _
& " Law Firm:" & rstbk.Fields("LawFirm") & " Attorney Name:" & rstbk.Fields("AttorneyName") & " Attorney Address:" & rstbk.Fields("AttorneyAddress") & " Attorney City:" & rstbk.Fields("AttorneyCity") & " Attorney State:" & rstbk.Fields("AttorneyState") & " Attorney Zip:" & rstbk.Fields("AttorneyZip") & " AttorneyPhone:" & rstbk.Fields("AttorneyPhone") & " 341 Date:" & rstbk.Fields("341Date") & " 341 Time:" & rstbk.Fields("341Time") & " 341 Location:" & rstbk.Fields("341Location") & " Trustee:" & rstbk.Fields("Trustee") & " Trustee Address:" & rstbk.Fields("TrusteeAddress") _
 & " Trustee City:" & rstbk.Fields("TrusteeCity") & " Trustee State:" & rstbk.Fields("TrusteeState") & " Trustee Zip:" & rstbk.Fields("TrusteeZip") & " Trustee Phone:" & rstbk.Fields("TrusteePhone") & " Judges Initials:" & rstbk.Fields("JudgesInitials") & " Court District:" & rstbk.Fields("CourtDistrict") & " Court Phone:" & rstbk.Fields("CourtPhone") & " Debtor Phone:" & rstbk.Fields("DebtorPhone") & " Voluntary or Involuntary Dismissal:" & rstbk.Fields("VoluntaryInvoluntaryDismissal") & " Proof Of Claim Date:" & rstbk.Fields("ProofOfClaimDate")
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & lstFiles & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

AddStatus lstFiles, Date, "Lexis/Nexis BK Hit Identified and Noted in File."

Set rstqueue = Nothing
Me.Refresh
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueBanko", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

OpenCaseDONTCloseForms_S2 lstFiles

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

On Error GoTo Err_cmdRemove_Click


Dim rstqueue As Recordset, rstJnl As Recordset, rstbk As Recordset
Set rstbk = CurrentDb.OpenRecordset("Select * FROM qryQueueBankoALLINFO where File=" & lstFiles, dbOpenDynaset, dbSeeChanges)
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueBankoNoSSN where File=" & lstFiles, dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!DataDisiminated = True
!WhenUpdated = Now
!WhoUpdatedIt = GetStaffID
.Update
End With

'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = lstFiles
'!JournalDate = Now
'!Who = GetFullName
'!Info = "False or Duplicate Bankruptcy Filing Found by Lexis/Nexis by" & GetFullName & ". information rendered inactive. Record Number: " & rstbk.Fields("ID") & " in table BKInputFromLexisNexis. PLEASE NOTE: a Bankuptcy note below denotes a duplicate."
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing

DoCmd.SetWarnings False
strinfo = "False or Duplicate Bankruptcy Filing Found by Lexis/Nexis by" & GetFullName & ". information rendered inactive. Record Number: " & rstbk.Fields("ID") & " in table BKInputFromLexisNexis. PLEASE NOTE: a Bankuptcy note below denotes a duplicate."
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & lstFiles & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

'AddStatus lstFiles, Date, "Lexis/Nexis BK Hit Not Germaine. Reference Deleted."

Set rstqueue = Nothing
Me.Refresh
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueBanko", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

Exit_cmdRemove_Click:
    Exit Sub

Err_cmdRemove_Click:
    MsgBox ("You Must First Select A File")
    ' MsgBox Err.Description
    Resume Exit_cmdRemove_Click

End Sub

Private Sub cmdReview_Click()

On Error GoTo Err_cmdReview_Click


OpenCaseDONTCloseForms_S2 lstFiles


Exit_cmdReview_Click:
    Exit Sub

Err_cmdReview_Click:
    MsgBox ("You Must First Select A File")
    ' MsgBox Err.Description
    Resume Exit_cmdReview_Click
    
End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)

    OpenCaseDONTCloseForms_S2 lstFiles
    
End Sub
Private Sub Form_Open(Cancel As Integer)

Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueBankoNoSSN", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
