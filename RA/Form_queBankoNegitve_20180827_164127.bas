VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queBankoNegitve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCorrections_Click()
Dim rstqueue As Recordset ', rstJnl As Recordset

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
Dim rstqueue As Recordset, rstJnl As Recordset, rstDE As Recordset, rstUpdateName As Recordset, rstDoc, rstJnlS, rstdocs As Recordset
On Error GoTo Err_cmdOK_Click
If Not IsNull(lstFiles) Then
If lstFiles.Column(4) <> "Locked" Then

'Set rstDE = CurrentDb.OpenRecordset("Select * FROM BankoQueueNegativeHits where File=" & lstFiles.Column(0))
'
''Set rstqueue = CurrentDb.OpenRecordset("Select * FROM BKInputFromLexisNexisNegative", dbOpenDynaset, dbSeeChanges)
'
'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where FileNumber=" & lstFiles.Column(0), dbOpenDynaset, dbSeeChanges)
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'Set rstUpdateName = CurrentDb.OpenRecordset("Select * FROM Names Where FileNumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)



Set rstDoc = CurrentDb.OpenRecordset("BKInputFromLexisNexisNegative", dbOpenDynaset, dbSeeChanges)
With rstDoc
.AddNew
!CaseFile = lstFiles.Column(0)
!ProjectName = lstFiles.Column(1)
!ClientName = lstFiles.Column(2)
!State = lstFiles.Column(3)
!Stage = lstFiles.Column(7)
!Date = lstFiles.Column(6)
!DaysDue = lstFiles.Column(8)
!DataDisiminated = True
!WhenUpdated = Now()
!WhoUpdatedIt = GetStaffID
.Update
End With
Set rstDoc = Nothing



'With rstJnl
'.AddNew
'!FileNumber = lstFiles.Column(0)
'!JournalDate = Now
'!Who = GetFullName
If lstFiles.Column(9) = "Not in File" Then
'!Info = "Loan # " & lstFiles.Column(10) & ", as of " & lstFiles.Column(6) & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no hits for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", as of " & lstFiles.Column(6) & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no hits for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Else
Select Case lstFiles.Column(7)
Case "At Referral"
'!Info = "Loan # " & lstFiles.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

strinfo = "Loan # " & lstFiles.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "First legal advertisement runs"
'!Info = "Loan # " & lstFiles.Column(10) & ", as of " & Format(lstFiles.Column(6), "mm/dd/yyyy") & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

strinfo = "Loan # " & lstFiles.Column(10) & ", as of " & Format(lstFiles.Column(6), "mm/dd/yyyy") & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "First legal Sent to Docket"
'!Info = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", file sent to docket first legal no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", file sent to docket first legal no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "2 Days Before Sale"
'!Info = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -2, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -2, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -4, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -2, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 2 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -2, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -2, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -4, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -2, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 2 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "3 Days After Sale"
'!Info = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 3, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 5, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", 3, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", 4, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 3 days after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 3, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 5, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", 3, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", 4, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 3 days after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "5 - 10 Days before Sale"
'!Info = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -10, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -8, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -10, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -9, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -10, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5-10 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -10, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -8, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -10, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -9, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -10, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5-10 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "Day Of Sale"
'!Info = "Loan # " & lstFiles.Column(10) & ", search conducted on " & Format(lstFiles.Column(6), "mm/dd/yyyy") & " 2.00 am" & ", day of sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & Format(lstFiles.Column(6), "mm/dd/yyyy") & " 2.00 am" & ", day of sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "After Sale "
'!Info = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 day after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 day after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "Day Before Sale"
'!Info = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -1, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 day Before sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -1, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 day Before sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
Case "7 Days Before Sale"
'!Info = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -7, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -5, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -7, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -6, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -7, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 7 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -7, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -5, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -7, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -6, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -7, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 7 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "5 Days  Before Sale"
'!Info = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -5, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -5, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -4, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -5, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -5, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -5, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -4, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -5, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "Borrower Served"
'!Info = "Loan # " & lstFiles.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "Sent Complaint"
strinfo = "Loan # " & lstFiles.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
'Entry of Judgment

Case "Entry of Judgment"
strinfo = "Loan # " & lstFiles.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."


Case "Judgment Entered"
strinfo = "Loan # " & lstFiles.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 0, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 2, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 0, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & lstFiles.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "First legal Sent to Docket-same day"
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & Format(lstFiles.Column(6), "mm/dd/yyyy") & " 2.00 am" & ", First legal Sent to Docket, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."


Case "1 Days After Sale"
'!Info = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 3, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 5, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", 3, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", 4, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 3 days after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 3, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", 1, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", 2, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 1 days after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "5 Business Days After Sale"
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 5, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", 7, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", 5, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", 6, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 5, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5 Business days after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "1 Business day prior to Sale"
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -1, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 business day prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "First legal ad - 5 business  days"
'strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -5, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -5, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -4, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -5, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5 business days prior to First publication date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & Format(DateAdd("d", -1 * (Businessdays(lstFiles.Column(6), 5)), lstFiles.Column(6)), "mm/dd/yyyy") & " 2.00 am" & ", 5 business days prior to First publication date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."


Case "First legal ad - 1 day Prior"
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -1, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -1, lstFiles.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 day prior to first publication , no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "3 Business day prior to Sale"
'strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -5, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -5, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -4, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -5, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5 business days prior to First publication date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
'strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -3, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -1, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -3, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -2, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 3 business days prior to Sale date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & Format(DateAdd("d", -1 * (Businessdays(lstFiles.Column(6), 3)), lstFiles.Column(6)), "mm/dd/yyyy") & " 2.00 am" & ", 3 business days prior to sale date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

Case "5 Business Days prior to scheduled "
'strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -5, lstFiles.Column(6))) = vbSaturday, Format(DateAdd("d", -3, lstFiles.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -5, lstFiles.Column(6))) = vbSunday, Format(DateAdd("d", -4, lstFiles.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -5, lstFiles.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5 business days prior to scheduled sale date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."
strinfo = "Loan # " & lstFiles.Column(10) & ", search conducted on " & Format(DateAdd("d", -1 * (Businessdays(lstFiles.Column(6), 5)), lstFiles.Column(6)), "mm/dd/yyyy") & " 2.00 am" & ", 5 business days prior to scheduled sale date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(lstFiles.Column(0), 2) & "."

End Select
End If


'!Info = IIf(lstFiles.Column(9) = "Not in File", "Not in File : Loan # " & lstFiles.Column(10) & ", as of " & lstFiles.Column(6) & ", " & lstFiles.Column(7) & ", there were no hits for " & BorrowerNames(lstFiles.Column(0)) & " located", " Negative Hit: For Loan # " & lstFiles.Column(10) & ", as of " & lstFiles.Column(6) & ", " & lstFiles.Column(7) & ", " & BorrowerNames(lstFiles.Column(0)) & ", have been added to the monitoring queue as of " & lstFiles.Column(6) & " , no parties have filed bankruptcy.")

'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing
'Set rstUpdateName = Nothing

    DoCmd.SetWarnings False
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & lstFiles.Column(0) & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

AddStatus lstFiles, Date, "Lexis/Nexis Banko Negative hits confirmed."

Set rstqueue = Nothing

Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM BankoQueueNegativeHits")
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
'Set rstDE = Nothing
OpenCaseDONTCloseForms_S2 lstFiles
Else

OpenCaseDONTCloseForms_S2 lstFiles
End If
Else

        If ListS.Column(4) <> "Locked" Then
    
    'Set rstDE = CurrentDb.OpenRecordset("Select * FROM BankoQueueNegativeHits where File=" & lstFiles.Column(0))
    '
    ''Set rstqueue = CurrentDb.OpenRecordset("Select * FROM BKInputFromLexisNexisNegative", dbOpenDynaset, dbSeeChanges)
    '
    'Set rstJnlS = CurrentDb.OpenRecordset("Select * FROM journal where FileNumber=" & ListS.Column(0), dbOpenDynaset, dbSeeChanges)
    'Set rstJnlS = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
    
    'Set rstUpdateName = CurrentDb.OpenRecordset("Select * FROM Names Where FileNumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
    
    
    
    Set rstdocs = CurrentDb.OpenRecordset("BKInputFromLexisNexisNegative", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !CaseFile = ListS.Column(0)
    !ProjectName = ListS.Column(1)
    !ClientName = ListS.Column(2)
    !State = ListS.Column(3)
    !Stage = ListS.Column(7)
    !Date = ListS.Column(6)
    !DaysDue = ListS.Column(8)
    !DataDisiminated = True
    !WhenUpdated = Now()
    !WhoUpdatedIt = GetStaffID
    .Update
    End With
    Set rstdocs = Nothing
    
    
'     With rstJnlS
'    .AddNew
'    !FileNumber = ListS.Column(0)
'    !JournalDate = Now
'    !Who = GetFullName
    If ListS.Column(9) = "Not in File" Then
'!Info = "Loan # " & ListS.Column(10) & ", as of " & ListS.Column(6) & " 2.00 am" & ", " & ListS.Column(7) & ", there were no hits for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."


strinfo = "Loan # " & ListS.Column(10) & ", as of " & ListS.Column(6) & " 2.00 am" & ", " & ListS.Column(7) & ", there were no hits for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."


Else
Select Case ListS.Column(7)
Case "At Referral"
'!Info = " Our File, " & ListS.Column(0) & " , involving borrowers " & BorrowerNames(ListS.Column(0)) & IIf(ListS.Column(9) = "Negative hit", ", has been added to the monitering queue of ", " , has been monitered by Lexis Nexis through Banko for bankruptcy data.  As of ") & Date & ", " & ListS.Column(7) & " , no parties have filed bankruptcy."
'!Info = " Our File, " & ListS.Column(0) & " , involving borrowers " & BorrowerNames(ListS.Column(0)) & ", has been added to the monitering queue of " & ListS.Column(7) & " , no parties have filed bankruptcy."
'!Info = IIf(ListS.Column(9) = "Not in File", "Loan # " & ListS.Column(10) & ", as of " & ListS.Column(6) & ", " & ListS.Column(7) & ", there were no hits for " & BorrowerNames(ListS.Column(0)), " Our File, " & ListS.Column(0) & " , involving borrowers " & BorrowerNames(ListS.Column(0)) & ", has been added to the monitering queue of " & ListS.Column(7) & " , no parties have filed bankruptcy.")
'!Info = "Loan # " & ListS.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & ListS.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & ListS.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "First legal advertisement runs"
'!Info = "Loan # " & ListS.Column(10) & ", as of " & Format(ListS.Column(6), "mm/dd/yyyy") & " 2.00 am" & ", " & ListS.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", as of " & Format(ListS.Column(6), "mm/dd/yyyy") & " 2.00 am" & ", " & ListS.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "First legal Sent to Docket"
'!Info = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", file sent to docket first legal no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", file sent to docket first legal no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "2 Days Before Sale"
'!Info = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -2, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", -3, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -2, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -4, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -2, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 2 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -2, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", -3, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -2, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -4, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -2, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 2 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "3 Days After Sale"
'!Info = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 3, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 5, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", 3, ListS.Column(6))) = vbSunday, Format(DateAdd("d", 4, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 3 days after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 3, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 5, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", 3, ListS.Column(6))) = vbSunday, Format(DateAdd("d", 4, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 3 days after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "5 - 10 Days before Sale"
'!Info = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -10, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", -8, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -10, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -9, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -10, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5-10 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -10, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", -8, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -10, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -9, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -10, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5-10 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "Day Of Sale"
'!Info = "Loan # " & ListS.Column(10) & ", search conducted on " & Format(ListS.Column(6), "mm/dd/yyyy") & " 2.00 am" & ", day of sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & Format(ListS.Column(6), "mm/dd/yyyy") & " 2.00 am" & ", day of sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "After Sale "
'!Info = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 day after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 day after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "Day Before Sale"
'!Info = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -1, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 day Before sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -1, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 day Before sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "7 Days Before Sale"
'!Info = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -7, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", -5, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -7, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -6, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -7, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 7 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -7, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", -5, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -7, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -6, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -7, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 7 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "5 Days  Before Sale"
strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -5, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", -3, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -5, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -4, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -5, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "Borrower Served"
'!Info = "Loan # " & ListS.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & ListS.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

strinfo = "Loan # " & ListS.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & ListS.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "Sent Complaint"
strinfo = "Loan # " & ListS.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & ListS.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "Entry of Judgment"
strinfo = "Loan # " & ListS.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & ListS.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."


Case "Judgment Entered"
strinfo = "Loan # " & ListS.Column(10) & ", as of " & IIf(Weekday(DateAdd("d", 0, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 2, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 0, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", " & ListS.Column(7) & ", there were no bankruptcy filings found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "First legal Sent to Docket-same day"
strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & Format(ListS.Column(6), "mm/dd/yyyy") & " 2.00 am" & ", First legal Sent to Docket, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."


Case "2 Days Before Sale"
strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -2, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 0, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -2, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -1, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -2, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 2 days prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."


Case "1 Days After Sale"
strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 3, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", 1, ListS.Column(6))) = vbSunday, Format(DateAdd("d", 2, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 1, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 1 days after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "5 Business Days After Sale"
strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", 5, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", 7, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", 5, ListS.Column(6))) = vbSunday, Format(DateAdd("d", 6, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", 5, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5 Business days after sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "1 Business day prior to Sale"
strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -1, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 business day prior to sale, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "First legal ad - 5 business  days"
strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & Format(DateAdd("d", -1 * (Businessdays(ListS.Column(6), 5)), ListS.Column(6)), "mm/dd/yyyy") & " 2.00 am" & ", 5 business days prior to first publication date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

'strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -5, ListS.Column(6))) = vbSaturday, Format(DateAdd("d", -3, ListS.Column(6)), "mm/dd/yyyy"), IIf(Weekday(DateAdd("d", -5, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -4, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -5, ListS.Column(6)), "mm/dd/yyyy"))) & " 2.00 am" & ", 5 business days prior to first publication date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

'Format(DateAdd("d", -1 * (Businessdays(lstFiles.Column(6), 5)), lstFiles.Column(6)), "mm/dd/yyyy")

Case "First legal ad - 1 day Prior"
strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & IIf(Weekday(DateAdd("d", -1, ListS.Column(6))) = vbSunday, Format(DateAdd("d", -3, ListS.Column(6)), "mm/dd/yyyy"), Format(DateAdd("d", -1, ListS.Column(6)), "mm/dd/yyyy")) & " 2.00 am" & ", 1 day prior to first publication , no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."


Case "3 Business day prior to Sale"
strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & Format(DateAdd("d", -1 * (Businessdays(ListS.Column(6), 3)), ListS.Column(6)), "mm/dd/yyyy") & " 2.00 am" & ", 3 business days prior to sale date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

Case "5 Business Days prior to scheduled "
'strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & Format(DateAdd("d", -1 * (Businessdays(ListS.Column(6), 3)), ListS.Column(6)), "mm/dd/yyyy") & " 2.00 am" & ", 3 business days prior to sale date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."
strinfo = "Loan # " & ListS.Column(10) & ", search conducted on " & Format(DateAdd("d", -1 * (Businessdays(ListS.Column(6), 5)), ListS.Column(6)), "mm/dd/yyyy") & " 2.00 am" & ", 5 business days prior to scheduled sale date, no bankruptcy filing found for " & BorrowerMortgagorOwnerName(ListS.Column(0), 2) & "."

End Select
End If
    
    '!Info = IIf(lstFiles.Column(9) = "Not in File", "Not in File : Loan # " & lstFiles.Column(10) & ", as of " & lstFiles.Column(6) & ", " & lstFiles.Column(7) & ", there were no hits for " & BorrowerNames(lstFiles.Column(0)) & " located", " Negative Hit: For Loan # " & lstFiles.Column(10) & ", as of " & lstFiles.Column(6) & ", " & lstFiles.Column(7) & ", " & BorrowerNames(lstFiles.Column(0)) & ", have been added to the monitoring queue as of " & lstFiles.Column(6) & " , no parties have filed bankruptcy.")
    
    'strInfo = IIf(lstFiles.Column(9) = "Not in File", "Not in File : Loan # " & lstFiles.Column(10) & ", as of " & lstFiles.Column(6) & ", " & lstFiles.Column(7) & ", there were no hits for " & BorrowerNames(lstFiles.Column(0)) & " located", " Negative Hit: For Loan # " & lstFiles.Column(10) & ", as of " & lstFiles.Column(6) & ", " & lstFiles.Column(7) & ", " & BorrowerNames(lstFiles.Column(0)) & ", have been added to the monitoring queue as of " & lstFiles.Column(6) & " , no parties have filed bankruptcy.")
    
'    !Color = 1
'    .Update
'    End With
'    Set rstJnlS = Nothing
    DoCmd.SetWarnings False
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & ListS.Column(0) & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    'Set rstUpdateName = Nothing
    
    AddStatus ListS, Date, "Lexis/Nexis Banko Negative hits confirmed."
   OpenCaseDONTCloseForms_S2 ListS
    
    Else

OpenCaseDONTCloseForms_S2 ListS
End If
End If

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
   ' MsgBox ("You Must First Select A File")
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
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM BankoQueueNegativeHits")
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
Me.Refresh
End Sub

Private Sub cmdRemove_Click()

Dim rstqueue As Recordset, rstJnl As Recordset, rstDE As Recordset, rstUpdateName As Recordset

On Error GoTo Err_cmdRemove_Click

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea where NameID=" & lstFiles.Column(1), dbOpenDynaset, dbSeeChanges)





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
'!Info = "False Info Found by Lexis/Nexis by" & GetFullName & " PLEASE NOTE: No Action will take to update the File."
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing

DoCmd.SetWarnings False
strinfo = "False Info Found by Lexis/Nexis by" & GetFullName & " PLEASE NOTE: No Action will take to update the File."
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & lstFiles & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

'AddStatus lstFiles, Date, "Lexis/Nexis BK Hit Not Germaine. Reference Deleted."

Set rstqueue = Nothing
Me.Refresh
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea", dbOpenDynaset, dbSeeChanges)
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

Private Sub ListS_Click()
Me.lstFiles = Null
ListS.SetFocus
Me.lstFiles = Null
End Sub

Private Sub ListS_DblClick(Cancel As Integer)
 OpenCaseDONTCloseForms_S2 ListS
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
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM BankoQueueNegativeHitsQ")
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
