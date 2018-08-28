VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ReportsWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub PrintDocs(PrintTo As Integer)
Dim ReportName As String
Dim rstStates As Recordset

Dim str_SQL As String
        
On Error GoTo Err_PrintDocs
        
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM WorkflowReportStates"
WorkflowStates = ""
        
'9-5-14
If chDC Then
str_SQL = "INSERT INTO WorkflowReportStates(ID,state) VALUES (1,'DC')"
            Debug.Print str_SQL
            RunSQL (str_SQL)
    WorkflowStates = WorkflowStates & "DC, "
        
End If
        
If chMaryland Then
str_SQL = "INSERT INTO WorkflowReportStates(ID,state) VALUES (2,'MD')"
            Debug.Print str_SQL
            RunSQL (str_SQL)
WorkflowStates = WorkflowStates & "MD, "
        
        
End If
        
If chVirginia Then
str_SQL = "INSERT INTO WorkflowReportStates(ID,state) VALUES (3,'VA')"
            Debug.Print str_SQL
            RunSQL (str_SQL)
WorkflowStates = WorkflowStates & "VA, "
        
End If
DoCmd.SetWarnings True


'Set rstStates = CurrentDb.OpenRecordset("WorkflowReportStates", dbOpenDynaset)
'If chDC Then
    'rstStates.AddNew
    'rstStates!State = "DC"
    'rstStates.Update
    'WorkflowStates = WorkflowStates & "DC, "
'End If
'If chMaryland Then
    'rstStates.AddNew
   ' rstStates!State = "MD"
    'rstStates.Update
    'WorkflowStates = WorkflowStates & "MD, "
'End If
'If chVirginia Then
    'rstStates.AddNew
    'rstStates!State = "VA"
   ' rstStates.Update
    'WorkflowStates = WorkflowStates & "VA, "
'End If
'rstStates.Close


If WorkflowStates = "" Then
    MsgBox "Select at least one state.", vbCritical
    Exit Sub
Else
    WorkflowStates = Left$(WorkflowStates, Len(WorkflowStates) - 2)
End If

If ch362ToBeFiled Then
    If PrintTo = -3 Then
    '\\fileserver\Applications\Database\Templates\326.xls"
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputReport, "Workflow 362ToBeFiled Excel", acFormatXLS, TemplatePath & "326.xlt", 1, , True
    DoCmd.SetWarnings True
    Else
    DoReport "Workflow 362 to be Filed", PrintTo
    End If
End If

If chTitleClaimsNotSent Then DoReport "Workflow Title Claims Not Sent", PrintTo

If chTitleClaimsOut Then DoReport "Workflow Title Claims Out", PrintTo

        
If chTitleToBeReviewed Then DoReport "Workflow Title To Be Reviewed", PrintTo


If ChFeesAndCosts Then
    If PrintTo = -3 Then
    '\\fileserver\Applications\Database\Templates\326.xls"
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputReport, "Workflow Fees and Costs", acFormatXLS, TemplatePath & "ReferralsReceived.xlt", 1, , True
    DoCmd.SetWarnings True
    Else
    DoReport "Workflow Fees and Costs", PrintTo
    End If
End If

If chPrepComplaint Then DoReport "Workflow Prepare Complaint", PrintTo
If chClientComplaintReturned Then DoReport "WOrkflow Client Complaint Returned", PrintTo
If chComplaintToCourt Then DoReport "Workflow Complaint to COurt", PrintTo
If chComplaintFiled Then DoReport "Workflow Complaint Filed", PrintTo
If chFiledLisPendens Then DoReport "WOrkflow Filed Lis Pendens", PrintTo

If chBankruptcy Then DoReport "Workflow Bankruptcy", PrintTo
If chMotionsOut Then DoReport "Workflow Motions Outstanding", PrintTo
If chAffDefault Then DoReport "Workflow Affidavits of Default", PrintTo
If chDefaultsCured Then DoReport "Workflow Affidavits of Default Monitor", PrintTo
If chHearings362 Then DoReport "Workflow Hearings 362", PrintTo
If chHearingsScheduled Then DoReport "Workflow Hearings Scheduled", PrintTo
If chConsentModified Then DoReport "Workflow Consent Orders Modified", PrintTo
If chPlanObjFiled Then DoReport "Workflow Objections Filed", PrintTo
If chBKtoFC Then DoReport "Workflow BK to FC", PrintTo
If chPOCToBeFiled Then
    If PrintTo = -3 Then
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputReport, "Workflow POC ToBeField", acFormatXLS, TemplatePath & "POC.xlt", 1, , False
     DoCmd.SetWarnings True
    Else
    DoReport "Workflow POC to be Filed", PrintTo
    
    End If
End If
If chTitleReviewOutstanding Then DoReport "Workflow BK Title Review Outstanding", PrintTo
If chPending Then DoReport "Workflow Files in Pending Status", PrintTo
If ChHearingScheduledFC Then DoReport "Workflow Hearing Scheduled FC", PrintTo

'added on 7_21_15

If ckPrepareDraftCom Then DoReport "Workflow Prepare Draft Complaint", PrintTo
If chApprovDrftComp Then DoReport "Workflow Approved Draft Complaint", PrintTo
If chDCHearingScheduled Then DoReport "Workflow Hearing Scheduled DC", PrintTo
If chFeesCostSent Then DoReport "Workflow LossMitigation_FeesCost", PrintTo
If ChServiceDeadline Then DoReport "Workflow Service Deadline DC", PrintTo
If chLinetostayaction Then DoReport "Workflow Line to stay action", PrintTo


If chAssignmentNotReceived Then DoReport "Workflow BK Assignments Not Received", PrintTo
If chAssignmentNotSentToCourt Then DoReport "Workflow BK Assignments Not Sent To Court", PrintTo
If chPOCWaitForStatus Then DoReport "Workflow POC Status", PrintTo
If chHearingsPOC Then DoReport "Workflow Hearings POC", PrintTo
If chPlansToReview Then DoReport "Workflow Plans to Review", PrintTo
If chPlanObjToFile Then DoReport "Workflow Objections to File", PrintTo
If chObjNoResp Then DoReport "Workflow POC Response", PrintTo
If chHearingsPlan Then DoReport "Workflow Hearings Plan", PrintTo
If chReaffToBeSent Then DoReport "Workflow Reaff To Send", PrintTo
If chReaffSentToClient Then DoReport "Workflow Reaff Send To Client", PrintTo
If chReaffToBeApproved Then DoReport "Workflow Reaff Not Approved", PrintTo
If chReaffFiled Then DoReport "Workflow Reaff to be Filed", PrintTo
If chCDNeedDeadline Then DoReport "Workflow CD Need Deadline", PrintTo
If chCDAnswered Then DoReport "Workflow CD To Be Answered", PrintTo
If chCDNeedHearing Then DoReport "Workflow CD Need Hearing", PrintTo
If Me.chCDHearingSet Then DoReport "Workflow CD Hearing Set", PrintTo
If chLoanModRefRecd Then DoReport "Workflow Loan Mods", PrintTo

If chTtlPayChgToBeSent Then DoReport "Workflow Title Pay Change To Send", PrintTo
If chFHLMCConvert Then DoReport "Workflow FHLMC BK To be Filed", PrintTo
If chBKFHLMCActive Then DoReport "Workflow BK FHLMC Active Cases", PrintTo
If chFHLMCChpt13 Then DoReport "Workflow BK FHLMC Chpt 13", PrintTo
If chFHLMCChpt7 Then DoReport "Workflow BK FHLMC Chpt 7", PrintTo
If chFHLMCDefCuredMon Then DoReport "Workflow BK FHLMC Def Cured Monitor", PrintTo
If chCashForKeys Then DoReport "Workflow Eviction Cash For Keys", PrintTo
If chLockoutScheduled Then DoReport "Workflow Eviction Lockout Scheduled", PrintTo
If chLockoutScheduledBalt Then DoReport "Workflow Eviction Lockout Scheduled BaltCi", PrintTo

If chEVRequestDocs Then DoReport "Workflow Eviction Request Docs", PrintTo
If chEVwritsdue Then DoReport "Workflow Eviction Writs Due", PrintTo
If chEVWaitForDocs Then DoReport "Workflow Eviction Wait for Docs", PrintTo
If chEVNotices Then DoReport "Workflow Eviction Notices", PrintTo
If chEVExpNotices Then DoReport "Workflow Eviction Notices Expired", PrintTo
If chEvictions Then
    If PrintTo = -3 Then
      DoReport "Eviction Status", PrintTo
    Else
      If chDC Then DoReport "Workflow Eviction DC", PrintTo
      If chVirginia Then DoReport "Workflow Eviction VA", PrintTo
      If chMaryland Then
        DoReport "Workflow Eviction MD Circuit", PrintTo
        'DoReport "Workflow Eviction MD District", PrintTo
      End If
      
    End If
End If

'If chEVComplaintFiled Then DoReport "Workflow Eviction Complaint Filed", PrintTo
'If chEVComplaintFiledHearingNotSet Then DoReport "Workflow Eviction Complaint Filed Hearing Date", PrintTo
'If chEVComplaintNotServed Then DoReport "Workflow Eviction Complaint Not Served", PrintTo
If chEvictionsComplete Then DoReport "Workflow Eviction Complete", PrintTo
If chEVHearings Then DoReport "Workflow Eviction Hearings", PrintTo
'If chEVShowCause Then DoReport "Workflow Eviction Balt Show Cause", PrintTo
'If chEVWaitJudge Then DoReport "Workflow Eviction Balt Wait Judgment", PrintTo
If chEVVADeedRecording Then DoReport "Workflow Eviction VA Deed Recording", PrintTo
If chRentLeasesSent Then DoReport "Workflow Rent Leases To Be Sent", PrintTo
If chWaitLeases Then DoReport "Workflow Rent Waiting For Leases", PrintTo
If chBalanceDue Then DoReport "Workflow Rent Balances", PrintTo

'added 4_28_15
                            
If chNoticeOfAppearance Then DoReport "Workflow Notice of Appearance", PrintTo
'---
If chRestartsInProgress Then DoReport "Workflow Restarts in Progress", PrintTo
If chCaseClose Then DoReport "Workflow Cases to be Closed", PrintTo
If chCaseDismiss Then DoReport "Workflow Cases to be Dismissed", PrintTo
If chWaitReferral Then DoReport "Workflow Waiting for Referral", PrintTo
If chSaleNotSet Then DoReport "Workflow Sale Not Set", PrintTo
'If chAcceleration Then DoReport "Workflow Acceleration", PrintTo 'Removed 2/24/14 PER DIANE  MC
If ChNOI Then DoReport "Workflow NOI Expires", PrintTo
If chDocsNotSent Then DoReport "Workflow Docs not Sent", PrintTo
If chDocsOut Then DoReport "Workflow Docs Out", PrintTo
If chTitleOut Then DoReport "Workflow Title Orders Out", PrintTo
'If chTitleClaimsNotSent Then DoReport "Workflow Title Claims Not Sent", PrintTo
'If chTitleClaimsOut Then DoReport "Workflow Title Claims Out", PrintTo

'If chAssignNotSent Then DoReport "Workflow Title Assign Not Sent", PrintTo
'If chAssignToBeRecorded Then DoReport "Workflow Title Assign Not Recorded", PrintTo
If chDOANotSent Then DoReport "Workflow DOA not Sent", PrintTo
If chDOANotRecorded Then DoReport "Workflow DOA not Recorded", PrintTo
If chNOINotSent Then DoReport "Workflow NOI", PrintTo
If chServiceNotSent Then DoReport "Workflow Service not Sent", PrintTo
'If chTitleToBeReviewed Then DoReport "Workflow Title To Be Reviewed", PrintTo
If chNotServed Then DoReport "Workflow not Served", PrintTo
If chSentToDocket Then DoReport "Workflow Sent to Docket", PrintTo
If chFinalLMA Then DoReport "Workflow Final LMA To Be Filed", PrintTo
If chDocket Then DoReport "Workflow No Docket", PrintTo
If chNotices Then DoReport "Workflow Notice Problems", PrintTo
If chSendNotices Then DoReport "Workflow Send Notices", PrintTo
If chIRSNotices Then DoReport "Workflow IRS Notices", PrintTo
'If chIRSNoticeSale Then DoReport "Workflow IRS Notices Sale Scheduled", PrintTo
If chFirstPub Then DoReport "Workflow First Pub", PrintTo
If ChCertPub Then DoReport "WorkflowCertOfPub", PrintTo
If chSaleNotScheduled Then DoReport "Workflow Sales not Scheduled", PrintTo
If chTitleOrderBeforeSale Then DoReport "Workflow Title Before Sale", PrintTo
If chDisposition Then DoReport "Workflow Disposition Missing", PrintTo
If chServicerRelease Then DoReport "Workflow Servicer Released", PrintTo
If chDispositionRescinded Then DoReport "Workflow Disposition Rescinded", PrintTo
If chTitleGood30Days Then DoReport "Workflow Title Good Through", PrintTo
If chSaleNoDocs Then DoReport "Workflow Sale No Docs", PrintTo
If chBidNeeded Then DoReport "Workflow Bid Needed", PrintTo
'If chBids Then DoReport "Workflow Bids", PrintTo
If chReportSale Then DoReport "Workflow Reports of Sale Due", PrintTo
If chNotRat Then DoReport "Workflow Not Ratified", PrintTo
If chStatePropReg Then DoReport "WorkFlow StatePropReg", PrintTo

If chPR Then
    If chMaryland Then DoReport "Workflow Property Registration", PrintTo
    If chVirginia Then DoReport "Workflow Property RegistrationVA", PrintTo
End If

If chClientNotPaid Then DoReport "Workflow Reinstated Client Not Paid", PrintTo
If chExceptionsFiled Then DoReport "Workflow Exceptions Filed", PrintTo
If chAuditsDue Then
    If chMaryland Then DoReport "Workflow Audits Due MD", PrintTo
    If chVirginia Then DoReport "Workflow Audits Due VA", PrintTo
End If
If chAuditsNotApproved Then DoReport "Workflow Audits not Approved", PrintTo
If chAudits3Pty Then DoReport "Workflow Audits 3rd Party Sales", PrintTo
If chDeedsNotSent Then DoReport "Workflow Deeds not Sent", PrintTo
If chDeedAppOut Then DoReport "WOrkflow DOA Out", PrintTo
If chRealPropTaxes Then DoReport "Workflow Real Property Taxes Outstanding", PrintTo
If chDeedsNotRecorded Then DoReport "Workflow Deeds not Recorded", PrintTo

'added 5/29/15
If chTitMonitorSale Then DoReport "Workflow MonitorSale", PrintTo
If chCancelServiceduetodisposition Then DoReport "Workflow Cancel Service Due to Disposition", PrintTo

If chFinalPackages Then
    DoReport "Workflow Final Packages HUD", PrintTo
    DoReport "Workflow Final Packages VA", PrintTo
End If
If chNiSi Then DoReport "Workflow NiSi", PrintTo
If ch3rdParty Then DoReport "Workflow 3rd Party Not Settled", PrintTo
If ch3PtyClientNotPaid Then DoReport "Workflow 3rd Party Client Not Paid", PrintTo
If chReSale Then DoReport "Workflow Resale", PrintTo
If chOnHold Then DoReport "Workflow On Hold", PrintTo
If chDeceased Then DoReport "Workflow Deceased", PrintTo
If chFCVADeedRecording Then DoReport "Workflow Eviction VA Deed Recording", PrintTo
If chTitleDeedCorrection Then DoReport "Workflow Title Deed Correction", PrintTo
If chFairDebtNeedTitleOrdered Then DoReport "Workflow Fair Debt Title Order Needed", PrintTo

If chAssignmentNeeded Then DoReport "Workflow Assignment Needed", PrintTo
If chAssignmentfromClient Then DoReport "Workflow Assignment Not Received from Client", PrintTo
If chAssignmentToBeSent Then DoReport "Workflow Assignment Not Sent", PrintTo
If chAssignmentNotRecorded Then DoReport "Workflow Assignment Not Recorded", PrintTo
'If chCancelServiceduetodisposition Then DoReport "Workflow Cancel Service Due to Disposition", PrintTo


If chColStatus Then DoReport "Workflow Collection Status", PrintTo
If chColNoComplaint Then DoReport "Workflow Collection No Complaint", PrintTo
If chColServiceDue Then DoReport "Workflow Collection Service", PrintTo
If chColAnswerDue Then DoReport "Workflow Collection Answer Due", PrintTo
If chColHearings Then DoReport "Workflow Collection Hearings", PrintTo
If chColPostJudgment Then DoReport "Workflow Collection Post-Judgment", PrintTo

If chReferrals Then DoReport "Workflow Referrals", PrintTo
If chReferralsBK Then DoReport "Workflow Referrals BK", PrintTo
If chMonSale Then DoReport "Workflow Monitor Sale", PrintTo
If chLimbo Then DoReport "Workflow Files in Limbo", PrintTo
If chFNMAFC Then DoReport "Workflow FNMA FC", PrintTo
If chFNMABK Then DoReport "Workflow FNMA BK", PrintTo
If chFNMAHolds Then DoReport "Workflow FNMA Holds", PrintTo
If chFNMAMissingDocs Then DoReport "Workflow FNMA Missing Docs", PrintTo
If chFNMAPostponements Then DoReport "Workflow FNMA Postponements", PrintTo
If chFHLMCOpenFiles Then DoReport "Workflow FHLMC Open Files", PrintTo
If chConflicts Then DoReport "Workflow Conflicts Pending", PrintTo
If chDocRequest Then DoReport "Workflow Document Request Pending", PrintTo

If chCompiledExhibits Then DoReport "Workflow Compiled Exhibits", PrintTo

If ACation Or AWaitingDoc Or AWitingBill Or ATitleIssue Or AStop Then DoReport "wkfAttribue", PrintTo

'Marko
If ChCIVAllLitigation Then
    Call OutputExcel("Workflow Civil ALl Litigation", "wkflCIVLitigation")

   ' DoEvents
   ' Call ModifyExportedExcelFileFormats("C:\Database" & "\Export_" & Format(Date, "yyyymmmdd") & ReportName & ".xls")
End If



If chFNMACombined Then
  If PrintTo = -3 Then 'excel with combining reports into one excel spreadsheet
    DoReport "Workflow FNMA Combined", PrintTo
  Else
    DoReport "Workflow FNMA FC", PrintTo
    DoReport "Workflow FNMA BK", PrintTo
    DoReport "Workflow FNMA Holds", PrintTo
    DoReport "Workflow FNMA Missing Docs", PrintTo
    DoReport "Workflow FNMA Postponements", PrintTo
  End If

End If

If chDCLossMediation Then
    DoCmd.SetWarnings False

    DoCmd.OpenQuery ("rqryLossMediation_DC")
    DoCmd.OpenQuery ("DeleteCDMediation")
    DoCmd.OpenQuery ("AppendDCMediation")
    DoCmd.OpenQuery ("UpdateMediationCOMMENT")
    DoCmd.OpenQuery ("UpdateMediation")
    DoCmd.OpenQuery ("UpdatePostMediation")
    DoCmd.OpenQuery ("UpdatePreMediation")
    DoCmd.OpenQuery ("UpdatePostMediation")
    DoCmd.OpenQuery ("UpdateMediationDis")
    DoCmd.OpenQuery ("UpdatePreMediationDis")
    DoCmd.OpenQuery ("UpdatePostMediationDis")
    
    DoReport "Workflow Loss Mediation_DC", PrintTo
End If

If chNeedToInvoiceFC Then DoReport "Workflow Need To Invoice FC", PrintTo
If chNeedToInvoiceFCnew Then DoReport "Workflow Need To Invoice FC New", PrintTo
If chNeedToInvoiceBK Then DoReport "Workflow Need To Invoice BK", PrintTo
If chNeedToInvoiceEV Then DoReport "Workflow Need To Invoice EV", PrintTo
If chNeedtoInvoiceServicerReleased Then DoReport "Workflow Need To Invoice Servicer Released", PrintTo
If chNeedToInvoiceRent Then DoReport "Workflow Need To Invoice Rent", PrintTo
If chNeedToInvoiceTR Then DoReport "Workflow Need To Invoice TR", PrintTo
If chNeedtoInvoiceTitle Then DoReport "Workflow Need To Invoice Title", PrintTo
If chDeedCalc Then DoReport "Workflow Deed Calc Not entered", PrintTo
If chReceivable_Litigation Then DoReport "Workflow Receivables_Litigation", PrintTo
If chReceivable_PSAdvanced Then DoReport "Workflow Receivables_PSAdvanced", PrintTo
If chAttribBills Then DoReport "Workflow Waiting for Bills", PrintTo
If chNeedInvoiceDIL Then DoReport "Workflow Need To Invoice DIL", PrintTo
'added on 5/11/15
If chFCMonitor Then DoReport "Workflow Need To Invoice FCMonitor", PrintTo
'-------
If chReceivables Then
  Select Case Me.cboARRecAll
    Case "Date", "File"
      DoReport "Workflow Receivables Unpaid Invoices Amount", PrintTo
    Case Else
     ' DoReport "Workflow Receivables_Group", PrintTo
  End Select

End If
If chReceivablesDate Then
  Select Case Me.cboARRecAll
    Case "Date", "File"
      DoReport "Workflow Receivables Unpaid Invoices date", PrintTo
    Case Else
     ' DoReport "Workflow Receivables_Group", PrintTo
  End Select

End If


If chReceivables_servicerreleased Then
DoReport "Workflow Receivables_ServicerReleased", PrintTo
End If

If chReceivables_FC Then
  Select Case Me.cboARRecAll
    Case "Date", "File"
      DoReport "Workflow Receivables_FC", PrintTo
    Case Else
      DoReport "Workflow Receivables_FC_Group", PrintTo
  End Select
End If

If chReceivables_BK Then
  Select Case Me.cboARRecAll
    Case "Date", "File"
      DoReport "Workflow Receivables_BK", PrintTo
    Case Else
      DoReport "Workflow Receivables_BK_Group", PrintTo
  End Select
End If

If chReceivables_EV Then
  Select Case Me.cboARRecAll
    Case "Date", "File"
      DoReport "Workflow Receivables_EV", PrintTo
    Case Else
      DoReport "Workflow Receivables_EV_Group", PrintTo
  End Select
End If

'added on 4/22/15 Lin
    
If chReceivables_DIL Then
  Select Case Me.cboARRecAll
    Case "Date", "File"
      DoReport "Workflow Receivables_DIL", PrintTo
    Case Else
      DoReport "Workflow Receivables_DIL_Group", PrintTo
  End Select
End If
'....
'chFCMonitorReceivable
'addedon 5/12/15
If chFCMonitorReceivable Then
  Select Case Me.cboARRecAll
    Case "Date", "File"
      DoReport "Workflow Receivables_FCMonitor", PrintTo
    Case Else
      DoReport "Workflow Receivables_FCMonitor_Group", PrintTo
  End Select
End If

If chReceivables_OTH Then
  Select Case Me.cboARRecAll
    Case "Date", "File"
      DoReport "Workflow Receivables_OTH", PrintTo
    Case Else
      DoReport "Workflow Receivables_OTH_Group", PrintTo
  End Select

End If

If chDisbursingSurplus Then DoReport "Workflow Disbursing Surplus", PrintTo
If chTaxRefShort Then DoReport "Workflow Accounting Tax Refund Shortage", PrintTo
If chWrittenOff Then DoReport "Workflow Written Off Invoices", PrintTo
If chCRPending Then DoReport "Workflow Check Request Pending", PrintTo
If chClientNotPaid2 Then DoReport "Workflow Reinstated Client Not Paid", PrintTo
If chREO_TitleOrderOut Then DoReport "Workflow REO Title Not Recd", PrintTo
If chREO_FC_Out Then DoReport "Workflow REO Docs Not Recd", PrintTo
If chREO_Commitment Then DoReport "Workflow REO Commit Not Sent", PrintTo
If chREO_Contract Then DoReport "Workflow REO Commit Sent No Contract", PrintTo
If chREO_Close Then DoReport "Workflow REO Waiting To Close", PrintTo
If chREO_FileClose Then DoReport "Workflow REO Files Not Closed", PrintTo

If chFairDebtDispute Then DoReport "Workflow Fair Debt Dispute", PrintTo
If chReinstatementRequested Then DoReport "Workflow Reinstatement Requested", PrintTo
If chPayoffRequested Then DoReport "Workflow Payoff Requested", PrintTo

If chDILSendToBorrower Then DoReport "Workflow DIL Send to Borrowers", PrintTo
If chDILSendToClient Then DoReport "Workflow DIL Send to Client", PrintTo
If chDILSendToRecord Then DoReport "Workflow DIL Send to Record", PrintTo
If chDILTitleReview Then DoReport "Workflow DIL Title To be Reviewed", PrintTo
If chDILRecordedLandRecords Then DoReport "Workflow DIL Records Waiting", PrintTo
If chDILReceiptFromCLient Then DoReport "Workflow DIl Client Waiting", PrintTo
If chDILFromBorrower Then DoReport "Workflow DIL From Borrowers", PrintTo
If chDIL Then DoReport "Workflow DIL All", PrintTo

If chLossMediation Then DoReport "Workflow Loss Mediation", PrintTo

If chHUDFirstLegal Then DoReport "Workflow HUD First Legal", PrintTo
If chVAappraisal Then DoReport "Workflow VA Appraisal", PrintTo

If chTROpen Then DoReport "Workflow TR Open", PrintTo
If chTRIntroLtrNotSent Then DoReport "Workflow TR Intro Not Sent", PrintTo
If chInitNegPending Then DoReport "Workflow TR Negot Pending", PrintTo
If chComplaintNotFiled Then DoReport "Workflow TR Complaint Not Filed", PrintTo
If chTRToBeClosed Then DoReport "Workflow TR Need Closed", PrintTo


'12/22/14 Linda
If ckCIV_tobeclose Then DoReport "Workflow CVI  Need Closed", PrintTo


If Me.chAuditorFollowUp Then DoReport "Workflow Auditor Follow Up", PrintTo
If chAnswersDue Then DoReport "Workflow Answers Due", PrintTo
If ckJudgEnteredNeedSetSale Then DoReport "Workflow Judgment Entered Need Set Sale", PrintTo
If chJudgmentEnteredMonitorSetSale Then DoReport "Workflow Judgment Entered Monitor Set Sale", PrintTo


Exit Sub

Err_PrintDocs:
    MsgBox Err.Description
    Exit Sub

End Sub

Private Sub cmdAcrobat_Click()
Call PrintDocs(-2)
End Sub

Private Sub cmdExcel_Click()
Call PrintDocs(-3)
End Sub


Private Sub cmdForm_Click()
Call PrintDocs(55)
End Sub

Private Sub cmdInvoiceAll_Click()
    If Not CheckDatesOK Then Exit Sub
    If IsNull(PrintTo) Then
    MsgBox ("select format before run the report")
    Exit Sub
    End If
    DoReport "workflowInvoiceAll", PrintTo

'If PrintTo = -3 Then
'
'    'On Error Resume Next
'    'Kill "S:\ProductionReporting\InvoicesbyOpenDate" & Format$(Now(), "yyyymmdd") & ".xls"
'    On Error GoTo 0
'    'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qrywkflInvoiceAll", "S:\ProductionReporting\InvoicesOpenDate" & Format$(Now(), "yyyymmdd") & ".xls"
'    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qrywkflInvoiceAllExcel", "H:\Accounting project\InvoicesbyOpenDate" & Format$(Now(), "yyyymmdd") & ".xls"
'
'    Dim ExcelObj As Object
'    Set ExcelObj = CreateObject("Excel.Application")
'    With ExcelObj
'    .Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModule.xlsm"
'    .Run "InvoicesbyOpenDate"
'    .ActiveWorkbook.Close
'    '.Visible
'    End With
'    'Set ExcelObj = Nothing
'    MsgBox "The report is now ready to view in Excel"
'
'    Else
'        DoReport "workflowInvoiceAll", PrintTo
'    End If
End Sub

Private Sub cmdLastMonth_Click()
On Error GoTo Err_cmdLastMonth_Click
DateFrom = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), 1)
DateThru = DateAdd("d", (Day(Date)) * -1, Date)

Exit_cmdLastMonth_Click:
    Exit Sub

Err_cmdLastMonth_Click:
    MsgBox Err.Description
    Resume Exit_cmdLastMonth_Click

End Sub

Private Sub cmdLastWeek_Click()
Dim mon As Date

On Error GoTo Err_cmdLastWeek_Click
'What was this past Monday?
mon = DateAdd("d", 1 - Weekday(Now(), vbMonday), Now())

DateFrom = DateAdd("d", -7, mon)
DateThru = DateAdd("d", -3, mon)

Exit_cmdLastWeek_Click:
    Exit Sub

Err_cmdLastWeek_Click:
    MsgBox Err.Description
    Resume Exit_cmdLastWeek_Click
End Sub

Private Sub cmdLossMitigation_Click()
On Error GoTo Err_cmdLossMitigation_Click

chFairDebtDispute = True
chReinstatementRequested = True
chPayoffRequested = True

chDIL = True
chDILRecordedLandRecords = True
chDILSendToRecord = True
chDILReceiptFromCLient = True
chDILSendToClient = True
chDILFromBorrower = True
chDILSendToBorrower = True
chDILTitleReview = True


chLossMediation = True

TabCtl.Value = 7

Exit_cmdLossMitigation_Click:
    Exit Sub

Err_cmdLossMitigation_Click:
    MsgBox Err.Description
    Resume Exit_cmdLossMitigation_Click
    

End Sub


Private Sub cmdPaid_Click()

If Not CheckDatesOK Then Exit Sub
    If IsNull(PrintTo) Then
    MsgBox ("select format before run the report")
    Exit Sub
    End If
    DoReport "workflowInvoicePaid", PrintTo

End Sub

Private Sub cmdSelTitleResolution_Click()
On Error GoTo Err_cmdSelTitleResolution_Click

chTROpen = True
chTRIntroLtrNotSent = True
chInitNegPending = True
chComplaintNotFiled = True
chTRToBeClosed = True

TabCtl.Value = 8

Exit_cmdSelTitleResolution_Click:
  Exit Sub
  
Err_cmdSelTitleResolution_Click:
  MsgBox Err.Description
  Resume Exit_cmdSelTitleResolution_Click
  
End Sub

Private Sub cmdThisMonth_Click()
On Error GoTo Err_cmdThisMonth_Click
DateFrom = DateAdd("d", (Day(Date) - 1) * -1, Date)
DateThru = DateAdd("d", -1, DateAdd("m", 1, DateFrom))

Exit_cmdThisMonth_Click:
    Exit Sub

Err_cmdThisMonth_Click:
    MsgBox Err.Description
    Resume Exit_cmdThisMonth_Click
    
End Sub

Private Sub cmdThisYear_Click()
On Error GoTo Err_cmdThisYear_Click
DateFrom = DateSerial(Year(Date), 1, 1)
Dim nextYear As Variant
nextYear = DateAdd("yyyy", 1, DateFrom)
DateThru = DateAdd("d", -1, nextYear)


Exit_cmdThisYear_Click:
    Exit Sub

Err_cmdThisYear_Click:
    MsgBox Err.Description
    Resume Exit_cmdThisYear_Click
End Sub

Private Sub cmdToday_Click()
On Error GoTo Err_cmdToday_Click
DateFrom = Date
DateThru = Date


Exit_cmdToday_Click:
    Exit Sub

Err_cmdToday_Click:
    MsgBox Err.Description
    Resume Exit_cmdToday_Click
End Sub

Private Sub cmdWord_Click()
Call PrintDocs(-1)
End Sub

Private Sub cmdPrint_Click()
Call PrintDocs(acViewNormal)
End Sub

Private Sub cmdView_Click()
Call PrintDocs(acViewPreview)
End Sub

Private Sub cmdAll_Click()
On Error GoTo Err_cmdAll_Click

Me.chAuditorFollowUp = True
chFiledLisPendens = True
chComplaintFiled = True
chComplaintToCourt = True
chClientComplaintReturned = True
chPrepComplaint = True
ch362ToBeFiled = True
chBankruptcy = True
chMotionsOut = True
chAffDefault = True
chDefaultsCured = True
chHearings362 = True
chHearingsScheduled = True
chConsentModified = True
chPlanObjFiled = True
chBKtoFC = True
chPOCToBeFiled = True
chTitleReviewOutstanding = True
chAssignmentNotReceived = True
chAssignmentNotSentToCourt = True
chPOCWaitForStatus = True
chHearingsPOC = True
chPlansToReview = True
chPlanObjToFile = True
chObjNoResp = True
chHearingsPlan = True
chFCVADeedRecording = True
chTitleDeedCorrection = True
chReaffToBeSent = True
chReaffSentToClient = True
chReaffToBeApproved = True
chReaffFiled = True
chCDNeedDeadline = True
chCDAnswered = True
chCDNeedHearing = True
chCDHearingSet = True
chLoanModRefRecd = True

chTtlPayChgToBeSent = True
chFHLMCConvert = True
chFairDebtNeedTitleOrdered = True
chBKFHLMCActive = True
chFHLMCChpt13 = True
chFHLMCChpt7 = True
chFHLMCDefCuredMon = True
chAssignmentNeeded = True
chAssignmentToBeSent = True
chAssignmentNotRecorded = True


chEVRequestDocs = True
chEVWaitForDocs = True
chEVNotices = True
chEVExpNotices = True
chEvictions = True
chEVComplaintFiled = True
chEVComplaintFiledHearingNotSet = True
chEVComplaintNotServed = True
chEvictionsComplete = True
chEVHearings = True
chEVShowCause = True
chEVWaitJudge = True
chEVVADeedRecording = True
chCashForKeys = True
chLockoutScheduled = True
Me.chLockoutScheduledBalt = True

chRestartsInProgress = True
chCaseClose = True
chCaseDismiss = True
ChHearingScheduledFC = True

chWaitReferral = True
chSaleNotSet = True
'chAcceleration = True
ChNOI = True
chDocsNotSent = True
chDocsOut = True
chTitleOut = True
chTitleClaimsNotSent = True
chTitleClaimsOut = True

'chAssignNotSent = True
'chAssignToBeRecorded = True
chTitleToBeReviewed = True
chDOANotSent = True
chDOANotRecorded = True
chNOINotSent = True
chServiceNotSent = True
chNotServed = True
chSentToDocket = True
chFinalLMA = True
chDocket = True
chNotices = True
chSendNotices = True
chIRSNotices = True
'chIRSNoticeSale = True
chFirstPub = True
ChCertPub = True
chSaleNotScheduled = True
chTitleOrderBeforeSale = True
chDisposition = True
chDispositionRescinded = True
chTitleGood30Days = True
chSaleNoDocs = True
chBidNeeded = True
'chBids = True
chReportSale = True
chNotRat = True
chStatePropReg = True
chClientNotPaid = True
chExceptionsFiled = True
chAuditsDue = True
chAuditsNotApproved = True
chAudits3Pty = True
chDeedsNotSent = True
chDeedAppOut = True
chRealPropTaxes = True
chDeedsNotRecorded = True
chFinalPackages = True
chNiSi = True
ch3rdParty = True
ch3PtyClientNotPaid = True
chReSale = True
chOnHold = True
chDeceased = True

chColStatus = True
chColNoComplaint = True
chColServiceDue = True
chColAnswerDue = True
chColHearings = True
chColPostJudgment = True

chReferrals = True
chReferralsBK = True
chMonSale = True
chLimbo = True
chFNMAFC = True
chFNMABK = True
chFNMACombined = True
chFHLMCOpenFiles = True
chConflicts = True
chDocRequest = True
chNeedInvoiceDIL = True

chNeedToInvoiceFC = True
chNeedToInvoiceBK = True
chNeedToInvoiceEV = True
chNeedToInvoiceRent = True
chNeedToInvoiceTR = True

chAttribBills = True
chReceivables = True
chReceivables_FC = True
chReceivables_BK = True
chReceivables_EV = True
chReceivables_OTH = True
chTaxRefShort = True
chCRPending = True
chClientNotPaid2 = True

chREO_TitleOrderOut = True
chREO_FC_Out = True
chREO_Commitment = True
chREO_Contract = True
chREO_Close = True
chREO_FileClose = True

chFairDebtDispute = True
chReinstatementRequested = True
chPayoffRequested = True

chDIL = True
chDILRecordedLandRecords = True
chDILSendToRecord = True
chDILReceiptFromCLient = True
chDILSendToClient = True
chDILFromBorrower = True
chDILSendToBorrower = True
chDILTitleReview = True

chLossMediation = True

chTROpen = True
chTRIntroLtrNotSent = True
chInitNegPending = True
chComplaintNotFiled = True
chTRToBeClosed = True
'12/22/14
ckCIV_tobeclose = True

chAnswersDue = True
chPR = True
ChFeesAndCosts = True
chReceivable_Litigation = True
chReceivable_PSAdvanced = True

chFCMonitorReceivable = True
chFCMonitor = True


Exit_cmdAll_Click:
    Exit Sub

Err_cmdAll_Click:
    MsgBox Err.Description
    Resume Exit_cmdAll_Click
    
End Sub

Private Sub cmdNone_Click()
On Error GoTo Err_cmdNone_Click

Me.chAuditorFollowUp = False
chFiledLisPendens = False
chComplaintFiled = False
chComplaintToCourt = False
chClientComplaintReturned = False
chPrepComplaint = False
ch362ToBeFiled = False
chBankruptcy = False
chMotionsOut = False
chAffDefault = False
chDefaultsCured = False
chHearings362 = False
chHearingsScheduled = False
chPlanObjFiled = False
chConsentModified = False
chBKtoFC = False
chPOCToBeFiled = False
chTitleReviewOutstanding = False
chAssignmentNotReceived = False
chAssignmentNotSentToCourt = False
chPOCWaitForStatus = False
chHearingsPOC = False
chPlansToReview = False
chPlanObjToFile = False
chObjNoResp = False
chHearingsPlan = False
chFCVADeedRecording = False
chTitleDeedCorrection = False
chReaffToBeSent = False
chReaffSentToClient = False
chReaffToBeApproved = False
chReaffFiled = False
chCDNeedDeadline = False
chCDAnswered = False
chCDNeedHearing = False
chCDHearingSet = False
chLoanModRefRecd = False
chTtlPayChgToBeSent = False
chFHLMCConvert = False
chFairDebtNeedTitleOrdered = False
chBKFHLMCActive = False
chFHLMCChpt13 = False
chFHLMCChpt7 = False
chFHLMCDefCuredMon = False

chAssignmentNeeded = False
chAssignmentToBeSent = False
chAssignmentNotRecorded = False


chEVRequestDocs = False
chEVWaitForDocs = False
chEVNotices = False
chEVExpNotices = False
chEvictions = False
chEVComplaintFiled = False
chEVComplaintFiledHearingNotSet = False
chEVComplaintNotServed = False
chEvictionsComplete = False
chEVHearings = False
chEVShowCause = False
chEVWaitJudge = False
chEVVADeedRecording = False
chCashForKeys = False
chLockoutScheduled = False
chLockoutScheduledBalt = False

chRestartsInProgress = False
chCaseClose = False
chCaseDismiss = False
ChHearingScheduledFC = False

chWaitReferral = False
chSaleNotSet = False
'chAcceleration = False
ChNOI = False
chDocsNotSent = False
chDocsOut = False
chTitleOut = False
chTitleClaimsNotSent = False
chTitleToBeReviewed = False
chTitleClaimsOut = False

'chAssignNotSent = False
'chAssignToBeRecorded = False
chDOANotSent = False
chDOANotRecorded = False
chNOINotSent = False
chServiceNotSent = False
chNotServed = False
chSentToDocket = False
chFinalLMA = False
chDocket = False
chNotices = False
chSendNotices = False
chIRSNotices = False
'chIRSNoticeSale = False
chFirstPub = False
ChCertPub = False
chSaleNotScheduled = False
chTitleOrderBeforeSale = False
chDisposition = False
chDispositionRescinded = False
chSaleNoDocs = False
chBidNeeded = False
'chBids = False
chTitleGood30Days = False
chReportSale = False
chNotRat = False
chStatePropReg = False
chClientNotPaid = False
chExceptionsFiled = False
chAuditsDue = False
chAuditsNotApproved = False
chAudits3Pty = False
chDeedsNotSent = False
chDeedAppOut = False
chRealPropTaxes = False
chDeedsNotRecorded = False
chFinalPackages = False
chNiSi = False
ch3rdParty = False
ch3PtyClientNotPaid = False
chReSale = False
chOnHold = False
chDeceased = False

chColStatus = False
chColNoComplaint = False
chColServiceDue = False
chColAnswerDue = False
chColHearings = False
chColPostJudgment = False

chReferrals = False
chReferralsBK = False
chMonSale = False
chLimbo = False
chFNMAFC = False
chFNMABK = False
chFNMACombined = False
chFHLMCOpenFiles = False
chConflicts = False
chDocRequest = False

chNeedToInvoiceFC = False
chNeedToInvoiceBK = False
chNeedToInvoiceEV = False
chNeedToInvoiceRent = False
chNeedToInvoiceTR = False
chNeedInvoiceDIL = False
chAttribBills = False
chReceivables = False
chReceivables_FC = False
chReceivables_BK = False
chReceivables_EV = False
chReceivables_OTH = False
chTaxRefShort = False
chCRPending = False
chClientNotPaid = False
chClientNotPaid2 = False

chREO_TitleOrderOut = False
chREO_FC_Out = False
chREO_Commitment = False
chREO_Contract = False
chREO_Close = False
chREO_FileClose = False

chFairDebtDispute = False
chReinstatementRequested = False
chPayoffRequested = False

chDIL = False
chDILRecordedLandRecords = False
chDILSendToRecord = False
chDILReceiptFromCLient = False
chDILSendToClient = False
chDILFromBorrower = False
chDILSendToBorrower = False
chDILTitleReview = False

chLossMediation = False

chTROpen = False
chTRIntroLtrNotSent = False
chInitNegPending = False
chComplaintNotFiled = False
chTRToBeClosed = False
'12/22/14

ckCIV_tobeclose = False

chAnswersDue = False
chPR = False
ChFeesAndCosts = False
chReceivable_Litigation = False
chReceivable_PSAdvanced = False
chFCMonitorReceivable = False
chFCMonitor = False



Exit_cmdNone_Click:
    Exit Sub

Err_cmdNone_Click:
    MsgBox Err.Description
    Resume Exit_cmdNone_Click
    
End Sub

Private Sub cmdSelectFC_Click()
On Error GoTo Err_cmdSelectFC_Click

chRestartsInProgress = True
chWaitReferral = True
chSaleNotSet = True
'chAcceleration = True
ChNOI = True
chDocsNotSent = True
chDocsOut = True
chTitleOut = True
chTitleClaimsNotSent = True
chTitleToBeReviewed = True
chTitleClaimsOut = True

'chAssignNotSent = True
'chAssignToBeRecorded = True
chDOANotSent = True
chDOANotRecorded = True
chNOINotSent = True
chServiceNotSent = True
chNotServed = True
chSentToDocket = True
chFinalLMA = True
chDocket = True
chNotices = True
chSendNotices = True
chIRSNotices = True
'chIRSNoticeSale = True
chFirstPub = True
ChCertPub = True
chSaleNotScheduled = True
chTitleOrderBeforeSale = True
chDisposition = True
chDispositionRescinded = True
chSaleNoDocs = True
chBidNeeded = True
'chBids = True
chTitleGood30Days = True
chReportSale = True
chNotRat = True
chStatePropReg = True
chClientNotPaid = True
chExceptionsFiled = True
chAuditsDue = True
chAuditsNotApproved = True
chAudits3Pty = True
chDeedsNotSent = True
chDeedAppOut = True
chRealPropTaxes = True
chDeedsNotRecorded = True
chFinalPackages = True
chNiSi = True
ch3rdParty = True
ch3PtyClientNotPaid = True
chReSale = True
chCaseClose = True
chCaseDismiss = True
ChHearingScheduledFC = True

chOnHold = True
chDeceased = True
chFCVADeedRecording = True
chTitleDeedCorrection = True
chFairDebtNeedTitleOrdered = True
chAssignmentNeeded = True
chAssignmentToBeSent = True
chAssignmentNotRecorded = True
chPR = True
Me.chAuditorFollowUp = True

TabCtl.Value = 0

Exit_cmdSelectFC_Click:
    Exit Sub

Err_cmdSelectFC_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectFC_Click
    
End Sub

Private Sub cmdSelectBK_Click()

On Error GoTo Err_cmdSelectBK_Click

ch362ToBeFiled = True
chBankruptcy = True
chMotionsOut = True
chAffDefault = True
chDefaultsCured = True
chHearings362 = True
chHearingsScheduled = True
chConsentModified = True
chPlanObjFiled = True
chBKtoFC = True
chPOCToBeFiled = True
chTitleReviewOutstanding = True
chAssignmentNotReceived = True
chAssignmentNotSentToCourt = True

chPOCWaitForStatus = True
chHearingsPOC = True
chPlansToReview = True
chPlanObjToFile = True
chObjNoResp = True
chHearingsPlan = True
chReaffToBeSent = True
chReaffSentToClient = True
chReaffToBeApproved = True
chReaffFiled = True
chCDNeedDeadline = True
chCDAnswered = True
chCDNeedHearing = True
chCDHearingSet = True
chLoanModRefRecd = True
chTtlPayChgToBeSent = True
chFHLMCConvert = True
chBKFHLMCActive = True
chFHLMCChpt13 = True
chFHLMCChpt7 = True
chFHLMCDefCuredMon = True

TabCtl.Value = 1

Exit_cmdSelectBK_Click:
    Exit Sub

Err_cmdSelectBK_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectBK_Click
    
End Sub

Private Sub cmdSelectEV_Click()

On Error GoTo Err_cmdSelectEV_Click

chEVRequestDocs = True
chEVWaitForDocs = True
chEVNotices = True
chEVExpNotices = True
chEvictions = True
chEVComplaintFiled = True
chEVComplaintFiledHearingNotSet = True
chEVComplaintNotServed = True
chEvictionsComplete = True
chEVHearings = True
chEVShowCause = True
chEVWaitJudge = True
chEVVADeedRecording = True
chCashForKeys = True
chLockoutScheduled = True
chLockoutScheduledBalt = True

TabCtl.Value = 2

Exit_cmdSelectEV_Click:
    Exit Sub

Err_cmdSelectEV_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectEV_Click
    
End Sub

Private Sub cmdSelectCOL_Click()

On Error GoTo Err_cmdSelectCOL_Click

chColStatus = True
chColNoComplaint = True
chColServiceDue = True
chColAnswerDue = True
chColHearings = True
chColPostJudgment = True

TabCtl.Value = 3

Exit_cmdSelectCOL_Click:
    Exit Sub

Err_cmdSelectCOL_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectCOL_Click
    
End Sub

Private Sub cmdSelOther_Click()

On Error GoTo Err_cmdOther_Click

chReferrals = True
chReferralsBK = True
chMonSale = True
chLimbo = True

chFNMAFC = True
chFNMABK = True
chFNMACombined = True
chFHLMCOpenFiles = True
chConflicts = True
chDocRequest = True

TabCtl.Value = 9

Exit_cmdOther_Click:
    Exit Sub

Err_cmdOther_Click:
    MsgBox Err.Description
    Resume Exit_cmdOther_Click
    
End Sub

Private Sub cmdSelectREO_Click()

On Error GoTo Err_cmdSelectREO_Click

chREO_TitleOrderOut = True
chREO_FC_Out = True
chREO_Commitment = True
chREO_Contract = True
chREO_Close = True
chREO_FileClose = True

TabCtl.Value = 4

Exit_cmdSelectREO_Click:
    Exit Sub

Err_cmdSelectREO_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectREO_Click
    
End Sub

Private Sub cmdAccounting_Click()

On Error GoTo Err_cmdAccounting_Click

chNeedToInvoiceFC = True
chNeedToInvoiceBK = True
chNeedToInvoiceEV = True
chNeedToInvoiceRent = True
chNeedToInvoiceTR = True
chNeedInvoiceDIL = True
chAttribBills = True
chReceivables = True
chReceivables_FC = True
chReceivables_BK = True
chReceivables_EV = True
chReceivables_OTH = True
chTaxRefShort = True
chCRPending = True
chClientNotPaid2 = True
chReceivable_Litigation = True
chReceivable_PSAdvanced = True


TabCtl.Value = 5

Exit_cmdAccounting_Click:
    Exit Sub

Err_cmdAccounting_Click:
    MsgBox Err.Description
    Resume Exit_cmdAccounting_Click
    
End Sub

Private Sub cmdCivil_Click()

On Error GoTo Err_cmdCivil_Click

chFiledLisPendens = True
chAnswersDue = True
chPrepComplaint = True
ChCIVAllLitigation = True
chClientComplaintReturned = True
chComplaintToCourt = True
chComplaintFiled = True

TabCtl.Value = 6

Exit_cmdCivil_Click:
    Exit Sub

Err_cmdCivil_Click:
    MsgBox Err.Description
    Resume Exit_cmdCivil_Click
    
End Sub


Private Sub Command460_Click()
On Error GoTo Err_Command460_Click

    Dim stDialStr As String
    Dim PrevCtl As Control
    Const ERR_OBJNOTEXIST = 2467
    Const ERR_OBJNOTSET = 91
    Const ERR_CANTMOVE = 2483

    Set PrevCtl = Screen.PreviousControl
    
    If TypeOf PrevCtl Is TextBox Then
      stDialStr = IIf(VarType(PrevCtl) > V_NULL, PrevCtl, "")
    ElseIf TypeOf PrevCtl Is ListBox Then
      stDialStr = IIf(VarType(PrevCtl) > V_NULL, PrevCtl, "")
    ElseIf TypeOf PrevCtl Is ComboBox Then
      stDialStr = IIf(VarType(PrevCtl) > V_NULL, PrevCtl, "")
    Else
      stDialStr = ""
    End If
    
    Application.Run "utility.wlib_AutoDial", stDialStr

Exit_Command460_Click:
    Exit Sub

Err_Command460_Click:
    If (Err = ERR_OBJNOTEXIST) Or (Err = ERR_OBJNOTSET) Or (Err = ERR_CANTMOVE) Then
      Resume Next
    End If
    MsgBox Err.Description
    Resume Exit_Command460_Click
    
End Sub
Private Sub cmdOpenFeesCosts_Click()
On Error Resume Next
Kill "S:\ProductionReporting\ReceivablesbyFeesCosts" & Format$(Now(), "yyyymmdd") & ".xls"
On Error GoTo 0
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryReceivablesbyFeeCost", "S:\ProductionReporting\ReceivablesbyFeesCosts" & Format$(Now(), "yyyymmdd") & ".xls"
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModule.xlsm"
.Run "ReceivablesbyFeesCosts"
.ActiveWorkbook.Close
'.Visible
End With
'Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"

End Sub
Private Sub cmdExportFinData_Click()
On Error Resume Next
Kill "S:\ProductionReporting\FinReports\FeesCostsbyVendorDaily" & Format$(Now(), "yyyymmdd") & ".xls"
Kill "S:\ProductionReporting\FinReports\FeesCostsDaily" & Format$(Now(), "yyyymmdd") & ".xls"
On Error GoTo 0
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryDailyFeesCostsbyVendor", "S:\ProductionReporting\FinReports\FeesCostsbyVendorDaily" & Format$(Now(), "yyyymmdd") & ".xls"
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryDailyFeesCosts", "S:\ProductionReporting\FinReports\FeesCostsDaily" & Format$(Now(), "yyyymmdd") & ".xls"
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Visible = True
.Workbooks.Open "S:\ProductionReporting\FinReports\FinancialReportingMenu.xlsm"

End With
End Sub

Private Sub cmdOpenAmtReport_Click()
On Error Resume Next
Kill "S:\ProductionReporting\ReceivablesbyOpenAmount" & Format$(Now(), "yyyymmdd") & ".xls"
On Error GoTo 0
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryReceivablesbyAmount", "S:\ProductionReporting\ReceivablesbyOpenAmount" & Format$(Now(), "yyyymmdd") & ".xls"
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModule.xlsm"
.Run "ReceivablesbyOpenAmount"
.ActiveWorkbook.Close
'.Visible
End With
'Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub
Private Sub cmdOpenDateReport_Click()
On Error Resume Next
Kill "S:\ProductionReporting\ReceivablesbyOpenDate" & Format$(Now(), "yyyymmdd") & ".xls"
On Error GoTo 0
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryReceivablesbyDate", "S:\ProductionReporting\ReceivablesbyOpenDate" & Format$(Now(), "yyyymmdd") & ".xls"
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModule.xlsm"
.Run "ReceivablesbyOpenDate"
.ActiveWorkbook.Close
'.Visible
End With
'Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub


Private Sub cmdYesterday_Click()
On Error GoTo Err_cmdYesterday_Click
DateFrom = Date - 1
DateThru = Date - 1


Exit_cmdYesterday_Click:
    Exit Sub

Err_cmdYesterday_Click:
    MsgBox Err.Description
    Resume Exit_cmdYesterday_Click
    
End Sub

Private Sub Command505_Click()
Me.ACation = True
Me.AWaitingDoc = True
Me.AWitingBill = True
Me.ATitleIssue = True
Me.AStop = True
End Sub

Private Sub Command510_Click()
Me.ACation = False
Me.AWaitingDoc = False
Me.AWitingBill = False
Me.ATitleIssue = False
Me.AStop = False
End Sub

Private Sub DateFrom_DblClick(Cancel As Integer)
DateFrom = Date

End Sub

Private Sub DateFrom1_DblClick(Cancel As Integer)

End Sub

Private Sub DateThru_DblClick(Cancel As Integer)
DateThru = Date
End Sub

Private Function CheckDatesOK() As Boolean
    Dim dt1 As Date, dt2 As Date
    Dim eMsg As String
    eMsg = ""
    On Error Resume Next
    dt1 = Forms!ReportsWorkflow!DateFrom
    dt2 = Forms!ReportsWorkflow!DateThru
    On Error GoTo 0

    If (1899 = Year(dt1)) And (1899 = Year(dt2)) Then
        eMsg = "Please fill-in dates or select date range."
    ElseIf (1899 = Year(dt1)) Then
        eMsg = "Please fill-in From Date, or select date range."
    ElseIf (1899 = Year(dt2)) Then
        eMsg = "Please fill-in Through Date, or select date range."
    ElseIf (dt1 > dt2) Then
        eMsg = "From Date must not be after Through Date."
    End If
    
    If "" <> eMsg _
    Then
        MsgBox eMsg, vbExclamation, "Valid date range must be supplied"
        CheckDatesOK = False
        Exit Function
    End If
    CheckDatesOK = True
End Function

