VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ForeclosureDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub AccelerationIssued_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    If BHproject Then
    AccelerationIssued = Now()
    AddStatus FileNumber, AccelerationIssued, "Demand Issued date"
    End If
    
End If
End Sub

Private Sub AddNewDate_Click()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    If StatusResults.Value = 1 Then
        StatusHearing.Enabled = True
        StatusHearing.Locked = False
        StatusHearingTime.Enabled = True
        StatusHearingTime.Locked = False
        StatusResults.Value = Null
    End If
            
    If IsNull(StatusResults) Then
        StatusHearing.Enabled = True
        StatusHearing.Locked = False
        StatusHearingTime.Enabled = True
        StatusHearingTime.Locked = False
    End If
End If

    
End Sub

Private Sub AddNewDateException_Click()
If cbxSustained.Value = 4 Then
ExceptionsHearing.Enabled = True
ExceptionsHearing.Locked = False
ExceptionsHearingTime.Enabled = True
ExceptionsHearingTime.Locked = False
cbxSustained.Value = Null
End If
End Sub

Private Sub AuditFile_Click()
If AuditFile.Locked = True Then
MsgBox ("You are not authorized to edit Audit Filed.")
End If

End Sub

Private Sub AuditRat_Click()
If AuditRat.Locked = True Then
MsgBox ("You are not authorized to edit Audit Ratified")
End If

End Sub

Private Sub BidAmount_AfterUpdate()
If Not BHproject Then

If IsNull(Me.DispositionDesc) And Not IsNull(Sale) And (Date <= Sale Or Format(Date, "mm/dd/yyyy") = Format(Sale, "mm/dd/yyyy")) And Not IsNull(Me.BidReceived) And Not IsNull(BidAmount) Then
    Me.SalePrice.Locked = False
    Me.Purchaser.Locked = False
    Me.PurchaserAddress.Locked = False
    Me.cmdPurchaserInvestor.Enabled = True
End If

End If

End Sub

Private Sub BondAmount_AfterUpdate()
If Not BHproject Then

Dim cost As Currency
cost = DLookup("ivalue", "db", "ID=" & 33) / 100
AddInvoiceItem FileNumber, "FC-BND", "Bond Issued", cost, 191, False, True, False, False

End If

End Sub

Private Sub BondReturned_AfterUpdate()
If Not BHproject Then
Dim cost As Currency
cost = -DLookup("ivalue", "db", "ID=" & 33) / 100

AddInvoiceItem FileNumber, "FC-BND", "Bond Returned", cost, 191, False, True, False, False
AddInvoiceItem FileNumber, "FC-BND", "Bond Returned Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, True, False, False
End If

End Sub



Private Sub btnNewDSurplus_Click()
DoCmd.OpenForm "sfrmdisbursingSurplusUpdate", , , , acFormAdd
'Me.sfrmDisbursingSurplusTable.Requery
End Sub

Private Sub chMannerofService_AfterUpdate()
If chMannerofService = True Then
AddInvoiceItem FileNumber, "FC-SVC", "Service Mailed Postage", 2 * Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, True, False, False
End If
End Sub

Private Sub chMannerofService_BeforeUpdate(Cancel As Integer)
If WizardSource <> "borrowerserved" And chMannerofService.OldValue = False Then
MsgBox "Manner of Service can only be changed through the Borrower Served Wizard", vbCritical
Cancel = 1
End If
If Not PrivSetSale Then
MsgBox "You must have Sale Setting privileges to change Manner of Service", vbCritical
Cancel = 1
End If
End Sub

Private Sub cmdAddMonitor_Click()
Dim rs As Recordset
                                                
If Not IsNull(Me.Sale) And Not IsNull(Me.SaleTime) And IsNull(Me.SaleCalendarEntryID) Then
                                                
    Set rs = CurrentDb.OpenRecordset("Select * FROM FCDetails where filenumber=" & Me.FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
        rs.Edit
        rs!Deposit = Null
        rs!SaleSet = Null
        rs!Disposition = Null
        rs!DispositionDate = Null
        rs!Sale = Null
        rs!SaleTime = Null
        'Forms!Foreclosuredetails!DispositionInitials = Null
        rs.Update
        rs.Close
        Set rs = Nothing
        'Forms!Foreclosuredetails.Requery
        Me.Requery
Else
                                                                                                                                                                             
        If (IsNull(Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced) Or (Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced = "")) And (Forms![Case List]!CaseTypeID = 8) Or (IsNull(Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced) Or (Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced = "")) And (Forms![Case List]!CaseTypeID = 1) Then
            Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced = Date
        End If
                                                                                                         
                Me!Deposit = Null
                Me!SaleSet = Null
                Me!Disposition = Null
                Me!DispositionDate = Null
                                                                                                                    
            If Not IsNull(Me!Sale) Then
                Me!Sale = Null
                Me!SaleTime = Null
                'Forms!Foreclosuredetails!.Requery
            End If
                                                                                                                      
            Me!Sale.Locked = False
            Me!SaleTime.Locked = False
End If
            'Forms!Foreclosuredetails!.Requery
                                                            
            If Not IsNull(Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced) Then
                AddStatus FileNumber, Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced, "Monitor Referral Recieved"
                                                            
                Dim FeeAmt As Currency
                    FeeAmt = Nz(DLookup("MonitorFee", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
                    AddInvoiceItem FileNumber, "FC-MON", "Monitor sale fee", Format$(FeeAmt, "Currency"), 0, True, True, False, False
            Else
                AddStatus FileNumber, Now(), "Moitor Referral Removed"
            End If
'End If

End Sub

Private Sub cmdAdjuEdit_Click()
'frmJudgmentsInfo

DoCmd.OpenForm ("frmJudgmentsInfo")
Forms!frmJudgmentsInfo.txtFilenum = Me.FileNumber
Forms!frmJudgmentsInfo.txtJudgments = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewJudgments

End Sub

Private Sub cmdchangedate_Click()
If MsgBox("Are you sure to change Monitor Referral date?", vbYesNo) = vbYes Then
    Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced.Locked = False
Else
    Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced.Locked = True
End If
    
End Sub

Private Sub cmdcloserestart_Click()
DoCmd.Save acForm, "foreclosuredetails"
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "Case List"
If IsLoadedF("EnterTitleOutReason") = True Then
DoCmd.Close acForm, "EnterTitleOutReason"
End If

If IsLoadedF("Limbo_Prosecc") = True Then
DoCmd.Close acForm, "Limbo_Prosecc"
End If
DoCmd.Close
End Sub

Private Sub cmdEdit_Click()

DoCmd.OpenForm ("frmDeedInfo")
Forms!frmDeedInfo.txtFilenum = Me.FileNumber
Forms!frmDeedInfo.txtnameof = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewNameOf

End Sub

Private Sub cmdLiensEdit_Click()
DoCmd.OpenForm ("frmLiensInfo")
Forms!frmLiensInfo.txtFilenum = Me.FileNumber

Forms!frmLiensInfo.txtSenior = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewLiens
Forms!frmLiensInfo.txtJunior = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewJunior
Forms!frmLiensInfo.txtLiens3 = Forms!foreclosuredetails!sfrmFCtitle!TitleReview3
Forms!frmLiensInfo.txtblank = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewBlank

End Sub

Private Sub cmdStatusEdit_Click()
DoCmd.OpenForm ("frmDeedStatus")
Forms!frmDeedStatus.txtFilenum = Me.FileNumber
Forms!frmDeedStatus.txtStatus = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewStatus
End Sub

Private Sub cmdTaxEdit_Click()

DoCmd.OpenForm ("frmDeedTaxes")
Forms!frmDeedTaxes.txtFilenum = Me.FileNumber
Forms!frmDeedTaxes.txtTaxes = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewTaxes

End Sub

Private Sub cmdWizComplete_Click()

Dim rstCase As Recordset, rstFC As Recordset, rstNames As Recordset
Dim rstsale As Recordset

Dim FileNum As Long, MissingTxt As String

On Error GoTo Err_cmdOK_Click

Select Case WizardSource
Case "Limbo_MDWhite", "Limbo_MDYellow", "Limbo_MDRed", "Limbo_VAWhite", "Limbo_VAYellow", "Limbo_VARed", "Limbo_DCWhite", "Limbo_DCYellow", "Limbo_DCRed"
Forms![Case List].SetFocus
Forms![Case List]!Page97.SetFocus
Call Limbo_Prosecc(WizardSource)



Case "HUDOCC"

If IsNull(Deposit) Then
MsgBox "The Wizard cannot be completed until a Deposit is entered.", vbCritical
Exit Sub
End If
If IsNull(Sale) Then
MsgBox "The Wizard cannot be completed until a Sale date is entered.", vbCritical
Exit Sub
End If
If State = "VA" Then
If IsNull(SaleTime) Then
MsgBox "The Wizard cannot be completed until a Sale time is entered.", vbCritical
Exit Sub
End If
End If
If IsNull(HUDOccLetter) Then
MsgBox "The Wizard cannot be completed until the HUD Occ Letter is printed."
Exit Sub
End If
Call HUDOccCompletionUpdate(FileNumber)

Case "VAappraisal"
If Forms![Case List]!SCRAID Is Null Or (Sale - VAAppraisal) < 180 Then Call VAappraisalCompletionUpdate(FileNumber)


Case "Restart"
'Removed per Diane
'If IsNull(TitleOrder) Then
'MsgBox "The Wizard cannot be completed until the title is ordered.", vbCritical
'Exit Sub
'End If
If IsNull(OccupancyStatusID) Then
MsgBox "The Wizard cannot be completed until Occupancy Status is entered.", vbCritical
Exit Sub
End If
If IsNull(LPIDate) Then
MsgBox "The Wizard cannot be completed until the LPI date is entered.", vbCritical
Exit Sub
End If
'If IsNull(DocstoClient) Then
'MsgBox "The Wizard cannot be completed until Docs to Client are sent.", vbCritical
'Exit Sub
'End If

Call RestartCompletionUpdate(FileNumber)

Forms!foreclosuredetails!DocstoClient = Now()
'        Dim rstqueue As Recordset
'        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM WizardSupportTwo where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
'
'        With rstqueue
'        .Edit
'
'       rstqueue!AttyMilestoneRestart = Null
'        rstqueue!AttyRestartRemark = ""
'        .Update
'        End With
'        Set rstqueue = Nothing
        


    Dim rstvalumeintake As Recordset
    Set rstvalumeintake = CurrentDb.OpenRecordset("Select * from ValumeRestart", dbOpenDynaset, dbSeeChanges)
    With rstvalumeintake
    .AddNew
    !CaseFile = FileNumber
    !Client = DLookup("ShortClientName", "ClientList", "ClientID = " & Forms![Case List]!ClientID)
    !RestartComplete = Now
    !RestartCompleteC = 1
 
    !Name = GetFullName()
    .Update
    End With
    Set rstvalumeintake = Nothing

    Dim lrs As Recordset
    Set lrs = CurrentDb.OpenRecordset("select * from journal where FileNumber=" & FileNumber & " AND warning = 100", dbOpenDynaset, dbSeeChanges)
    With lrs
    '.Edit
    Do Until .EOF
    .Edit
    ![Warning] = 0
    .Update
    .MoveNext
    
    Loop
    '.Update
    End With
    lrs.Close
    
DoCmd.SetWarnings False
strinfo = "Restart wizard complete."
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True


Call ReleaseFile(FileNumber)
MsgBox "Restart Wizard complete", vbInformation

DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

Case "Intake"
Forms!foreclosuredetails!DocstoClient = Now()
'Sarab stopped as there is no need for at after Atty approved 6/21/2015
'Dim rstIntakeDocs As Recordset
'Set rstIntakeDocs = CurrentDb.OpenRecordset("select * from wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
'
'If Not rstIntakeDocs!IntakeDocsRecdFlag And Not IsNull(rstIntakeDocs!IntakeWaiting) Then
'MsgBox "The Wizard cannot be completed because there are documents outstanding", vbCritical
'Exit Sub
'End If

DocstoClient = Date
Call DocstoClient_AfterUpdate
Call IntakeCompletionUpdate(FileNumber)
Call ReleaseFile(FileNumber)
MsgBox "Intake Wizard complete", vbInformation

DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

Case "Docketing"

If MsgBox("Update Sent to Docket date to " & Date & "?", vbYesNo) = vbYes Then
SentToDocket = Date
DoCmd.OpenForm "EnterLMAOption", , , , , acDialog
If MsgBox("Are you recording the SOT with this package?", vbYesNo) = vbYes Then
DeedAppSentToRecord = Date
End If
Call SentToDocket_AfterUpdate
End If
Call DocketingCompletionUpdate(FileNumber)



Call ReleaseFile(FileNumber)
MsgBox "Docketing Wizard complete", vbInformation
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

Case "FLMA"
FLMASenttoCourt = Date
Call FLMACompletionUpdate(FileNumber)

DoCmd.SetWarnings False
Set rstsale = CurrentDb.OpenRecordset("select * from FCdetails where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
        With rstsale
                strSQLJournal = "Insert into MoniterFLMAComplete (FileNumber,CompleteDate,UserName,FLMA) Values(!Filenumber,Now,GetFullName(),!FLMASenttoCourt )"
                DoCmd.RunSQL strSQLJournal
        End With
    Set rstsale = Nothing
DoCmd.SetWarnings True
'end Ticket 1253

Call ReleaseFile(FileNumber)
MsgBox "FLMA Wizard complete", vbInformation
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

Case "Service"

If IsNull(Me.CourtCaseNumber) Or IsNull(Me.Docket) Then

MsgBox "The Wizard cannot be completed until Case number and Docket date are entered.", vbCritical
Exit Sub
End If


ServiceSent = Date
Call ServiceSent_AfterUpdate
Call ServiceCompletionUpdate(FileNumber)

'Project Missing date ticket1253 10/23/14 Sarab

DoCmd.SetWarnings False
Set rstsale = CurrentDb.OpenRecordset("select * from FCdetails where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
        With rstsale
                strSQLJournal = "Insert into MoniterServiceServiceComplete (FileNumber,CompleteDate,UserName,ServiceSent) Values(!Filenumber,Now,GetFullName(),!ServiceSent )"
                DoCmd.RunSQL strSQLJournal
        End With
    Set rstsale = Nothing
DoCmd.SetWarnings True
'end Ticket 1253

Call ReleaseFile(FileNumber)
MsgBox "Service Wizard complete", vbInformation
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

Case "ServiceMailed"
ServiceMailed = Date
Call ServiceMailed_AfterUpdate
Call ServiceMailedCompletionUpdate(FileNumber)
'Project Missing date ticket1253 10/24/14 Sarab

DoCmd.SetWarnings False
Set rstsale = CurrentDb.OpenRecordset("select * from FCdetails where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
        With rstsale
                strSQLJournal = "Insert into MoniterNoticeToOccupantComplete (FileNumber,CompleteDate,UserName,ServiceMailed) Values(!Filenumber,Now,GetFullName(),!ServiceMailed )"
                DoCmd.RunSQL strSQLJournal
        End With
    Set rstsale = Nothing
DoCmd.SetWarnings True
'end Ticket 1253

Call ReleaseFile(FileNumber)
MsgBox "Service Mailed Wizard complete", vbInformation
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

Case "SaleSetting"
Call SaleSettingCompletionUpdate(FileNumber)
    DoCmd.SetWarnings False
    'Sarab10/16 Micheck remove the only line below in case problem in sale setting in md only
    'DoCmd.RunSQL ("update FCdetails set sale = " & Forms!ForeclosureDetails!Sale & " where [FileNumber] = " & Forms!ForeclosureDetails!FileNumber & " and current=true")

    strinfo = "MD Sale Setting wizard completed, Sale Scheduled for " & Forms!foreclosuredetails!Sale & IIf(IsNull(Forms!foreclosuredetails!SaleTime), "", " ,at time " & Forms!foreclosuredetails!SaleTime)
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    
    
    Set rstsale = CurrentDb.OpenRecordset("select * from FCdetails where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
        With rstsale
                strSQLJournal = "Insert into MoniterSaleComplete (FileNumber,CompleteDate,UserName,Sale, SaleTime) Values(!Filenumber,Now,GetFullName(),!sale,!Saletime )"
                DoCmd.RunSQL strSQLJournal
        End With
    Set rstsale = Nothing
    DoCmd.SetWarnings True
    
Call ReleaseFile(FileNumber)
MsgBox "Sale Setting Wizard complete", vbInformation
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

Case "VASaleSetting"
Call VAsalesettingCompletionUpdate(FileNumber)
    DoCmd.SetWarnings False
    strinfo = "VA Sale Setting wizard completed, Sale Scheduled for " & Forms!foreclosuredetails!Sale & IIf(IsNull(Forms!foreclosuredetails!SaleTime), "", " at time " & Forms!foreclosuredetails!SaleTime)
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    
    Set rstsale = CurrentDb.OpenRecordset("select * from FCdetails where Current=true and FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
        With rstsale
                strSQLJournal = "Insert into MoniterSaleComplete (FileNumber,CompleteDate,UserName,Sale, SaleTime) Values(!Filenumber,Now,GetFullName(),!sale,!Saletime )"
                DoCmd.RunSQL strSQLJournal
        End With
    Set rstsale = Nothing
    DoCmd.SetWarnings True

Call ReleaseFile(FileNumber)
MsgBox "VA Sale Setting Wizard complete", vbInformation
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

Case "VALNNSetting"
Call VALNNSetting(FileNumber)
Call ReleaseFile(FileNumber)
'MsgBox "VA Lost Note Notice complete", vbInformation
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

Case "BorrowerServed"
Call BorrowerServedCompletionUpdate(FileNumber)

'Project Missing date ticket1253 10/24/14 Sarab

DoCmd.SetWarnings False
Set rstsale = CurrentDb.OpenRecordset("select * from FCdetails where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
        With rstsale
                strSQLJournal = "Insert into MonterBorrowerServedComplete (FileNumber,CompleteDate,UserName,BorrowerServed) Values(!Filenumber,Now,GetFullName(),!BorrowerServed )"
                DoCmd.RunSQL strSQLJournal
        End With
    Set rstsale = Nothing
DoCmd.SetWarnings True
'end Ticket 1253


Call ReleaseFile(FileNumber)
MsgBox "Borrower Served Wizard complete", vbInformation
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

Case "Title"

Call TitleOrderCompletionUpdate(FileNumber)
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "Case List"
DoCmd.Close


Case "TitleOut"
TitleOutComplete (FileNumber)
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "EnterTitleoutReason"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "Case List"
If IsLoadedF("ForeclosureDetails") = True Then
DoCmd.Close acForm, "ForeclosureDetails"
End If


Case "TitleReview"
TitleReviewComplete (FileNumber)
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "EnterTitleoutReason"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "Case List"
If IsLoadedF("queTitleReview") = True Then

Dim cntr As Integer
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueTiteReview", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queTitleReview!QueueCount = cntr
Set rstqueue = Nothing

Forms!queTitleReview!lstFiles.Requery
End If

End Select

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub
Private Sub cmdWaiting1_Click()
Dim rstDocketingDocs As Recordset
Dim rstwiz As Recordset
Select Case WizardSource
Case "Docketing"
Set rstDocketingDocs = CurrentDb.OpenRecordset("select * from wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If Not rstDocketingDocs!DocketingDocsRecdFlag And Not IsNull(rstDocketingDocs!DocketingWaiting) Then
MsgBox "The Wizard cannot be completed because there are documents outstanding", vbCritical
Exit Sub
End If
Call DocketingAttyCompletionUpdate(FileNumber)
Call ReleaseFile(FileNumber)
MsgBox "Docketing Wizard complete", vbInformation

Case "vasalesetting"
Set rstDocketingDocs = CurrentDb.OpenRecordset("select * from wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If Not rstDocketingDocs!VASaleSettingDocsRecdFlag And Not IsNull(rstDocketingDocs!VASaleSettingWaiting) Then
MsgBox "There are documents outstanding", vbCritical

    DoCmd.SetWarnings False
    strinfo = "There are documents outstanding"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

End If
Call VAsalesettingAttyCompletionUpdate(FileNumber)
Call ReleaseFile(FileNumber)
MsgBox "VA Sale Setting Wizard complete", vbInformation

Case "Intake"

DoCmd.OpenForm "Atty_Intake", , , WhereCondition:="FileNumber= " & Forms!foreclosuredetails!FileNumber
GoTo A

Case "Restart"
DoCmd.OpenForm "Atty_Restart", , , WhereCondition:="FileNumber= " & Forms!foreclosuredetails!FileNumber
GoTo A

'Set rstwiz = CurrentDb.OpenRecordset("select * from wizardqueuestats where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
'With rstwiz
'.Edit
'!DateSentAttyIntake = Now()
'!AttyMilestoneMgr2 = Null
'!AttyMilestone2 = Null
'!AttyMilestone2Reject = False
'If IsNull(rstwiz!IntakeWaiting) Then rstwiz!IntakeWaiting = Now()
'
'.Update
'End With
'Set rstwiz = Nothing

End Select

DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"
A:
End Sub
Private Sub cmdWaiting_Click()
If WizardSource <> "Title" Then
DoCmd.Close acForm, "DocsWindow"
'DoCmd.Close acForm, "Journal"
'DoCmd.Close acForm, "Case List"
End If
Dim rstdocs As Recordset
'If Dirty Then DoCmd.RunCommand acCmdSaveRecord
Select Case WizardSource

Case "Title"
If IsLoadedF("Print Title Order") = True Then
DoCmd.Close acForm, "Print Title Order"
Else
DoCmd.OpenForm "EnterTitleReason"
Forms!EnterTitleReason!FileNumber = FileNumber
End If

Case "titleOut"
DoCmd.OpenForm "EnterTitleOutReason"
Forms!EnterTitleOutReason!FileNumber = FileNumber

Case "Intake"
DoCmd.OpenForm "EnterIntakeDocs"
Forms!enterintakedocs!FileNumber = FileNumber

Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded where docreceived is null AND filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstdocs
Do Until .EOF
If Not IsNull(!DocName) Then
Select Case !DocName
Case "SOD"
Forms!enterintakedocs!btn1 = True
Case "MA"
Forms!enterintakedocs!btn2 = True
Case "Va LNA"
Forms!enterintakedocs!btn3 = True
Case "SOT"
Forms!enterintakedocs!btn4 = True
Case "ANO"
Forms!enterintakedocs!btn5 = True
Case "NOI Aff"
Forms!enterintakedocs!btn6 = True
Case "PLMA"
Forms!enterintakedocs!btn7 = True
'Case "FLMA"
'Forms!enterintakedocs!btn8 = True
Case "Note"
Forms!enterintakedocs!btn9 = True
Case "DOT"
Forms!enterintakedocs!btn10 = True
Case "LoanMod"
Forms!enterintakedocs!btn11 = True
Case "SSN"
Forms!enterintakedocs!btn12 = True
Case "Jfigs"
Forms!enterintakedocs!btn13 = True
Case "Title"
Forms!enterintakedocs!btn14 = True
Case "NOI"
Forms!enterintakedocs!btn15 = True
Case "Other"
Forms!enterintakedocs!btn16 = True
End Select
End If
.MoveNext
Loop
End With

Dim rstJnl As Recordset, Comment As String, rstwiz As Recordset
Case "FLMA"
Comment = InputBox("Please enter comment", "Incompletion Comment Entry")

    '2/11/14
    DoCmd.SetWarnings False
    strinfo = Comment
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True



'Set rstJnl = CurrentDb.OpenRecordset("select * from journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = Comment
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing


Set rstwiz = CurrentDb.OpenRecordset("select flmacomment from wizardqueuestats where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstwiz
.Edit
!FLMAComment = Comment
.Update
End With
Set rstwiz = Nothing

Case "Service"
Comment = InputBox("Please enter comment", "Incompletion Comment Entry")
'2/11/14
    DoCmd.SetWarnings False
    strinfo = Comment
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True



'Set rstJnl = CurrentDb.OpenRecordset("select * from journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = Comment
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing

Set rstwiz = CurrentDb.OpenRecordset("select servicecomment from wizardqueuestats where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstwiz
.Edit
!ServiceComment = Comment
.Update
End With
Set rstwiz = Nothing

Case "SaleSetting"
Comment = InputBox("Please enter comment", "Incompletion Comment Entry")
'2/11/14
    DoCmd.SetWarnings False
    strinfo = Comment
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True


'Set rstJnl = CurrentDb.OpenRecordset("select * from journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Waiting for sale confirmation date.  " & Comment
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing


Set rstwiz = CurrentDb.OpenRecordset("select * from wizardqueuestats where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstwiz
.Edit
!SaleSettingWaiting = Now
!SaleSettingWaitingUser = GetStaffID
!SaleSettingComment = Comment
.Update
End With
Set rstwiz = Nothing

Case "BorrowerServed"
Comment = InputBox("Please enter comment", "Incompletion Comment Entry")

    DoCmd.SetWarnings False
    strinfo = Comment
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

'Set rstJnl = CurrentDb.OpenRecordset("select * from journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = Comment
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing

Set rstwiz = CurrentDb.OpenRecordset("select * from wizardqueuestats where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstwiz
.Edit
!BorrowerServedLastEdited = Now
!BorrowerServedUser = GetStaffID
!BorrowerServedComment = Comment
.Update
End With
Set rstwiz = Nothing


Case "Restart"
DoCmd.OpenForm "EnterRestartReason"
Forms!enterrestartreason!FileNumber = FileNumber

Case "Docketing"
DoCmd.OpenForm "EnterDocketingDocs"
Forms!enterdocketingdocs!FileNumber = FileNumber

Set rstdocs = CurrentDb.OpenRecordset("select * from docketingdocsneeded where docreceived is null AND filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstdocs
Do Until .EOF
If Not IsNull(!DocName) Then
Select Case !DocName
Case "SOD"
Forms!enterdocketingdocs!btn9 = True
Case "SOT"
Forms!enterdocketingdocs!btn10 = True
Case "Fair Debt"
Forms!enterdocketingdocs!btn11 = True
Case "Acceleration"
Forms!enterdocketingdocs!btn12 = True
Case "ANO"
Forms!enterdocketingdocs!btn13 = True
Case "NOI Aff"
Forms!enterdocketingdocs!btn14 = True
Case "NOI"
Forms!enterdocketingdocs!btn15 = True
'Case "LMA"
''Forms!enterdocketingdocs!btn17 = True As per Diane 10/28 Sarab
Case "Other"
Forms!enterdocketingdocs!btn16 = True
End Select
End If
.MoveNext
Loop
End With

Case "VAsalesetting"
DoCmd.OpenForm "EnterVASalesettingDocs"
Forms!enterVAsalesettingdocs!FileNumber = FileNumber

Set rstdocs = CurrentDb.OpenRecordset("select * from VAsalesettingdocsneeded where docreceived is null AND filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
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

Case "TitleReview"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"
Forms!queTitleReview!lstFiles.Requery


End Select

ExitProc:

End Sub


Private Sub Audit2File_BeforeUpdate(Cancel As Integer)
If Not BHproject Then

Cancel = CheckFutureDate(Audit2File)
End If

End Sub

Private Sub Audit2Rat_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(Audit2Rat)
End If

End Sub

Private Sub AuditFile_AfterUpdate()
AddStatus FileNumber, AuditFile, "Audit filed"
End Sub

Private Sub AuditFile_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(AuditFile)
End If

End Sub

Private Sub AuditFile_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    AuditFile = Now()
    AddStatus FileNumber, AuditFile, "Audit filed"
End If

End Sub

Private Sub AuditRat_AfterUpdate()
AddStatus FileNumber, AuditRat, "Audit ratified"
End Sub

Private Sub AuditRat_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(AuditRat)
End If

End Sub

Private Sub AuditRat_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    AuditRat = Now()
    AddStatus FileNumber, AuditRat, "Audit ratified"
End If

End Sub

Private Sub BidReceived_AfterUpdate()
If BHproject Then
AddStatus FileNumber, BidReceived, "Bid received"
Else

AddStatus FileNumber, BidReceived, "Bid received"
If IsNull(Me.DispositionDesc) And Not IsNull(Sale) And (Date <= Sale Or Format(Date, "mm/dd/yyyy") = Format(Sale, "mm/dd/yyyy")) And Not IsNull(Me.BidReceived) And Not IsNull(BidAmount) Then
    Me.SalePrice.Locked = False
    Me.Purchaser.Locked = False
    Me.PurchaserAddress.Locked = False
    Me.cmdPurchaserInvestor.Enabled = True
End If
End If

End Sub

Private Sub BidReceived_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
  Cancel = CheckFutureDate(BidReceived)
End If

End Sub

Private Sub BidReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    BidReceived = Date
    AddStatus FileNumber, BidReceived, "Bid received"
End If
End Sub

Private Sub BorrowerServed_AfterUpdate()
If BHproject Then
AddStatus FileNumber, BorrowerServed, "Borrower served by posting"
Else

Dim rstJnl As Recordset, Comment As String, FeeAmount As Currency


If Not IsNull(BorrowerServed) Then
If DLookup("milestonebilling", "clientlist", "clientid=" & Forms![Case List]!ClientID) Then
Forms![Case List]!BillCase = True
Forms![Case List]!BillCaseUpdateUser = GetStaffID()
Forms![Case List]!BillCaseUpdateDate = Date
Forms![Case List]![BillCaseUpdateReasonID] = 5
Forms![Case List]!lblBilling.Visible = True
Forms![Case List].SetFocus
DoCmd.RunCommand acCmdSaveRecord
Forms![foreclosuredetails].SetFocus

Dim rstBillReasons As Recordset
Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstBillReasons
.AddNew
!FileNumber = FileNumber
!billingreasonid = 5
!UserID = GetStaffID
!Date = Date
.Update
End With


'Milestone Billing for Referral Fee
Dim InvPct As Double
If State = "MD" Then
    Select Case LoanType
    Case 4
    FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=177"))
    Case 5
    FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=263"))
    Case Else
    FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
    End Select
       
        If FeeAmount > 0 Then
            InvPct = DLookup("MDservicecomppct", "clientlist", "clientid=" & Forms![Case List]!ClientID)
            If InvPct < 1 Then
            AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when borrower served of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
            Else
                AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
            End If
        End If
End If
End If

DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter costs for Process Server|FC-PS|Estimated Process Server Costs|Process Server"
FeeAmount = InputBox("Please re-enter costs for Process Server")

If MsgBox("Was the borrower PERSONALLY served?", vbYesNo) = vbYes Then
chMannerofService = True
AddStatus FileNumber, BorrowerServed, "Borrower served by personal service"
Comment = "Borrower served by personal service; costs were $"
Else
AddStatus FileNumber, BorrowerServed, "Borrower served by posting"
Comment = "Borrower served by posting; costs were $"
End If


Set rstJnl = CurrentDb.OpenRecordset("select * from journal", dbOpenDynaset, dbSeeChanges)
With rstJnl
.AddNew
!FileNumber = FileNumber
!JournalDate = Now
!Who = GetFullName
!Info = Comment & FeeAmount
!Color = 1
.Update
End With
Set rstJnl = Nothing

cmdWizComplete.Enabled = True
End If

End If


End Sub

Private Sub BorrowerServed_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
  Cancel = CheckFutureDate(BorrowerServed)
  If (Cancel = 1) Then Exit Sub
  
  If (DateDiff("d", Docket, BorrowerServed) < 0) Then
    Cancel = 1
    MsgBox "Date cannot be before Docket Date.", vbCritical
  End If
End If

End Sub

Private Sub BorrowerServed_DblClick(Cancel As Integer)
If FileReadOnly Or IsNull(ServiceSent) Then
   DoCmd.CancelEvent
ElseIf (Me.WizardSource = "None" And Me.State = "MD") Or (Me.State = "MD" And Len(Me.WizardSource & "") = 0) Then
    DoCmd.CancelEvent
    

Else

    BorrowerServed = Now()
    AddStatus FileNumber, BorrowerServed, "Borrower served"
    Call BorrowerServed_AfterUpdate
End If

If EditDispute And Not IsNull(Me.ServiceSent) Then
        Me.BorrowerServed.Locked = False
        BorrowerServed = Now()
    AddStatus FileNumber, BorrowerServed, "Borrower served"
    Call BorrowerServed_AfterUpdate
End If

End Sub

Private Sub cbxAddTrustee_AfterUpdate()
Dim t As Recordset

On Error GoTo Err_cbxAddTrustee

Set t = CurrentDb.OpenRecordset("Trustees", dbOpenDynaset, dbSeeChanges)
t.AddNew
t!FileNumber = Me!FileNumber
t!Trustee = Me!cbxAddTrustee
t!Name = Me!cbxAddTrustee.Column(1)
t!Assigned = Now()

t.Update
t.Close
Me!lstTrustees.Requery
TrusteeWordFile = 0         ' invalidate cache

Exit_cbxAddTrustee:
    Exit Sub

Err_cbxAddTrustee:
    MsgBox Err.Description
    Resume Exit_cbxAddTrustee

End Sub

Private Sub cmdEditPropertyDetails_Click()
On Error GoTo Err_EditPropertyDetails_Click


DoCmd.OpenForm "EditPropertyDetails", , , WhereCondition:="FileNumber= " & Forms!foreclosuredetails!FileNumber & " And Current = true"

'    Dim stDocName As String
'    Dim stLinkCriteria As String
'
'    stDocName = "EditPropertyDetails"
'    DoCmd.OpenForm stDocName, , , stLinkCriteria
'    With Forms!EditPropertyDetails
'        .txtFileNumber = FileNumber
'        .txtPrimaryFirstName = PrimaryFirstName
'        .txtPrimaryLastName = PrimaryLastName
'        .txtSecondaryFirstName = SecondaryFirstName
'        .txtSecondaryLastName = SecondaryLastName
'        .txtPropertyAddress = PropertyAddress
'        .txtCity = City
'        .txtState = State
'        .txtZipCode = ZipCode
'
'        .txtCourtCaseNumber = CourtCaseNumber
'        .txtTaxID = TaxID
'
'
'    End With
'
'If Dirty Then DoCmd.RunCommand acCmdSaveRecord

Exit_EditPropertyDetails_Click:
    Exit Sub

Err_EditPropertyDetails_Click:
    MsgBox Err.Description
    Resume Exit_EditPropertyDetails_Click
End Sub

Private Sub ComAddName_Click()
DoCmd.OpenForm "sfrmNamesUpdate", , , , acFormAdd

End Sub



Private Sub Command681_Click()
Call RemoveDates
End Sub







Private Sub ComEdit_Click()
checkCmdEdit = True

Me.SalePrice.Locked = False
Me.Purchaser.Locked = False
Me.PurchaserAddress.Locked = False
Me.cmdPurchaserInvestor.Enabled = True
End Sub



Private Sub CommdEdit_Click()

DoCmd.OpenForm "sfrmNamesUpdate", , , WhereCondition:="ID= " & Forms!foreclosuredetails!sfrmNames!ID

''If Not CheckNameEdit() Then
'Dim ctrl As Control
'For Each ctrl In Me.sfrmNames.Form.Controls
'
'If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then ' TypeOf ctrl Is CommandButton Then
''If Not ctrl.Locked) Then
'ctrl.Locked = False
'ctrl.Enabled = True
''Else
''ctrl.Locked = True
'End If
''End If
'Next
'With Me.sfrmNames.Form
'
'.AllowAdditions = True
'.AllowEdits = True
'.AllowDeletions = True
'.cmdCopyClient.Enabled = True
'.cmdCopy.Enabled = True
'.cmdTenant.Enabled = True
'.cmdMERS.Enabled = True
'.cmdEnterSSN.Enabled = True
'.cmdNoNotice.Enabled = True
'.cmdPrintNotice.Enabled = True
'.cmdPrintLabel.Enabled = True
'.cbxNotice.Enabled = True
'.cmdDelete.Enabled = True
'.cmdNoNotice.Enabled = True
'.cbxNotice.Enabled = True
'.cbxNotice.Locked = False
'End With
''Exit Sub
''Else
'
''End If
End Sub

Private Sub ComRemoveCase_Click()
If Me.CourtCaseNumber = "" Then
MsgBox ("There is no case court number")
Exit Sub
End If

If Me.State = "MD" Then
    If IsNull(Me.Disposition) Then
    MsgBox ("Must enter dispostion first")
    Exit Sub
    Else
    
   
    
            CaseNuUpdate = True
             
                If MsgBox("Do you want to remove the case number and all fields for this sale? ", vbYesNo) = vbYes Then
              
                Call RemoveCaseFiled
             
                DoCmd.SetWarnings False
            
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails.FileNumber,Now,GetFullName(),'" & "Removed Court Case Number: " & CourtCaseNumber.OldValue & ". " & "',1 )"
                DoCmd.RunSQL strSQLJournal
                
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails.FileNumber,Now,GetFullName(),'" & "Case:  " & CourtCaseNumber.OldValue & " Dismissed. Dates associated with this case have been removed from the file. " & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
            
            
            '    DoCmd.Close
         
                End If
    End If

End If
End Sub

Private Sub RemoveCaseFiled()

Dim rstDisbursingSurplus As Recordset
Set rstDisbursingSurplus = CurrentDb.OpenRecordset("Select * from DisbursingSurplus where Filenumber = " & FileNumber, dbOpenDynaset, dbSeeChanges)
    Do While Not rstDisbursingSurplus.EOF
        rstDisbursingSurplus.Delete
        rstDisbursingSurplus.MoveLast
    Loop
rstDisbursingSurplus.Close


Dim rstFCDIL As Recordset
Set rstFCDIL = CurrentDb.OpenRecordset("Select * from FCDIL where Filenumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)
If Not rstFCDIL.EOF Then
With rstFCDIL
.Edit
!CertOfPubField = Null
!SoftDate = Null
!SoftStaffInitial = Null
!AuditorLetterReceived = Null
!AuditorRespondDeadline = Null
!AuditorResponseSent = Null

.Update
End With

End If


Set rstFCDIL = Nothing

With Forms!foreclosuredetails
'mdeatinn
!CourtCaseNumber = ""
!MedCaseNumber = Null
!MedRequestDate = Null
!MedRecDocDate = Null
!MedReqDocDate = Null
!MedDocSentDate = Null
!MedHearingLocation = Null
!MedHearingResults = Null
!MedHearingClientContactID = Null



If Not IsNull(Forms!foreclosuredetails.Form!sfrmLMHearing!txtHearing) And Not IsNull(Forms!foreclosuredetails!LMDispositionDesc) Then
    Dim t As Recordset
    
    Set t = CurrentDb.OpenRecordset("SELECT * FROM LMHearings WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
       
  '  If Not t.EOF Then t.Delete

    Do While Not t.EOF
       t.Delete
       t.MoveLast
    Loop
    t.Close
End If

!LMDispositionStaffID = Null
!LMDisposition = Null
!LMDispositionDate = Null

'LMHearings

'Dim ss As Boolean
'ss = True
'End If
'If ss Then
'Forms!foreclosuredetails.Form!sfrmLMHearing!txtHearing = Null
'Forms!forclosuredetails!LMDispositionDesc = Null
'ss = False
'End If


'DocBackNoteOwnership
' presale
!HUDOccLetter = Null
!DocstoClient = Null
!DocsBack = Null
!DocBackMilAff = False
!DocBackSOD = False
!DocBackNoteOwnership = False
!DocBackLossMitPrelim = False
!DocBackLossMitFinal = False
!DocBackAff7105 = False
!StatementOfDebtDate = Null
!StatementOfDebtAmount = Null
!StatementOfDebtPerDiem = Null
!SentToDocket = Null
!Docket = Null
!LienCert = Null
!FLMASenttoCourt = Null
!LossMitFinalDate = Null
!ServiceSent = Null
!BorrowerServed = Null
!ServiceMailed = Null
!FirstPub = Null
!IRSNotice = Null
!Notices = Null
!UpdatedNotices = Null
!BidReceived = Null
!BidAmount = Null

!Sale = Null
!SaleTime = Null
!Deposit = Null
!SaleSet = Null
!BondNumber = Null
!BondAmount = Null
!BondReturned = Null
!chMannerofService = False
!ReviewAdProof = Null
!SaleCert = Null
!PayoffAmount = Null

'Post sale tab
!Report = Null
!StatePropReg = Null
!NiSiEnd = Null
!SaleRat = Null
!PropReg = Null
!DeedtoRec = Null
!DeedtoTitleCo = 0
!DeedtoRec = Null
!RecordDeed = Null
!RecordDeedLiber = Null
!RecordDeedFolio = Null
!FinalPkg = Null
!AuditFile = Null
!AuditRat = Null
!Audit2File = Null
!Audit2Rat = Null
!SalePrice = Null
!Purchaser = ""
!PurchaserAddress = ""
!SubstitutePurchaser = Null
!OrderSubsPurch = Null
If Not IsNull(Forms!foreclosuredetails!ExceptionsHearing) And Not IsNull(Forms!foreclosuredetails!cbxSustained) Then
Forms!foreclosuredetails!ExceptionsHearing = Null
Forms!foreclosuredetails!ExceptionsHearingTime = Null
Forms!foreclosuredetails!ExceptionsFiled = False
Forms!foreclosuredetails!cbxSustained = Null
End If
'
If Not IsNull(Forms!foreclosuredetails!StatusHearing) And Not IsNull(Forms!foreclosuredetails!StatusResults) Then
Forms!foreclosuredetails!StatusHearing = Null
Forms!foreclosuredetails!StatusHearingTime = Null
Forms!foreclosuredetails!StatusResults = Null
End If

!ResellMotion = Null
!ResellServed = Null
!ResellShowCauseExpires = Null
!ResellAnswered = Null
!ResellGranted = Null
!SaleConductedTrusteeID = Null
'!SaleCompleted = Null
!Disposition = Null
!DispositionInitials = Null
!DispositionDate = Null
!DismissalSent = Null
!DismissalDate = Null 'here the action
!CorrectiveDeedSent = False
!CorrectiveDeedRecorded = False
!RescindClientReq = False
!Settled = Null
!ClientPaid = Null
!chEviction = False
!REO = False


!AmmDocBackSOD = False
!AmmStatementOfDebtDate = Null
!AmmStatementOfDebtAmount = Null
!AmmStatementOfDebtPerDiem = Null
!MonitorMotionSurplusFiled = Null
!MonitorOrderSurplus = Null
!MonitorClientPaid = Null

'Mediation tab


End With
'Forms!foreclosuredetails.Requery

End Sub




Private Sub CourtCaseNumber_AfterUpdate()
'If IsNull(CourtCaseNumber) Then
'If MsgBox("Is the case no longer active?", vbYesNo) = vbYes Then
' LienCert = Null
'  CourtCaseNumber = Null
'  Docket = Null
'  FLMASenttoCourt = Null
'  LossMitFinalDate = Null
'  ServiceSent = Null
'  BorrowerServed = Null
'  ServiceMailed = Null
'  SentToDocket = Null
'  CourtCaseNumber = Null
'End If
'End If
End Sub

Private Sub DeedAppReceived_AfterUpdate()
    If IsNull(DeedAppReceived) = False Then
        Me.DocBackDOA = True
        AddStatus FileNumber, Date, "Received Substitution of Trustee"
        
        DoCmd.SetWarnings False
        strinfo = "Received Substition of Trustee by " & GetFullName
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
    
    Else
        Me.DocBackDOA = False
        AddStatus FileNumber, Date, "Removed Substitution of Trustee"
        
        DoCmd.SetWarnings False
        strinfo = "Removed Substition of Trustee by " & GetFullName
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
    End If
End Sub



Private Sub DispositionDate_AfterUpdate()

Dim Reason As Integer
'Fix the dispostion error and missing after 1/5 SA 1/13/2015
Select Case SelectedDispositionID
'Select Case Disposition

Case 1
Reason = 10
Case 2
Reason = 11
Case 3
Reason = 12
Case 4
Reason = 13
Case 5
Reason = 14
Case 6
Reason = 15
Case 7
Reason = 9
Case 8
Reason = 17
Case 9
Reason = 18
Case 10
Reason = 19
Case 11
Reason = 20
Case 12
Reason = 21
Case 13
Reason = 22
Case 16
Reason = 23
Case 26
Reason = 24
Case 27
Reason = 25
Case 28
Reason = 26
Case 29
Reason = 27
Case 30
Reason = 28
Case 31
Reason = 29
End Select

DoCmd.SetWarnings False

DoCmd.RunSQL ("UPDATE CaseList set BillCase = True ,BillCaseUpdateUser = " & GetStaffID() & " ,BillCaseUpdateDate = Date() ,BillCaseUpdateReasonID ='" & Reason & "' WHERE [FileNumber] = " & Me.FileNumber)
DoCmd.SetWarnings True


'Forms![Case List]!BillCase = True
'Forms![Case List]!BillCaseUpdateUser = GetStaffID()
'Forms![Case List]!BillCaseUpdateDate = Date
'Forms![Case List]![BillCaseUpdateReasonID] = Reason
Forms![Case List]!lblBilling.Visible = True
Forms![Case List].SetFocus
'
DoCmd.RunCommand acCmdSaveRecord
Forms![foreclosuredetails].SetFocus

Dim rstBillReasons As Recordset
Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstBillReasons
.AddNew
!FileNumber = FileNumber
!billingreasonid = Reason
!UserID = GetStaffID
!Date = Date
.Update
End With

End Sub

Private Sub DocBackLossMitFinal_AfterUpdate()

If DocBackLossMitFinal Then
    AddStatus FileNumber, Date, "Received FLMA"
Else
    AddStatus FileNumber, Date, "Removed FLMA"
End If

End Sub

Private Sub DocBackLossMitPrelim_AfterUpdate()

If DocBackLossMitPrelim Then
    AddStatus FileNumber, Date, "Received PLMA"
Else
    AddStatus FileNumber, Date, "Removed PLMA"
End If

End Sub

Private Sub DocBackNoteOwnership_AfterUpdate()

If DocBackNoteOwnership Then
    AddStatus FileNumber, Date, "Received Affidavit of Note Ownership"
Else
    AddStatus FileNumber, Date, "Removed Affidavit of Note Ownership"
End If

End Sub

Private Sub DOTrecorded_AfterUpdate()

If Not IsNull(DOTrecorded) Then
 AddStatus FileNumber, DOTrecorded, "Recorded Deed of Trust on " & Format(DOTrecorded, "mm/dd/yyyy")
Else
 AddStatus FileNumber, Now(), "Removed Recorded Deed of Trust Date"
End If

End Sub

Private Sub DOTrecorded_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DOTrecorded = Now()
    Call DOTrecorded_AfterUpdate
End If
End Sub

Private Sub ExceptionsHearing_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
    If Not ExceptionsFiled Then
    MsgBox " Check the Exceptions Filed box"
    Me.Undo
    
    Exit Sub
    Else
    If HearingCheking(ExceptionsHearing, 1) = 1 Then
        Cancel = 1
    End If
    
    If Weekday(ExceptionsHearing) = vbSunday Or Weekday(ExceptionsHearing) = vbSaturday Then
        
        MsgBox "Exceptions Hearing date cannot be Saturday or Sunday", vbCritical
       Cancel = 1
    End If
    
    End If

End If

End Sub

Private Sub ExceptionsHearingTime_AfterUpdate()
If Not BHproject Then
    If IsNull(ExceptionsHearing) Or IsNull(ExceptionsHearingTime) Then Exit Sub
    
'    If Hour(ExceptionsHearingTime) < 8 Or Hour(ExceptionsHearingTime) > 18 Then
'        ExceptionsHearingTime = DateAdd("h", 12, ExceptionsHearingTime)
'        If Hour(ExceptionsHearingTime) < 8 Or Hour(ExceptionsHearingTime) > 18 Then
'            MsgBox "Invalid Exceptions Hearing time: " & Format$(ExceptionsHearingTime, "h:nn am/pm")
'            ExceptionsHearingTime = Null
'            Exit Sub
'        End If
'    End If

'edited on 9/8/15
    If Hour(ExceptionsHearingTime) >= 8 And Hour(ExceptionsHearingTime) < 13 Then
            ExceptionsHearingTime = Format$(ExceptionsHearingTime, "h:nn am/pm")
    ElseIf Hour(ExceptionsHearingTime) >= 1 And Hour(ExceptionsHearingTime) <= 6 Then
            ExceptionsHearingTime = DateAdd("h", 12, ExceptionsHearingTime)
    Else
            MsgBox "Hearing time must be between 8:00 AM and 6:00 PM" ': " & Format$(DCHearingTime, "h:nn am/pm")
            ExceptionsHearingTime = Null
            Exit Sub
    End If

    
    
    AddStatus FileNumber, Now(), "Exceptions Hearing scheduled time for " & Format$(ExceptionsHearing, "m/d/yyyy") & " at " & Format$(ExceptionsHearingTime, "h:nn am/pm")
    If Not IsNull(ExceptionsHearing) Then Call UpdateCalendarExceptionHearing
    Call Visuals
    ExceptionsHearingTime.Locked = True
End If

End Sub

Private Sub FHLMCLoanNumber_BeforeUpdate(Cancel As Integer)
If Len([FHLMCLoanNumber]) <> 9 Then
MsgBox ("The FHLMCL loan Number is not 9 digits")
Cancel = 1
Else
Cancel = 0
End If



End Sub

Private Sub FirstLegal_AfterUpdate()
AddStatus FileNumber, FirstLegal, "First Legal Due Date"
End Sub

Private Sub FirstLegal_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    FirstLegal = Date
    Call FirstLegal_AfterUpdate
End If

End Sub



Private Sub FNMALoanNumber_BeforeUpdate(Cancel As Integer)
If Len([FNMALoanNumber]) <> 10 Then
    MsgBox ("Loan Number should be only 10 digits ")
    Me.Undo
End If

End Sub

Private Sub ForebearanceAgreementSend_AfterUpdate()
AddInvoiceItem FileNumber, "FC-FOR", "Forbearance Fee", DLookup("ForbearanceFee", "clientlist", "clientid=" & Forms![Case List]!ClientID), 0, True, True, False, False

End Sub

Private Sub Form_Deactivate()


'DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub Form_Load()
If Len(Me.OpenArgs) > 0 Then
If Me.OpenArgs = "Title" Then
 Me.TitleBack.BackColor = vbYellow
 Me.TitleOrder.BackColor = vbYellow
 Me.TitleThru.BackColor = vbYellow
 cmdWaiting.Caption = "TITLE"
 Forms!foreclosuredetails!cmdWaiting.FontBold = True
Forms!foreclosuredetails!cmdWaiting.ForeColor = vbBlack

 Else
    If Me.OpenArgs = "TitleOut" Then
    Me.TitleBack.BackColor = vbGreen
    Me.TitleOrder.BackColor = vbGreen
    Me.TitleThru.BackColor = vbGreen
    cmdWaiting.Caption = "TITLE Outstanding"
    Else
    
 Me.TitleBack.BackColor = vbWhite
 Me.TitleOrder.BackColor = vbWhite
 Me.TitleThru.BackColor = vbWhite
 cmdWaiting.Caption = "Waiting for Restart"
 'Waiting for Restart
 
 End If
 End If
 End If

End Sub

Private Sub Form_Open(Cancel As Integer)
CaseNuUpdate = False

Dim C_Name As Recordset
checkCmdEdit = False
Set C_Name = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE Name =""" & GetLoginName() & """", dbOpenSnapshot)

  If C_Name!AuditFile = 0 Then
  AuditFile.Locked = True
  End If
  
  If C_Name!AuditRatified = 0 Then
      AuditRat.Locked = True
  End If
  
'8/25/2014
If PrivitLimitedView = True Then

Me.Trustees.Visible = False
Me.pgNOI.Visible = False
Me.Page256.Visible = False
Me.pgRealPropTaxes.Visible = False

End If

  
  If Disposition > 2 And Disposition <> 7 And Disposition <> 11 And (Date - DispositionDate) > 14 Then
'Property.Enabled = False
'Page96.Enabled = False
'Trustees.Enabled = False
'Page195.Enabled = False
'Page412.Enabled = False 'Removed 7/12 in support of DIL Printable documents (MC)
pgNOI.Enabled = False
'Page256.Enabled = False
If C_Name!SaleSetter = False Then
pgMediation.Enabled = False
Else
pgMediation.Enabled = True

End If

'[pre-sale].Enabled = False


End If

'Sarab10/16 Michael this is the loop that might messup with the staff.
  
If FileReadOnly Or EditDispute Then

    Dim ctl As Control
    Dim lngI As Long
    Dim bSkip As Boolean

    For Each ctl In Form.Controls
    Select Case ctl.ControlType
    Case acTextBox, acComboBox, acListBox, acOptionGroup, acCheckBox, acSubform, acOptionButton
        
            If Not (ctl.Locked) Then ctl.Locked = True
            
    Case acCommandButton
        bSkip = False
            If ctl.Name = "cmdClose" Then bSkip = True
            If ctl.Name = "cmdSelectFile" Then bSkip = True
            If Not bSkip Then ctl.Enabled = False
         
       
    End Select
    Next
End If

If Forms![Case List]!CaseTypeID = 8 Then


                   ' Forms!foreclosureDetails!Page96.Visible = False
                   ' Forms!foreclosureDetails!Trustees.Visible = False
                    Forms!foreclosuredetails!Page412.Visible = False
                    Forms!foreclosuredetails!pgNOI.Visible = False
                '    Forms!foreclosureDetails!Page256.Visible = False
                    Forms!foreclosuredetails!pgMediation.Visible = False
                 '   Forms!foreclosureDetails![Pre-Sale].Visible = False
                 '   Forms!foreclosureDetails![Post-Sale].Visible = False
                    Forms!foreclosuredetails!pgRealPropTaxes.Visible = False
                   ' Forms!foreclosureDetails!pageStatus.Visible = False
                   
End If

 'added on 7/9/15
 
If DCTabView = False Then
'dc tab locked
'    Me.txtFistPub.Enabled = False
'    Me.txtSale.Enabled = False
'    Me.txtSaleTime.Enabled = False
'    Me.txtDeposit.Enabled = False
'    Me.txtSaleSet.Enabled = False
'    Me.txtreviewadproof.Enabled = False
'    Me.txtNewAdVendor.Enabled = False

End If
  
End Sub

Private Sub LienPosition_AfterUpdate()

If Me.LienPosition = 1 Then
    Forms!foreclosuredetails!sfrmFCtitle!ckSenior = True
Else
    Forms!foreclosuredetails!sfrmFCtitle!ckSenior = False
End If
 
If Me.LienPosition = 2 Then
    Forms!foreclosuredetails!sfrmFCtitle!ckJunior = True
Else
    Forms!foreclosuredetails!sfrmFCtitle!ckJunior = False
End If

If Me.LienPosition = 3 Then
    Forms!foreclosuredetails!sfrmFCtitle!ck3 = True
Else
    Forms!foreclosuredetails!sfrmFCtitle!ck3 = False

End If

If IsNull(Me.LienPosition) Or Me.LienPosition = "" Then
    Forms!foreclosuredetails!sfrmFCtitle!ckOther = True
'Else
    'Forms!ForeclosureDetails!sfrmFCtitle!ckOther = True
End If

Forms!foreclosuredetails!sfrmFCtitle.Requery
End Sub

Private Sub LossMitFinalDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    LossMitFinalDate = Date
    Call LossMitFinalDate_AfterUpdate
End If

End Sub

Private Sub LostNoteAffSent_AfterUpdate()
If Not IsNull([LostNoteAffSent]) Then
AddStatus FileNumber, LostNoteAffSent, "Lost Note Affidavit Sent"
Else
AddStatus FileNumber, Now(), "Lost Note Affidavit Removed"
End If
End Sub

Private Sub LostNoteNotice_AfterUpdate()
If Not IsNull([LostNoteNotice]) Then
AddStatus FileNumber, LostNoteNotice, "Lost Note Notice Sent"
Else
AddStatus FileNumber, Now(), "Lost Note Notice Removed"
End If
End Sub

Private Sub New45Notice_Click()

If (IsNull(Disposition) Or (Disposition = 1 Or Disposition = 2)) Then

    If Not IsNull([NOI]) Or ([ClientSentNOI] = "C") Then
    
    If MsgBox(" You are about to remove dates ? ", vbOKCancel) = vbOK Then
    
        '45 Day Notice sent
        AddStatus FileNumber, Now(), "Removed 45 Days Notice (" & [NOI] & ") by " & GetFullName
    
        DoCmd.SetWarnings False
        strinfo = "Removed 45 Days Notice (" & [NOI] & ") by " & GetFullName
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
    
        'Put file in FairDebt queue
        Dim rstNOI As Recordset
        Set rstNOI = CurrentDb.OpenRecordset("Select * From WizardQueueStats Where FileNumber = " & FileNumber & "And Current = True", dbOpenDynaset, dbSeeChanges)
            If Not rstNOI.EOF Then
                With rstNOI
                .Edit
               ' If IsNull(!FairDebtComplete) Then !FairDebtComplete = #1/2/2012# NOICompleteDocsMsng
                !NOIcomplete = Null
                !DateInWaiitingQueueNOI = Null
                !DateInQueueNOI = Null
               ' !Add45 = "45day"
                .Update
                End With
                Else
                MsgBox ("There is no Currrent Wizard Record for this File, Please Contact the IT")
            End If
        Set rstNOI = Nothing
        Forms!foreclosuredetails![NOI] = Null
        
                  
            If Forms!foreclosuredetails!txtClientSentNOI = "C" Then
                AddStatus FileNumber, Now(), "Removed C Of NOI by " & GetFullName
                
                DoCmd.SetWarnings False
                strinfo = "Removed C Of NOI by " & GetFullName
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery
            
                Forms!foreclosuredetails!txtClientSentNOI = ""
            End If
    
    End If
    Else
    Exit Sub
    End If
Else
MsgBox ("The File has dispsotion not buy in or 3rd party, proceduer canceld")
Exit Sub
End If


End Sub

Private Sub NewDemand_Click()

If (IsNull(Disposition) Or (Disposition = 1 Or Disposition = 2)) Then

    If Not IsNull([AccelerationIssued]) Or Not IsNull([AccelerationLetter]) Or ([ClientSentAcceleration] = "C") Then
        If MsgBox(" You are about to remove dates ? ", vbOKCancel) = vbOK Then
        
            AddStatus FileNumber, Now(), "Removed Demand Issued (" & [AccelerationIssued] & ") by " & GetFullName
        
            DoCmd.SetWarnings False
            strinfo = "Removed Demand Issued (" & [AccelerationIssued] & ") by " & GetFullName
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery
        
            'Put file in Demand queue
            Dim rstAccelerationIssued As Recordset
            Set rstAccelerationIssued = CurrentDb.OpenRecordset("Select * From WizardQueueStats Where FileNumber = " & FileNumber & "And Current = True", dbOpenDynaset, dbSeeChanges)
                If Not rstAccelerationIssued.EOF Then
                    With rstAccelerationIssued
                    .Edit
                    !DemandComplete = Null
                    !DemandWaiting = Null
                    !DemandQueue = Null
                 
                    .Update
                    End With
                    Else
                    MsgBox ("There is no Currrent Wizard Record for this File, Please Contact the IT")
                End If
            Set rstAccelerationIssued = Nothing
            Forms!foreclosuredetails![AccelerationIssued] = Null
            
                If Not IsNull(Forms!foreclosuredetails![AccelerationLetter]) Then
                    AddStatus FileNumber, Now(), "Removed Demand Expires date (" & [AccelerationLetter] & ") by " & GetFullName
                    
                    DoCmd.SetWarnings False
                    strinfo = "Removed Demand Expires date (" & [AccelerationLetter] & ") by " & GetFullName
                    strinfo = Replace(strinfo, "'", "''")
                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                    DoCmd.RunSQL strSQLJournal
                    DoCmd.SetWarnings True
                    Forms!Journal.Requery
                    
                    Forms!foreclosuredetails!AccelerationLetter = Null
                End If
                
                If Forms!foreclosuredetails!txtClientSentAcceleration = "C" Then
                    AddStatus FileNumber, Now(), ":  Removed C from the Demand Field by " & GetFullName
                    
                    DoCmd.SetWarnings False
                    strinfo = ":  Removed C from the Demand Field by " & GetFullName
                    strinfo = Replace(strinfo, "'", "''")
                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                    DoCmd.RunSQL strSQLJournal
                    DoCmd.SetWarnings True
                    Forms!Journal.Requery
                
                    Forms!foreclosuredetails!txtClientSentAcceleration = ""
                End If
        End If
        Else
        Exit Sub
    End If
    
Else
MsgBox ("The File has dispsotion not buy in or 3rd party, proceduer canceld")
Exit Sub
End If


End Sub

Private Sub NewFairDebt_Click()
If (IsNull(Disposition) Or (Disposition = 1 Or Disposition = 2)) Then

    If Not IsNull([FairDebt]) Then
        If MsgBox(" You are about to remove dates ? ", vbOKCancel) = vbOK Then
            
            AddStatus FileNumber, Now(), "Removed Fair Debt (" & [FairDebt] & ") by " & GetFullName
        
            DoCmd.SetWarnings False
            strinfo = "Removed Fair Debt (" & [FairDebt] & ") by " & GetFullName
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!Foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery
        
            'Put file in FairDebt queue
            Dim rstWFairDebt As Recordset
            Set rstWFairDebt = CurrentDb.OpenRecordset("Select * From WizardQueueStats Where FileNumber = " & FileNumber & "And Current = True", dbOpenDynaset, dbSeeChanges)
                If Not rstWFairDebt.EOF Then
                    With rstWFairDebt
                    .Edit
                   ' If IsNull(!RSIIcomplete) Then !RSIIcomplete = #1/1/2011#
                    If Not IsNull(!FairDebtComplete) Then !FairDebtComplete = Null
                    If Not IsNull(!FairDebtWaiting) Then !FairDebtWaiting = Null
                    'If Not IsNull(!NOIcomplete) Then !NOIcomplete = Null
                    '!DateInQueueNOI = Null
                    !AddFair = "Fair"
                  '  If Not IsNull(!RestartQueue) Then !FairDebtRestart = #1/1/2011#
                   .Update
                    End With
                    Else
                    MsgBox ("There is no Currrent Wizard Record for this File, Please Contact the IT")
                End If
            Set rstWFairDebt = Nothing
        
            Forms!foreclosuredetails![FairDebt] = Null
          
        End If
        Else
        Exit Sub
    End If
Else
MsgBox ("The File has dispsotion not buy in or 3rd party, proceduer canceld")
Exit Sub
End If

End Sub

Private Sub OriginalTrustee_AfterUpdate()
If Left([OriginalTrustee], 21) = "Commonwealth Trustees" Then
SubstituteTrustees = False
End If

End Sub

Private Sub PropReg_AfterUpdate()
If Not IsNull(PropReg) Then
AddStatus FileNumber, PropReg, "Property Registration sent"
'AddInvoiceItem FileNumber, "FC-Reg", "County Property Registration Fee", DLookup("Registrationcounty", "clientlist", "clientid=" & Forms![case list]!ClientID), 0, True, True, False, False
End If
End Sub

Private Sub PropReg_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    PropReg = Now()
    'AddStatus FileNumber, PropReg, "Property Registration sent"
    Call PropReg_AfterUpdate
End If

End Sub

Private Sub PurchaserAddress_AfterUpdate()
If Not BHproject Then
Dim strinfo As String
Dim strSQLJournal As String

If checkCmdEdit Then

    If Nz(PurchaserAddress) <> Nz(PurchaserAddress.OldValue) Then
        If IsNull(PurchaserAddress) And Not IsNull(PurchaserAddress.OldValue) Then
        MsgBox ("You should put a PurchaserAddress")
        Me.Undo
        Else
        DoCmd.SetWarnings False
        strinfo = " Edit Purchaser Address "
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!foreclosuredetails!FileNumber & ", #" & Now & "# , '" & GetFullName() & "', '" & strinfo & "'," & 1 & ")"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
        End If
    End If
Else
    If Not IsNull(PurchaserAddress) Then
            DoCmd.SetWarnings False
            strinfo = " Edit Purchaser Address "
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!foreclosuredetails!FileNumber & ", #" & Now & "# , '" & GetFullName() & "', '" & strinfo & "'," & 1 & ")"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery
    End If

End If

'If Not IsNull(SalePrice) And Not IsNull(Purchaser) And Not IsNull(PurchaserAddress) And Me.PurchaserAddress.Enabled = False Then ComEdit.Enabled = True
End If

End Sub

Private Sub ReviewAdProof_AfterUpdate()
AddStatus FileNumber, ReviewAdProof, "Reviewed Proof of Advertising"

If IsNull(NewspaperVendor) Or NewspaperVendor = "" Then
DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter Advertising(Newspapers) cost|FC-ADV|Advertising|Advertising"
End If
End Sub

Private Sub ReviewAdProof_DblClick(Cancel As Integer)
ReviewAdProof = Date
Call ReviewAdProof_AfterUpdate
End Sub

Private Sub SaleCert_AfterUpdate()
If Not IsNull(SaleCert) Then
AddStatus FileNumber, Date, "Sale Certification Complete"
End If

End Sub

Private Sub SalePrice_AfterUpdate()
If Not BHproject Then
Dim strinfo As String
Dim strSQLJournal As String

If checkCmdEdit Then

    If Nz(SalePrice) <> Nz(SalePrice.OldValue) Then
        If IsNull(SalePrice) And Not IsNull(SalePrice.OldValue) Then
        MsgBox ("You should put a sale price")
        Me.Undo
        
        Else
        DoCmd.SetWarnings False
        strinfo = " Edit Sale price "
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!foreclosuredetails!FileNumber & ", #" & Now & "# , '" & GetFullName() & "', '" & strinfo & "'," & 1 & ")"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
        End If
    End If
Else
    If Not IsNull(SalePrice) Then
            DoCmd.SetWarnings False
            strinfo = " Edit Sale price "
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!foreclosuredetails!FileNumber & ", #" & Now & "# , '" & GetFullName() & "', '" & strinfo & "'," & 1 & ")"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery
            
    End If
End If


'If Not IsNull(SalePrice) And Not IsNull(Purchaser) And Not IsNull(PurchaserAddress) And SalePrice.Enabled = False Then ComEdit.Enabled = True

End If


End Sub

Private Sub ServiceMailed_AfterUpdate()
If BHproject Then
If Not chMannerofService Then
AddStatus FileNumber, ServiceMailed, "Service Mailed"
End If
Else

If Not chMannerofService Then
AddStatus FileNumber, ServiceMailed, "Service Mailed"
AddInvoiceItem FileNumber, "FC-SVC", "Service Mailed Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, False, False, False
End If
End If

End Sub

Private Sub ServiceMailed_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ServiceMailed = Now()
    AddStatus FileNumber, ServiceMailed, "Service Mailed"
End If

End Sub

Private Sub StatePropReg_AfterUpdate()

If Not IsNull(StatePropReg) Then
AddStatus FileNumber, StatePropReg, "State property registration filed"
'AddInvoiceItem FileNumber, "FC-Reg", "State Property Registration Fee", DLookup("RegistrationState", "clientlist", "clientid=" & Forms![case list]!ClientID), 0, True, True, False, False
'AddInvoiceItem FileNumber, "FC-Reg", "State Property Registration Cost", DLookup("ivalue", "db", "id=44"), 0, True, True, False, False
End If
End Sub

Private Sub StatePropReg_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    StatePropReg = Now()
    Call StatePropReg_AfterUpdate
End If

End Sub

Private Sub StatusHearing_AfterUpdate()
StatusHearing.Locked = True
End Sub

Private Sub StatusHearing_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
If Not IsNull([StatusHearing]) Then
    If HearingCheking(StatusHearing, 1) = 1 Then
    Cancel = 1
 Else


If Weekday(StatusHearing) = vbSunday Or Weekday(StatusHearing) = vbSaturday Then
   
    MsgBox "Status Hearing date cannot be Saturday or Sunday", vbCritical
Cancel = 1
Else
Cancel = 0
End If
StatusHearingTime = Null
End If
End If
End If

End Sub

Private Sub StatusHearing_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    StatusHearing = Now()
    AddStatus FileNumber, Date, "Status Hearing filed " & Format$(StatusHearing, "m/d/yyyy")
End If

End Sub

Private Sub StatusHearingTime_AfterUpdate()
If Not BHproject Then

    If IsNull(StatusHearing) Or IsNull(StatusHearingTime) Then Exit Sub
    
'    If Hour(StatusHearingTime) < 8 Or Hour(StatusHearingTime) > 18 Then
'        StatusHearingTime = DateAdd("h", 12, StatusHearingTime)
'        If Hour(StatusHearingTime) < 8 Or Hour(StatusHearingTime) > 18 Then
'            MsgBox "Invalid Status Hearing Time: " & Format$(StatusHearingTime, "h:nn am/pm")
'            ExceptionsHearingTime = Null
'            Exit Sub
'        End If
'    End If

    If Hour(StatusHearingTime) >= 8 And Hour(StatusHearingTime) < 13 Then
        StatusHearingTime = Format$(StatusHearingTime, "h:nn am/pm")
    ElseIf Hour(StatusHearingTime) >= 1 And Hour(StatusHearingTime) <= 6 Then
        StatusHearingTime = DateAdd("h", 12, StatusHearingTime)
    Else
        MsgBox "Hearing time must be between 8:00 AM and 6:00 PM" ': " & Format$(DCHearingTime, "h:nn am/pm")
        StatusHearingTime = Null
        Exit Sub
    End If


    AddStatus FileNumber, Now(), "Status Hearing scheduled time for " & Format$(StatusHearing, "m/d/yyyy") & " at " & Format$(StatusHearingTime, "h:nn am/pm")
    'If Not IsNull(StatusHearing) Then Call UpdateCalendarStatusHearing
    StatusHearingTime.Locked = True

End If
    'AddStatus FileNumber, Now(), "Status Hearing scheduled time for " & Format$(StatusHearing, "m/d/yyyy") & " at " & Format$(StatusHearingTime, "h:nn am/pm")
End Sub

Private Sub StatusResults_AfterUpdate()
Select Case StatusResults

Case 1

AddStatus FileNumber, Date, " Status Hearing Continue"
    
     If Now() < StatusHearing.OldValue Then
      If Not IsNull(StatusHearingEntryID) Then
    Call DeleteCalendarEvent(StatusHearingEntryID)
     StatusHearingEntryID = Null
    StatusHearing.Value = Null
    StatusHearing.Enabled = False
    StatusHearingTime.Value = Null
    StatusHearingTime.Enabled = False
    AddNewDate.Enabled = True
    Else
    StatusHearing.Value = Null
    StatusHearing.Enabled = False
    StatusHearingTime.Value = Null
    StatusHearingTime.Enabled = False
    AddNewDate.Enabled = True
    End If
    Else
    StatusHearing.Value = Null
    StatusHearing.Enabled = False
    StatusHearingTime.Value = Null
    StatusHearingTime.Enabled = False
    AddNewDate.Enabled = True
    End If
    
  
Case 2
    AddStatus FileNumber, Date, " Status Hearing Dismiss"
    If Now() < StatusHearing Then
    If Not IsNull(StatusHearingEntryID) Then
    Call DeleteCalendarEvent(StatusHearingEntryID)
    StatusHearingEntryID = Null
    End If
    End If
Case 3
    AddStatus FileNumber, Date, " Status Hearing Resolved"
    If Now() < StatusHearing Then
    If Not IsNull(StatusHearingEntryID) Then
    Call DeleteCalendarEvent(StatusHearingEntryID)
    StatusHearingEntryID = Null
    End If
    End If
End Select



'If IsNull(StatusResults) Then Exit Sub
'AddStatus FileNumber, Date, StatusResults
'If Now() < StatusHearing Then
'If StatusResults = 1 Then
'   Call DeleteCalendarEvent(StatusHearingEntryID)
'   StatusHearingEntryID = Null
'   Exit Sub
'End If
'Else
'Exit Sub
'End If

End Sub



Private Sub TitleBack_AfterUpdate()
Me.Requery
If Not IsNull(Me.TitleBack) Then
DoCmd.SetWarnings False
Dim rstsql As String
rstsql = "Insert InTo TitleReceivedArchive (FileNumber, TitleRecieved, DateEntered) Values ( " & FileNumber & ", '" & Date & "' , '" & Now() & "')"
DoCmd.RunSQL rstsql
DoCmd.SetWarnings True
End If


End Sub

Private Sub TitleClaimDate_AfterUpdate()
AddStatus FileNumber, TitleClaimDate, "Title Claim Needed"
TitleClaim = True
Call Visuals
End Sub

Private Sub TitleClaimSent2_AfterUpdate()
If Not IsNull(TitleClaimSent2) Then
AddStatus FileNumber, TitleClaimSent2, "Sent title claim"
If IsNull(Comment) Then
    Comment = Format$(Date, "m/d/yyyy") & " Resolved title claim"
Else
    Comment = Comment & vbNewLine & Format$(Date, "m/d/yyyy") & " Resolved title claim"
End If
FeeAmount = DLookup("titleclaim", "clientlist", "clientid=" & Forms![Case List]!ClientID)
If MsgBox("Do you want to override the standard fee of $" & FeeAmount & " for this client?", vbYesNo) = vbYes Then
FeeAmount = InputBox("Please enter fee, then rememeber to note the journal")
MsgBox "Please upload fee approval to documents"
End If
        
AddInvoiceItem FileNumber, "FC-TC", "Title Claim", FeeAmount, 0, True, False, True, True
End If
Call Visuals
End Sub


Private Sub TitleThru_AfterUpdate()
Me.Requery
If Not IsNull(Me.TitleBack) Then
DoCmd.SetWarnings False
Dim rstsql As String
rstsql = "Insert InTo TitleThroughArchive (FileNumber, TitleThrough, DateEntered) Values ( " & FileNumber & ", '" & TitleThru & "' , '" & Now() & "')"
DoCmd.RunSQL rstsql
DoCmd.SetWarnings True
End If

End Sub

Private Sub txtClientSentAcceleration_AfterUpdate()

If Not IsNull(AccelerationLetter) Then
Dim rstJnl As Recordset, ClientSentAccelerationtxt As String
If ClientSentAcceleration = "C" Then
ClientSentAccelerationtxt = "Acceleration Letter noted as sent by client"
Else
ClientSentAcceleration = "Acceleration Letter noted as sent by Rosenberg & Associates"
End If
Set rstJnl = CurrentDb.OpenRecordset("select * from journal", dbOpenDynaset, dbSeeChanges)
With rstJnl
.AddNew
!FileNumber = FileNumber
!JournalDate = Now
!Who = GetFullName
!Info = ClientSentAccelerationtxt
!Color = 1
.Update
End With
Set rstJnl = Nothing
End If
AddStatus FileNumber, Date, ClientSentAccelerationtxt
End Sub

Private Sub ClientPaid_AfterUpdate()
If StaffID = 0 Then Call GetLoginName
Me.ReinstateClientPaidStaffID = StaffID
AddStatus FileNumber, Now(), "Proceeds sent to client"
sfrmDisbursingSurplus.Visible = True
End Sub

Private Sub ClientPaid_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ClientPaid = Now()
    AddStatus FileNumber, Now(), "Proceeds sent to client"
End If

End Sub
Private Sub cmdEditLegal_Click()
DoCmd.OpenForm "EditLegalDesc"
Forms!EditLegalDesc.txtFileNumber = FileNumber
Forms!EditLegalDesc.txtLegalDesc = LegalDescription
If Dirty Then DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub cmdSetLMDisposition_Click()
If Not PrivSetDisposition Then
MsgBox ("You do  not have permission to enter a disposition, see your Manager")
Exit Sub
End If

On Error GoTo Err_cmdSetLMDisposition_Click



If IsNull(LMDisposition) And PrivSetDisposition Then
     
    Call SetLMDisposition(0)
    Call DeleteFutureHearings(LMDispositionDate)
   ' Call RemoveSoftHold(FileNumber)
    
End If
    
Exit_cmdSetLMDisposition_Click:
    Exit Sub

Err_cmdSetLMDisposition_Click:
    MsgBox Err.Description
    Resume Exit_cmdSetLMDisposition_Click

End Sub



Private Sub DateOfDefault_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DateOfDefault)
If (Cancel = 1) Then Exit Sub

If Not IsNull(DateOfDefault) Then
  If (Not IsNull(Me.LastPaymentApplied)) Then
  
    If (DateDiff("d", DateOfDefault, LastPaymentApplied) >= 1) Then
    
        Cancel = 1
        MsgBox "Date must be on or after Last Payment Applied.", vbCritical
    End If
  End If
  
End If
End Sub

Private Sub DeedAppDate_AfterUpdate()
AddStatus FileNumber, DeedAppDate, "Deed App Dated"

End Sub

Private Sub DeedAppDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DeedAppDate = Now()
    Call DeedAppDate_AfterUpdate
End If
End Sub

Private Sub DeedAppReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DeedAppReceived = Date
    Me.DocBackDOA = True
    AddStatus FileNumber, Date, "Received Substitution of Trustee"
    
    DoCmd.SetWarnings False
    strinfo = "Received Substition of Trustee by " & GetFullName
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If

End Sub

Private Sub DeedAppRecorded_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DeedAppRecorded = Date
End If

End Sub

Private Sub DeedAppSentToRecord_AfterUpdate()
If BHproject Then
    If Not (IsNull(DeedAppSentToRecord)) Then
    AddStatus FileNumber, DeedAppSentToRecord, "Deed App sent to record"
    Exit Sub
    End If

Else
        If Not (IsNull(DeedAppSentToRecord)) Then
        AddStatus FileNumber, DeedAppSentToRecord, "Deed App sent to record"
          If State = "VA" Then
        DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total Overnight costs|FC-SOT|SOT Overnight Delivery Costs"
            AddInvoiceItem FileNumber, "FC-SOT", "Substitution of Trustees Mailing", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, True, False, True
          End If
        End If
End If

End Sub


Private Sub DeedAppSentToRecord_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DeedAppSentToRecord = Date
    Call DeedAppSentToRecord_AfterUpdate
End If

End Sub

Private Sub DeedtoRec_AfterUpdate()
If Not IsNull(DeedtoRec) And Me.DeedtoTitleCo = False Then
    AddStatus FileNumber, DeedtoRec, "Deed sent to record"
ElseIf Not IsNull(DeedtoRec) And Me.DeedtoTitleCo = True Then
    AddStatus FileNumber, DeedtoRec, "Deed sent to Title Company"
End If
'If Not IsNull(DeedtoRec) And Me.DeedtoTitleCo = False Then
If Not IsNull(DeedtoRec) Then
       
    Dim InvPct As Double
    Dim cbxClient As Integer
    cbxClient = Nz(DLookup("ClientID", "Caselist", "FileNumber=" & FileNumber))
    
    If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
        If State = "VA" Then
            InvPct = DLookup("VAPostsale", "clientlist", "clientid=" & cbxClient)
            FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
        ElseIf State = "MD" Then
            InvPct = DLookup("MDPostSale", "clientlist", "clientid=" & cbxClient)
            FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & cbxClient))

        ElseIf State = "DC" Then
            InvPct = DLookup("FeeDCReferral", "clientlist", "clientid=" & cbxClient)
            FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))

        End If

    End If
    
    If InvPct > 0 And FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "FC-REF", "Post Sale Attorney Fee- " & Format(InvPct, "percent") & " of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
       
    End If
    
    
'add to need invoice FC
DoCmd.SetWarnings False
DoCmd.RunSQL ("UPDATE CaseList set BillCase = True ,BillCaseUpdateUser = " & GetStaffID() & " ,BillCaseUpdateDate = Date() ,BillCaseUpdateReasonID = 35 WHERE [FileNumber] = " & Me.FileNumber)
DoCmd.SetWarnings True

Dim rstBillReasons As Recordset
Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstBillReasons
.AddNew
!FileNumber = FileNumber
!billingreasonid = 35
!UserID = GetStaffID
!Date = Date
.Update
End With


End If
'addedmilestone atty fee
End Sub

Private Sub DeedtoRec_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(DeedtoRec)
End If

End Sub

Private Sub DeedtoRec_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DeedtoRec = Now()
    'AddStatus FileNumber, DeedtoRec, "Deed sent to record"
    Call DeedtoRec_AfterUpdate
End If

End Sub

Private Sub DeedtoTitleCo_AfterUpdate()
Call Visuals
End Sub


Private Sub DismissalDate_AfterUpdate()
    
             
If MsgBox("Do you want to dismiss the case? ", vbYesNo) = vbYes Then
 AddStatus FileNumber, DismissalDate, "Case " & [CourtCaseNumber] & " Dismissed"
    CaseNuUpdate = True
    
        DoCmd.SetWarnings False
        
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & "Case " & [CourtCaseNumber] & " Dismissed" & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
              
    Call RemoveCaseFiled
    
     
      'Filing Fee
      AddInvoiceItem FileNumber, "FC-Oth", "Filing Fee for Dismissal", 15, 187, False, False, False, True
      'Overnight costs
      AddInvoiceItem FileNumber, "FC-Oth", "Overnight costs for Dismissal", 7, 77, False, False, False, True

End If

End Sub

Private Sub DismissalDate_BeforeUpdate(Cancel As Integer)
If Not PrivPrintPostSale Then
Cancel = True
Me.Undo
Else
Cancel = CheckFutureDate(DismissalDate)
End If

End Sub

Private Sub DismissalDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DismissalDate = Date
    Call DismissalDate_AfterUpdate
End If

End Sub

Private Sub DismissalSent_AfterUpdate()
AddStatus FileNumber, DismissalSent, "Dismissal Sent"

'added on 8_14_15 for Dismissal check to be invoiced
Dim newfilename As String
Dim selecteddoctype As Long
Dim fileextension As String
Dim DocDate As Date
Dim strSQL As String
Dim strSQLValues As String
Dim DocIDNo As Long
Dim clientShor As String
Dim rstBillReasons As Recordset

strSQL = ""
strSQLValues = ""
'DocDate = Now
selecteddoctype = 113

DoCmd.SetWarnings False

'newfilename = "Dismissal" & " " & Format$(Now(), "yyyymmdd hhnnss")

'If Dir$(DocLocation & DocBucket(txtFilenum) & "\" & txtFilenum & "\" & newfilename) <> "" Then
    'MsgBox txtFilenum & " already exists.", vbCritical
    'Exit Sub
'End If
'DoCmd.OutputTo acOutputForm, "AdvPostSaleCostPkg", acFormatPDF, DocLocation & DocBucket(txtFilenum) & "\" & txtFilenum & "\" & newfilename & ".pdf", False, "", 0 '-put doc like PDF form copy

'newfilename = newfilename & ".pdf"
'
'strSQLValues = FileNumber & "," & selecteddoctype & ",'" & "B" & "'," & GetStaffID() & ",'" & DocDate & "','" & Replace(newfilename, "'", "''") & "','" & Replace(newfilename, "'", "''") & "'"
'strsql = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
'DoCmd.RunSQL (strsql)

'Dim GroupCode As String

'GroupCode = Nz(DLookup("GroupCode", "DocumentTitles", "ID=" & selecteddoctype))
'If IsNull(GroupCode) Then


clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)
DocIDNo = GetDismissalDocIDNo(GetStaffID(), selecteddoctype, FileNumber)

'strsql = "Insert into Accou_PSAdvancedCostsPackageQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, Hold, MangNotic, DocIndexID, DocumentId, StaffID, StaffName) Values (" & FileNumber & ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & _
'        clientShor & " '," & Forms![Case List]!ClientID & ", Now(), '','', " & DocIDNo & ", " & selecteddoctype & ", " & GetStaffID() & ", '" & GetFullName() & "'" & ")"
If Not IsNull(DocIDNo) Then
    strSQL = "Insert into Accou_PSAdvancedCostsPackageQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, Hold, MangNotic, DocIndexID, DocumentId, StaffID, StaffName) Values (" & FileNumber & ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & _
    clientShor & " '," & Forms![Case List]!ClientID & ", Now(), '','', " & DocIDNo & ", " & selecteddoctype & ", " & GetStaffID() & ", '" & GetFullName() & "'" & ")"
Else
    strSQL = "Insert into Accou_PSAdvancedCostsPackageQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, Hold, MangNotic, DocIndexID, DocumentId, StaffID, StaffName) Values (" & FileNumber & ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & _
    clientShor & " '," & Forms![Case List]!ClientID & ", Now(), '','', 0, '', " & GetStaffID() & ", '" & GetFullName() & "'" & ")"
End If

DoCmd.RunSQL strSQL

DoCmd.SetWarnings True

End Sub

Private Sub DismissalSent_BeforeUpdate(Cancel As Integer)
If Not PrivPrintPostSale Then
Cancel = True
Me.Undo
Else
Cancel = CheckFutureDate(DismissalSent)
End If

End Sub

Private Sub DismissalSent_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DismissalSent = Date
    Call DismissalSent_AfterUpdate
End If
End Sub

Private Sub DispositionRescinded_AfterUpdate()
If Not IsNull(DispositionRescinded) Then
Forms![Case List]!BillCase = True
Forms![Case List]!BillCaseUpdateUser = GetStaffID()
Forms![Case List]!BillCaseUpdateDate = Date
Forms![Case List]![BillCaseUpdateReasonID] = 6
Forms![Case List]!lblBilling.Visible = True
Forms![Case List].SetFocus
DoCmd.RunCommand acCmdSaveRecord
Forms![foreclosuredetails].SetFocus

Dim rstBillReasons As Recordset
Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstBillReasons
.AddNew
!FileNumber = FileNumber
!billingreasonid = 6
!UserID = GetStaffID
!Date = Date
.Update
End With

AddStatus FileNumber, DispositionRescinded, "Disposition Rescinded"
AddInvoiceItem FileNumber, "FC", "Disposition Rescinded- Placeholder entry", 0, 0, True, False, False, False
End If
End Sub

Private Sub DispositionRescinded_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DispositionRescinded)

End Sub

Private Sub DispositionRescinded_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DispositionRescinded = Date
    Call DispositionRescinded_AfterUpdate
End If

End Sub

Private Sub DocBackAff7105_AfterUpdate()

If DocBackAff7105 Then
    AddStatus FileNumber, Date, "Received Affidavit Pursuant to MD Real Property Code 7-105.1(D)(II)"
Else
    AddStatus FileNumber, Date, "Removed Affidavit Pursuant to MD Real Property Code 7-105.1(D)(II)"
End If

End Sub

Private Sub DocBackDOA_AfterUpdate()
If DocBackDOA Then
    DeedAppReceived = Date
    AddStatus FileNumber, Date, "Received Substitution of Trustee"
    
    DoCmd.SetWarnings False
    strinfo = "Received Substition of Trustee by " & GetFullName
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
    
Else
    DeedAppReceived = Null
    AddStatus FileNumber, Date, "Removed Substitution of Trustee"

    DoCmd.SetWarnings False
    strinfo = "Removed Substition of Trustee by " & GetFullName
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!foreclosuredetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery

End If
End Sub

Private Sub DocBackLostNote_AfterUpdate()

If DocBackLostNote Then
    AddStatus FileNumber, Date, "Received Lost Note Affidavit"
Else
    AddStatus FileNumber, Date, "Removed Lost Note Affidavit"
End If

End Sub

Private Sub DocBackMilAff_AfterUpdate()

If DocBackMilAff Then
    AddStatus FileNumber, Date, "Received Military Affidavit"
Else
    AddStatus FileNumber, Date, "Removed Military Affidavit"
End If

End Sub

Private Sub DocBackOrigNote_AfterUpdate()

If DocBackOrigNote Then
    AddStatus FileNumber, Date, "Received Original Note"
Else
    AddStatus FileNumber, Date, "Removed Original Note"
End If

End Sub

Private Sub DocBackSOD_AfterUpdate()

If DocBackSOD Then
    AddStatus FileNumber, Date, "Received Statement of Debt"
Else
    AddStatus FileNumber, Date, "Removed Statement of Debt"
End If

End Sub

Private Sub Docket_AfterUpdate()
If BHproject Then
AddStatus FileNumber, Docket, "First Legal filed"
Else

AddStatus FileNumber, Docket, "First Legal filed"
If State = "DC" Then
    AddInvoiceItem FileNumber, "FC-DKT", "Substitution of Trustees", 26.5, 0, False, True, False, False
    AddInvoiceItem FileNumber, "FC-DKT", "Filing Fee", 26.5, 0, False, True, False, False
End If
End If

End Sub

Private Sub Docket_BeforeUpdate(Cancel As Integer)
If Not BHproject Then

If WizardSource <> "Service" Then
MsgBox "You do not have privileges to edit the Docket date", vbCritical
Cancel = 1
Exit Sub
End If

Cancel = CheckFutureDate(Docket)
If (Cancel = 1) Then Exit Sub

If Not IsNull(Disposition) And Nz(SaleCompleted) = 0 Then
    If Docket > DateAdd("d", 7, DispositionDate) Then
        Cancel = 1
        MsgBox "Date must be within 7 days of setting the Disposition.  Otherwise you probably need to add a foreclosure.", vbCritical
    End If
ElseIf (DateDiff("d", NOI, Docket) < 46) Then
  Cancel = 1
  MsgBox "Date must be 45 days after 45 Day Notice", vbCritical
End If
End If

End Sub

Private Sub Docket_DblClick(Cancel As Integer)

If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    If IsNull(Disposition) Then
        Docket = Now()
        Call Docket_AfterUpdate
    End If
End If

End Sub

Private Sub DocsBack_AfterUpdate()
AddStatus FileNumber, DocsBack, "Received all affidavits from " & Forms![Case List]!ClientID.Column(1)
If IsNull(DeedAppReceived) Then DeedAppReceived = Date
End Sub

Private Sub DocsBack_BeforeUpdate(Cancel As Integer)

Cancel = CheckFutureDate(DocsBack)
If (Cancel = 1) Then Exit Sub

If Not IsNull(Disposition) And Nz(SaleCompleted) = 0 Then
    If DocsBack > DateAdd("d", 7, DispositionDate) Then
        Cancel = 1
        MsgBox "Date must be within 7 days of setting the Disposition.  Otherwise you probably need to add a foreclosure.", vbCritical
    End If
End If
End Sub

Private Sub DocsBack_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    If IsNull(Disposition) Then
        DocsBack = Date
        AddStatus FileNumber, DocsBack, "Received executed documents"
        If IsNull(DeedAppReceived) Then DeedAppReceived = Date
    End If
End If

End Sub

Private Sub DocstoClient_AfterUpdate()
If BHproject Then
AddStatus FileNumber, DocstoClient, "Affidavits sent to " & Forms![Case List]!ClientID.Column(1)
Else

AddStatus FileNumber, DocstoClient, "Affidavits sent to " & Forms![Case List]!ClientID.Column(1)

'REMOVED 1099 - 9/8/2014 MC
    'If Forms![Case List]!ClientID = 446 Then
    'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee - Bill 30% @ Affidavits", 0, 0, True, True, False, False
    'Forms![Case List]!BillCase = True
    'Forms![Case List]!BillCaseUpdateUser = GetStaffID()
    'Forms![Case List]!BillCaseUpdateDate = Date
    'Forms![Case List]![BillCaseUpdateReasonID] = 7
    'Forms![Case List]!lblBilling.Visible = True
    'Forms![Case List].SetFocus
    'DoCmd.RunCommand acCmdSaveRecord
    'Forms![ForeclosureDetails].SetFocus
    
    'Dim rstBillReasons As Recordset
    'Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
    'With rstBillReasons
    '.AddNew
    '!FileNumber = FileNumber
    '!billingreasonid = 7
    '!userid = GetStaffID
    '!Date = Date
    '.Update
    'End With
    'End If
End If
End Sub

Private Sub DocstoClient_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DocstoClient)
End Sub

Private Sub DocstoClient_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DocstoClient = Now()
    'AddStatus FileNumber, DocstoClient, "Overnighted documents to " & Forms![Case List]!ClientID.Column(1)
    Call DocstoClient_AfterUpdate
End If

End Sub

Private Sub ExceptionsFiled_Click()
If ExceptionsFiled Then
    AddStatus FileNumber, Date, "Exceptions filed to sale"
Else
    If MsgBox("Really ""undo"" exceptions filed?  This will reset the hearing date and hearing status.  You should check the status report.", vbQuestion + vbYesNo) <> vbYes Then
        ExceptionsFiled = 1
        Exit Sub
    End If
End If
Call Visuals
End Sub

Private Sub ExceptionsHearing_AfterUpdate()

'AddStatus FileNumber, Date, "Exceptions to foreclosure sale scheduled for " & Format$(ExceptionsHearing, "m/d/yyyy")

Call Visuals
ExceptionsHearing.Locked = True

End Sub

Private Sub ExceptionsHearing_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ExceptionsHearing = Now()
    AddStatus FileNumber, Date, "Exceptions to foreclosure sale scheduled for " & Format$(ExceptionsHearing, "m/d/yyyy")
    Call Visuals
End If

End Sub

Private Sub Visuals()
Dim X As Boolean, lt As Long

X = (UCase$(Nz(State)) = "MD")
StatementOfDebtAmount.Enabled = X
Report.Enabled = (X And (Nz(Disposition) <> 7))
NiSiEnd.Enabled = (X And (Nz(Disposition) <> 7))
SaleRat.Enabled = (X And (Nz(Disposition) <> 7))
pgNOI.Enabled = X
NOI.Enabled = X
txtNOIExpires.Enabled = X

X = (UCase$(Nz(State)) = "MD" Or UCase$(Nz(State)) = "DC")
If X Then
    'SentToDocket.Enabled = True
    'SentToDocket.Locked = False
    SentToDocket.BackStyle = 1
    Docket.Enabled = True
    Docket.Locked = False
    Docket.BackStyle = 1
    
    If (IsNull(LossMitFinalDate)) Then
      
      Call SetObjectAttributes(LossMitFinalDate, True)
    Else
      
      Call SetObjectAttributes(LossMitFinalDate, False)
    End If
End If

BorrowerServed.Enabled = (UCase$(Nz(State)) = "MD")

X = (UCase$(Nz(State)) = "MD" Or UCase$(Nz(State)) = "VA" Or UCase$(Nz(State)) = "DC")

cmdAudit.Enabled = X
AuditFile.Enabled = (X And (Nz(Disposition) <> 7))
AuditRat.Enabled = (X And (Nz(Disposition) <> 7))

AssessedValue.Enabled = (UCase$(Nz(State)) = "VA")

TitleClaimSent.Enabled = TitleClaim
TitleClaimResolved.Enabled = TitleClaim

If IsNull(LoanType) Then
    lt = 0
Else
    lt = LoanType
End If
FHALoanNumber.Enabled = (lt = 2 Or lt = 3)    ' enable for VA or HUD
FirstLegal.Enabled = (lt = 2 Or lt = 3)
FNMALoanNumber.Enabled = (lt = 4)
FHLMCLoanNumber.Enabled = (lt = 5)
FNMAHoldReason.Enabled = (lt = 4)
FNMAHoldReasonDate.Enabled = (lt = 4)
sfrmFNMAMissingDocs.Enabled = (lt = 4)

FNMAPostponeReason.Enabled = (lt = 4)
FNMAPostponeReasonDate.Enabled = (lt = 4)

HUDOccLetter.Enabled = (lt = 3)
VAAppraisal.Enabled = (lt = 2)

FinalPkg.Enabled = (Nz(Disposition) <> 2) And (lt = 2 Or lt = 3)     ' not 3rd party and (HUD or VA)

'Disabled 9/24/2014 MC  - Ticket 1178 will handle all Ground rents etc
'GroundRentAmount.Enabled = IIf(optLeasehold = 1, -1, 0) ' it is changes from optLeasehold because of DC project Ticket 866 SA 06/3
'GroundRentPayable.Enabled = IIf(optLeasehold = 1, -1, 0) ' it is changes from optLeasehold because of DC project Ticket 866 SA 06/3

If ExceptionsFiled Then
    ExceptionsHearing.Enabled = True
    ExceptionsHearingTime.Enabled = True
    cbxSustained.Enabled = True
     
       
       
Else
'    If Not FileReadOnly Then ExceptionsHearing = Null ' there is no need to change the form in case FileReadOnly (this is made mistack) Sarab
'    If Not FileReadOnly Then ExceptionsHearingTime = Null
    ExceptionsHearing.Enabled = False
    ExceptionsHearingTime.Enabled = False
'    If Not FileReadOnly Then cbxSustained = 0
    cbxSustained.Enabled = False
End If

       
    

IRSNotice.Enabled = IRSLiens

'RecordDeed.Enabled = Not DeedtoTitleCo
'RecordDeedLiber.Enabled = Not DeedtoTitleCo
'RecordDeedFolio.Enabled = Not DeedtoTitleCo

Settled.Enabled = (Nz(Disposition) = 2)     ' enable only for 3rd party

ResellMotion.Enabled = Resell
ResellServed.Enabled = Resell
ResellShowCauseExpires.Enabled = Resell
ResellAnswered.Enabled = Resell
ResellGranted.Enabled = Resell

AmmDocBackSOD.Enabled = (Nz(Disposition = 2) Or Nz(Disposition = 1))  ' enable if 3rd party or BI
AmmStatementOfDebtDate.Enabled = (Nz(Disposition = 2) Or Nz(Disposition = 1))
AmmStatementOfDebtAmount.Enabled = (Nz(Disposition = 2) Or Nz(Disposition = 1))
AmmStatementOfDebtPerDiem.Enabled = (Nz(Disposition = 2) Or Nz(Disposition = 1))

Audit2File.Enabled = (Nz(JurisdictionID) = 17)
Audit2Rat.Enabled = (Nz(JurisdictionID) = 17)

lstTrustees.Visible = (CaseTypeID <> 8)
cbxAddTrustee.Visible = (CaseTypeID <> 8)
cmdRemoveTrustee.Visible = (CaseTypeID <> 8)

MonitorTrusteeName.Visible = (CaseTypeID = 8)
MonitorTrusteeAddress.Visible = (CaseTypeID = 8)
MonitorTrusteeContact.Visible = (CaseTypeID = 8)
MonitorMotionSurplusFiled.Visible = (CaseTypeID = 8)
MonitorOrderSurplus.Visible = (CaseTypeID = 8)
MonitorClientPaid.Visible = (CaseTypeID = 8)

SubstitutePurchaser.Enabled = (CaseTypeID <> 8)
OrderSubsPurch.Enabled = (CaseTypeID <> 8)
ExceptionsFiled.Enabled = (CaseTypeID <> 8)
ExceptionsHearing.Enabled = (CaseTypeID <> 8)
ExceptionsHearingTime.Enabled = (CaseTypeID <> 8)
cbxSustained.Enabled = (CaseTypeID <> 8)

ClientPaid.Enabled = Nz((Disposition = 2 And Not IsNull(Settled)) Or (Disposition = 4 Or Disposition = 26))
 
If SubstituteTrustees Then
DeedAppReceived.Enabled = True
DeedAppDate.Enabled = True
DeedAppSentToRecord.Enabled = True
DeedAppRecorded.Enabled = True
DeedAppLiber.Enabled = True
DeedAppFolio.Enabled = True
End If

If Disposition > 2 And Disposition <> 7 And Disposition <> 11 And (Date - DispositionDate) > 14 Then
DeedAppSentToRecord.Enabled = False
End If

X = Not ((Nz(Disposition) = 2) Or (Nz(Disposition) = 1))
If PrivSetSale And IsNull(Notices) And X Then
    Sale.Enabled = True
 '   Sale.Locked = False
    Sale.BackStyle = 1
    SaleTime.Enabled = True
    SaleTime.Locked = False
    SaleTime.BackStyle = 1
'    Deposit.Enabled = True
'    Deposit.Locked = False
'    Deposit.BackStyle = 1
Else
    Sale.Enabled = False
 '   Sale.Locked = True
    Sale.BackStyle = 0
    SaleTime.Enabled = False
    SaleTime.Locked = True
    SaleTime.BackStyle = 0
'    Deposit.Enabled = False
'    Deposit.Locked = True
'    Deposit.BackStyle = 0
End If


If (PrivAdjustDeposit = False Or (PrivAdjustDeposit = True And Me.Disposition <> 2)) Then
    Deposit.Enabled = False
    Deposit.Locked = True
    Deposit.BackStyle = 0
Else
    Deposit.Enabled = True
    Deposit.Locked = False
    Deposit.BackStyle = 1
End If
'stopped below by Sarab for Dismissal project 6/26/2015
'If ((UCase$(Nz(State)) = "MD") And (Disposition = 3 Or Disposition = 4) And Not IsNull([SentToDocket])) Then
'  Call SetObjectAttributes(DismissalSent, True)
'  Call SetObjectAttributes(DismissalDate, True)
'Else
'  Call SetObjectAttributes(DismissalSent, False)
'  Call SetObjectAttributes(DismissalDate, False)
'End If

'If (Me.Disposition = 4 Or Me.Disposition = 26) Then
'  Call SetObjectAttributes(Me.ReinstateClientPaidDate, True)
'Else
'  Call SetObjectAttributes(Me.ReinstateClientPaidDate, False)
'End If
'If (IsNull(Forms!sfrmFCtitle!TitleAssignRecordedDate)) Then
'  Call SetObjectAttributes(RecordedFolio, False)
'  Call SetObjectAttributes(RecordedLiber, False)

'Else
'  Call SetObjectAttributes(RecordedFolio, True)
'  Call SetObjectAttributes(RecordedLiber, True)

'End If
On Error Resume Next
LossMitSolicitationDate.Enabled = (LoanType = 4) Or (LoanType = 5)
Resume Next

End Sub

Private Sub FairDebt_AfterUpdate()
AddStatus FileNumber, FairDebt, "Fair Debt Letter sent"
End Sub

Private Sub FairDebt_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(FairDebt)
End Sub

Private Sub FairDebt_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    If BHproject Then
    FairDebt = Now()
    AddStatus FileNumber, FairDebt, "Fair Debt Letter sent"
    End If
    
End If
End Sub

Private Sub AccelerationLetter_AfterUpdate()
If BHproject Then
 AddStatus FileNumber, Date, "Acceleration Payoff Due " & Format$(AccelerationLetter, "mm/dd/yyyy")
Else


Dim FeeAmount As Currency

If Not IsNull(AccelerationLetter) Then
    AddStatus FileNumber, Date, "Acceleration Payoff Due " & Format$(AccelerationLetter, "mm/dd/yyyy")
    Select Case LoanType
    Case 4
    FeeAmount = Nz(DLookup("FeeAcceleration", "ClientList", "ClientID=177"))
    Case 5
    FeeAmount = Nz(DLookup("FeeAcceleration", "ClientList", "ClientID=263"))
    Case Else
    FeeAmount = Nz(DLookup("FeeAcceleration", "ClientList", "ClientID=" & Forms![Case List]!ClientID), 0)
    End Select
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "FC-ACC", "Acceleration Letter", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "FC-ACC", "Acceleration Letter", 1, 0, True, True, False, False
    End If
End If
 DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter postage for mailing the Acceleration Letter|FC-ACC|Acceleration Letter Postage"

End If

End Sub

Private Sub FairDebtDispute_AfterUpdate()
AddStatus FileNumber, FairDebtDispute, "Debt dispute received"
End Sub

Private Sub FairDebtDispute_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(FairDebtDispute)
End Sub

Private Sub FairDebtDispute_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    FairDebtDispute = Date
    Call FairDebtDispute_AfterUpdate
End If
End Sub

Private Sub FairDebtVerified_AfterUpdate()
If BHproject Then
  If Not IsNull(FairDebtVerified) Then
    AddStatus FileNumber, FairDebtVerified, "Debt verified and notice sent"
  Exit Sub
  End If
Else

  

            If Not IsNull(FairDebtVerified) Then
            
                If IsNull(FairDebtDispute) Then
                   MsgBox ("Please Add the Fair Debt Disputed Received Date")
                   FairDebtVerified = Null
                   Exit Sub
                   Else
                   
                        AddStatus FileNumber, FairDebtVerified, "Debt verified and notice sent"
                        Dim strInsert As String
                        Dim clientShor As String
                        Dim StrJuirs As String
                        StrJuirs = DLookup("Jurisdiction", "JurisdictionList", "JurisdictionID= " & Forms![Case List]!JurisdictionID)
                        StrJuirs = Replace(StrJuirs, "'", "''")
                        clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)
                        DoCmd.SetWarnings False
                        strInsert = "Insert Into Tracking_DebtVerified (CaseFile,ProjectName,ClientShortName,Juris,Client,DIT,DebtDisputeRec,StaffID,StaffName) Values (" & FileNumber & ",'" & Forms![Case List]!PrimaryDefName & "','" & clientShor & "','" & StrJuirs & "', " & Forms![Case List]!ClientID & ", #" & Now() & "#, #" & Forms![foreclosuredetails]!FairDebtDispute & "#," & GetStaffID & ",'" & GetFullName() & "')"
                        DoCmd.RunSQL strInsert
                        DoCmd.SetWarnings True
                   
                    
                    End If
            End If
        
End If



End Sub

Private Sub FairDebtVerified_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(FairDebtVerified)
End Sub

Private Sub FairDebtVerified_DblClick(Cancel As Integer)
If FileReadOnly Then
   DoCmd.CancelEvent
Else

    If IsNull(FairDebtDispute) Then
    MsgBox ("Please Add the Fair Debt Disputed Received Date")
    Exit Sub
    Else
    FairDebtVerified = Date
    Call FairDebtVerified_AfterUpdate
    End If
    
End If
End Sub

Private Sub FCTab_Change()
If Not CaseNuUpdate Then
If Forms![Case List]!CaseType = "Foreclosure" And (Not IsNull(Me.SalePrice) Or Not IsNull(Me.Purchaser)) And IsNull(Me.DispositionDesc) Then
     MsgBox ("You have entered sale information but have not entered a buy in or third party sale disposition. You will lose your information unless you enter the disposition now")
     Me.Post_Sale.SetFocus
     Me.SalePrice = Null
     Me.Purchaser = Null
     Me.PurchaserAddress = Null
     DoCmd.RunCommand acCmdSaveRecord
     
    ' Call cmdSetDisposition_Click
 
     

End If

'Me.cmdSetDisposition.Enabled = False








Select Case FCTab.Value     ' 0 based
'    Case 6      ' title
'        If Not IsNull(OriginalPBal) Then
'            If IsNull(TitleReviewNameOf) Then TitleReviewNameOf = GetNames(FileNumber, 2, "Owner=True") & _
'                    " Owner" & IIf(CountNames(FileNumber, "Owner=True") > 1, "s", "") & " of a " & _
'                    IIf(Leasehold, "leasehold property with an annual ground rent of " & _
'                    Format$(GroundRentAmount, "Currency") & " payable " & GroundRentPayable, _
'                    "fee simple property") & " by Deed dated"
'            If IsNull(TitleReviewLiens) Then TitleReviewLiens = "1.  " & DOTWord(DOT) & " dated " & _
'                Format$(DOTdate, "mmmm d, yyyy") & " securing " & OriginalBeneficiary & " in the original amount of " & _
'                Format$(OriginalPBal, "Currency") & " and recorded on " & Format$(DOTrecorded, "mmmm d, yyyy") & _
'                " " & LiberFolio(Liber, Folio, State)
'            If IsNull(TitleReviewStatus) Then If Not IsNull(LienPosition) Then TitleReviewStatus = Investor & " is foreclosing in " & Ordinal(LienPosition) & " position."
'        End If
    Case 11      ' status

        If Dirty Then DoCmd.RunCommand acCmdSaveRecord

        sfrmStatus.Requery
End Select

End If

End Sub

Private Sub FinalPkg_AfterUpdate()
AddStatus FileNumber, FinalPkg, "Final package sent"
End Sub

Private Sub FinalPkg_BeforeUpdate(Cancel As Integer)
If BHproject Then
Cancel = CheckFutureDate(FinalPkg)
End If
End Sub

Private Sub FinalPkg_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    FinalPkg = Now()
    AddStatus FileNumber, FinalPkg, "Final package sent"
End If

End Sub

Private Sub FirstPub_AfterUpdate()

'If FirstPub >= DateAdd("D", 120, LPIDate) Then
   
    'Milestone Billing for Referral Fee
    Dim InvPct As Double
    If State = "VA" And Not IsNull(FirstPub) Then
    Dim cbxClient As Integer
    cbxClient = Nz(DLookup("ClientID", "Caselist", "FileNumber=" & FileNumber))
   Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
      Case 1 'Conventional
        FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
      Case 2 'VA or Veteran's Affairs
        FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
      Case 3 'FHA or HUD
        FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=570")) 'HUD/FHA
      Case 4
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=177")) 'Fannie Mae
      Case 5
        FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=263")) 'Freddie Mac
      Case Else
        FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
    End Select
    If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
      InvPct = DLookup("VA1stactionpct", "clientlist", "clientid=" & cbxClient)
    Else
      InvPct = 1
    End If
    If FeeAmount > 0 Then
      If InvPct <= 1 Then
        AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due at 1st action of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      Else
        'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
      End If
    End If
    'Removed per Diane 1/30, do not milestone bill VA files
    'Forms![case list]!BillCase = True
    'Forms![case list]!BillCaseUpdateUser = GetStaffID()
    'Forms![case list]!BillCaseUpdateDate = Date
    'Forms![case list]!BillCaseUpdateReasonID = 3
    'Forms![case list]!lblBilling.Visible = True
    
    'Dim rstBillReasons As Recordset
    'Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
    'With rstBillReasons
    '.AddNew
    '!FileNumber = FileNumber
    '!billingreasonid = 3
    '!userid = GetStaffID
    '!Date = Date
    '.Update
    'End With
    
    End If
    AddStatus FileNumber, FirstPub, "1st Legal Advertisement Runs"
'Else
 '   MsgBox "First Publication Dates Must Be 120 days past the LPI Date"
  '  FirstPub = Null
  '  Exit Sub
'End If

End Sub

Private Sub FirstPub_BeforeUpdate(Cancel As Integer)
'Dim Info As Date
If Not BHproject Then
If Not IsNull(Disposition) And Nz(SaleCompleted) = 0 Then
    If FirstPub > DateAdd("d", 7, DispositionDate) Then
        Cancel = 1
        MsgBox "Date must be within 7 days of setting the Disposition.  Otherwise you probably need to add a foreclosure.", vbCritical
    End If
End If

If FirstPub <= DateDiff("d", -120, LPIDate) And Forms![Case List]!ClientID <> 567 Then
    Cancel = 1
    MsgBox "First Publication Dates Must Be 120 days past the LPI Date.", vbCritical
    'Info = DateDiff("d", -120, LPIDate)
End If
End If

End Sub

Private Sub FirstPub_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    If IsNull(Disposition) Then
        FirstPub = Now()
        Call FirstPub_AfterUpdate
    End If
End If

End Sub

Private Sub FNMAHoldReason_AfterUpdate()
  If (Not IsNull(FNMAHoldReason)) Then
    Me.FNMAHoldReasonDate = Date
  Else
    Me.FNMAHoldReasonDate = Null
  End If
End Sub

Private Sub FNMAMissingDoc_AfterUpdate()
  If (Not IsNull(FNMAMissingDoc)) Then
    Me.FNMAMissingDocDate = Date
  Else
    Me.FNMAMissingDocDate = Null
  End If
End Sub

Private Sub FNMAPostponeReason_AfterUpdate()
  If (Not IsNull(FNMAPostponeReason)) Then
    Me.FNMAPostponeReasonDate = Date
  Else
    Me.FNMAPostponeReasonDate = Null
  End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
If BHproject Or CaseNuUpdate = True Then
    Exit Sub
Else
    
    If Me.State = "DC" Then
        
      'DC StatusHearing
        If Not IsNull(StatusHearing) And Not IsNull(StatusHearingTime) Then
            If (Nz(StatusHearing) <> Nz(StatusHearing.OldValue)) Or (Nz(StatusHearingTime) <> Nz(StatusHearingTime.OldValue)) Then Call UpdateCalendarStatusHearing
        End If
    ' DC Sale
        If (Nz(txtSale) <> Nz(txtSale.OldValue)) Or (Nz(txtSaleTime) <> Nz(txtSaleTime.OldValue)) Then Call UpdateCalendarDC
        Exit Sub
          
    Else
        If (Nz(Sale) <> Nz(Sale.OldValue)) Or (Nz(SaleTime) <> Nz(SaleTime.OldValue)) Then Call UpdateCalendar
        If cbxSustained.Value = Null Then
            If (Nz(ExceptionsHearing) <> Nz(ExceptionsHearing.OldValue)) Or (Nz(ExceptionsHearingTime) <> Nz(ExceptionsHearingTime.OldValue)) Then Call UpdateCalendarExceptionHearing
        End If
        
    'If StatusResults.Value = Null Then
        If Not IsNull(StatusHearing) And Not IsNull(StatusHearingTime) Then
            If (Nz(StatusHearing) <> Nz(StatusHearing.OldValue)) Or (Nz(StatusHearingTime) <> Nz(StatusHearingTime.OldValue)) Then Call UpdateCalendarStatusHearing
        End If
    
    End If

End If

End Sub

Private Sub Form_Current()
DismissalDate.Enabled = PrivPrintPostSale
ComRemoveCase.Enabled = PrivPrintPostSale
If BHproject Then Call PHproejctFC
If Not BHproject Then

    If PrivNewNOIFDDemaind Then
    NewFairDebt.Visible = True
    NewDemand.Visible = True
    New45Notice.Visible = True
    End If
    
        
        
        If Me.StatusResults.Value = 1 Then
        StatusHearing.Enabled = False
        StatusHearingTime.Enabled = False
        AddNewDate.Enabled = True
        End If
    
        
        
        
        If (Me!State = "VA" Or Me!State = "DC") Then
        PropReg.Enabled = False
        StatePropReg.Enabled = False
          
        End If
End If

    
Dim FC As Recordset, rstEV As Recordset, EVFileNumbers As String


If FileReadOnly Or (Not IsNull(FairDebtDispute) And IsNull(FairDebtVerified)) Then
    Me.AllowEdits = False
    cmdAddFC.Enabled = False
    cmdAudit.Enabled = False
    cmdPrint.Enabled = False
    cmdEditPropertyDetails.Enabled = False
    cmdEditLegal.Enabled = False
    sfrmNames.Form.AllowEdits = False
    sfrmNames.Form.AllowAdditions = False
    sfrmNames.Form.AllowDeletions = False
    sfrmNames!cmdCopyClient.Enabled = False
    sfrmNames!cmdCopy.Enabled = False
    sfrmNames!cmdTenant.Enabled = False
    sfrmNames!cmdDelete.Enabled = False
    sfrmNames!cmdNoNotice.Enabled = False
    sfrmFCtitle.Form.AllowEdits = False
    sfrmFCtitle.Form.AllowAdditions = False
    sfrmFCtitle.Form.AllowDeletions = False
    sfrmFCDIL.Form.AllowEdits = False
    sfrmFCDIL.Form.AllowAdditions = False
    sfrmFCDIL.Form.AllowDeletions = False
    cmdRemoveTrustee.Enabled = False
    cmdCalcPerDiem.Enabled = False
    cmdPurchaserInvestor.Enabled = False
    sfrmStatus.Form.AllowEdits = False
    sfrmStatus.Form.AllowAdditions = False
    sfrmStatus.Form.AllowDeletions = False
    CommdEdit.Enabled = False
    ComAddName.Enabled = False
        
    
   ' Forms!foreclosuredetails!sfrmDCComplaintNew.cmdNew.Enabled = False
    
    If (Not IsNull(FairDebtDispute) And IsNull(FairDebtVerified)) Then
        Me.AllowEdits = True
        
        Dim ctl As Control
        Dim lngI As Long
        Dim bSkip As Boolean

         For Each ctl In Form.Controls
        Select Case ctl.ControlType
         Case acTextBox, acComboBox, acListBox, acOptionGroup, acCheckBox, acSubform, acOptionButton
        
            If Not (ctl.Locked) Then ctl.Locked = True
            
        Case acCommandButton
        bSkip = False
            If ctl.Name = "cmdClose" Then bSkip = True
            If ctl.Name = "cmdSelectFile" Then bSkip = True
            If Not bSkip Then ctl.Enabled = False
         
     
    End Select
    Next
         If Me!State = "DC" And Not IsNull(Me.Child794!SentClientComplaint) Then
                If Me.FairDebtDispute > Me.Child794!SentClientComplaint Then
                    Me.Child794.Enabled = True
                    Me.Child794.Locked = False
    
                        For Each ctl In Me.Child794.Controls
                            Select Case ctl.ControlType
                            Case acTextBox
                            
                              If ctl.Name = "ReceivedClientComplaintSinged" Then
                                Me.Child794!ReceivedClientComplaintSinged.Locked = False
                                Else
                                   If Not (ctl.Locked) Then ctl.Locked = True
                        
                              End If
                            End Select
                        Next
                End If
        End If
     
            
      If Me!State = "DC" And Not IsNull(Me.Child794!SentComplaintToCourt) Then
     
            If Me.FairDebtDispute >= Me.Child794!SentComplaintToCourt Then
                Me.Child794.Enabled = True
                Me.Child794.Locked = False

                        For Each ctl In Me.Child794.Controls
                            Select Case ctl.ControlType
                             Case acTextBox
                             
                              If ctl.Name = "ComplaintFiled" Then
                                Me.Child794!ComplaintFiled.Locked = False
                                ElseIf ctl.Name = "LisPendensFiled" Then
                                Me.Child794!LisPendensFiled.Locked = False
                              
                              Else
                                 If Not (ctl.Locked) Then ctl.Locked = True
                        
                              End If
                            End Select
                        Next
            End If
       End If
       
       
    If Me!State = "DC" And Not IsNull(Me.Child794!ServiceSent) Then
     
            If Me.FairDebtDispute >= Me.Child794!ServiceSent Then
                Me.Child794.Enabled = True
                Me.Child794.Locked = False

                        For Each ctl In Me.Child794.Controls
                            Select Case ctl.ControlType
                             Case acTextBox
                             
                              If ctl.Name = "AllBorrowerServed" Then
                                Me.Child794!AllBorrowerServed.Locked = False
                                                            
                              Else
                                 If Not (ctl.Locked) Then ctl.Locked = True
                        
                              End If
                            End Select
                        Next
            End If
       End If
            
            
 
                 
       
                
                
    
        EditDispute = True
        
        Me.DocBackMilAff.Locked = False
        Me.DocBackDOA.Locked = False
        Me.DocBackSOD.Locked = False
        Me.DocBackLossMitPrelim.Locked = False
        Me.DocBackLossMitFinal.Locked = False
        Me.DocBackLostNote.Locked = False
        Me.DocBackOrigNote.Locked = False
        Me.DocBackNoteOwnership.Locked = False
        Me.DocBackAff7105.Locked = False
        Me.FairDebtVerified.Locked = False
        Me.ReinstatementRequested.Locked = False
        Me.ReinstatementSent.Locked = False
        Me.PayoffRequested.Locked = False
        Me.PayoffSent.Locked = False
        Me.ForebearanceAgreementReceived.Locked = False
        Me.ForebearanceAgreementSend.Locked = False
        Me.Child794!cmdNew.Enabled = False
        
        
        
        
        If FairDebtColor = 0 Then FairDebtColor = DLookup("iValue", "DB", "Name='FairDebtColor'")
        Detail.BackColor = FairDebtColor
  
        cmdPrint.Enabled = True
        MsgBox "CAUTION! File is locked due to Fair Debt dispute.  See loss mitigation, managers or attorneys for direction.", vbExclamation
    Else
        Detail.BackColor = ReadOnlyColor
        
    End If
Else
    Me.AllowEdits = True
    cmdAddFC.Enabled = True
    cmdAudit.Enabled = True
    cmdPrint.Enabled = True
    
'    If Not CheckNameEdit() Then
'    sfrmNames.Form.AllowEdits = False
'    sfrmNames.Form.AllowAdditions = False
'    sfrmNames.Form.AllowDeletions = False
'    sfrmNames!cmdCopyClient.Enabled = False
'    sfrmNames!cmdCopy.Enabled = False
'    sfrmNames!cmdTenant.Enabled = False
'    sfrmNames!cmdDelete.Enabled = False
'    sfrmNames!cmdNoNotice.Enabled = False
'    Else
'    sfrmNames.Form.AllowEdits = True
'    sfrmNames.Form.AllowAdditions = True
'    sfrmNames.Form.AllowDeletions = True
'    sfrmNames!cmdCopyClient.Enabled = True
'    sfrmNames!cmdCopy.Enabled = True
'    sfrmNames!cmdTenant.Enabled = True
'    sfrmNames!cmdDelete.Enabled = True
'    sfrmNames!cmdNoNotice.Enabled = True
'    End If
    cmdRemoveTrustee.Enabled = True
    cmdCalcPerDiem.Enabled = True
   ' cmdPurchaserInvestor.Enabled = True
    sfrmStatus.Form.AllowEdits = True
    sfrmStatus.Form.AllowAdditions = True
    sfrmStatus.Form.AllowDeletions = True
    Detail.BackColor = -2147483633
     
 If Not BHproject Then
 
    If (IsNull(Disposition) And PrivSetDisposition) Then
      cmdSetDisposition.Enabled = True
      sfrmFCDIL!DILSentRecord.Enabled = True
    End If
    
    If (IsNull(LMDisposition) And PrivSetDisposition) Then
      cmdSetLMDisposition.Enabled = True
      End If
End If

End If

'Remove per Diane 10/11
'If IsNull(TitleReviewToClient) Then
'cmdEditLegal.Enabled = False
'End If

If IsDate(Me.ClientPaid) = True Then
    Me.sfrmDisbursingSurplus.Visible = True  'The checkbox
 Else
    Me.sfrmDisbursingSurplus.Visible = False
End If


If Me.sfrmDisbursingSurplus!DisbursingSurplus = True Then
    Me.sfrmDisbursingSurplusTable.Visible = True
    Me.btnNewDSurplus.Visible = True
Else
    Me.btnNewDSurplus.Visible = False
    Me.sfrmDisbursingSurplusTable.Visible = False
End If




If Me.NewRecord Then    ' fill in info from previous FC, if any
    Set FC = CurrentDb.OpenRecordset("SELECT * FROM FCDetails WHERE FileNumber = " & FileNumber & " AND Current = True", dbOpenDynaset, dbSeeChanges)
    If Not FC.EOF Then
        NewFC = Date
        Referral = Date
        PrimaryFirstName = FC("PrimaryFirstName")
        PrimaryLastName = FC("PrimaryLastName")
        SecondaryFirstName = FC("SecondaryFirstName")
        SecondaryLastName = FC("SecondaryLastName")
        PropertyAddress = FC("PropertyAddress")
        City = FC("City")
        State = FC("State")
        ZipCode = FC("ZipCode")
        TaxID = FC("TaxID")
        optLeasehold = FC("Leasehold")
        GroundRentAmount = FC("GroundRentAmount")
        GroundRentPayable = FC("GroundRentPayable")
        LegalDescription = FC("LegalDescription")
        Comment = FC("Comment")
        DOT = FC("DOT")
        DOTdate = FC("DOTdate")
        OriginalTrustee = FC("OriginalTrustee")
        OriginalBeneficiary = FC("OriginalBeneficiary")
        Liber = FC("Liber")
        Folio = FC("Folio")
        OriginalMortgagors = FC("OriginalMortgagors")
        OriginalPBal = FC("OriginalPBal")
        RemainingPBal = FC("RemainingPBal")
        LoanNumber = FC("LoanNumber")
        LoanType = FC("LoanType")
        LienPosition = FC("LienPosition")
        FHALoanNumber = FC("FHALoanNumber")
        FNMALoanNumber = FC("FNMALoanNumber")
        AbstractorCaseNumber = FC("AbstractorCaseNumber")
        CourtCaseNumber = FC("CourtCaseNumber")
        TitleThru = FC("TitleThru")
        If State <> "DC" Then Docket = FC("Docket")
       
'        TitleReviewNameOf = FC("TitleReviewNameOf")
'        TitleReviewLiens = FC("TitleReviewLiens")
'        TitleReviewJudgments = FC("TitleReviewJudgments")
'        TitleReviewTaxes = FC("TitleReviewTaxes")
'        TitleReviewStatus = FC("TitleReviewStatus")
        TitleClaim = FC("TitleClaim")
        TitleClaimSent = FC("TitleClaimSent")
        TitleClaimResolved = FC("TitleClaimResolved")
        DeedAppReceived = FC!DeedAppReceived
        DeedAppDate = FC!DeedAppDate
        DeedAppSentToRecord = FC!DeedAppSentToRecord
        DeedAppRecorded = FC!DeedAppRecorded
        DeedAppLiber = FC!DeedAppLiber
        DeedAppFolio = FC!DeedAppFolio
        SentToDocket = FC!SentToDocket
        Docket = FC!Docket
        ServiceSent = FC!ServiceSent
        BorrowerServed = FC!BorrowerServed
        IRSLiens = FC!IRSLiens
        NOI = FC!NOI
        DOTrecorded = FC!DOTrecorded
        LastPaymentDated = FC!LastPaymentDated
        AmountOwedNOI = FC!AmountOwedNOI
        DateOfDefault = FC!DateOfDefault
        SecuredParty = FC!SecuredParty
        SecuredPartyPhone = FC!SecuredPartyPhone
        TypeOfDefault = FC!TypeOfDefault
        OtherDefault = FC!OtherDefault
        MortgageLender = FC!MortgageLender
        MortgageLenderLicense = FC!MortgageLenderLicense
        MortgageOriginator = FC!MortgageOriginator
        MortgageOriginatorLicense = FC!MortgageOriginatorLicense
        
        'Current = True
        If MsgBox("Should the Fair Debt letter be re-sent?", vbYesNo + vbQuestion) = vbNo Then FairDebt = FC!FairDebt
        NewFC = Now()
        If StaffID = 0 Then Call GetLoginName
        NewFCBy = StaffID ' and make this record current
        
        If (State = "VA") Then
          AddInvoiceItem FileNumber, "FC-REF", "Attorney fee", 600, 0, True, False, False, False
        End If
        
        Do While Not FC.EOF     ' make all previously current records not current
            FC.Edit
            FC("Current") = False
            FC.Update
            FC.MoveNext
        Loop
    End If
    FC.Close
             
End If

'Current.Locked = Current.Value
'Me.Caption = IIf(CaseTypeID = 8, "Monitor ", "") & "Foreclosure File " & Me![FileNumber] & " " & [PrimaryDefName]
Me.Caption = IIf(CaseTypeID = 8 Or (CaseTypeID = 1 And Forms![Case List]!cbxDetails.Column(0) = 8), "Monitor ", "") & "Foreclosure File " & Me![FileNumber] & " " & [PrimaryDefName]

If Not BHproject Then

    Call Visuals
    Set FC = CurrentDb.OpenRecordset("select Investor from caselist where filenumber=" & FileNumber)
    If State = "MD" Then
        If IsNull(SecuredParty) And Len(FC!Investor) <= 255 Then SecuredParty = Forms![Case List]!Investor
    End If
    
    If Not IsNull(Disposition) And PrivRescindDisposition Then
        With DispositionRescinded
            .Enabled = True
            .Locked = False
            .BackStyle = 1
        End With
        
        With RescindClientReq
            .Enabled = True
            .Locked = False
        End With
        
    Else
        With DispositionRescinded
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With RescindClientReq
            .Enabled = False
            .Locked = True
        End With
        
        
    End If
    
    If IsNull(Disposition) Then
        lblDisposition.Visible = False
    '    With SalePrice
    '        .Enabled = True
    '        .Locked = False
    '        .BackStyle = 1
    '    End With
        If Not IsNull(Notices) Then
            With Notices
                .Enabled = False
                .Locked = True
                .BackStyle = 0
            End With
            If TimesUpdatedNotice <> 0 Then
            
            With UpdatedNotices
                .Enabled = True
                .Locked = False
                .BackStyle = 1
            End With
        End If
        
        Else
            With Notices
                .Enabled = True
                .Locked = False
            End With
            
            With UpdatedNotices
                .Enabled = False
                .Locked = True
                .BackStyle = 0
            End With
        
        End If
        
    Else
        Disposition.Locked = True
        
    '    With SalePrice
    '        .Enabled = False
    '        .Locked = True
    '        .BackStyle = 0
    '    End With
        With FairDebt
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With AccelerationLetter
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With NOI
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With HUDOccLetter
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With VAAppraisal
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With [FirstLegal] ' 2012.03.08
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With StatementOfDebtDate
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With StatementOfDebtAmount
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With StatementOfDebtPerDiem
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With BondNumber
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With BondAmount
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With SentToDocket
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        If IsNull(SentToDocket) Then
            With Docket
                .Enabled = False
                .Locked = True
                .BackStyle = 0
            End With
        End If
        'Sarab and Durreth  on 7/30/2015 for ticket 1738 dispute
        
        
'        With ServiceSent
'            .Enabled = False
'            .Locked = True
'            .BackStyle = 0
'        End With
    
    
        If IsNull(ServiceSent) Then
            With BorrowerServed
                .Enabled = False
                .Locked = True
                .BackStyle = 0
            End With
        End If
    
        With IRSNotice
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With Notices
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With BidReceived
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With BidAmount
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With PayoffAmount
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With Sale
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        With SaleTime
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With
        
        If (PrivAdjustDeposit = False Or (PrivAdjustDeposit = True And Me.Disposition <> 2)) Then
          With Deposit
            .Enabled = False
            .Locked = True
            .BackStyle = 0
          End With
        End If
        If Nz(SaleCompleted) = 0 Then
            With DocstoClient
                .Enabled = False
                .Locked = True
                .BackStyle = 0
            End With
            With TitleOrder
                .Enabled = False
                .Locked = True
                .BackStyle = 0
            End With
            'With IRSLiens
            '    .Enabled = False
            '    .Locked = True
            'End With
        End If
        
    End If
    
    'If SaleCalendarEntryID = "X" Then
    '    lblSharedCal1.Caption = "Shared Calendar must be updated manually"
    '    lblSharedCal2.Caption = "Shared Calendar must be updated manually"
    '    lblSharedCal1.ForeColor = vbRed
    '    lblSharedCal2.ForeColor = vbRed
    'Else
    '    lblSharedCal1.Caption = "Shared Calendar updates are automatic"
    '    lblSharedCal2.Caption = "Shared Calendar updates are automatic"
    '    lblSharedCal1.ForeColor = 10040115
    '    lblSharedCal2.ForeColor = 10040115
    'End If


    'Converted to rst 10/10 to accomodate restart wizard
    Dim rstCaseList As Recordset
    Set rstCaseList = CurrentDb.OpenRecordset("Select * from caselist where filenumber=" & FileNumber)
    If (rstCaseList![CaseTypeID] = 1 Or rstCaseList![CaseTypeID] = 7) Then ' foreclosure or eviction
      Call SetObjectAttributes(State, False) ' cannot edit
    Else
      Call SetObjectAttributes(State, True)
    End If
    Set rstCaseList = Nothing
    
    'If IsNull(TitleReviewToClient) Then 'As per Diane request on 03/04/2014
    '  LegalDescription.Locked = False
    '  Else
    '  LegalDescription.Locked = True
    'End If
    
    If State = "VA" Then
    chkFullLegal.Enabled = True
    End If
      
    If IsNull(LoanNumber) Then
      LoanNumber.Locked = False
      LoanNumber.BackStyle = 1
      Call SetObjectAttributes(LoanNumber, True)
    Else  ' this allows for copying
      '----- temporary fix so that loan numbers can be changed, requested by Angela 9/29/2011
      LoanNumber.Locked = True
      LoanNumber.BackStyle = 1
      '------------
      
      'Call SetObjectAttributes(LoanNumber, True)
    End If
      
    'Sale Cert
    If LoanType = 5 Or LoanType = 4 Or Forms![Case List]!ClientID = 97 Then
    SaleCert.Enabled = True
    End If
      
      
    If IsNull(FNMALoanNumber) Then
      FNMALoanNumber.Locked = False
      FNMALoanNumber.BackStyle = 1
      
    Else  ' this allows for copying
      FNMALoanNumber.Locked = True
      FNMALoanNumber.BackStyle = 0
    End If
    
    If IsNull(FHLMCLoanNumber) Then
      FHLMCLoanNumber.Locked = False
      FHLMCLoanNumber.BackStyle = 1
      
    Else  ' this allows for copying
      FHLMCLoanNumber.Locked = True
      FHLMCLoanNumber.BackStyle = 0
    End If
      
      
    If PrivNotices = False Then
          Me.Notices.Enabled = False
        End If
        If PrivNotices = False Then
          Me.UpdatedNotices.Enabled = False
          End If
    

End If 'BHproject

txtEvictionBroker = Null
txtEvictionFileNum = Null
EVFileNumbers = ""
Set rstEV = CurrentDb.OpenRecordset("SELECT FileNumber,Brokers.BrokerName,Brokers.BrokerPhone,Brokers.BrokerEMail FROM EVDetails LEFT JOIN Brokers ON EVDetails.BrokerID = Brokers.BrokerID WHERE FCFileNumber=" & FileNumber, dbOpenSnapshot)
Do While Not rstEV.EOF
    EVFileNumbers = EVFileNumbers & rstEV!FileNumber & ", "
    txtEvictionBroker = rstEV!BrokerName
    txtBrokerPhone = rstEV!BrokerPhone
    lblBrokerEMail.Caption = Nz(rstEV!BrokerEMail)
    If Not IsNull(rstEV!BrokerEMail) Then lblBrokerEMail.HyperlinkAddress = "mailto:" & rstEV!BrokerEMail & "?Subject=" & PropertyAddress
    rstEV.MoveNext
Loop
rstEV.Close
If EVFileNumbers <> "" Then txtEvictionFileNum = Left$(EVFileNumbers, Len(EVFileNumbers) - 2)

If Me.cbxSustained.Value = 4 Then
    ExceptionsHearing.Enabled = False
    ExceptionsHearingTime.Enabled = False
    AddNewDateException.Enabled = True
End If
    
If Me.State = "VA" Then
    Me.optLeasehold = 0
    Me.Option131.Enabled = False
    GroundRentAmount.Enabled = False ' changes from optLeasehold becausse of DC project ticket no866 SA 06/03
    GroundRentPayable.Enabled = False ' changes from optLeasehold becausse of DC project ticket no866 SA 06/03
    Me.Deposit.Locked = True
End If

If Me.State = "DC" Then
Call VisibleDCForeclosureDetailsForm
End If

If Me.LienPosition = 1 Then
    Forms!foreclosuredetails!sfrmFCtitle!ckSenior = True
ElseIf Me.LienPosition = 2 Then
    Forms!foreclosuredetails!sfrmFCtitle!ckJunior = True
ElseIf Me.LienPosition = 3 Then
    Forms!foreclosuredetails!sfrmFCtitle!ck3 = True
ElseIf IsNull(Me.LienPosition) Or Me.LienPosition = "" Then
    Forms!foreclosuredetails!sfrmFCtitle!ckOther = True
Else
    Forms!foreclosuredetails!sfrmFCtitle!ckOther = True
End If

    
    Set rstCaseList = CurrentDb.OpenRecordset("Select * from caselist where filenumber=" & FileNumber)
        If (rstCaseList![CaseTypeID]) = 1 Then  'Only foreclosure
'            If IsNull(Me.DispositionDesc) And Not IsNull(Sale) And (Date <= Sale Or Format(Date, "mm/dd/yyyy") = Format(Sale, "mm/dd/yyyy")) And Not IsNull(Me.BidReceived) And Not IsNull(BidAmount) Then
'                Me.SalePrice.Enabled = True
'                Me.Purchaser.Enabled = True
'                Me.PurchaserAddress.Enabled = True
'                Me.cmdPurchaserInvestor.Enabled = True
'                 Else
            If IsNull(Me.DispositionDesc) And Not IsNull(Sale) And Not IsNull(Me.BidReceived) And Not IsNull(BidAmount) Then
                Me.SalePrice.Locked = False
                Me.Purchaser.Locked = False
                Me.PurchaserAddress.Locked = False
                Me.cmdPurchaserInvestor.Enabled = True
'                End If
                
            End If
            
            If Not IsNull(SalePrice) And Not IsNull(Purchaser) And Not IsNull(PurchaserAddress) And SalePrice.Enabled = False Then
            ComEdit.Enabled = True
            Else
                If (Me.DispositionDesc = "Buy-In" Or Me.DispositionDesc = "3rd Party") And Not IsNull(Sale) And Not IsNull(Me.BidReceived) And Not IsNull(BidAmount) Then
                  ComEdit.Enabled = True
                End If
            End If
            
     End If
    Set rstCaseList = Nothing
    
If MonitorChoose Then
Call MonitorVisiable
MonitorChoose = False
End If


If BHproject Then
Me.pgNOI.Enabled = True
Me.NOI.Enabled = True
Me.NOI.Locked = False
End If
   
    If Not PrivPrintPostSale Then
    DismissalSent.Locked = True
    DismissalSent.Enabled = False
    End If
    
    
    If Not PrivPrintPostSale Then
    DismissalDate.Locked = True
    DismissalDate.Enabled = False
    End If
    
 'added on 7/9/15
 
If DCTabView = False Then
'dc tab locked
    Me.txtFistPub.Enabled = False
    Me.txtSale.Enabled = False
    Me.txtSaleTime.Enabled = False
    Me.txtDeposit.Enabled = False
    Me.txtSaleSet.Enabled = False
    Me.txtreviewadproof.Enabled = False
    Me.txtNewAdVendor.Enabled = False

End If
    
   If (Not IsNull(FairDebtDispute) And IsNull(FairDebtVerified)) Then
   Me.ServiceSent.Locked = True
        If Not IsNull(Me.ServiceSent) Then
        Me.BorrowerServed.Locked = False
        Else
        Me.BorrowerServed.Locked = True
        End If
     
   End If
   
'added on 9/3/15
If Not IsNull(txtSale) And Not IsNull(txtSaleTime) Then
    txtSale.Locked = True
    txtSaleTime.Locked = True
End If


 
    
    
End Sub

Private Sub MonitorVisiable()

    Call SetObjectAttributes(FairDebt, False)
    Me.NewFairDebt.Enabled = False
    Call SetObjectAttributes(AccelerationIssued, False)
    Me.NewDemand.Enabled = False
    Call SetObjectAttributes(AccelerationLetter, False)
    Call SetObjectAttributes(txtClientSentAcceleration, False)
    Call SetObjectAttributes(txtNOIExpires, False)
    'txtNOIExpires
    Me.New45Notice.Enabled = False
    Call SetObjectAttributes(NOI, False)
    Call SetObjectAttributes(txtClientSentNOI, False)
    Call SetObjectAttributes(DocstoClient, False)
    Call SetObjectAttributes(LossMitSolicitationDate, False)
    Call SetObjectAttributes(VAAppraisal, False)
    Call SetObjectAttributes(FirstLegal, False)
    Call SetObjectAttributes(DocsBack, False)
'    Call SetObjectAttributes(DocBackMilAff, False)
    Me.DocBackMilAff.Enabled = False
'    Call SetObjectAttributes(DocBackDOA, False)
    Me.DocBackDOA.Enabled = False
'    Call SetObjectAttributes(DocBackLossMitFinal, False)
    Me.DocBackLossMitFinal.Enabled = False
'    Call SetObjectAttributes(DocBackLostNote, False)
    Me.DocBackLostNote.Enabled = False
'    Call SetObjectAttributes(DocBackNoteOwnership, False)
    Me.DocBackNoteOwnership.Enabled = False
'    Call SetObjectAttributes(DocBackAff7105, False)
    Me.DocBackAff7105.Enabled = False
'    Call SetObjectAttributes(DocBackLossMitFinal, False)
    Me.DocBackLossMitFinal.Enabled = False
    Call SetObjectAttributes(SentToDocket, False)
    Call SetObjectAttributes(Docket, False)
    Call SetObjectAttributes(LienCert, False)
    Call SetObjectAttributes(FLMASenttoCourt, False)
    Call SetObjectAttributes(LossMitFinalDate, False)
    Call SetObjectAttributes(ServiceSent, False)
    Call SetObjectAttributes(ServiceMailed, False)
    Call SetObjectAttributes(FirstPub, False)
    Call SetObjectAttributes(IRSNotice, False)
    Call SetObjectAttributes(Notices, False)
    Call SetObjectAttributes(UpdatedNotices, False)
    Call SetObjectAttributes(TimesUpdatedNotice, False)
    Call SetObjectAttributes(SaleSet, False)
    Call SetObjectAttributes(BondNumber, False)
    Call SetObjectAttributes(BondAmount, False)
    Call SetObjectAttributes(BondReturned, False)
    Call SetObjectAttributes(Sale, True)
    Call SetObjectAttributes(SaleTime, True)
   
    Me.chMannerofService.Enabled = False
    Me.DocBackLossMitPrelim.Enabled = False
    Call SetObjectAttributes(ReviewAdProof, False)
    Call SetObjectAttributes(NewspaperVendorName, False)
    Call SetObjectAttributes(ReviewAdProof, False)
    Call SetObjectAttributes(SaleCert, False)
  
    Call SetObjectAttributes(BorrowerServed, False)
    Call SetObjectAttributes(StatementOfDebtDate, True)
    Call SetObjectAttributes(StatementOfDebtAmount, True)
    Call SetObjectAttributes(TitleOrder, True)
    Call SetObjectAttributes(TitleDue, True)
    Call SetObjectAttributes(StatementOfDebtPerDiem, True)
    Call SetObjectAttributes(LostNoteAffSent, False)
    Call SetObjectAttributes(LostNoteNotice, False)
    'Post sale
    Call SetObjectAttributes(Report, False)
    Call SetObjectAttributes(NiSiEnd, False)
    Call SetObjectAttributes(FinalPkg, False)
    Call SetObjectAttributes(AuditFile, False)
    Call SetObjectAttributes(AuditRat, True)
    Call SetObjectAttributes(SalePrice, False)
    Call SetObjectAttributes(Purchaser, False)
    Call SetObjectAttributes(PurchaserAddress, False)
    Me.cmdPurchaserInvestor.Enabled = False
    Me.ExceptionsFiled.Enabled = False
    Call SetObjectAttributes(ExceptionsHearing, False)
    Call SetObjectAttributes(ExceptionsHearingTime, False)
    Me.cbxSustained.Enabled = False
    Me.Resell.Enabled = False
    Call SetObjectAttributes(ResellMotion, False)
    Call SetObjectAttributes(ResellServed, False)
    Call SetObjectAttributes(ResellShowCauseExpires, False)
    Call SetObjectAttributes(ResellAnswered, False)
    Call SetObjectAttributes(ResellGranted, False)
    Me.AddNewDateException.Enabled = False
    Me.AddNewDate.Enabled = False
    Call SetObjectAttributes(StatusHearing, False)
    Call SetObjectAttributes(StatusHearingTime, False)
    Me.StatusResults.Enabled = False
   ' Me.cmdSetDisposition.Enabled = False
    Call SetObjectAttributes(Settled, False)
    Call SetObjectAttributes(ClientPaid, False)
    Me.chEviction.Enabled = False
    Me.REO.Enabled = False
    Call SetObjectAttributes(RecordDeed, False)
    Call SetObjectAttributes(RecordDeedLiber, False)
    Call SetObjectAttributes(RecordDeedFolio, False)
    Call SetObjectAttributes(Audit2File, True)
    Call SetObjectAttributes(Audit2Rat, True)
    Me.AmmDocBackSOD.Enabled = True
    Call SetObjectAttributes(AmmStatementOfDebtDate, True)
    Call SetObjectAttributes(AmmStatementOfDebtAmount, True)
    Call SetObjectAttributes(AmmStatementOfDebtPerDiem, True)
    
    
 '  Forms!foreclosureDetails!Page96.Visible = False
'    Forms!foreclosureDetails!Trustees.Visible = False
    Forms!foreclosuredetails!Page412.Visible = False
    Forms!foreclosuredetails!pgNOI.Visible = False
  '  Forms!foreclosureDetails!Page256.Visible = False
    Forms!foreclosuredetails!pgMediation.Visible = False
 '   Forms!foreclosureDetails![Pre-Sale].Visible = False
'    Forms!foreclosureDetails![Post-Sale].Visible = False
    Forms!foreclosuredetails!pgRealPropTaxes.Visible = False
    Forms!foreclosuredetails!pageStatus.Visible = True
    Me.[Trustees:_Label].Visible = True
    Me.lstTrustees.Visible = True
    Me.[Add Trustee:_Label].Visible = True
    Me.cbxAddTrustee.Visible = True
    Me.cmdAddDefaultTrustees.Visible = True
    Me.cmdRemoveTrustee.Visible = True
    
    
    
    
    
    
 
    
End Sub


Private Sub cmdClose_Click()

Dim rs As Recordset
Dim skip As Boolean
                
Set rs = CurrentDb.OpenRecordset("SELECT * FROM LMHearings_DC WHERE FileNumber =" & Me!FileNumber, dbOpenDynaset, dbSeeChanges)
skip = False
                
Do Until rs.EOF
    If Not IsNull(rs!CondactedTypeID) And rs!CondactedTypeID <> 6 And IsNull(rs!HearingAttyID) Then
                
        skip = True
        MsgBox ("Who conducted the mediation is missing")
        Exit Sub
    End If
                    
rs.MoveNext
Loop
                
rs.Close
Set rs = Nothing

'added on 9/3/15

If Not IsNull(txtSale) And Not IsNull(txtSaleTime) Then
    txtSale.Locked = True
    txtSaleTime.Locked = True
End If

If Me.State = "DC" Then
            
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCDIL WHERE FileNumber =" & Me!FileNumber, dbOpenDynaset, dbSeeChanges)
        If Not IsNull(rs!InitialHearingConference) And IsNull(rs!DCHearingTime) Then
        MsgBox ("DC Initia Hearing time is missing")
        Exit Sub
        End If
    rs.Close
    Set rs = Nothing
End If


On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdPrint_Click()

On Error GoTo Err_cmdPrint_Click

If Disposition > 2 And Disposition <> 7 And Disposition <> 11 And (Date - DispositionDate) > 14 And Not PrivPrintPostSale Then
DoCmd.OpenForm "foreclosureprintdisposition", , , "[CaseList].[FileNumber]=" & Me![FileNumber]
Exit Sub
End If

If (Not IsNull(FairDebtDispute) And IsNull(FairDebtVerified)) Then
    MsgBox "File is locked due to Fair Debt dispute.  Only payoff, reinstatements and accelerations can be printed.", vbCritical
    Exit Sub
End If

If CaseTypeID = 8 Or Forms![Case List]!cbxDetails = 8 Then
    DoCmd.OpenForm "MonitorPrint", , , "[Caselist].[Filenumber]=" & Me![FileNumber]
    Exit Sub
End If

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
If IsNull(WizardSource) Then WizardSource = "None"
DoCmd.OpenForm "ForeclosurePrint", , , "[CaseList].[FileNumber]=" & Me![FileNumber], , , WizardSource


Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    
    Resume Exit_cmdPrint_Click
    
End Sub

Private Sub cmdRemoveTrustee_Click()
Dim t As Recordset

On Error GoTo Err_cmdRemoveTrustee_Click

If IsNull(Me!lstTrustees) Then
    MsgBox "Select a Trustee to remove", vbCritical
    Exit Sub
End If

Set t = CurrentDb.OpenRecordset("SELECT * FROM Trustees WHERE ID=" & Me!lstTrustees, dbOpenDynaset, dbSeeChanges)
If Not t.EOF Then t.Delete
t.Close
Me!lstTrustees.Requery
TrusteeWordFile = 0         ' invalidate cache

Exit_cmdRemoveTrustee_Click:
    Exit Sub

Err_cmdRemoveTrustee_Click:
    MsgBox Err.Description
    Resume Exit_cmdRemoveTrustee_Click
    
End Sub

Private Sub cmdPurchaserInvestor_Click()

On Error GoTo Err_cmdPurchaserInvestor_Click
Me!Purchaser = Investor
Me!PurchaserAddress = InvestorAddress
AddStatus FileNumber, Sale, "Property sold to " & Purchaser & " for " & Format$(SalePrice, "Currency")

Exit_cmdPurchaserInvestor_Click:
    Exit Sub

Err_cmdPurchaserInvestor_Click:
    MsgBox Err.Description
    Resume Exit_cmdPurchaserInvestor_Click
    
End Sub

Private Sub HUDOccLetter_AfterUpdate()
AddStatus FileNumber, HUDOccLetter, "Sent HUD Occupancy Letter"
End Sub

Private Sub HUDOccLetter_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(HUDOccLetter)
End Sub

Private Sub HUDOccLetter_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    HUDOccLetter = Now()
    AddStatus FileNumber, HUDOccLetter, "Sent HUD Occupancy Letter"
End If
End Sub

Private Sub InterestRate_AfterUpdate()
'If InterestRate > 1 Then InterestRate = InterestRate / 100#
End Sub

Private Sub IRSLiens_AfterUpdate()
If Me.State = "VA" Then
If Not IsNull(Sale) Then
MsgBox ("Notify sale setting and Notice managers to confirm sale can proceed")
Dim rstJnl As Recordset
Set rstJnl = CurrentDb.OpenRecordset("select * from journal", dbOpenDynaset, dbSeeChanges)
With rstJnl
.AddNew
!FileNumber = FileNumber
!JournalDate = Now
!Who = GetFullName
!Info = "Notify sale setting and Notice managers to confirm sale can proceed"
!Color = 1
.Update
End With
Set rstJnl = Nothing
End If
End If

Call Visuals
End Sub

Private Sub IRSNotice_AfterUpdate()
If BHproject Then
AddStatus FileNumber, IRSNotice, "IRS Notice sent"
Else

AddStatus FileNumber, IRSNotice, "IRS Notice sent"
AddInvoiceItem FileNumber, "FC-IRS", "IRS Notice Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 8)), 76, False, True, False, True
End If

End Sub

Private Sub IRSNotice_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
  Cancel = CheckFutureDate(IRSNotice)
End If

End Sub

Private Sub IRSNotice_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    IRSNotice = Now()
    Call IRSNotice_AfterUpdate
End If

End Sub

Private Sub LastPaymentApplied_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(LastPaymentApplied)
End Sub

Private Sub LoanType_AfterUpdate()
If Not BHproject Then
Call Visuals
End If

End Sub

Private Sub LossMitFinalDate_AfterUpdate()
AddStatus FileNumber, LossMitFinalDate, "Final Loss Mitigation Affidavit filed."
End Sub

Private Sub LossMitFinalDate_BeforeUpdate(Cancel As Integer)
If BHproject Then
Else

'
Dim retval As Boolean
'
'
Cancel = CheckFutureDate(LossMitFinalDate)
If (Cancel = 1) Then Exit Sub

'If (IsNull(LossMitPrelimDate)) Then
'  retval = False
'ElseIf (DateDiff("d", LossMitPrelimDate, LossMitFinalDate) < 18) Then
'  retval = False
'Else
'  retval = True
'End If
'
'If (retval = False) Then
'    Cancel = 1
'    MsgBox "Final LMA Date must be 18 days after Preliminary LMA date.", vbCritical
'End If
End If

End Sub
'
'
'Private Sub LossMitPrelimDate_AfterUpdate()
'AddStatus FileNumber, LossMitPrelimDate, "Preliminary Loss Mitigation Affidavit filed."
'End Sub
'
'
'Private Sub LossMitPrelimDate_BeforeUpdate(Cancel As Integer)
'  Cancel = CheckFutureDate(LossMitPrelimDate)
'
'End Sub
'
'Private Sub LossMitPrelimDate_DblClick(Cancel As Integer)
'LossMitPrelimDate = Date
'Call LossMitPrelimDate_AfterUpdate


'End Sub

Private Sub LPIDate_AfterUpdate()
AddStatus FileNumber, LPIDate, "LPI Date"
If LoanType = 3 Then FirstLegal = (LPIDate + 180)
Call FirstLegal_AfterUpdate
End Sub

Private Sub LPIDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    LPIDate = Date
    Call LPIDate_AfterUpdate
End If
End Sub

Private Sub MedDocSentDate_AfterUpdate()
AddStatus FileNumber, MedDocSentDate, "Docs Sent to Borrower and OAH"
End Sub

Private Sub MedDocSentDate_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(MedDocSentDate)
End Sub

Private Sub MedDocSentDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    MedDocSentDate = Date
    Call MedDocSentDate_AfterUpdate
End If

End Sub

Private Sub MedRecDocDate_AfterUpdate()
AddStatus FileNumber, MedRecDocDate, "Received Mediation Docs from Client"
End Sub

Private Sub MedRecDocDate_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(MedRecDocDate)
End Sub

Private Sub MedRecDocDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    MedRecDocDate = Date
    Call MedRecDocDate_AfterUpdate
End If

End Sub

Private Sub MedReqDocDate_AfterUpdate()
AddStatus FileNumber, MedReqDocDate, "Request Mediation Docs from Client"
End Sub

Private Sub MedReqDocDate_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(MedReqDocDate)
End Sub

Private Sub MedReqDocDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    MedReqDocDate = Date
    Call MedReqDocDate_AfterUpdate
End If

End Sub

Private Sub MedRequestDate_AfterUpdate()
If BHproject Then
If Not IsNull(MedRequestDate) Then
AddStatus FileNumber, MedRequestDate, "Mediation Date Requested"
End If
Else

If Not IsNull(MedRequestDate) Then
AddStatus FileNumber, MedRequestDate, "Mediation Date Requested"
 AddInvoiceItem FileNumber, "FC-MED", "Mediation Fee", DLookup("mediationfee", "clientlist", "clientid=" & Forms![Case List]!ClientID), 0, True, True, False, True
 End If
End If
End Sub

Private Sub MedRequestDate_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(MedRequestDate)
End Sub

Private Sub MedRequestDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    MedRequestDate = Date
    Call MedRequestDate_AfterUpdate
End If

End Sub

Private Sub MonitorClientPaid_AfterUpdate()
AddStatus FileNumber, MonitorClientPaid, "Payment sent"
End Sub

Private Sub MonitorClientPaid_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(MonitorClientPaid)

End Sub

Private Sub MonitorClientPaid_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    MonitorClientPaid = Now()
    AddStatus FileNumber, MonitorClientPaid, "Payment sent"
End If

End Sub

Private Sub MonitorMotionSurplusFiled_AfterUpdate()
AddStatus FileNumber, MonitorMotionSurplusFiled, "Filed Motion To Intervene of Right and Petition for Claim to Surplus Proceeds"
End Sub

Private Sub MonitorMotionSurplusFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(MonitorMotionSurplusFiled)

End Sub

Private Sub MonitorMotionSurplusFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    MonitorMotionSurplusFiled = Now()
    AddStatus FileNumber, MonitorMotionSurplusFiled, "Filed Motion To Intervene of Right and Petition for Claim to Surplus Proceeds"
End If

End Sub

Private Sub MonitorOrderSurplus_AfterUpdate()
AddStatus FileNumber, MonitorOrderSurplus, "Order entered to Intervene of Right and Petition for Claim to Surplus Proceeds"
End Sub

Private Sub MonitorOrderSurplus_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(MonitorOrderSurplus)

End Sub

Private Sub MonitorOrderSurplus_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    MonitorOrderSurplus = Now()
    AddStatus FileNumber, Now(), "Order entered to Intervene of Right and Petition for Claim to Surplus Proceeds"
End If

End Sub

Private Sub NiSiEnd_AfterUpdate()
AddStatus FileNumber, Now(), "Order NiSi entered, expires " & Format$(NiSiEnd, "m/d/yyyy")
End Sub


Private Sub NiSiEnd_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    NiSiEnd = Now()
    AddStatus FileNumber, Now(), "Order NiSi entered, expires " & Format$(NiSiEnd, "m/d/yyyy")
End If

End Sub

Private Sub NOI_AfterUpdate()
AddStatus FileNumber, NOI, "45 Day Notice sent"
End Sub

Private Sub NOI_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(NOI)
End Sub

Private Sub NOI_DblClick(Cancel As Integer)

If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    If BHproject Then
    NOI = Date
    AddStatus FileNumber, Date, "45 Day Notice sent"
    End If
    
End If


End Sub

Private Sub Notices_AfterUpdate()
If BHproject Then
AddStatus FileNumber, Notices, "Mailed Notice of Foreclosure Sale"
Else

AddStatus FileNumber, Notices, "Mailed Notice of Foreclosure Sale"
Dim noticecnt As Integer
noticecnt = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and nz(NoticeType,0) > 0")
If Me.State = "MD" Then noticecnt = noticecnt + 1
If (noticecnt > 0) Then

   AddInvoiceItem FileNumber, "FC-NOT", "Sale Notices - Certified Postage", (Nz(DLookup("Value", "StandardCharges", "ID=" & 8))) * noticecnt, 76, False, False, False, True
   AddInvoiceItem FileNumber, "FC-NOT", "Sale Notices - First Class Postage", (Nz(DLookup("Value", "StandardCharges", "ID=" & 1))) * noticecnt, 76, False, False, False, True
    
End If
End If

End Sub

Private Sub Notices_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
  Cancel = CheckFutureDate(Notices)
End If

End Sub

Private Sub Notices_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Notices = Now()
    Call Notices_AfterUpdate
End If
End Sub

Private Sub optLeasehold_Click()
Call Visuals
End Sub

Private Sub cbxSustained_AfterUpdate()

Select Case cbxSustained
    Case 1      ' sustained
        Call SetDisposition(9)
        If Now() < ExceptionsHearing Then
        If Not IsNull(ExceptionsHearingEntryID) Then
        Call DeleteCalendarEvent(ExceptionsHearingEntryID)
         ExceptionsHearingEntryID = Null
         
        End If
        End If
        AddStatus FileNumber, Date, " Exception Hearing Sustained"
    Case 2      ' overrruled
        AddStatus FileNumber, Date, "Exceptions overruled, sale should ratify shortly"
        If Now() < ExceptionsHearing Then
        If Not IsNull(ExceptionsHearingEntryID) Then
        Call DeleteCalendarEvent(ExceptionsHearingEntryID)
         ExceptionsHearingEntryID = Null
        End If
        End If
    Case 3
         AddStatus FileNumber, Date, " Exception Hearing Whithdrawn "
        If Now() < ExceptionsHearing Then
        If Not IsNull(ExceptionsHearingEntryID) Then
        Call DeleteCalendarEvent(ExceptionsHearingEntryID)
         ExceptionsHearingEntryID = Null
        End If
        End If
    
    Case 4
    
    AddStatus FileNumber, Date, "Exception Hearing Continue "
    If Now() < ExceptionsHearing.OldValue Then
    If Not IsNull(ExceptionsHearingEntryID) Then
    Call DeleteCalendarEvent(ExceptionsHearingEntryID)
    ExceptionsHearingEntryID = Null
    
    ExceptionsHearing.Value = Null
    ExceptionsHearing.Enabled = False
    ExceptionsHearingTime.Value = Null
    ExceptionsHearingTime.Enabled = False
    AddNewDateException.Enabled = True
    
    Else
    ExceptionsHearing.Value = Null
    ExceptionsHearing.Enabled = False
    ExceptionsHearingTime.Value = Null
    ExceptionsHearingTime.Enabled = False
    AddNewDateException.Enabled = True
    End If
    Else
    ExceptionsHearing.Value = Null
    ExceptionsHearing.Enabled = False
    ExceptionsHearingTime.Value = Null
    ExceptionsHearingTime.Enabled = False
    AddNewDateException.Enabled = True
    End If
    
    
    
   End Select
        
        
        
        

End Sub

Private Sub OrderSubsPurch_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(OrderSubsPurch)
End Sub

Private Sub OrderSubsPurch_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    OrderSubsPurch = Now()
End If

End Sub

Private Sub OriginalBeneficiary_AfterUpdate()
If IsNull(MortgageLender) Then MortgageLender = OriginalBeneficiary
End Sub

Private Sub PayoffRequested_AfterUpdate()
AddStatus FileNumber, PayoffRequested, "Payoff Requested"
End Sub

Private Sub PayoffRequested_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(PayoffRequested)
End Sub

Private Sub PayoffRequested_DblClick(Cancel As Integer)
If FileReadOnly Then
   DoCmd.CancelEvent
Else
    PayoffRequested = Date
    Call PayoffRequested_AfterUpdate
End If

End Sub

Private Sub PayoffSent_AfterUpdate()

If BHproject Then
If Not IsNull(PayoffSent) Then
 AddStatus FileNumber, PayoffSent, "Payoff Sent"
 Exit Sub
 End If
 
Else

        If Not IsNull(PayoffSent) Then
        
            If IsNull(PayoffRequested) Then
                MsgBox ("Please Add Payoff Requested date")
                PayoffSent = Null
                
                Exit Sub
                Else
            
                AddStatus FileNumber, PayoffSent, "Payoff Sent"
            
                Dim strInsert As String
                Dim clientShor As String
                Dim StrJuirs As String
                StrJuirs = DLookup("Jurisdiction", "JurisdictionList", "JurisdictionID= " & Forms![Case List]!JurisdictionID)
                StrJuirs = Replace(StrJuirs, "'", "''")
                clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)
                clientShor = Replace(clientShor, "'", "''")
                DoCmd.SetWarnings False
                strInsert = "Insert Into Tracking_PayoffSent (CaseFile,ProjectName,ClientShortName,Juris,Client,DIT, PayoffRequested ,StaffID,StaffName) Values (" & FileNumber & ",'" & Forms![Case List]!PrimaryDefName & "','" & clientShor & "','" & StrJuirs & "', " & Forms![Case List]!ClientID & ", #" & Now() & "#,#" & Forms![foreclosuredetails]!PayoffRequested & "#, " & GetStaffID & ",'" & GetFullName() & "')"
                DoCmd.RunSQL strInsert
                DoCmd.SetWarnings True
            End If
        
        End If
End If

End Sub

Private Sub PayoffSent_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(PayoffSent)
End Sub

Private Sub PayoffSent_DblClick(Cancel As Integer)
If FileReadOnly Then
   DoCmd.CancelEvent
Else
    If IsNull(PayoffRequested) Then
    MsgBox ("Please Add Payoff Requested date")
    Exit Sub
    Else
    PayoffSent = Date
    Call PayoffSent_AfterUpdate
    End If
End If

End Sub

Private Sub Purchaser_AfterUpdate()
If Not BHproject Then
Dim strinfo As String
Dim strSQLJournal As String

If checkCmdEdit Then

    If Nz(Purchaser) <> Nz(Purchaser.OldValue) Then
    
      If (Nz(Purchaser, "")) = "" Then
        MsgBox ("You should put a Purchaser")
        Me.Undo
        
        Else
        DoCmd.SetWarnings False
        strinfo = " Edit Purchaser "
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!foreclosuredetails!FileNumber & ", #" & Now & "# , '" & GetFullName() & "', '" & strinfo & "'," & 1 & ")"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
       End If
    End If
    
Else
If (Nz(Purchaser, "")) = "" Then
        DoCmd.SetWarnings False
        strinfo = " Edit Sale price "
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!foreclosuredetails!FileNumber & ", #" & Now & "# , '" & GetFullName() & "', '" & strinfo & "'," & 1 & ")"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
End If

End If

AddStatus FileNumber, Sale, "Property sold to " & Purchaser & " for " & Format$(SalePrice, "Currency")
'If Not IsNull(SalePrice) And Not IsNull(Purchaser) And Not IsNull(PurchaserAddress) And PurchaserAddress.Enabled = False Then ComEdit.Enabled = True
End If

End Sub

Private Sub RecordDeed_AfterUpdate()
AddStatus FileNumber, RecordDeed, "Deed recorded"
End Sub

Private Sub RecordDeed_BeforeUpdate(Cancel As Integer)
If BHproject Then
Cancel = CheckFutureDate(RecordDeed)
End If

End Sub

Private Sub RecordDeed_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    RecordDeed = Now()
    AddStatus FileNumber, RecordDeed, "Deed recorded"
End If

End Sub





Private Sub ReinstatementRequested_AfterUpdate()
AddStatus FileNumber, ReinstatementRequested, "Reinstatement Requested"
End Sub


Private Sub ReinstatementRequested_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ReinstatementRequested)
End Sub

Private Sub ReinstatementRequested_DblClick(Cancel As Integer)
If FileReadOnly Then
   DoCmd.CancelEvent
Else
    ReinstatementRequested = Date
    Call ReinstatementRequested_AfterUpdate
End If

End Sub

Private Sub ReinstatementSent_AfterUpdate()
If BHproject Then
        If Not IsNull(ReinstatementSent) Then
          AddStatus FileNumber, ReinstatementSent, "Reinstatement Sent"
        Exit Sub
        End If

Else


    If Not IsNull(ReinstatementSent) Then
    
    
        If IsNull(ReinstatementRequested) Then
            MsgBox ("Please Add Reinstratment Requeested date")
            ReinstatementSent = Null
            Exit Sub
            Else
                AddStatus FileNumber, ReinstatementSent, "Reinstatement Sent"
                
        
                    Dim strInsert As String
                    Dim clientShor As String
                    Dim StrJuirs As String
                    
                    StrJuirs = DLookup("Jurisdiction", "JurisdictionList", "JurisdictionID= " & Forms![Case List]!JurisdictionID)
                    StrJuirs = Replace(StrJuirs, "'", "''")
                    clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)
                    clientShor = Replace(clientShor, "'", "''")
                    DoCmd.SetWarnings False
                    strInsert = "Insert Into Tracking_ReinstatementSent (CaseFile,ProjectName,ClientShortName,Juris,Client,ReinstRequested,DIT,StaffID,StaffName) Values (" & FileNumber & ",'" & Forms![Case List]!PrimaryDefName & "','" & clientShor & "','" & StrJuirs & "', " & Forms![Case List]!ClientID & ",#" & Forms![foreclosuredetails]!ReinstatementRequested & "#, #" & Now() & "#," & GetStaffID & ",'" & GetFullName() & "')"
                    DoCmd.RunSQL strInsert
                    DoCmd.SetWarnings True
          End If
          
    End If

End If


End Sub

Private Sub ReinstatementSent_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ReinstatementSent)
End Sub

Private Sub ReinstatementSent_DblClick(Cancel As Integer)
If FileReadOnly Then
   DoCmd.CancelEvent
Else
    If IsNull(ReinstatementRequested) Then
    MsgBox ("Please Add Reinstratment Requeested date")
    Exit Sub
    Else
    ReinstatementSent = Date
    Call ReinstatementSent_AfterUpdate
    End If
End If
End Sub


Private Sub REO_AfterUpdate()
Dim rstREO As Recordset

If REO Then
    Set rstREO = CurrentDb.OpenRecordset("SELECT * FROM REOdetails WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
    If rstREO.EOF Then
        If MsgBox("Really create a REO record for this file?", vbQuestion + vbYesNo) <> vbYes Then
            REO = False
        Else
            rstREO.AddNew
            rstREO!FileNumber = FileNumber
            rstREO!Referred = Date
            rstREO.Update
        End If
    End If
    rstREO.Close
End If
End Sub

Private Sub Report_AfterUpdate()
AddStatus FileNumber, Report, "Report of Sale filed"
'Dim FeeAmount As Currency, Update As Boolean, Jurisdiction As Long

'Order Lien Cert
'If State = "MD" Then
'If LoanType = 2 Or LoanType = 3 Then
'Jurisdiction = Forms![case list]!JurisdictionID
'If DLookup("LienCert", "JurisdictionList", "jurisdictionid=" & Jurisdiction) = True Then
'If Not IsNull(LienCert) Then
'Update = True
'End If
'If MsgBox("This file is in a jurisdiction that requires a Lien Cert.  Did you want to order one?", vbYesNo) = vbYes Then
'LienCert = Date
'Call PrintLienCerts(FileNumber, Jurisdiction, Update)
'End If
'End If
'End If
'End If
End Sub

Private Sub Report_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(Report)
End If

End Sub

Private Sub Report_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Report = Now()
    'Call Report_AfterUpdate
    
    AddStatus FileNumber, Now(), "Report of Sale filed"
End If
End Sub

Private Sub RescindClientReq_Click()
If (RescindClientReq = True) Then
  AddInvoiceItem FileNumber, "FC-DISPRESC", "Sale Rescind", 350, 0, True, True, False, False
End If
End Sub

Private Sub Resell_AfterUpdate()
Call Visuals
End Sub

Private Sub Resell_Click()
If Not Resell Then
    If MsgBox("Really ""undo"" Resell?  This will reset the motion, served, answered, and granted dates.  You should also check the status report.", vbQuestion + vbYesNo) <> vbYes Then
        Resell = 1
        Exit Sub
    End If
End If
Call Visuals

End Sub

Private Sub ResellAnswered_AfterUpdate()
AddStatus FileNumber, ResellAnswered, "Motion to Resell answered"
End Sub

Private Sub ResellAnswered_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(ResellAnswered)
End If

End Sub

Private Sub ResellAnswered_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ResellAnswered = Date
    AddStatus FileNumber, ResellAnswered, "Motion to Resell answered"
End If

End Sub

Private Sub ResellGranted_AfterUpdate()
AddStatus FileNumber, ResellGranted, "Motion to Resell granted"
End Sub

Private Sub ResellGranted_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(ResellGranted)
End If

End Sub

Private Sub ResellGranted_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ResellGranted = Date
    AddStatus FileNumber, ResellGranted, "Motion to Resell granted"
End If

End Sub

Private Sub ResellMotion_AfterUpdate()
AddStatus FileNumber, ResellMotion, "Motion to Resell"
End Sub

Private Sub ResellMotion_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(ResellMotion)
End If

End Sub

Private Sub ResellMotion_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ResellMotion = Date
    AddStatus FileNumber, ResellMotion, "Motion to Resell"
End If

End Sub

Private Sub ResellServed_AfterUpdate()
AddStatus FileNumber, ResellServed, "Motion to Resell served"
End Sub

Private Sub ResellServed_BeforeUpdate(Cancel As Integer)
If Not BHproject Then

Cancel = CheckFutureDate(ResellServed)
End If

End Sub

Private Sub ResellServed_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ResellServed = Date
    AddStatus FileNumber, ResellServed, "Motion to Resell served"
End If

End Sub

Private Sub Sale_AfterUpdate()
If BHproject Then
Exit Sub
End If




If Forms![Case List]!CaseTypeID = 8 Or Not IsNull(Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced) Then

     Dim strTrustee As Recordset
     Set strTrustee = CurrentDb.OpenRecordset("SELECT DISTINCTROW Trustees.ID from Trustees WHERE Trustees.FileNumber =" & [Forms]![foreclosuredetails]![FileNumber] & " ORDER BY Trustees.Assigned; ", dbOpenDynaset, dbSeeChanges)
     If Not strTrustee.EOF Then
            strTrustee.Close
            Set strTrustee = Nothing
            If State = "MD" Or State = "DC" Then
                    If IsNull(CourtCaseNumber) Then
                        MsgBox ("There is not Case court number")
                        Me.Undo
                        Exit Sub
                    Else
                     AddStatus FileNumber, Date, "Sale Monitor scheduled for " & Format$(Sale, "m/d/yyyy")
                     Exit Sub
                    End If
            AddStatus FileNumber, Date, "Sale Monitor scheduled for " & Format$(Sale, "m/d/yyyy")
            Exit Sub
            End If
    Else
    MsgBox ("There is no Trustee")
    Me.Undo
    Exit Sub
    End If
End If


If IsNull(Sale) Then Exit Sub
If Me.State = "VA" Then Me.Deposit = 20000
Dim FeeAmount As Currency, Update As Boolean, Jurisdiction As Long, cost As Currency

FeeAmount = Nz(DLookup("auctioneerfee", "jurisdictionlist", "jurisdictionid=" & Forms![Case List]!JurisdictionID), 0)
If Me.State = "md" Then
If IsNull(Sale.OldValue) Then AddInvoiceItem FileNumber, "FC-Auc", "Co-Counsel/Auctioneer's Fee", FeeAmount, 85, False, False, False, True
ElseIf DLookup("Auctioneercocounsel", "JurisdictionList", "jurisdictionid=" & Jurisdiction) <> 196 Then
DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Please verify the Co-Counsel and Amount, override if needed|FC-Auc|Co-Counsel|cocounsel"
End If
FeeAmount = 0

'Order Lien Cert
If Me.State = "MD" Then
 If DLookup("LienCert", "JurisdictionList", "jurisdictionid=" & Jurisdiction) = True Then
  'If MsgBox("This file is in a jurisdiction that requires a Lien Cert.  Did you want to order one?", vbYesNo) = vbYes Then
  If Not IsNull(LienCert) Then Update = True
  '  LienCert = Date  ' done on 05/14 As per Diane request SA
        Select Case Jurisdiction
        Case 4 'Balto City
        If Update = False Then
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=4")
        Else
        cost = DLookup("UpdateCost", "LienCertCosts", "jurisdictionid=4")
        End If
        Case 5 'Balto County
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=5")
        Case 14 'Harford
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=14")
        Case 15 ' Howard
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=15")
        Case 3 'Anne Arundel
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=3")
        Case 8 'Carroll
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=8")
        Case 12 'Frederick
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=12")
        Case 10 'Charles
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=10")
        Case Else
        MsgBox "Lien cert cost not found for this jurisdiction!  Please see your manager to have this added and notify accounting", vbExclamation
        End Select

    
        If LoanType = 1 Or LoanType = 4 Or LoanType = 5 Then
        AddInvoiceItem FileNumber, "FC-LIEN", "Lien Certificate Ordered", cost, 189, False, False, False, True
        AddInvoiceItem FileNumber, "FC-Lien", "Lien Cert Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, False, False, False
        Else
        AddInvoiceItem FileNumber, "FC-LIEN", "Lien Certificate Ordered", cost * 2, 189, False, False, False, True
        AddInvoiceItem FileNumber, "FC-Lien", "Lien Cert Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)) * 2, 76, False, False, False, False
        End If
   ' Call PrintLienCerts(FileNumber, Jurisdiction, Update)
'End If
End If
cmdWizComplete.Enabled = True
End If
     
    If Me.State = "VA" Then
                
         If Sale.Value <= Now() + 13 Then
         MsgBox ("Sale must be set 14 days out")
         Sale.Value = Null
         Else
        AddStatus FileNumber, Date, "Sale scheduled for " & Format$(Sale, "m/d/yyyy")
        SaleSet = Now()
        If WizardSource = "vasalesetting" Then cmdWizComplete.Enabled = True
        FNMAHoldReason = Null
        FNMAHoldReasonDate = Null
        FNMAMissingDoc = Null
        FNMAMissingDocDate = Null
        FNMAPostponeReason = Null
        FNMAPostponeReasonDate = Null
        End If
    Else
      
   ' AddStatus FileNumber, Date, "Sale scheduled for " & Format$(Sale, "m/d/yyyy")
    SaleSet = Now()

    FNMAHoldReason = Null
    FNMAHoldReasonDate = Null
    FNMAMissingDoc = Null
    FNMAMissingDocDate = Null
    FNMAPostponeReason = Null
    FNMAPostponeReasonDate = Null
  End If

End Sub

Private Sub Sale_BeforeUpdate(Cancel As Integer)

If Not BHproject Then
If Forms![Case List]!CaseTypeID = 8 Or Not IsNull(Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced) Then
     
Exit Sub
End If


If BHproject Then
Exit Sub
End If



If (IsNull(Sale)) Then Exit Sub
If Weekday(Sale) = vbSunday Or Weekday(Sale) = vbSaturday Then
    Cancel = 1
    MsgBox "Sale date cannot be Saturday or Sunday", vbCritical
End If

If Not IsNull(DLookup("Desc", "holidays", "Holiday=#" & Sale & "#")) Then
    Cancel = 1
    MsgBox "Sale date cannot be on " & DLookup("Desc", "holidays", "Holiday=#" & Sale & "#"), vbCritical
End If

'If Not IsNull(Forms!Foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced) Then GoTo aa:
    If IsNull(WizardSource) Then
    Cancel = 1
    MsgBox "Sales can only be entered using the Sale Setting wizard."
    End If
'End If
'aa:
'Virginia Logic
If (State = "VA") Then
If (Me.IRSLiens = True And Not IsNull(Me.IRSNotice)) Then
  If (DateDiff("d", Me.IRSNotice, Sale) < 36) Then
    Cancel = 1
    MsgBox "Sale Date must be 35 days after IRS Notice.", vbCritical
  End If
End If

If (DateDiff("d", FairDebt, Sale) < 31) Then
    Cancel = 1
    MsgBox "Sale Date must be 30 days after sending the Fair Debt Letter", vbCritical
End If

If Not IsNull(LostNoteNotice) And (Date - LostNoteNotice) < 13 Then
    Cancel = 1
    MsgBox "Sale Date must be 14 days after sending the Lost Note Notice", vbCritical
End If

If Not IsNull(TitleClaimDate) And IsNull(TitleClaimResolved) Then
    Cancel = 1
    MsgBox "Sale Date cannot be set with an outstanding Title Claim", vbCritical
End If

If Not IsNull(sfrmFCtitle!TitleAssignNeededDate) And IsNull(sfrmFCtitle!TitleAssignReceivedDate) Then
    Cancel = 1
    MsgBox "Sale Date cannot be set without an Assignment received", vbCritical
End If


Select Case LoanType
Case 1  'Conv
    If Sale < Date + 19 Then
    Cancel = 1
    MsgBox "Sale date must be at least 19 days in the future."
    End If
Case 2 'VA
    If IsNull(VAAppraisal) Then
    MsgBox "Sale cannot be set until VA Appraisal is ordered."
    Cancel = 1
    Else
    Select Case Sale
    Case Is < (VAAppraisal + 63)
    If Forms![Case List]!ClientID = 97 Then
    Cancel = 1
    MsgBox "Sale Date must be at least 63 days after the VA appraisal date.  Please order the VA appraisal using the wizard."
    End If
    
    Case Is < (VAAppraisal + 45)
    Cancel = 1
    MsgBox "Sale Date must be at least 45 days after the VA appraisal date.  Please order the VA appraisal using the wizard."
    'End If
    Case Is > (VAAppraisal + 180)
    Cancel = 1
    MsgBox "Sale is attempting to be set after the VA Appraisal will expire.  See manager to re-order the appraisal or schedule sale within 180 days of appraisal date."
    End Select
    End If
Case 3 'HUD
    If Sale < Date + 30 Then
    Cancel = 1
    MsgBox "Sale date must be at least 30 days in the future."
    End If
Case 4  'FNMA
    If Sale < Date + 26 Then
    Cancel = 1
    MsgBox "Sale date must be at least 25 days in the future."
    End If
Case 5 'FHLMC
    DoCmd.OpenForm "EnterSaleSettingOption", , , , , acDialog
    If Autovalue = "Autovalue" Then
    If Sale < Date + 23 Then
    Cancel = 1
    MsgBox "Sale date must be at least 22 days in the future."
    End If
    End If
    If Autovalue = "BPO" Then
        If Sale < Date + 31 Then
        Cancel = 1
        MsgBox "Sale date must be at least 30 days in the future."
        End If
    End If
End Select

    If HearingCheking(Sale, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(Sale, 2) = 1 Then
    Cancel = 1
    End If

If Not IsNull(AccelerationLetter) Then
  If (AccelerationLetter >= Date) Then
   Cancel = 1
   MsgBox "Cannot set sale until Demand Letter expires", vbCritical
  End If
End If
  
End If

'MD Logic
If (State = "MD") Then
If (Me.IRSLiens = True And Not IsNull(Me.IRSNotice)) Then
  If (DateDiff("d", Me.IRSNotice, Sale) < 45) Then
    Cancel = 1
    MsgBox "Sale Date must be 45 days after IRS Notice.", vbCritical
  End If
End If

Select Case LoanType

Case 2 'VA type
Select Case Sale

Case Is < (VAAppraisal + 63)
 If Forms![Case List]!ClientID = 97 Then
 Cancel = 1
MsgBox "Sale Date must be at least 63 days after the VA appraisal date.  Please order the VA appraisal using the wizard."
End If

Case Is < (VAAppraisal + 45)
Cancel = 1
MsgBox "Sale Date must be at least 45 days after the VA appraisal date.  Please order the VA appraisal using the wizard."

Case Is > (VAAppraisal + 180)
Cancel = 1
MsgBox "Sale is attempting to be set after the VA Appraisal will expire.  See manager to re-order the appraisal or schedule sale within 180 days of appraisal date."
Case IsNull(VAAppraisal) = True
Cancel = 1
MsgBox "Sale cannot be set until VA Appraisal is ordered."
End Select


Case 3 'HUD
If Sale < Date + 30 Then
Cancel = 1
MsgBox "Sale date must be at least 30 days in the future."
End If
Case 4 Or 1 'Conv/FNMA
If Sale < Date + 22 Then
Cancel = 1
MsgBox "Sale date must be at least 22 days in the future."
End If
Case 5 'FHLMC
DoCmd.OpenForm "EnterSaleSettingOption", , , , , acDialog
If Autovalue = "Autovalue" Then
If Sale < Date + 22 Then
Cancel = 1
MsgBox "Sale date must be at least 22 days in the future."
End If
End If
If Autovalue = "BPO" Then
If Sale < Date + 35 Then
Cancel = 1
MsgBox "Sale date must be at least 35 days in the future."
End If
End If
End Select

    If HearingCheking(Sale, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(Sale, 2) = 1 Then
    Cancel = 1
    End If

If Not IsNull(AccelerationLetter) Then
  If (AccelerationLetter >= Sale) Then
   Cancel = 1
   MsgBox "Sale Date must be greater than Acceleration.", vbCritical
  End If
End If

  If IsNull(Me.LossMitFinalDate) Or (DateDiff("d", Me.LossMitFinalDate, Sale) < 15) Then
    Cancel = 1
    MsgBox "Sale Date must be 15 days after Final Loss Mitigation Affidavit Date.", vbCritical
  End If
  
' comment as per Diane request on 10/22/2013
'  If IsNull(Me.ServiceSent) Or (DateDiff("d", ServiceSent, Sale) < 46) Then
'    Cancel = 1
'    MsgBox "Date must be 45 days after Service Sent", vbCritical
'  End If
'
'  If IsNull(Me.BorrowerServed) Or (DateDiff("d", BorrowerServed, Sale) < 46) Then
'    Cancel = 1
'    MsgBox "Date must be 45 days after Borrower Served", vbCritical
'  End If
  
  If IsNull(Me.ServiceMailed) Or (DateDiff("d", ServiceMailed, Sale) < 45) Then
    Cancel = 1
    MsgBox "Date must be 45 days after Notice to Occupant", vbCritical
  End If
  End If
End If ' for bhproject

'End If
End Sub

Private Sub SaleRat_AfterUpdate()
AddStatus FileNumber, SaleRat, "Sale ratified/confirmed"
Dim FeeAmount As Currency, Update As Boolean, Jurisdiction As Long

'Order Lien Cert
If State = "MD" Then
Jurisdiction = Forms![Case List]!JurisdictionID
If DLookup("LienCert", "JurisdictionList", "jurisdictionid=" & Jurisdiction) = True Then
If Not IsNull(LienCert) Then
Update = True
End If
End If
End If
' As per Diane request 05/14 SA.
''If MsgBox("This file is in a jurisdiction that requires a Lien Cert.  Did you want to order one?", vbYesNo) = vbYes Then 'stope as per Daine request 05/13 SA
''LienCert = Date
''Call PrintLienCerts(FileNumber, Jurisdiction, Update)
'End If
'End If
'End If


End Sub

Private Sub SaleRat_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(SaleRat)
End If

End Sub

Private Sub SaleRat_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    SaleRat = Now()
    AddStatus FileNumber, SaleRat, "Sale ratified/confirmed"
End If

End Sub

Private Sub SaleTime_AfterUpdate()
If Not BHproject Then

'    If IsNull(Sale) Or IsNull(SaleTime) Then Exit Sub
'
'        If Hour(SaleTime) < 8 Or Hour(SaleTime) > 19 Then
'           ' SaleTime = DateAdd("h", 12, SaleTime)  'As per Diane Request 9/23 SA
'            'If Hour(SaleTime) < 8 Or Hour(SaleTime) > 19 Then
'                MsgBox "Invalid sale time: " & Format$(SaleTime, "h:nn am/pm")
'                SaleTime = Null
'                Exit Sub
'            'End If
'        End If
If IsNull(Sale) Or IsNull(SaleTime) Then Exit Sub

If Hour(SaleTime) >= 8 And Hour(SaleTime) < 13 Then
        SaleTime = Format$(SaleTime, "h:nn am/pm")
    ElseIf Hour(SaleTime) >= 1 And Hour(SaleTime) <= 7 Then
        SaleTime = DateAdd("h", 12, SaleTime)
    Else
        MsgBox "Invalid sale time: " & Format$(SaleTime, "h:nn am/pm")
        SaleTime = Null
        Exit Sub
    End If


    AddStatus FileNumber, Now(), "Scheduled foreclosure sale for " & Format$(Sale, "m/d/yyyy") & " at " & Format$(SaleTime, "h:nn am/pm")
Else
    AddStatus FileNumber, Now(), "Scheduled foreclosure sale for " & Format$(Sale, "m/d/yyyy") & " at " & Format$(SaleTime, "h:nn am/pm")

End If

End Sub

Private Sub SentToDocket_AfterUpdate()
If BHproject Then
AddStatus FileNumber, SentToDocket, "Case sent to court for filing"
Else

Dim FeeAmount As Currency, Update As Boolean, Jurisdiction As Long, cost As Currency

AddStatus FileNumber, SentToDocket, "Case sent to court for filing"
If State = "MD" Then
    
    If IsNull(txtNOIExpires) Then
     MsgBox "Date must be 45 days after 45 Day Notice", vbCritical
     Me.SentToDocket = ""
     Exit Sub
        Else
    
    If SentToDocket < txtNOIExpires Then
    MsgBox "Date must be 45 days after 45 Day Notice", vbCritical
    Me.SentToDocket = ""
    Exit Sub
    End If
    End If


 '   AddInvoiceItem FileNumber, "FC-DKT", "Substitution of Trustees", 40, False, True, False, True
 '   FeeAmount = Nz(DLookup("FeeDocket", "JurisdictionList", "JurisdictionID=" & Forms![Case List]!JurisdictionID))
 '   If FeeAmount > 0 Then
 '       AddInvoiceItem FileNumber, "FC-DKT", "Filing Fee", FeeAmount, False, True, False, True
 '   Else
 '       AddInvoiceItem FileNumber, "FC-DKT", "Filing Fee", GetFeeAmount("Filing Fee"), False, True, False, True
 '   End If
    If Forms![Case List]!JurisdictionID = 18 Then   ' PG County
        AddInvoiceItem FileNumber, "FC-ENV", "Environmental Letter Mailing", Nz(DLookup("Value", "StandardCharges", "ID=" & 8)), 76, False, True, False, True
    End If

'Milestone Billing for Referral Fee
Dim InvPct As Double
If State = "MD" Then


'    Select Case LoanType
'    Case 4
'    FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=177"))
'    Case 5
'    FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=263"))
'    Case Else
'    FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & Forms![case list]!ClientID))
'    End Select
'
'        If FeeAmount > 0 Then
'            InvPct = DLookup("MDComplaintFiledpct", "clientlist", "clientid=" & Forms![case list]!ClientID)
'            If InvPct < 1 Then
'            AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when docketed of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
'            Else
'                AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
'            End If
'        End If
Dim cbxClient As Integer
cbxClient = Nz(DLookup("ClientID", "Caselist", "FileNumber=" & FileNumber))
   Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
      Case 1 'Conventional
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & cbxClient))
      Case 2 'VA or Veteran's Affairs
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
      Case 3 'FHA or HUD
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=570")) 'HUD/FHA
      Case 4
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=177")) 'Fannie Mae
      Case 5
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=263")) 'Freddie Mac
      Case Else
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & cbxClient))
    End Select
    If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
      InvPct = DLookup("MDComplaintFiledPct", "clientlist", "clientid=" & cbxClient)
    Else
      InvPct = 1
    End If
    If FeeAmount > 0 Then
      If InvPct <= 1 Then
        AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when docketed of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      Else
        'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
      End If
    End If
        'Removed per Diane 1/30, file does not go to Need to Invoice at docketing
'Forms![case list]!BillCase = True
'Forms![case list]!BillCaseUpdateUser = GetStaffID()
'Forms![case list]!BillCaseUpdateDate = Date
'Forms![case list]!BillCaseUpdateReasonID = 4
'Forms![case list]!lblBilling.Visible = True

'Dim rstBillReasons As Recordset
'Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
'With rstBillReasons
'.AddNew
'!FileNumber = FileNumber
'!billingreasonid = 4
'!userid = GetStaffID
'!Date = Date
'.Update
'End With

End If

End If

End If

End Sub

Private Sub SentToDocket_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
If Not PrivDataManager Or WizardSource = "Service" Then
MsgBox "You must have admin privileges to change the Docket date"
     
Cancel = CheckFutureDate(SentToDocket)
If (Cancel = 1) Then Exit Sub


'If (DateDiff("d", NOI, SentToDocket) < 40) Then
'   Cancel = 1
'   MsgBox "Sent to Docket Date must be at least 40 days from 45 Day Notice.", vbCritical
'ElseIf (DateDiff("d", NOI, SentToDocket) < 46) Then
'
'   Dim retval As Integer
'   retval = MsgBox("Sent to Docket Date is less the 45 day from 45 Day Notice.  Do you really want to send?", vbYesNo)
'   If retval = vbNo Then
'     Cancel = 1
'   End If

'End If
End If
End If

End Sub

Private Sub SentToDocket_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    SentToDocket = Date
    Call SentToDocket_AfterUpdate
End If

End Sub

Private Sub ServiceSent_AfterUpdate()
If BHproject Then
If Not IsNull(ServiceSent) Then
AddStatus FileNumber, ServiceSent, "Service sent"
End If
Else

If Not IsNull(ServiceSent) Then
AddStatus FileNumber, ServiceSent, "Service sent"
AddInvoiceItem FileNumber, "FC-SVC", "Postage for Service Sent", 1, 76, False, False, False, False
End If
End If

End Sub

Private Sub ServiceSent_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(ServiceSent)
If (Cancel = 1) Then Exit Sub


  If (DateDiff("d", Docket, ServiceSent) < 0) Then
    Cancel = 1
    MsgBox "Date cannot be before Docket Date.", vbCritical
  End If
End If

End Sub

Private Sub ServiceSent_DblClick(Cancel As Integer)

If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
ElseIf (Me.WizardSource = "None" And Me.State = "MD") Or (Me.State = "MD" And Len(Me.WizardSource & "") = 0) Then
    DoCmd.CancelEvent
Else
    ServiceSent = Date
    Call ServiceSent_AfterUpdate
End If
End Sub

Private Sub Settled_AfterUpdate()
AddStatus FileNumber, Now(), "3rd party settled"
Call Visuals
End Sub

Private Sub Settled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Settled = Now()
    AddStatus FileNumber, Now(), "3rd party settled"
    Call Visuals
End If

End Sub



Private Sub State_AfterUpdate()
Call Visuals
End Sub



Private Sub StatementOfDebtDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    StatementOfDebtDate = Date
End If

End Sub


Private Sub Text562_DblClick(Cancel As Integer)

End Sub


Private Sub TitleClaim_AfterUpdate()
Call Visuals
End Sub

Private Sub TitleClaimResolved_AfterUpdate()
AddStatus FileNumber, TitleClaimResolved, "Resolved title claim"

If Not IsNull([TitleClaimResolved]) And (Forms![Case List].[Active] And [OnStatusReport]) Then
  Forms![Case List]!RestartReceived = Date
  AddStatus FileNumber, Date, "Restart Received"
  
End If

End Sub

Private Sub TitleClaimResolved_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
  Cancel = CheckFutureDate(TitleClaimResolved)
End If

End Sub

Private Sub TitleClaimResolved_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    TitleClaimResolved = Now()
    
    If IsNull(Comment) Then
        Comment = Format$(Date, "m/d/yyyy") & " Resolved title claim"
    Else
        Comment = Comment & vbNewLine & Format$(Date, "m/d/yyyy") & " Resolved title claim"
    End If
    
    Call TitleClaimResolved_AfterUpdate
End If

End Sub

Private Sub TitleClaimResolved2_AfterUpdate()
AddStatus FileNumber, TitleClaimResolved, "Resolved title claim"
If IsNull(Comment) Then
    Comment = Format$(Date, "m/d/yyyy") & " Resolved title claim"
Else
    Comment = Comment & vbNewLine & Format$(Date, "m/d/yyyy") & " Resolved title claim"
End If
End Sub

Private Sub TitleClaimResolved2_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleClaimResolved2)
End Sub

Private Sub TitleClaimResolved2_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    TitleClaimResolved = Now()
    AddStatus FileNumber, TitleClaimResolved, "Resolved title claim"
    If IsNull(Comment) Then
        Comment = Format$(Date, "m/d/yyyy") & " Resolved title claim"
    Else
        Comment = Comment & vbNewLine & Format$(Date, "m/d/yyyy") & " Resolved title claim"
    End If
End If
End Sub

Private Sub TitleClaimSent_AfterUpdate()
If BHproject Then
If Not IsNull(TitleClaimSent) Then
AddStatus FileNumber, TitleClaimSent2, "Sent title claim"
End If
Else

If Not IsNull(TitleClaimSent) Then
AddStatus FileNumber, TitleClaimSent2, "Sent title claim"
If IsNull(Comment) Then
    Comment = Format$(Date, "m/d/yyyy") & " Resolved title claim"
Else
    Comment = Comment & vbNewLine & Format$(Date, "m/d/yyyy") & " Resolved title claim"
End If
FeeAmount = DLookup("titleclaim", "clientlist", "clientid=" & Forms![Case List]!ClientID)
If MsgBox("Do you want to override the standard fee of $" & FeeAmount & " for this client?", vbYesNo) = vbYes Then
FeeAmount = InputBox("Please enter fee, then rememeber to note the journal")
MsgBox "Please upload fee approval to documents"
End If
        
AddInvoiceItem FileNumber, "FC-TC", "Title Claim", FeeAmount, 0, True, True, True, True
End If
End If

End Sub

Private Sub TitleClaimSent_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
Cancel = CheckFutureDate(TitleClaimSent)
End If

End Sub

Private Sub TitleClaimSent_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    TitleClaimSent = Now()
    Call TitleClaimSent_AfterUpdate
End If

End Sub

Private Sub TitleClaimSent2_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleClaimSent2)
End Sub


Private Sub TitleOrder_AfterUpdate()
AddStatus FileNumber, TitleOrder, "Ordered title"
'TitleBack = "" 'stop by sarab on 09/10/14 as make no since , it is date and using ""
'TitleThru = ""
'TitleReviewToClient = ""
'Me.Requery
End Sub

Private Sub cmdSelectFile_Click()

On Error GoTo Err_cmdSelectFile_Click
DoCmd.Close acForm, Me.Name
DoCmd.OpenForm "Select File"

Exit_cmdSelectFile_Click:
    Exit Sub

Err_cmdSelectFile_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectFile_Click
    
End Sub

Private Sub TitleOrder_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleOrder)
End Sub

Private Sub TitleOrder_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    TitleOrder = Now()
    AddStatus FileNumber, TitleOrder, "Ordered title"
End If

End Sub

Private Sub TitleOrder_LossMit_AfterUpdate()
Call TitleOrder_AfterUpdate
End Sub


Private Sub TitleOrder_LossMit_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleOrder)
End Sub

Private Sub TitleOrder_LossMit_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    TitleOrder_LossMit = Now()
    AddStatus FileNumber, TitleOrder, "Ordered title"
End If

End Sub

Private Sub TitleReviewToClient_AfterUpdate()
Me.Requery
If Not IsNull(Me.TitleReviewToClient) Then
DoCmd.SetWarnings False
Dim rstsql As String
rstsql = "Insert InTo TitleReviewArchive (FileNumber, TitleReviewToClient, DateEntered) Values ( " & FileNumber & ", '" & TitleReviewToClient & "' , '" & Now() & "')"
DoCmd.RunSQL rstsql
DoCmd.SetWarnings True
End If
AddStatus FileNumber, TitleReviewToClient, "Title Review Completed"
End Sub

Private Sub TitleReviewToClient_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleReviewToClient)
End Sub

Private Sub txtFistPub_AfterUpdate()
Dim InvPct As Double
    If State = "VA" And Not IsNull(txtFistPub) Then
        Dim cbxClient As Integer
        cbxClient = Nz(DLookup("ClientID", "Caselist", "FileNumber=" & FileNumber))
        Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
          Case 1 'Conventional
            FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
          Case 2 'VA or Veteran's Affairs
            FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
          Case 3 'FHA or HUD
            FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=570")) 'HUD/FHA
          Case 4
            FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=177")) 'Fannie Mae
          Case 5
            FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=263")) 'Freddie Mac
          Case Else
            FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
        End Select
        If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
          InvPct = DLookup("VA1stactionpct", "clientlist", "clientid=" & cbxClient)
        Else
          InvPct = 1
        End If
        
        If FeeAmount > 0 Then
            If InvPct <= 1 Then
              AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due at 1st action of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
            Else
              'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
            End If
        End If
    'Removed per Diane 1/30, do not milestone bill VA files
    'Forms![case list]!BillCase = True
    'Forms![case list]!BillCaseUpdateUser = GetStaffID()
    'Forms![case list]!BillCaseUpdateDate = Date
    'Forms![case list]!BillCaseUpdateReasonID = 3
    'Forms![case list]!lblBilling.Visible = True
    
    'Dim rstBillReasons As Recordset
    'Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
    'With rstBillReasons
    '.AddNew
    '!FileNumber = FileNumber
    '!billingreasonid = 3
    '!userid = GetStaffID
    '!Date = Date
    '.Update
    'End With
    
    End If
    
    AddStatus FileNumber, txtFistPub, "1st Legal Advertisement Runs"
    
    DoCmd.SetWarnings False
    strinfo = "1st Legal Advertisement Runs"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End Sub

Private Sub txtFistPub_BeforeUpdate(Cancel As Integer)

If Not BHproject Then
    If Not IsNull(Disposition) And Nz(SaleCompleted) = 0 Then
        If txtFistPub > DateAdd("d", 7, DispositionDate) Then
            Cancel = 1
            MsgBox "Date must be within 7 days of setting the Disposition.  Otherwise you probably need to add a foreclosure.", vbCritical
        End If
    End If

    If txtFistPub <= DateDiff("d", -120, LPIDate) And Forms![Case List]!ClientID <> 567 Then
        Cancel = 1
        MsgBox "First Publication Dates Must Be 120 days past the LPI Date.", vbCritical
    'Info = DateDiff("d", -120, LPIDate)
    End If
End If

End Sub

Private Sub txtFistPub_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    If IsNull(Disposition) Then
        txtFistPub = Now()
        Call txtFistPub_AfterUpdate
    End If
End If
End Sub

Private Sub txtreviewadproof_AfterUpdate()
AddStatus FileNumber, txtreviewadproof, "Reviewed Proof of Advertising"

'If IsNull(txtNewAdVendor) Or txtNewAdVendor = "" Then
DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter Advertising(Newspapers) cost|FC-ADV|Advertising|Advertising"
'End If
End Sub

Private Sub txtreviewadproof_DblClick(Cancel As Integer)
txtreviewadproof = Date
Call txtreviewadproof_AfterUpdate
End Sub

Private Sub txtSale_AfterUpdate()
If Forms![Case List]!CaseTypeID = 8 Or Not IsNull(Forms!foreclosuredetails!sfrmMonitorReferRecd!Monitor_Refer_reced) Then

     Dim strTrustee As Recordset
     Set strTrustee = CurrentDb.OpenRecordset("SELECT DISTINCTROW Trustees.ID From Trustees WHERE Trustees.FileNumber =" & [Forms]![foreclosuredetails]![FileNumber] & " ORDER BY Trustees.Assigned; ", dbOpenDynaset, dbSeeChanges)
     If Not strTrustee.EOF Then
            strTrustee.Close
            Set strTrustee = Nothing
            If State = "MD" Or State = "DC" Then
                    If IsNull(CourtCaseNumber) Then
                        MsgBox ("There is not Case court number")
                        Me.Undo
                        Exit Sub
                    Else
                        AddStatus FileNumber, Date, "Sale Monitor scheduled for " & Format$(txtSale, "m/d/yyyy")
                    Exit Sub
                    End If
            AddStatus FileNumber, Date, "Sale Monitor scheduled for " & Format$(txtSale, "m/d/yyyy")
            Exit Sub
            End If
    Else
        MsgBox ("There is no Trustee")
        Me.Undo
    Exit Sub
    End If
End If


If IsNull(txtSale) Then Exit Sub
If Not IsNull(Me.Disposition) And Not IsNull(txtSale) Then
    MsgBox ("Can not set sale with a disposition")
    txtSale = Null
End If
Exit Sub


'    If Me.State = "VA" Then Me.Deposit = 20000
'        Dim FeeAmount As Currency, Update As Boolean, Jurisdiction As Long, cost As Currency
'
'        FeeAmount = Nz(DLookup("auctioneerfee", "jurisdictionlist", "jurisdictionid=" & Forms![Case List]!JurisdictionID), 0)
'    If Me.State = "md" Then
'        If IsNull(Sale.OldValue) Then AddInvoiceItem FileNumber, "FC-Auc", "Co-Counsel/Auctioneer's Fee", FeeAmount, 85, False, False, False, True
'
'        ElseIf DLookup("Auctioneercocounsel", "JurisdictionList", "jurisdictionid=" & Jurisdiction) <> 196 Then
'            DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Please verify the Co-Counsel and Amount, override if needed|FC-Auc|Co-Counsel|cocounsel"
'        End If
'        FeeAmount = 0

'Order Lien Cert
'If Me.State = "MD" Then
'  If DLookup("LienCert", "JurisdictionList", "jurisdictionid=" & Jurisdiction) = True Then
'  If Not IsNull(LienCert) Then Update = True
'        Select Case Jurisdiction
'        Case 4 'Balto City
'        If Update = False Then
'        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=4")
'        Else
'        cost = DLookup("UpdateCost", "LienCertCosts", "jurisdictionid=4")
'        End If
'        Case 5 'Balto County
'        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=5")
'        Case 14 'Harford
'        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=14")
'        Case 15 ' Howard
'        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=15")
'        Case 3 'Anne Arundel
'        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=3")
'        Case 8 'Carroll
'        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=8")
'        Case 12 'Frederick
'        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=12")
'        Case 10 'Charles
'        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=10")
'        Case Else
'        MsgBox "Lien cert cost not found for this jurisdiction!  Please see your manager to have this added and notify accounting", vbExclamation
'        End Select
'
'
'        If LoanType = 1 Or LoanType = 4 Or LoanType = 5 Then
'        AddInvoiceItem FileNumber, "FC-LIEN", "Lien Certificate Ordered", cost, 189, False, False, False, True
'        AddInvoiceItem FileNumber, "FC-Lien", "Lien Cert Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, False, False, False
'        Else
'        AddInvoiceItem FileNumber, "FC-LIEN", "Lien Certificate Ordered", cost * 2, 189, False, False, False, True
'        AddInvoiceItem FileNumber, "FC-Lien", "Lien Cert Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)) * 2, 76, False, False, False, False
'        End If
'   ' Call PrintLienCerts(FileNumber, Jurisdiction, Update)
''End If
'End If
'cmdWizComplete.Enabled = True
'End If
     
'    If Me.State = "VA" Then
'
'         If Sale.Value <= Now() + 13 Then
'         MsgBox ("Sale must be set 14 days out")
'         Sale.Value = Null
'         Else
        'AddStatus FileNumber, Date, "DC Sale scheduled for " & Format$(txtSale, "m/d/yyyy")
        txtSaleSet = Now()
''        If WizardSource = "vasalesetting" Then cmdWizComplete.Enabled = True
''        FNMAHoldReason = Null
''        FNMAHoldReasonDate = Null
''        FNMAMissingDoc = Null
''        FNMAMissingDocDate = Null
''        FNMAPostponeReason = Null
''        FNMAPostponeReasonDate = Null
''        End If
''    Else
''
''   ' AddStatus FileNumber, Date, "Sale scheduled for " & Format$(Sale, "m/d/yyyy")
''    SaleSet = Now()
''
''    FNMAHoldReason = Null
''    FNMAHoldReasonDate = Null
''    FNMAMissingDoc = Null
''    FNMAMissingDocDate = Null
''    FNMAPostponeReason = Null
''    FNMAPostponeReasonDate = Null
  ''End If
End Sub

Private Sub txtSale_BeforeUpdate(Cancel As Integer)

If Not BHproject Then

    If (IsNull(txtSale)) Then Exit Sub
        If Weekday(txtSale) = vbSunday Or Weekday(txtSale) = vbSaturday Then
            Cancel = 1
            MsgBox "Sale date cannot be Saturday or Sunday", vbCritical
        End If
    
    If Not IsNull(DLookup("Desc", "holidays", "Holiday=#" & txtSale & "#")) Then
        Cancel = 1
        MsgBox "Sale date cannot be on " & DLookup("Desc", "holidays", "Holiday=#" & txtSale & "#"), vbCritical
    End If

    'If IsNull(WizardSource) Then
    'Cancel = 1
    'MsgBox "Sales can only be entered using the Sale Setting wizard."
    'End If

'If (State = "VA") Then
'    If (Me.IRSLiens = True And Not IsNull(Me.IRSNotice)) Then
'      If (DateDiff("d", Me.IRSNotice, Sale) < 36) Then
'        Cancel = 1
'        MsgBox "Sale Date must be 35 days after IRS Notice.", vbCritical
'      End If
'    End If
'
'    If (DateDiff("d", FairDebt, Sale) < 31) Then
'        Cancel = 1
'        MsgBox "Sale Date must be 30 days after sending the Fair Debt Letter", vbCritical
'    End If
'
'    If Not IsNull(LostNoteNotice) And (Date - LostNoteNotice) < 13 Then
'        Cancel = 1
'        MsgBox "Sale Date must be 14 days after sending the Lost Note Notice", vbCritical
'    End If
'
'    If Not IsNull(TitleClaimDate) And IsNull(TitleClaimResolved) Then
'        Cancel = 1
'        MsgBox "Sale Date cannot be set with an outstanding Title Claim", vbCritical
'    End If
'
'    If Not IsNull(sfrmFCtitle!TitleAssignNeededDate) And IsNull(sfrmFCtitle!TitleAssignReceivedDate) Then
'        Cancel = 1
'        MsgBox "Sale Date cannot be set without an Assignment received", vbCritical
'    End If
'
'
'    Select Case LoanType
'    Case 1  'Conv
'        If Sale < Date + 19 Then
'        Cancel = 1
'        MsgBox "Sale date must be at least 19 days in the future."
'        End If
'    Case 2 'VA
'        If IsNull(VAAppraisal) Then
'        MsgBox "Sale cannot be set until VA Appraisal is ordered."
'        Cancel = 1
'        Else
'        Select Case Sale
'        Case Is < (VAAppraisal + 63)
'        If Forms![Case List]!ClientID = 97 Then
'        Cancel = 1
'        MsgBox "Sale Date must be at least 63 days after the VA appraisal date.  Please order the VA appraisal using the wizard."
'        End If
'
'        Case Is < (VAAppraisal + 45)
'        Cancel = 1
'        MsgBox "Sale Date must be at least 45 days after the VA appraisal date.  Please order the VA appraisal using the wizard."
'        'End If
'        Case Is > (VAAppraisal + 180)
'        Cancel = 1
'        MsgBox "Sale is attempting to be set after the VA Appraisal will expire.  See manager to re-order the appraisal or schedule sale within 180 days of appraisal date."
'        End Select
'        End If
'    Case 3 'HUD
'        If Sale < Date + 30 Then
'        Cancel = 1
'        MsgBox "Sale date must be at least 30 days in the future."
'        End If
'    Case 4  'FNMA
'        If Sale < Date + 26 Then
'        Cancel = 1
'        MsgBox "Sale date must be at least 25 days in the future."
'        End If
'    Case 5 'FHLMC
'        DoCmd.OpenForm "EnterSaleSettingOption", , , , , acDialog
'        If Autovalue = "Autovalue" Then
'        If Sale < Date + 23 Then
'        Cancel = 1
'        MsgBox "Sale date must be at least 22 days in the future."
'        End If
'        End If
'        If Autovalue = "BPO" Then
'            If Sale < Date + 31 Then
'            Cancel = 1
'            MsgBox "Sale date must be at least 30 days in the future."
'            End If
'        End If
'    End Select
'
'        If HearingCheking(Sale, 1) = 1 Then
'        Cancel = 1
'        End If
'        If HearingCheking(Sale, 2) = 1 Then
'        Cancel = 1
'        End If
'
'    If Not IsNull(AccelerationLetter) Then
'      If (AccelerationLetter >= Date) Then
'       Cancel = 1
'       MsgBox "Cannot set sale until Demand Letter expires", vbCritical
'      End If
'    End If
'
'End If

'MD Logic
'    If (State = "MD") Then
'        If (Me.IRSLiens = True And Not IsNull(Me.IRSNotice)) Then
'          If (DateDiff("d", Me.IRSNotice, Sale) < 45) Then
'            Cancel = 1
'            MsgBox "Sale Date must be 45 days after IRS Notice.", vbCritical
'          End If
'        End If
'
'        Select Case LoanType
'
'        Case 2 'VA type
'        Select Case Sale
'
'        Case Is < (VAAppraisal + 63)
'         If Forms![Case List]!ClientID = 97 Then
'         Cancel = 1
'        MsgBox "Sale Date must be at least 63 days after the VA appraisal date.  Please order the VA appraisal using the wizard."
'        End If
'
'        Case Is < (VAAppraisal + 45)
'        Cancel = 1
'        MsgBox "Sale Date must be at least 45 days after the VA appraisal date.  Please order the VA appraisal using the wizard."
'
'        Case Is > (VAAppraisal + 180)
'        Cancel = 1
'        MsgBox "Sale is attempting to be set after the VA Appraisal will expire.  See manager to re-order the appraisal or schedule sale within 180 days of appraisal date."
'        Case IsNull(VAAppraisal) = True
'        Cancel = 1
'        MsgBox "Sale cannot be set until VA Appraisal is ordered."
'        End Select
'
'
'        Case 3 'HUD
'        If Sale < Date + 30 Then
'        Cancel = 1
'        MsgBox "Sale date must be at least 30 days in the future."
'        End If
'        Case 4 Or 1 'Conv/FNMA
'        If Sale < Date + 22 Then
'        Cancel = 1
'        MsgBox "Sale date must be at least 22 days in the future."
'        End If
'        Case 5 'FHLMC
'        DoCmd.OpenForm "EnterSaleSettingOption", , , , , acDialog
'        If Autovalue = "Autovalue" Then
'        If Sale < Date + 22 Then
'        Cancel = 1
'        MsgBox "Sale date must be at least 22 days in the future."
'        End If
'        End If
'        If Autovalue = "BPO" Then
'        If Sale < Date + 35 Then
'        Cancel = 1
'        MsgBox "Sale date must be at least 35 days in the future."
'        End If
'        End If
'        End Select
'
'        If HearingCheking(Sale, 1) = 1 Then
'        Cancel = 1
'        End If
'        If HearingCheking(Sale, 2) = 1 Then
'        Cancel = 1
'        End If
'
'        If Not IsNull(AccelerationLetter) Then
'          If (AccelerationLetter >= Sale) Then
'           Cancel = 1
'           MsgBox "Sale Date must be greater than Acceleration.", vbCritical
'          End If
'        End If
'
'        If IsNull(Me.LossMitFinalDate) Or (DateDiff("d", Me.LossMitFinalDate, Sale) < 15) Then
'          Cancel = 1
'          MsgBox "Sale Date must be 15 days after Final Loss Mitigation Affidavit Date.", vbCritical
'        End If
'
'    ' comment as per Diane request on 10/22/2013
'    '  If IsNull(Me.ServiceSent) Or (DateDiff("d", ServiceSent, Sale) < 46) Then
'    '    Cancel = 1
'    '    MsgBox "Date must be 45 days after Service Sent", vbCritical
'    '  End If
'    '
'    '  If IsNull(Me.BorrowerServed) Or (DateDiff("d", BorrowerServed, Sale) < 46) Then
'    '    Cancel = 1
'    '    MsgBox "Date must be 45 days after Borrower Served", vbCritical
'    '  End If
'
'        If IsNull(Me.ServiceMailed) Or (DateDiff("d", ServiceMailed, Sale) < 45) Then
'          Cancel = 1
'          MsgBox "Date must be 45 days after Notice to Occupant", vbCritical
'        End If
'      End If
End If ' for bhproject
End Sub

Private Sub txtSaleTime_AfterUpdate()
'If Not BHproject Then
'
'    If IsNull(txtSale) Or IsNull(txtSaleTime) Then Exit Sub
'
'        If Hour(txtSaleTime) < 8 Or Hour(txtSaleTime) > 19 Then
'           ' SaleTime = DateAdd("h", 12, SaleTime)  'As per Diane Request 9/23 SA
'            'If Hour(SaleTime) < 8 Or Hour(SaleTime) > 19 Then
'                MsgBox "Invalid sale time: " & Format$(SaleTime, "h:nn am/pm")
'                txtSaleTime = Null
'                Exit Sub
'            'End If
'        End If
'
'    AddStatus FileNumber, Now(), "Scheduled DC foreclosure sale for " & Format$(txtSale, "m/d/yyyy") & " at " & Format$(txtSaleTime, "h:nn am/pm")
'Else
'    AddStatus FileNumber, Now(), "Scheduled DC foreclosure sale for " & Format$(txtSale, "m/d/yyyy") & " at " & Format$(txtSaleTime, "h:nn am/pm")
'
'End If

'Modified on 9_2_15
If Not BHproject Then

    If IsNull(txtSale) Or IsNull(txtSaleTime) Then
        txtSaleTime = Null
        txtSale = Null
    Exit Sub
    End If
    
    If Hour(txtSaleTime) >= 8 And Hour(txtSaleTime) < 13 Then
        txtSaleTime = Format$(txtSaleTime, "h:nn am/pm")
    ElseIf Hour(txtSaleTime) >= 1 And Hour(txtSaleTime) <= 7 Then
        txtSaleTime = DateAdd("h", 12, txtSaleTime)
    Else
        MsgBox "Invalid sale time: " & Format$(txtSaleTime, "h:nn am/pm")
        txtSaleTime = Null
        Exit Sub
    End If

    AddStatus FileNumber, Now(), "Scheduled DC foreclosure sale for " & Format$(txtSale, "m/d/yyyy") & " at " & Format$(txtSaleTime, "h:nn am/pm")
Else
    AddStatus FileNumber, Now(), "Scheduled DC foreclosure sale for " & Format$(txtSale, "m/d/yyyy") & " at " & Format$(txtSaleTime, "h:nn am/pm")

End If

End Sub



Private Sub UpdatedNotices_AfterUpdate()
If BHproject Then
AddStatus FileNumber, Notices, "Updated notices sent"
Else

AddStatus FileNumber, Notices, "Updated notices sent"
Dim noticecnt As Integer
noticecnt = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and nz(NoticeType,0) > 0")
If Me.State = "MD" Then noticecnt = noticecnt + 1
If (noticecnt > 0) Then

   AddInvoiceItem FileNumber, "FC-NOT", "Sale Notices - Certified Postage", (Nz(DLookup("Value", "StandardCharges", "ID=" & 8))) * noticecnt, 76, False, False, False, True
   AddInvoiceItem FileNumber, "FC-NOT", "Sale Notices - First Class Postage", (Nz(DLookup("Value", "StandardCharges", "ID=" & 1))) * noticecnt, 76, False, False, False, True
    
End If
End If

End Sub

Private Sub UpdatedNotices_BeforeUpdate(Cancel As Integer)
If Not BHproject Then
  Cancel = CheckFutureDate(UpdatedNotices)
End If

End Sub

Private Sub VAAppraisal_AfterUpdate()
AddStatus FileNumber, VAAppraisal, "Ordered VA Appraisal"
End Sub

Private Sub VAAppraisal_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(VAAppraisal)
'If (Date - VAAppraisal) < 180 Then
'MsgBox "Appraisal has already been ordered within the last 6 months", vbCritical
'Cancel = 1
'End If
    



End Sub

Private Sub VAAppraisal_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    VAAppraisal = Now()
    AddStatus FileNumber, VAAppraisal, "Ordered VA Appraisal"
End If

End Sub

Private Sub cmdAddFC_Click()
On Error GoTo Err_cmdAddFC_Click

If (Nz(Disposition) = 2) Or (Nz(Disposition) = 1) Then
    If PrivAdmin Then
        If MsgBox("The property has already been sold! Are you sure you want to add another Foreclosure?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
    Else
        MsgBox "You cannot add another Foreclosure because the property has already been sold.  (Management can override this for you.)", vbCritical
        Exit Sub
    End If
Else
    If MsgBox("Are you sure you want to add another Foreclosure?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
End If

Forms![Case List]!ReferralDate = Date
Forms![Case List]!ReferralDocsReceived = Null
Forms![Case List]!RestartReceived = Null
Forms![Case List]!Active = True
Forms![Case List]!OnStatusReport = True

Call AddStatus(FileNumber, Now(), "Referral Date")

Me.AllowAdditions = True
DoCmd.GoToRecord , , acNewRec
Me.AllowAdditions = False

Exit_cmdAddFC_Click:
    Exit Sub

Err_cmdAddFC_Click:
    MsgBox Err.Description
    Resume Exit_cmdAddFC_Click
    
End Sub

Private Sub cmdAudit_Click()
Dim A As Recordset, J As Recordset

On Error GoTo Err_cmdAudit_Click

Set A = CurrentDb.OpenRecordset("SELECT ForeclosureID FROM Audits WHERE ForeclosureID=" & Me![ForeclosureID], dbOpenSnapshot)
If A.EOF Then   ' create new record
    A.Close
    Set A = CurrentDb.OpenRecordset("Audits", dbOpenDynaset, dbSeeChanges)
    A.AddNew
    A!ForeclosureID = Me![ForeclosureID]
    Set J = CurrentDb.OpenRecordset("SELECT Newspaper FROM JurisdictionList INNER JOIN CaseList ON (JurisdictionList.JurisdictionID = CaseList.JurisdictionID) WHERE FileNumber = " & Me!FileNumber, dbOpenSnapshot)
    A!NewspaperPreSale = J!Newspaper
    A!NewspaperNiSi = J!Newspaper
    J.Close
    A!Proceeds = 0
    A!Interest = 0
    A!UnpaidBalance = 0
    A!PropertyTaxes = 0
    A!DelinquentTaxes = 0
    A!TaxSaleRedemption = 0
    A!Water = 0
    A!FilingFee = 0
    A!DeedApp = 0
    A!BondFiling = 0
    A!Bond = 0
    A!AdvertisingPreSale = 0
    A!AdvertisingNiSi = 0
    A!Commission = 0
    A!AuditorFee = 0
    A!AttorneyFee = 0
    A!TitleReport = 0
    A!AuctioneerFee = 0
    A!NoticesMailingCount = 0
    A!NoticesMailingAmount = 0
    A!ChicagoJudgement = 0
    A!OtherAmount = 0
    A!Other2Amount = 0
    A!Other3Amount = 0
    A!Other4Amount = 0
    A!Other5Amount = 0
    A!InterestPerDiem = 0
    A!GrantorsTax = 0
    A!CommissionerOfAccounts = 0
    A!Mail = 0
    A!RecordTrusteesDeed = 0
    A.Update
End If
A.Close
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm "Audit - " & State, , , "[ForeclosureID]=" & Me![ForeclosureID]

Exit_cmdAudit_Click:
    Exit Sub

Err_cmdAudit_Click:
    MsgBox Err.Description
    Resume Exit_cmdAudit_Click
    
End Sub

Private Sub cmdCalcPerDiem_Click()

On Error GoTo Err_cmdCalcPerDiem_Click
PerDiem = RemainingPBal * InterestRate / 100 / 365

Exit_cmdCalcPerDiem_Click:
    Exit Sub

Err_cmdCalcPerDiem_Click:
    MsgBox Err.Description
    Resume Exit_cmdCalcPerDiem_Click
    
End Sub

Private Sub cmdAddDefaultTrustees_Click()
Dim rstTrustees As Recordset, rstStaff As Recordset

On Error GoTo Err_cmdAddDefaultTrustees_Click

Call AddDefaultTrustees(Me!FileNumber)
Me!lstTrustees.Requery

Exit_cmdAddDefaultTrustees_Click:
    Exit Sub

Err_cmdAddDefaultTrustees_Click:
    MsgBox Err.Description
    Resume Exit_cmdAddDefaultTrustees_Click
    
End Sub
Private Sub cmdSetDisposition_Click()
If FileReadOnly Or EditDispute Then
   Exit Sub
End If

If BHproject Then
 Call SetDisposition(0)
 GoTo A:
End If


If Not PrivSetDisposition Then
MsgBox ("You do not have permission to enter a disposition, see your Manager")
Exit Sub
End If

FCDis = True
CheckCancelDisposition = True

Dim cost As Currency, Update As Boolean
'On Error GoTo Err_cmdSetDisposition_Click

If Not CurrentProject.AllForms("case list").IsLoaded Then
  'Removed by JE 07-14-2014
  'Dim rstLocks As Recordset
  'Set rstLocks = CurrentDb.OpenRecordset("select * from locksarchive", dbOpenDynaset, dbSeeChanges)
  'With rstLocks
  '.AddNew
  '!FileNumber = FileNumber
  '!StaffID = GetStaffID()
  '!Type = "X"
  '.Update
  'End With
  'Added by JE 07-14-2014
  Dim str_SQL As String
  str_SQL = "INSERT INTO LocksArchive(FileNumber,StaffID,[TimeStamp],[Type]) VALUES (" & FileNumber & "," & GetStaffID() & ",'" & Now() & "','X')"
  Debug.Print str_SQL
  RunSQL (str_SQL)
  MsgBox "You cannot set a disposition without the case window open.  Please re-open the file", vbCritical
  Exit Sub
End If


If IsNull(Disposition) And PrivSetDisposition Then
    
    Call SetDisposition(0)
    If Not CheckCancelDisposition Then
    Exit Sub
    End If
    
    If Sale > Date Then     ' if the sale is in the future then try to remove it from the shared calendar
        If Not IsNull(SelectedDispositionID) And Not IsNull(SaleCalendarEntryID) Then
            Call DeleteCalendarEvent(SaleCalendarEntryID)
            Dim rstFCdetail As Recordset
            Set rstFCdetail = CurrentDb.OpenRecordset("Select * From FCdetails Where FileNumber=" & Me.FileNumber & " AND FCdetails.Current= True", dbOpenDynaset, dbSeeChanges)
            With rstFCdetail
            rstFCdetail.Edit
            rstFCdetail!SaleCalendarEntryID = Null
            rstFCdetail.Update
            End With
            Set rstFCdetail = Nothing
        End If
    End If
  
End If

'Milestone Billing for Referral Fee

Dim InvPct As Double
Dim cbxClient As Integer
cbxClient = Nz(DLookup("ClientID", "Caselist", "FileNumber=" & FileNumber))
Select Case Nz(DLookup("State", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
  Case "VA"
   Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
    Case 1 'Conventional
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
    Case 2 'VA or Veteran's Affairs
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
    Case 3 'FHA or HUD
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=570")) 'HUD/FHA
    Case 4
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=177")) 'Fannie Mae
    Case 5
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=263")) 'Freddie Mac
    Case Else
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
    End Select
    If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
      InvPct = DLookup("VASalepct", "clientlist", "clientid=" & cbxClient)
    Else
      InvPct = 1
    End If
    If FeeAmount > 0 Then
      If InvPct <= 1 Then
        AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due at sale received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      Else
        'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
      End If
    End If
  Case "MD"
    Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
      Case 1 'Conventional
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & cbxClient))
      Case 2 'VA or Veteran's Affairs
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
      Case 3 'FHA or HUD
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=570")) 'HUD/FHA
      Case 4
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=177")) 'Fannie Mae
      Case 5
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=263")) 'Freddie Mac
      Case Else
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & cbxClient))
    End Select
    If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
      InvPct = DLookup("MDSalepct", "clientlist", "clientid=" & cbxClient)
    Else
      InvPct = 1
    End If
    If FeeAmount > 0 Then
      If InvPct <= 1 Then
        AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due at sale received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      Else
        'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
      End If
    End If
    Case "DC"
      Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
        Case 1 'Conventional
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=" & cbxClient))
        Case 2 'VA or Veteran's Affairs
          FeeAmount = Nz(DLookup("FeeDcReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
        Case 3 'FHA or HUD
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=570")) 'HUD/FHA
        Case 4
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=177")) 'Fannie Mae
        Case 5
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=263")) 'Freddie Mac
        Case Else
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=" & cbxClient))
      End Select
      If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
        InvPct = DLookup("DCSalepct", "clientlist", "clientid=" & cbxClient)
      Else
        InvPct = 1
      End If
      If FeeAmount > 0 Then
        If InvPct <= 1 Then
          AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due at sale received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
        Else
          'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
        End If
      End If
  End Select
''1/5/15 SA
A: Forms![Case List].Requery
Forms![foreclosuredetails].Requery


'1/31/15 SA Add Tracking disp



'DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_cmdSetDisposition_Click:
    Exit Sub

Err_cmdSetDisposition_Click:
    MsgBox Err.Description
    Resume Exit_cmdSetDisposition_Click
    
End Sub

Private Sub SetDisposition(DispositionID As Long)
Dim StatusText As String, FeeAmount As Currency, cost As Currency, Jurisdiction As Long, Update As Boolean, ctr As Integer, rstNames As Recordset
Dim txtType As String
 txtType = "FC"
 If IsLoaded("Case List") = True Then
 If Forms![Case List]!CaseTypeID = 8 Then txtType = "Mo"
 End If
 
If DispositionID = 0 Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    SelectedDispositionID = 0
    DoCmd.OpenForm "SetDisposition", , , , , acDialog, txtType
Else
'1/5/15 SA
 DoCmd.SetWarnings False
 DoCmd.RunSQL ("update FCdetails set Disposition = " & DispositionID & " where [FileNumber] = " & Me.FileNumber & " and current=true")
 DoCmd.SetWarnings True
'    Disposition = DispositionID
End If


'12/18/14
If SelectedDispositionID = 8 Then

DoCmd.SetWarnings False

Dim strtext As String
strtext = InputBox("Please enter the reason for cancel disposition")

strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strtext & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery


End If

DoCmd.SetWarnings True


If SelectedDispositionID > 0 Then   ' if it was actually set

    Dim disComplete As Boolean
    disComplete = DLookup("[Completed]", "FCDisposition", "[ID] = " & SelectedDispositionID)
    If (disComplete = True And (IsNull(Me!BidAmount) And IsNull(Me!BidReceived))) Then
      MsgBox "Cannot enter sales results until bid received and bid amount is completed.", vbCritical, "Set Disposition"
      Exit Sub
    End If
    '1/5/15 SA
    cmdClose.SetFocus
    cmdSetDisposition.Enabled = False ' don't allow any changes
    sfrmFCDIL!DILSentRecord.Enabled = False
  
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("update FCdetails set Disposition = " & SelectedDispositionID & " where [FileNumber] = " & Me.FileNumber & " and current=true")
 
  '  Disposition = SelectedDispositionID
  
    Forms!foreclosuredetails!Disposition.Requery
    
    'Disposition.Requery
    If MsgBox("Is the disposition date today?", vbYesNo, "Disposition Date Entry") = vbYes Then
    '1/5/15 SA
  
    DoCmd.RunSQL ("update FCdetails set DispositionDate = Date() where [FileNumber] = " & Me.FileNumber & " and current=true")
    DoCmd.SetWarnings True

    'DispositionDate = Date
    Else
    DoCmd.RunSQL ("update FCdetails set DispositionDate = #" & Format(InputBox("Please enter disposition date", "Disposition Date Entry"), "mm/dd/yyyy") & "# where [FileNumber] = " & Me.FileNumber & " and current=true")
    'DispositionDate = InputBox("Please enter disposition date", "Disposition Date Entry")
    End If
    Call DispositionDate_AfterUpdate
    
    If StaffID = 0 Then Call GetLoginName
    
    '1/5/15 SA
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("update FCdetails set DispositionStaffID = " & StaffID & " where [FileNumber] = " & Me.FileNumber & " and current=true")
    DoCmd.SetWarnings True
    'DispositionStaffID = StaffID
   ' DoCmd.RunCommand acCmdSaveRecord
    
    'DispositionDesc.Requery
    'DispositionInitials.Requery
    Jurisdiction = Forms![Case List]!JurisdictionID
    
    
    
'Varginal post sale SA 11/12/2014
If State = "VA" And disComplete = True Then
  If IsNull(SaleConductedTrusteeID) Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    SelectedTrusteeID = 0
    DoCmd.OpenForm "SetTrusteeVA", , , , , acDialog
  End If

  If SelectedTrusteeID > 0 Then
      DoCmd.SetWarnings False
      DoCmd.RunSQL ("update FCdetails set SaleConductedTrusteeID = " & SelectedTrusteeID & " where [FileNumber] = " & Me.FileNumber & " and current=true")
      DoCmd.SetWarnings True
'    SaleConductedTrusteeID = SelectedTrusteeID
    txtTrusteeConductedSale.Requery
    DoCmd.RunCommand acCmdSaveRecord
  End If
End If


'DC post sale SA 7/6/2015

If State = "DC" And disComplete = True Then
  If IsNull(SaleConductedTrusteeID) Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    SelectedTrusteeID = 0
    DoCmd.OpenForm "SetTrustee", , , , , acDialog
  End If

  If SelectedTrusteeID > 0 Then
    DoCmd.SetWarnings False
      DoCmd.RunSQL ("update FCdetails set SaleConductedTrusteeID = " & SelectedTrusteeID & " where [FileNumber] = " & Me.FileNumber & " and current=true")
      DoCmd.SetWarnings True
'    SaleConductedTrusteeID = SelectedTrusteeID
    txtTrusteeConductedSale.Requery
    DoCmd.RunCommand acCmdSaveRecord
  End If
End If
  

   
'MARYLAND POST SALE FEES/COSTS
If (State = "MD") And disComplete = True Then
  If IsNull(SaleConductedTrusteeID) Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    SelectedTrusteeID = 0
    DoCmd.OpenForm "SetTrustee", , , , , acDialog
  End If

  If SelectedTrusteeID > 0 Then
    DoCmd.SetWarnings False
      DoCmd.RunSQL ("update FCdetails set SaleConductedTrusteeID = " & SelectedTrusteeID & " where [FileNumber] = " & Me.FileNumber & " and current=true")
      DoCmd.SetWarnings True
'    SaleConductedTrusteeID = SelectedTrusteeID
    txtTrusteeConductedSale.Requery
    DoCmd.RunCommand acCmdSaveRecord
  End If

If SelectedDispositionID = 1 Then

  If Not IsNull(SalePrice) Then
  FeeAmount = SalePrice * 0.01
  If FeeAmount > 0 Then
  AddInvoiceItem FileNumber, "FC-COMM", "Trustees Commission", FeeAmount, 0, False, True, False, False
  End If
  Else
  MsgBox (" Please add Sale Price ")
  End If
  

'Transfer/Registration fees/costs/taxes for buy ins
AddInvoiceItem FileNumber, "FC-Tax", "Recordation Tax", Int(SalePrice / 500 + 0.999) * DLookup("RecordationTaxMD", "JurisdictionList", "jurisdictionid=" & Jurisdiction), 187, False, False, False, True
AddInvoiceItem FileNumber, "FC-Tax", "State Transfer Tax", (SalePrice * DLookup("StateTransferTaxMD", "JurisdictionList", "jurisdictionid=" & Jurisdiction)) / 100, 187, False, False, False, True
If Jurisdiction <> 9 Then 'Cecil county is flat $10 per deed
AddInvoiceItem FileNumber, "FC-Tax", "County Transfer Tax", (SalePrice * DLookup("CountyTransferTaxMD", "JurisdictionList", "jurisdictionid=" & Jurisdiction)) / 100, 187, False, False, False, True
Else
AddInvoiceItem FileNumber, "FC-Tax", "County Transfer Tax", 10, 187, False, False, False, True
End If

AddInvoiceItem FileNumber, "FC-Tax", "Deed Abstractor Recording Cost", DMax("DeedRecordingCost", "JurisdictionAbstractorDeed", "jurisdictionid=" & Jurisdiction), 0, False, False, False, True

AddInvoiceItem FileNumber, "FC-Tax", "Deed Recording Overnight Cost", 8, 77, False, False, False, True
AddInvoiceItem FileNumber, "FC-Tax", "Deed Recording Filing Fee", 60, 187, False, False, False, True
AddInvoiceItem FileNumber, "FC-Reg", "State Property Registration Attorney Fee", DLookup("RegistrationState", "clientlist", "clientid=" & Forms![Case List]!ClientID), 0, True, True, False, False
AddInvoiceItem FileNumber, "FC-Reg", "State Property Registration Filing Fee", DLookup("ivalue", "db", "id=44"), 187, False, False, False, False
AddInvoiceItem FileNumber, "FC-Reg", "County Property Attorney Fee", DLookup("Registrationcounty", "clientlist", "clientid=" & Forms![Case List]!ClientID), 0, True, True, False, False
AddInvoiceItem FileNumber, "FC-Reg", "Postage for Property Registration", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, False, False, False
'AddInvoiceItem FileNumber, "Trustee-Comm", "Postage for Property Registration", Nz(DLookUp("Value","StandardCharges","ID=" & 1)), 76, False, False, False, False



'HUD title policy
If CaseTypeID = 2 Or CaseTypeID = 3 Then
AddInvoiceItem FileNumber, "FC-Title", "Title Policy", 0, 72, False, False, False, True
End If

'calc 1% Trustee commission
                If Not IsNull(SalePrice) And LoanType = 1 Then
                    If DLookup("TrusteeComm", "JurisdictionList", "JurisdictionID=" & Jurisdiction) = True Then
                        FeeAmount = SalePrice * 0.01
                        If FeeAmount > 0 Then
                            AddInvoiceItem FileNumber, "FC-COMM", "Trustees Commission", FeeAmount, 0, False, True, False, False
                        End If
                    End If
                End If

'Order Lien Cert for buy ins
If DLookup("LienCert", "JurisdictionList", "jurisdictionid=" & Jurisdiction) = True Then
  
  If Not IsNull(LienCert) Then Update = True
  '    LienCert = Date  ' This is was the wrong 05/13 and Diane requested removed it on 05/14 SA.
            
        Select Case Jurisdiction
        Case 4 'Balto City
        If Update = False Then
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=4")
        Else
        cost = DLookup("UpdateCost", "LienCertCosts", "jurisdictionid=4")
        End If
        Case 5 'Balto County
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=5")
        Case 14 'Harford
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=14")
        Case 15 ' Howard
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=15")
        Case 3 'Anne Arundel
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=3")
        Case 8 'Carroll
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=8")
        Case 12 'Frederick
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=12")
        Case 10 'Charles
        cost = DLookup("OrderCost", "LienCertCosts", "jurisdictionid=10")
        Case Else
        MsgBox "Lien cert cost not found for this jurisdiction!  Please see your manager to have this added and notify accounting", vbExclamation
        End Select

    
        If LoanType = 1 Or LoanType = 4 Or LoanType = 5 Then
        AddInvoiceItem FileNumber, "FC-LIEN", "Lien Certificate Ordered", cost, 189, False, False, False, True
        AddInvoiceItem FileNumber, "FC-Lien", "Lien Cert Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, False, False, False
        Else
        AddInvoiceItem FileNumber, "FC-LIEN", "Lien Certificate Ordered", cost * 2, 189, False, False, False, True
        AddInvoiceItem FileNumber, "FC-Lien", "Lien Cert Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)) * 2, 76, False, False, False, False
        End If
 '   Call PrintLienCerts(FileNumber, Jurisdiction, Update) 'As per Diane request
End If
End If


'BuyIn/3rd party
If SelectedDispositionID = 1 Or SelectedDispositionID = 2 Then

  If SelectedDispositionID = 2 Then
  If Not IsNull(SalePrice) Then
  FeeAmount = SalePrice * 0.05
  If FeeAmount > 0 Then
  AddInvoiceItem FileNumber, "FC-COMM", "Trustees Commission", FeeAmount, 0, False, True, False, False
  End If
  Else
  MsgBox (" Please add Sale Price ")
  End If
  End If


 'Title update for buy in/3rd party
    If SelectedDispositionID < 3 And (DispositionDate - TitleThru) > 30 Then
    AddInvoiceItem FileNumber, "FC-Title", "Title update- post sale", 55, 0, False, False, False, True
        If Jurisdiction = 4 Or Jurisdiction = 18 Then
        AddInvoiceItem FileNumber, "FC-Title", "Judgment update- post sale", 20, 0, False, False, False, True
        End If
    End If
 
 'JPM SCRA  - **recode, needs to be per party
  If Forms![Case List]!ClientID = 97 Then
    Set rstNames = CurrentDb.OpenRecordset("SELECT Names.FileNumber, Names.mortgagor, Names.First, Names.Last, Names.SSN FROM [Names] GROUP BY Names.FileNumber, Names.mortgagor, Names.First, Names.Last, Names.SSN HAVING (((Names.FileNumber)=" & FileNumber & ") AND ((Names.mortgagor)=Yes) AND ((Names.SSN) Is Not Null)) OR (((Names.FileNumber)=" & FileNumber & ") AND ((Names.mortgagor)=Yes) AND ((Names.SSN)<>""999999999""))", dbOpenDynaset, dbSeeChanges)
        With rstNames
        .MoveLast
        ctr = .RecordCount
        .MoveFirst
        End With
    cost = ctr * DLookup("ivalue", "db", "ID=" & 32)
    AddInvoiceItem FileNumber, "FC-DOD", "DOD Search- JPM Post Sale", cost, 0, True, True, False, False
    AddInvoiceItem FileNumber, "FC-DOD", "DOD Search- JPM Deed", cost, 0, True, True, False, False
  End If
 
    'Auditor Fee
    AddInvoiceItem FileNumber, "FC-Tax", "Auditor Fee", DMax("AuditorFee", "JurisdictionAuditorFee", "jurisdictionid=" & Jurisdiction), 0, False, False, False, True
    AddInvoiceItem FileNumber, "FC-Aud", "Auditor Package", 8, 77, False, False, False, True
    'add NIsi costs by jurisdiction
    AddInvoiceItem FileNumber, "FC-RPT", "Report of Sale Overnight Cost", 8, 77, False, False, False, True
    
    If Jurisdiction <> 4 And Jurisdiction <> 18 Then 'No PG or Balt City Nisi cert mailing
    AddInvoiceItem FileNumber, "FC-NiSi", "NiSi Cert Overnight Cost", 8, 77, False, False, False, True
    AddInvoiceItem FileNumber, "FC-NiSi", "NiSi Ad costs", Nz(DLookup("NisiOrder", "JurisdictionList", "jurisdictionID=" & Jurisdiction), 1), 0, False, False, False, False
    End If
    If Jurisdiction = 10 Or Jurisdiction = 20 Then 'St Mary's and charles county Nisi ad mailing
    AddInvoiceItem FileNumber, "FC-NiSi", "NiSi Ad Overnight Cost", 8, 77, False, False, False, True
    End If
End If

'3rd Party
'If Not IsNull(SalePrice) And Disposition = 2 Then
If Not IsNull(SalePrice) And SelectedDispositionID = 2 Then
'Rider Bond calc
    If (SalePrice - 25000) <= 150000 Then
    FeeAmount = (SalePrice - 25000) * 0.00365
    Else
    'FeeAmount = (150000 * 0.00365) + (SalePrice - 25000 - 150000) * 0.003
    FeeAmount = (SalePrice - 25000) * 0.003
    End If
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "FC-BND", "Rider Bond Issued", FeeAmount, 0, False, True, False, False
    End If
    
'calc 5% Trustee commission
    If Not IsNull(SalePrice) And LoanType = 1 Then
        If DLookup("TrusteeComm", "JurisdictionList", "JurisdictionID=" & Jurisdiction) = True Then
            FeeAmount = SalePrice * 0.05
            If FeeAmount > 0 Then
                AddInvoiceItem FileNumber, "FC-COMM", "Trustees Commission (less all attorney fees and co-counsel fees)*", FeeAmount, 0, False, True, False, False
            End If
        End If
    End If

End If

End If

'Motion to Dismiss
If SelectedDispositionID > 2 And Not IsNull(Docket) Then
AddInvoiceItem FileNumber, "FC-Motion", "Motion to Dismiss Filing Fee", 15, 187, False, False, False, True
End If
''ssss
    StatusText = Nz(DLookup("StatusInfo", "FCDisposition", "ID =" & DLookup("Disposition", "Fcdetails", "[FileNumber] = " & Me.FileNumber & " and current=True")))
    If StatusText <> "" Then AddStatus FileNumber, Now(), StatusText
    If SelectedDispositionID = 6 Then     ' Bankruptcy
        If MsgBox("Do you want to change this file to Bankruptcy?", vbYesNo + vbDefaultButton2) = vbYes Then
            '1/5/15 sa
            DoCmd.SetWarnings False
            DoCmd.RunSQL ("update CaseList set CaseTypeID = " & 2 & " where [FileNumber] = " & Me.FileNumber)
            DoCmd.SetWarnings True

           'CaseTypeID = 2
        End If
    End If
    If CaseTypeID = 8 And Nz(SaleCompleted) = 0 Then    ' Monitor Sale cancelled
        If MsgBox("Do you want to change this file to Foreclosure?", vbYesNo + vbDefaultButton2) = vbYes Then
         'sarabalani      5/12/15    CaseTypeID = 1
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("UPDATE CaseList set CaseTypeID = 1 WHERE [FileNumber] = " & Me.FileNumber)
        DoCmd.SetWarnings True
        
 
        End If
    End If
 
'Virginia
    
If (State = "VA") Then
     Dim varFeeAmt As Variant
    Dim currAmt As Currency
    
    If Disposition = 1 Then
    
    'Deed costs for buy ins
    If LoanType = 4 Or LoanType = 5 Then
    AddInvoiceItem FileNumber, "FC-Tax", "Deed Recordation", 23, 187, False, False, False, True
    Else
    AddInvoiceItem FileNumber, "FC-Tax", "Deed Recordation", 43, 187, False, False, False, True
    End If
    AddInvoiceItem FileNumber, "FC-Deed", "Record Deed- Overnight Costs", 7, 77, False, False, False, False
    AddInvoiceItem FileNumber, "FC-Deed", "Receipt of Recorded Deed Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, True, False, False
                
                'calc 1% Trustee commission
                If Not IsNull(SalePrice) And LoanType = 1 Then
                    If DLookup("TrusteeComm", "JurisdictionList", "JurisdictionID=" & Jurisdiction) = True Then
                        FeeAmount = SalePrice * 0.01
                        If FeeAmount > 0 Then
                            AddInvoiceItem FileNumber, "FC-COMM", "Trustees Commission to Commonwealth Trustees", FeeAmount, 0, False, True, False, False
                        End If
                    End If
                End If
    End If

    If SelectedDispositionID < 3 Then 'Buyin and 3rd party fees/costs
        'JPM SCRA  - **recode, needs to be per party
  If Forms![Case List]!ClientID = 97 Then
    Set rstNames = CurrentDb.OpenRecordset("SELECT Names.FileNumber, Names.mortgagor, Names.First, Names.Last, Names.SSN FROM [Names] GROUP BY Names.FileNumber, Names.mortgagor, Names.First, Names.Last, Names.SSN HAVING (((Names.FileNumber)=" & FileNumber & ") AND ((Names.mortgagor)=Yes) AND ((Names.SSN) Is Not Null)) OR (((Names.FileNumber)=" & FileNumber & ") AND ((Names.mortgagor)=Yes) AND ((Names.SSN)<>""999999999""))", dbOpenDynaset, dbSeeChanges)
        With rstNames
        .MoveLast
        ctr = .RecordCount
        .MoveFirst
        End With
    cost = ctr * DLookup("ivalue", "db", "ID=" & 32)
    AddInvoiceItem FileNumber, "FC-DOD", "DOD Search- JPM Post Sale", cost, 0, True, True, False, False
    AddInvoiceItem FileNumber, "FC-DOD", "DOD Search- JPM Deed", cost, 0, True, True, False, False
  End If
        'Title Update
        AddInvoiceItem FileNumber, "FC-Title", "Title update- post sale", 55, 0, False, False, False, True
        'Auditor Fee
             Dim VAAuditorFee As Currency
            If (IsNull(SalePrice)) Then
              MsgBox "Sales price is missing. Auditor Fee was not added to the billing sheet.", vbCritical
            Else
              VAAuditorFee = CalculateVAAuditorFee(SalePrice)
              AddInvoiceItem FileNumber, "FC-AUD", "Commissioner of Accounts", VAAuditorFee, 188, False, False, False, True
            End If
            
        'LNA cost
        If (DocBackLostNote = True) Then
               AddInvoiceItem FileNumber, "FC-AUD", "LNA to Commissioner ", DLookup("costsLostnoteaudit", "clientlist", "clientid=" & Forms![Case List]!ClientID), 188, False, False, False, False
        End If
            
        'DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter overnight costs to the Commissioner|FC-AUD|Overnight Costs to Commissioner"
        AddInvoiceItem FileNumber, "FC-AUD", "Return mail postage for fee audit", 1.5, 76, False, True, False, False  'set as $1.50 true up later
        'Prompt for Liens/Property taxes
        If MsgBox("Are there liens or real property taxes to enter?") = vbYes Then
        DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter Total Amount of Liens & Property Taxes Paid|FC-TAX|Liens and Property Taxes|"
        End If
    End If
End If

'3rd party
If SelectedDispositionID = 2 Then
'calc 5% Trustee commission
            If Not IsNull(SalePrice) And LoanType = 1 Then
                If DLookup("TrusteeComm", "JurisdictionList", "JurisdictionID=" & Jurisdiction) = True Then
                    FeeAmount = SalePrice * 0.05
                    If FeeAmount > 0 Then
                        AddInvoiceItem FileNumber, "FC-COMM", "Trustees Commission to Commonwealth Trustees (less all attorney fees and co-counsel fees)*", FeeAmount, 0, False, True, False, False
                    End If
                End If
            End If
End If

  Call Visuals
End If


If SelectedDispositionID > 0 Then
    Dim StrDis As String
    Dim clientShor As String
    Dim strInsert As String
    StrDis = DLookup("Disposition", "FCDisposition", "ID= " & SelectedDispositionID)
    clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)
    DoCmd.SetWarnings False
    strInsert = "Insert Into Tracking_FCDispositions (CaseFile,ProjectName,ClientShortName,State,Disp,Client,DIT,StaffID,StaffName) Values (" & FileNumber & ",'" & Forms![Case List]!PrimaryDefName & "','" & clientShor & "','" & Forms![foreclosuredetails]!State & "','" & StrDis & "', " & Forms![Case List]!ClientID & ", #" & Now() & "#," & GetStaffID & ",'" & GetFullName() & "')"
    DoCmd.RunSQL strInsert
    DoCmd.SetWarnings True
End If



End Sub

Private Sub SetLMDisposition(DispositionID As Long)
Dim StatusText As String, FeeAmount As Currency

If DispositionID = 0 Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    
    SelectedDispositionID = 0
    DoCmd.OpenForm "SetDisposition", , , , , acDialog, "LM"
Else
    Disposition = DispositionID
End If

If SelectedDispositionID > 0 Then   ' if it was actually set
    If DateValue(sfrmLMHearing!Hearing) = Date Then
        AddInvoiceItem FileNumber, "FC-MED", "Mediation Fee- Appearance made", DLookup("MediationFee", "clientlist", "clientid=" & Forms![Case List]!ClientID), 0, True, True, False, True
    End If
    cmdClose.SetFocus
    cmdSetLMDisposition.Enabled = False ' don't allow any changes
    
    LMDisposition = SelectedDispositionID
    LMDisposition.Requery
    LMDispositionDate = Date
    AddStatus FileNumber, Date, "Add LM Disposition Date"
    
    If StaffID = 0 Then Call GetLoginName
    
    LMDispositionStaffID = StaffID
    DoCmd.RunCommand acCmdSaveRecord
    
    LMDispositionDesc.Requery
    
    StatusText = Nz(DLookup("Disposition", "LMDisposition", "LMDispositionID=" & LMDisposition))
    If StatusText <> "" Then AddStatus FileNumber, Now(), StatusText
    
    Call Visuals
    DoCmd.OpenForm "SetLMTrustee", , , , , acDialog
    
End If

'If SelectedDispositionID > 0 Then    ""

'Adds who conducted Mediation to the text box
'If SelectedLMTrusteeID > 0 Then
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL "INSERT INTO LMTrustees(Filenumber, LMConductedTrusteeID) VALUES(" & Me.FileNumber & "," & SelectedLMTrusteeID & ");"
'    DoCmd.SetWarnings True
'    Me.txtTrusteeConductedLM.Requery
'    DoCmd.RunCommand acCmdSaveRecord
'End If
End Sub




Private Sub UpdateCalendar()

Dim emailGroup As String
'If IsNull(SaleCalendarEntryID) Then Exit Sub 'sarab

If Nz(SaleCalendarEntryID) = "X" Then Exit Sub

If IsNull(Sale) And Not IsNull(SaleCalendarEntryID) Then
    Call DeleteCalendarEvent(SaleCalendarEntryID)
    SaleCalendarEntryID = Null
    Exit Sub
End If
'DLookup("ShortClientName", "qryClientAddress")

'getTrustees(Forms!ForeclosureDetails!lstTrustees, Forms!ForeclosureDetails!lstTrustees)

If Forms![Case List]!CaseTypeID = 8 Then
    emailGroup = "SharedCalRecipMO"
    
    If Nz(SaleCalendarEntryID) = "" Then   ' new event on calendar
    SaleCalendarEntryID = AddCalendarEvent(Sale + Nz(SaleTime, 0), IsNull(SaleTime), getTrustees(Forms!foreclosuredetails!lstTrustees, Forms!foreclosuredetails!lstTrustees) & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ") SR SALE", Forms![Case List]!JurisdictionID.Column(1), 9, emailGroup)
    Else                                    ' change existing event on calendar
        Call UpdateCalendarEvent(SaleCalendarEntryID, Sale + Nz(SaleTime, 0), IsNull(SaleTime), getTrustees(Forms!foreclosuredetails!lstTrustees, Forms!foreclosuredetails!lstTrustees) & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ") SR SALE", Forms![Case List]!JurisdictionID.Column(1), 9)
    End If


Else
    Select Case State
    
    Case "MD"
    emailGroup = "SharedCalRecipFC-MD"
    Case "DC"
    emailGroup = "SharedCalRecipFC-DC"
    Case "VA"
    emailGroup = "SharedCalRecipFC-VA"
    Case Else
    emailGroup = "SharedCalRecip"
    End Select


    If Nz(SaleCalendarEntryID) = "" Then   ' new event on calendar
        SaleCalendarEntryID = AddCalendarEvent(Sale + Nz(SaleTime, 0), IsNull(SaleTime), DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")", Forms![Case List]!JurisdictionID.Column(1), 1, emailGroup)
    Else                                    ' change existing event on calendar
        Call UpdateCalendarEvent(SaleCalendarEntryID, Sale + Nz(SaleTime, 0), IsNull(SaleTime), DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")", Forms![Case List]!JurisdictionID.Column(1), 1)
    End If
End If

End Sub

Private Sub DeleteFutureHearings(pDispositionDate As Date)

Dim rstHearings As Recordset
Dim i As Integer

Set rstHearings = CurrentDb.OpenRecordset("SELECT * FROM LMHearings WHERE FileNumber=" & Me!FileNumber & " AND datediff('d',#" & pDispositionDate & "#, Hearing) > 0", dbOpenDynaset, dbSeeChanges)

With rstHearings

  If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
  
      If (Not IsNull(rstHearings![HearingCalendarEntryID])) Then
      
          DeleteCalendarEvent (rstHearings![HearingCalendarEntryID])
      End If
  
      .Delete
      .MoveNext
    Loop
  End If
  
  .Close
End With

Me.sfrmLMHearing.Requery


End Sub

Private Sub cmdShowCalendarData_Click()

On Error GoTo Err_cmdShowCalendarData_Click
If Len(Nz(SaleCalendarEntryID)) > 2 Then Call CalendarEventInfo(SaleCalendarEntryID)

Exit_cmdShowCalendarData_Click:
    Exit Sub

Err_cmdShowCalendarData_Click:
    MsgBox Err.Description
    Resume Exit_cmdShowCalendarData_Click
    
End Sub

Private Sub ZipCode_AfterUpdate()
  FetchZipCodeCityState ZipCode, Me.City, Me.State
End Sub



Private Sub UpdateCalendarExceptionHearing()

Dim emailGroup As String

If Nz(ExceptionsHearingEntryID) = "X" Then Exit Sub

If IsNull(ExceptionsHearing) And Not IsNull(ExceptionsHearingEntryID) Then
    Call DeleteCalendarEvent(ExceptionsHearingEntryID)
    ExceptionsHearingEntryID = Null
    Exit Sub
End If

Select Case State
Case "MD"
emailGroup = "SharedCalRecipFC-MD"
Case "DC"
emailGroup = "SharedCalRecipFC-DC"
Case "VA"
emailGroup = "SharedCalRecipFC-VA"
Case Else
emailGroup = "SharedCalRecip"
End Select

If Nz(ExceptionsHearingEntryID) = "" Then   ' new event on calendar
    ExceptionsHearingEntryID = AddCalendarEvent(ExceptionsHearing + Nz(ExceptionsHearingTime, 0), IsNull(ExceptionsHearingTime), DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " Exception Hearing " & " (" & FileNumber & ")", Forms![Case List]!JurisdictionID.Column(1), 8, emailGroup)
Else                                    ' change existing event on calendar
    Call UpdateCalendarEvent(ExceptionsHearingEntryID, ExceptionsHearing + Nz(ExceptionsHearingTime, 0), IsNull(ExceptionsHearingTime), DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " Exception Hearing " & " (" & FileNumber & ")", Forms![Case List]!JurisdictionID.Column(1), 8)
End If
    
End Sub


Private Sub UpdateCalendarStatusHearing()

Dim emailGroup As String

If Nz(StatusHearingEntryID) = "X" Then Exit Sub

If IsNull(StatusHearing) And Not IsNull(StatusHearingEntryID) Then
    Call DeleteCalendarEvent(StatusHearingEntryID)
    StatusHearingEntryID = Null
    Exit Sub
End If

Select Case State
Case "MD"
emailGroup = "SharedCalRecipFC-MD"
Case "DC"
emailGroup = "SharedCalRecipFC-DC"
Case "VA"
emailGroup = "SharedCalRecipFC-VA"
Case Else
emailGroup = "SharedCalRecip"
End Select

If Nz(StatusHearingEntryID) = "" And Not IsNull(StatusHearing) And Not IsNull(StatusHearingTime) Then   ' new event on calendar
    StatusHearingEntryID = AddCalendarEvent(StatusHearing + Nz(StatusHearingTime, 0), IsNull(StatusHearingTime), DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " Status Hearing " & " (" & FileNumber & ")", Forms![Case List]!JurisdictionID.Column(1), 8, emailGroup)
Else                                    ' change existing event on calendar
    Call UpdateCalendarEvent(StatusHearingEntryID, StatusHearing + Nz(StatusHearingTime, 0), IsNull(StatusHearingTime), DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " Status Hearing " & " (" & FileNumber & ")", Forms![Case List]!JurisdictionID.Column(1), 8)
End If
    
End Sub
Private Sub UpdateCalendarDC()

Dim emailGroup As String
'If IsNull(SaleCalendarEntryID) Then Exit Sub 'sarab

If Nz(SaleCalendarEntryID) = "X" Then Exit Sub

If IsNull(txtSale) And Not IsNull(SaleCalendarEntryID) Then
    Call DeleteCalendarEvent(SaleCalendarEntryID)
    SaleCalendarEntryID = Null
    Exit Sub
End If
'DLookup("ShortClientName", "qryClientAddress")

'getTrustees(Forms!ForeclosureDetails!lstTrustees, Forms!ForeclosureDetails!lstTrustees)

If Forms![Case List]!CaseTypeID = 8 Then
   ' emailGroup = "SharedCalRecipMO"
    
    If Nz(SaleCalendarEntryID) = "" Then   ' new event on calendar
    SaleCalendarEntryID = AddCalendarEvent(txtSale + Nz(txtSaleTime, 0), IsNull(txtSaleTime), getTrustees(Forms!foreclosuredetails!lstTrustees, Forms!foreclosuredetails!lstTrustees) & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ") SR SALE", Forms![Case List]!JurisdictionID.Column(1), 9, emailGroup)
Else                                    ' change existing event on calendar
        Call UpdateCalendarEvent(SaleCalendarEntryID, txtSale + Nz(txtSaleTime, 0), IsNull(txtSaleTime), getTrustees(Forms!foreclosuredetails!lstTrustees, Forms!foreclosuredetails!lstTrustees) & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ") SR SALE", Forms![Case List]!JurisdictionID.Column(1), 9)
End If


Else
    Select Case State
    
    Case "MD"
    emailGroup = "SharedCalRecipFC-MD"
    Case "DC"
    emailGroup = "SharedCalRecipFC-DC"
    Case "VA"
    emailGroup = "SharedCalRecipFC-VA"
    Case Else
    emailGroup = "SharedCalRecip"
    End Select


    If Nz(SaleCalendarEntryID) = "" Then   ' new event on calendar
        SaleCalendarEntryID = AddCalendarEvent(txtSale + Nz(txtSaleTime, 0), IsNull(txtSaleTime), DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")", Forms![Case List]!JurisdictionID.Column(1), 1, emailGroup)
    Else                                    ' change existing event on calendar
        Call UpdateCalendarEvent(SaleCalendarEntryID, txtSale + Nz(txtSaleTime, 0), IsNull(txtSaleTime), DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")", Forms![Case List]!JurisdictionID.Column(1), 1)
    End If
End If

End Sub



Private Sub PHproejctFC()
Me.LoanNumber.Locked = False
Me.FairDebt.Locked = False
Me.AccelerationIssued.Locked = False
Me.txtClientSentAcceleration.Locked = False
Me.NOI.Locked = False
Me.txtClientSentNOI.Locked = False
Me.HUDOccLetter.Enabled = True
Me.VAAppraisal.Enabled = True
Me.FirstLegal.Enabled = True
Me.StatementOfDebtAmount.Enabled = True
Me.LostNoteAffSent.Locked = False
Me.LostNoteNotice.Locked = False
Me.TitleBack.Locked = False
Me.TitleThru.Locked = False
Me.SentToDocket.Enabled = True
Me.LienCert.Locked = False
Me.FLMASenttoCourt.Locked = False
Me.ServiceSent.Locked = False
Me.BorrowerServed.Locked = False
Me.ServiceMailed.Locked = False
Me.IRSNotice.Enabled = True
Me.Notices.Enabled = True
Me.Notices.Locked = False
Me.Sale.Enabled = True
Me.SaleTime.Enabled = True
Me.SaleTime.Locked = False
Me.SaleSet.Locked = False
Me.SaleSet.Enabled = True
Me.BorrowerServed.Enabled = True
Me.BorrowerServed.Locked = False
Me.Deposit.Locked = False
Me.Report.Enabled = True
Me.NiSiEnd.Enabled = True
Me.SaleRat.Enabled = True
Me.FinalPkg.Enabled = True
Me.Audit2File.Enabled = True
Me.Audit2Rat.Enabled = True
Me.SalePrice.Locked = False
Me.Purchaser.Locked = False
Me.PurchaserAddress.Locked = False
Me.StatusHearing.Enabled = True
Me.StatusHearingTime.Enabled = True
Me.ResellMotion.Enabled = True
Me.ResellServed.Enabled = True
Me.ResellShowCauseExpires.Enabled = True
Me.ResellAnswered.Enabled = True
Me.ResellGranted.Enabled = True
Me.DismissalSent.Enabled = True
Me.DismissalSent.Locked = False
Me.VAAppraisal.Enabled = True
Me.VAAppraisal.Locked = False
Me.DispositionRescinded.Enabled = True
Me.DispositionRescinded.Locked = False
Me.VAAppraisal.Enabled = True
Me.ClientPaid.Enabled = True
Me.AmmDocBackSOD.Enabled = True
Me.AmmStatementOfDebtDate.Enabled = True
Me.AmmStatementOfDebtAmount.Enabled = True
Me.AmmStatementOfDebtPerDiem.Enabled = True
Me.cmdSetDisposition.Enabled = True


End Sub

