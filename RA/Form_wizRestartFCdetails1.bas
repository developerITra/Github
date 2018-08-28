VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_wizRestartFCdetails1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub AcceptRestart_Click()
'    Forms!wizRestartCaseList1.cmdDetails.Visible = False
    Me.AllowEdits = True
    sfrmNames.Form.AllowEdits = True
    sfrmNames.Form.AllowAdditions = True
    sfrmNames.Form.AllowDeletions = True
    sfrmNames!cmdCopyClient.Enabled = True
    sfrmNames!cmdCopy.Enabled = True
    sfrmNames!cmdTenant.Enabled = True
    sfrmNames!cmdDelete.Enabled = True
    sfrmNames!cmdNoNotice.Enabled = True
    cmdCalcPerDiem.Enabled = True
    cmdPurchaserInvestor.Enabled = True
    Detail.BackColor = -2147483633
    
    Forms!wizRestartCaseList1.AllowEdits = True
    'tabCase
   
    
    Forms!wizRestartCaseList1.Detail.BackColor = -2147483633
    
    cmdAddtoQueue.Visible = True
    cmdReturn.Visible = False
    cmdCancel.Visible = False
    cmdAddtoQueue.SetFocus
    AcceptRestart.Visible = False
    cmdRejected.Visible = True
    cmdRejected.SetFocus
  
 
  
End Sub

Private Sub cmdAddtoQueue_Click()

Dim rstFCdetails As Recordset, rstCase As Recordset, Reason As Long, FileNum As Long, FC As Recordset, rsFCDIL As Recordset, ctr As Integer, rstNames As Recordset, cost As Currency, rstJnl As Recordset
Dim rstqueue As Recordset

Set rstCase = CurrentDb.OpenRecordset("select * from caselist where filenumber = " & FileNumber, dbOpenDynaset)
'Set rstqueue = CurrentDb.OpenRecordset("select RestartReasonRestartQueue from wizardqueuestats where filenumber = " & FileNumber, dbOpenDynaset)
Set rstFCdetails = CurrentDb.OpenRecordset("SELECT DispositionRescinded FROM FCDetails WHERE FileNumber = " & FileNumber & " AND Current = True", dbOpenDynaset, dbSeeChanges)


If IsLoadedF("queRSIReview") Then
    If Not LockFile(FileNumber) Then
    MsgBox "This file is currently locked and cannot be processed", vbCritical
    Exit Sub
    End If
    
    Set rstFCdetails = CurrentDb.OpenRecordset("SELECT * FROM FCDetails WHERE FileNumber = " & FileNumber & " AND Current = True", dbOpenDynaset, dbSeeChanges)
        If Not IsNull(Forms!wizRestartFCdetails1!DispositionDesc) Then
        Dim textdispostion As String
        textdispostion = Forms!wizRestartFCdetails1!DispositionDesc
        With rstFCdetails
        .Edit
        !Disposition = Null
        !DispositionDate = Null
        !DispositionStaffID = Null
        .Update
        End With
        End If
    rstFCdetails.Close
   
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
     With rstqueue
    .Edit
    !RestartRSIreviewDateIn = Now
    !RestartRSIreviewUser = StaffID
    !RestartRSIreviewReason = Null
    If Not IsNull(rstqueue!RestartWaiting) Then rstqueue!RestartWaiting = Null
    If Not IsNull(rstqueue!RestartComplete) Then rstqueue!RestartComplete = Null
    .Update
    End With
    Set rstqueue = Nothing
    
    
    
    
      
    DoCmd.SetWarnings False
    strinfo = "File sent to Restart queue from Restart RSI Mgr queue" & IIf(Not IsNull(textdispostion), " And removed the dispostion; " & textdispostion, "")
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    
    Dim xx As Boolean
    xx = False
    If Forms!wizRestartFCdetails1.Dirty = True Then Forms!wizRestartFCdetails1.Dirty = False
    
    'Forms!wizRestartFCdetails1.Requery
    
    GoTo continueWithRestarFromReivewQ

DoCmd.Close acForm, "wizrestartfcdetails1"
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "Journal"
Call ReleaseFile(FileNumber)

Exit Sub
End If







If Not LockFile(FileNumber) Then
MsgBox "This file is currently locked and cannot be processed", vbCritical
Exit Sub
End If

Select Case Disposition & ""
Case 1
If CaseTypeID <> 11 Then
    If IsNull(rstFCdetails!DispositionRescinded) Then
    Reason = 2
    Call RestartRSICompletionUpdate(FileNumber, Reason)
    MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it is a Buy In with no rescinded date."
    rstFCdetails.Close
    GoTo Exit_Proc
    End If
End If
Case 2
If CaseTypeID <> 11 Then
    If IsNull(rstFCdetails!DispositionRescinded) Then
    Reason = 2
    Call RestartRSICompletionUpdate(FileNumber, Reason)
    MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it is a 3rd Party with no rescinded date."
    rstFCdetails.Close
    GoTo Exit_Proc
    End If
End If
Case ""
If CaseTypeID <> 11 Then
    Reason = 3
    Call RestartRSICompletionUpdate(FileNumber, Reason)
    MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it has no disposition."
    rstFCdetails.Close
    GoTo Exit_Proc
End If
Case 7
If CaseTypeID <> 11 Then
    Reason = 8
    Call RestartRSICompletionUpdate(FileNumber, Reason)
    MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it has Deed In Lieu."
    rstFCdetails.Close
    GoTo Exit_Proc
End If
Case 6
If CaseTypeID = 11 Then
    If CheckIfFileWasFCFirst(FileNumber) = True Then
    Reason = 9
        Call RestartRSICompletionUpdate(FileNumber, Reason)
        MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it is Peniding type and has BK disposition."
        rstFCdetails.Close
        GoTo Exit_Proc
    End If
End If

Case 6
If CaseTypeID = 1 Then
   ' If CheckIfFileWasFCFirst(FileNumber) = True Then
    Reason = 11
        Call RestartRSICompletionUpdate(FileNumber, Reason)
        MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it is BK disposition."
        rstFCdetails.Close
        GoTo Exit_Proc
End If







End Select

Set rstJnl = CurrentDb.OpenRecordset("select * from Journal where filenumber=" & FileNumber & " and warning=400", dbOpenDynaset, dbSeeChanges)
If Not rstJnl.EOF Then
Reason = 6
    Call RestartRSICompletionUpdate(FileNumber, Reason)
    MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it has a Stop attribute."
    rstJnl.Close
    GoTo Exit_Proc
End If
rstJnl.Close

Select Case CaseTypeID
Case 2
Reason = 1
    Call RestartRSICompletionUpdate(FileNumber, Reason)
    MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it is an active Bankruptcy."
    rstFCdetails.Close
Case 5
    Reason = 4
    Call RestartRSICompletionUpdate(FileNumber, Reason)
MsgBox "File " & FileNumber & "File has been added to the RSI Review Queue and may not be restarted at this time b/c of litigation"
    GoTo Exit_Proc
Case 10
    Reason = 5
    Call RestartRSICompletionUpdate(FileNumber, Reason)
MsgBox "File " & FileNumber & "File has been added to the RSI Review Queue and may not be restarted at this time b/c of title resolution"
    GoTo Exit_Proc
End Select
    
    
continueWithRestarFromReivewQ:
    
    With rstCase
    .Edit
    !Active = True
    !OnStatusReport = True
    !RestartReceived = Now
    .Update
    End With
If MsgBox("Are you sure you want to add another Foreclosure?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub

'Charge for SCRA check
Set rstNames = CurrentDb.OpenRecordset("select * from Names where filenumber=" & FileNumber & " and noteholder=yes", dbOpenDynaset, dbSeeChanges)

With rstNames
If Not .EOF Then
.MoveLast
ctr = .RecordCount
.MoveFirst
End If
.Close
cost = ctr * DLookup("ivalue", "db", "ID=" & 32)
End With

'change from PND to FC
If CaseTypeID = 11 Then
Dim rstBKdetails As Recordset, rstTrustees As Recordset, AttyID As Integer, rstBKAtty As Recordset
Dim AttyFirst As String, AttyLast As String, AttyFirm As String, AttyAddress As String, AttyPhone As String, AttyCity As String, AttyState As String, AttyZip As String, Platform As String
Dim TrusteeID As String, TrusteePhone As String, TrusteeFirst As String, TrusteeLast As String, TrusteeAddress As String, TrusteeAddress2 As String, TrusteeCity As String, TrusteeState As String, TrusteeZip As String

Set rstBKdetails = CurrentDb.OpenRecordset("select * from BKdetails where Filenumber=" & FileNumber & "And Current=True", dbOpenDynaset, dbSeeChanges)

On Error Resume Next

With rstBKdetails
If .RecordCount > 0 Then
AttyID = !AttorneyID
TrusteeID = !Trustee
End If
End With
rstBKdetails.Close

If TrusteeID <> "" Then
Set rstTrustees = CurrentDb.OpenRecordset("select * from BKTrustees where ID=" & TrusteeID, dbOpenDynaset, dbSeeChanges)
On Error Resume Next
With rstTrustees
TrusteeFirst = !First
TrusteeLast = !Last
TrusteeAddress = !Address
If Not IsNull(!TrusteeAddress2) Then
TrusteeAddress2 = !Address2
End If
TrusteeCity = !City
TrusteeState = !State
TrusteeZip = !Zip
End With
rstTrustees.Close
End If
On Error GoTo 0

If AttyID > 0 Then
Set rstNames = CurrentDb.OpenRecordset("select * from Names", dbOpenDynaset, dbSeeChanges)
Set rstBKAtty = CurrentDb.OpenRecordset("select * from BKAttorneys where attyid=" & AttyID, dbOpenDynaset, dbSeeChanges)
On Error Resume Next
With rstNames
.AddNew
!FileNumber = FileNumber
!First = rstBKAtty!FirstName
!Last = rstBKAtty!LastName
!Company = rstBKAtty!AttorneyFirm
!Address = rstBKAtty!Address
!City = rstBKAtty!City
!State = rstBKAtty!State
!Zip = rstBKAtty!Zip
!NoticeType = 10
.Update
End With
rstNames.Close
End If

If TrusteeID <> "" Then
Set rstNames = CurrentDb.OpenRecordset("select * from Names", dbOpenDynaset, dbSeeChanges)
On Error Resume Next
With rstNames
.AddNew
!FileNumber = FileNumber
!Company = "BK Trustee"
!First = TrusteeFirst
!Last = TrusteeLast
!Address = TrusteeAddress
!Address2 = TrusteeAddress2
!City = TrusteeCity
!State = TrusteeState
!Zip = TrusteeZip
!NoticeType = 14
.Update
End With
rstNames.Close
End If

Call AddStatus(FileNumber, Now(), "File type changed to FC")
End If

'Emulate Add FC here
If (Nz(Disposition) = 2) Or (Nz(Disposition) = 1) Then
    If PrivAdmin Then
        If MsgBox("The property has already been sold! Are you sure you want to add another Foreclosure?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
    Else
        MsgBox "You cannot add another Foreclosure because the property has already been sold.  (Management can override this for you.)", vbCritical
        Exit Sub
    End If
End If

'Convert these to rst
Dim rstCaseList As Recordset
Set rstCaseList = CurrentDb.OpenRecordset("SELECT * FROM CaseList WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstCaseList
.Edit
If CaseTypeID = 11 Then
!CaseTypeID = 1
End If
!ReferralDate = Date
!ReferralDocsReceived = Null
!RestartReceived = Null
.Update
End With
rstCaseList.Close


Call AddStatus(FileNumber, Now(), "Referral Date")
FileNum = FileNumber
Set FC = CurrentDb.OpenRecordset("SELECT * FROM FCDetails WHERE FileNumber = " & FileNum & " AND Current = True", dbOpenDynaset, dbSeeChanges)
Me.AllowAdditions = True
DoCmd.GoToRecord , , acNewRec

Me.AllowEdits = True

    sfrmNames.Form.AllowEdits = True
    sfrmNames.Form.AllowAdditions = True
    sfrmNames.Form.AllowDeletions = True
    sfrmNames!cmdCopyClient.Enabled = True
    sfrmNames!cmdCopy.Enabled = True
    sfrmNames!cmdTenant.Enabled = True
    sfrmNames!cmdDelete.Enabled = True
    sfrmNames!cmdNoNotice.Enabled = True

    If Not FC.EOF Then

        FileNumber = FileNum
        NewFC = Date
        ReferralDay = Date
        PrimaryFirstName = FC("PrimaryFirstName")
        PrimaryLastName = FC("PrimaryLastName")
        SecondaryFirstName = FC("SecondaryFirstName")
        SecondaryLastName = FC("SecondaryLastName")
        PropertyAddress = FC("PropertyAddress")
        [Fair Debt] = FC("[fair debt]")
        City = FC("City")
        State = FC("State")
        ZipCode = FC("ZipCode")
        TaxID = FC("TaxID")
        optLeasehold = FC("Leasehold")
        GroundRentAmount = FC("GroundRentAmount")
        GroundRentPayable = FC("GroundRentPayable")
        LegalDescription = FC("LegalDescription")
        Status = FC("Status")
        ShortLegal = FC("ShortLegal")
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
        'added TitleClaimDate on 1/13/15 Linda
        TitleClaimDate = FC("TitleClaimDate")
       
        If State <> "DC" Then Docket = FC("Docket")
        'added on 10_30_15 for DC restar file
        If State = "DC" Then
                FirstPub = Null
                ReviewAdProof = Null
                SaleSet = Null
        End If
        'TitleReviewNameOf = FC("TitleReviewNameOf")
        'TitleReviewLiens = FC("TitleReviewLiens")
        'TitleReviewJudgments = FC("TitleReviewJudgments")
        'TitleReviewTaxes = FC("TitleReviewTaxes")
        'TitleReviewStatus = FC("TitleReviewStatus")
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
        AccelerationLetter = FC!AccelerationLetter
        DocBackLostNote = FC!DocBackLostNote
        DocBackOrigNote = FC!DocBackOrigNote
        DocBackNoteOwnership = FC!DocBackNoteOwnership
        DocBackAff7105 = FC!DocBackAff7105
        DocBackMilAff = FC!DocBackMilAff
        DocBackDOA = FC!DocBackDOA
        DocBackSOD = FC!DocBackSOD
        DocBackLossMitPrelim = FC!DocBackLossMitPrelim
        DocBackLossMitFinal = FC!DocBackLossMitFinal
        VAAppraisal = FC!VAAppraisal
       
        
        FHALoanNumber = FC!FHALoanNumber
        FHLMCLoanNumber = FC!FHLMCLoanNumber
        FNMALoanNumber = FC!FNMALoanNumber
        FLMASenttoCourt = FC!FLMASenttoCourt
        LossMitFinalDate = FC!LossMitFinalDate
        MedCaseNumber = FC("MedCaseNumber")
        MedRequestDate = FC!MedRequestDate
        MedReqDocDate = FC!MedReqDocDate
        MedRecDocDate = FC!MedRecDocDate
        MedDocSentDate = FC!MedDocSentDate
        MedHearingLocation = FC!MedHearingLocation
        LMDisposition = FC!LMDisposition
        LMDispositionStaffID = FC!LMDispositionStaffID
        MedHearingResults = FC!MedHearingResults
        LMDispositionDate = FC!LMDispositionDate
        MedHearingLocation = FC!MedHearingLocation
        
        UpdatedNotices = FC("UpdatedNotices") ' as per Diane request on 4/29/2014 ticket no. 808
        AccelerationIssued = FC("AccelerationIssued")
        txtClientSentAcceleration = FC("ClientSentAcceleration")
        txtClientSentNOI = FC("ClientSentNOI")
        ServiceMailed = FC("ServiceMailed")
        
        'added on 4_14_15 Lin
        ExceptionsHearing = FC!ExceptionsHearing
        ExceptionsHearingTime = FC!ExceptionsHearingTime
        ExceptionsStatus = FC!ExceptionsStatus
        
        StatusHearing = FC!StatusHearing
        StatusHearingTime = FC!StatusHearingTime
        StatusResults = FC!StatusResults
        
        ExceptionsHearingEntryID = FC!ExceptionsHearingEntryID
        StatusHearingEntryID = FC!StatusHearingEntryID

        
        
        Current = True
        If FC!Disposition = 3 Or FC!Disposition = 4 Then
'        Dim lrs As Recordset,
        Dim rstwizqueue As Recordset
'      Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'      With lrs
'      .AddNew
'      ![FileNumber] = FileNumber
'      ![JournalDate] = Now
'      ![Who] = GetFullName()
'      ![Warning] = 100
'      ![Info] = "Fair Debt Letter needed for reinstatement/payoff" & vbCrLf
'      ![Color] = 1
'      .Update
'      .Close
'      End With
      DoCmd.SetWarnings False
strinfo = "Fair Debt Letter needed for reinstatement/payoff" & vbCrLf
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Warning,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),100,'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
      Set rstwizqueue = CurrentDb.OpenRecordset("select * from wizardqueuestats where filenumber=" & FileNumber & " and current = true", dbOpenDynaset, dbSeeChanges)
      With rstwizqueue
      .Edit
      !FairDebtComplete = Null
      !FairDebtRestart = Date
      .Update
      .Close
      End With
        Else
        FairDebt = FC!FairDebt
        End If
        NewFC = Now()
        If StaffID = 0 Then Call GetLoginName
        NewFCBy = StaffID
        FC.Close
'        If Me.Dirty Then
'        DoCmd.RunCommand acCmdSaveRecord
'        End If
        DoCmd.SetWarnings (False)
        Set FC = CurrentDb.OpenRecordset("SELECT Current FROM FCDetails WHERE FileNumber = " & FileNum & " AND Current = True", dbOpenDynaset, dbSeeChanges)
            ' make all previously current records not current
            With FC
            .Edit
            !Current = False
            .Update
            .Close
            End With
        DoCmd.SetWarnings (True)
    End If

    'Me!Current = True
    
'for DC restar file on 10_30_15
If State = "DC" Then
    Set rsFCDIL = CurrentDb.OpenRecordset("SELECT * FROM FCDIL WHERE FileNumber = " & FileNumber, dbOpenDynaset, dbSeeChanges)
    
    'WHERE FileNumber = " & FileNumber"
    
        If Not rsFCDIL.EOF Then
            rsFCDIL.Edit
            rsFCDIL!ServiceCancelled = Null
            rsFCDIL!LineStayingcase = Null
        rsFCDIL.Update
        rsFCDIL.Close
       Set rsFCDIL = Nothing
       End If
End If

   
' Restart Atty fee pulls from Restart Fee Approval button under Fee Requests tab only. 2/2/2015
   Select Case State
   Case "VA"
   'Select Case LoanType
    'Case 4
    'FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=177"))
    'Case 5
    'FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=263"))
    'Case Else
    'FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & Forms![wizRestartCaseList1]!ClientID))
    'End Select
            'If FeeAmount > 0 Then
               ' AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", FeeAmount, 0, True, True, False, False
            'Else
                'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", 1, 0, True, True, False, False
            'End If
    
    Case "MD"
    
    If Me.ClientID = 531 Then
    AddInvoiceItem FileNum, "FC-MD-M&T", "Usage costs", 100, 0, False, True, False, True
    End If
        
    'Select Case LoanType
    'Case 4
    'FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=177"))
    'Case 5
    'FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=263"))
    'case Else
    'FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & Forms![wizRestartCaseList1]!ClientID))
    'End Select

            'If FeeAmount > 0 Then
                'If Forms![wizrestartcaselist1]!ClientID = 446 Then
                'AddInvoiceItem FileNum, "FC-REF", "Attorney Fee - 30% due when Affidavits sent of " & FeeAmount & " Total", FeeAmount * 0.3, True, True, False, False
                'AddInvoiceItem FileNum, "FC-REF", "Attorney Fee - 50% due when Docketed of " & FeeAmount & " Total", FeeAmount * 0.2, True, True, False, False
                'AddInvoiceItem FileNum, "FC-REF", "Attorney Fee - remaining 50% due of " & FeeAmount & " Total", FeeAmount * 0.5, True, True, False, False
                'Else
                'AddInvoiceItem FileNum, "FC-REF", "Attorney Fee", FeeAmount, 0, True, True, False, False
                'End If
            'Else
                'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", 1, 0, True, True, False, False
            'End If
    Case "DC"
        
        'FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=" & Forms![wizRestartCaseList1]!ClientID))
            'If FeeAmount > 0 Then
                'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", FeeAmount, 0, True, True, False, False
            'Else
                'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", 1, 0, True, True, False, False
           ' End If
   End Select

'Add technology fee for LPS/Lenstar/Vendorscape for Freddie/Fannie loans
'MsgBox "Please update internet sources"
'DoCmd.OpenForm "Internet Sources", acDialog, "[FileNumber]=" & Me![FileNumber]
'If DLookup("Lenstar", "internetsites", "filenumber=" & FileNumber) = True And LoanType > 3 Then Platform = "Lenstar"
'If DLookup("Vendorscape", "internetsites", "filenumber=" & FileNumber) = True And LoanType > 3 Then Platform = "Vendorscape"
'If DLookup("LPS_Desktop", "internetsites", "filenumber=" & FileNumber) = True And LoanType > 3 Then Platform = "LPS"
'If MsgBox("Does this file require a new registration on LPS, Lenstar, or Vendorscape?", vbYesNo) = vbYes Then
'Select Case Platform
'Case "LPS"
'AddInvoiceItem FileNum, "FC-Oth", "Technology Fee- LPS", 80, 192, False, True, False, True
'Case "Lenstar"
'AddInvoiceItem FileNum, "FC-Oth", "Technology Fee- Lenstar", 55, 193, False, True, False, True
'Case "Vendorscape"
'AddInvoiceItem FileNum, "FC-Oth", "Technology Fee- Vendorscape", 80, 194, False, True, False, True
'End Select
'End If
            
Dim JnlNote As String
JnlNote = "RSI review complete- file to Restart queue for processing"

'Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'    lrs.AddNew
'    lrs![FileNumber] = FileNumber
'    lrs![JournalDate] = Now
'    lrs![Who] = GetFullName()
'    lrs![Info] = JnlNote & vbCrLf
'    lrs![Color] = 1
'    lrs.Update
'    lrs.Close

DoCmd.SetWarnings False
strinfo = JnlNote & vbCrLf
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

'DoCmd.Close

'Dim rstFCdil As Recordset
'Set rstFCdil = CurrentDb.OpenRecordset("Select CertOfPubField from FCDIL Where FileNumber = " & FileNumber, dbOpenDynaset, dbSeeChanges)
'If Not rstFCdil.EOF Then
'If Not IsNull(rstFCdil!CertOfPubField) Then
'With rstFCdil
'.Edit
'!CertOfPubField = Null
'.Update
'End With
'End If
'End If
'Set rstFCdil = Nothing


'Adding to ValuemRI  05/20

DoCmd.SetWarnings False
Dim typetext As String
If IsLoadedF("queRSIReview") Then
typetext = "Restart Mgr Review"
Else
typetext = "Restart"
End If
Dim Shortclient As String
Dim cbxLoanTypetext As String
Dim NewCaseType As String
Dim ClientIDn As Integer
Dim CaseTypeIDn As String
CaseTypeIDn = DLookup("CaseType", "CaseTypes", "CaseTypeID=" & Forms![wizRestartCaseList1]!CaseTypeID)
ClientIDn = Forms![wizRestartCaseList1]!ClientID
Shortclient = DLookup("ShortClientName", "ClientList", "ClientID=" & Forms![wizRestartCaseList1]!ClientID)
cbxLoanTypetext = DLookup("LoanType", "LoanTypes", "ID=" & LoanType)

NewCaseType = "Insert Into ValumeRSI (FileNumber,ShortClientName,Completiondate,State,Type,IdUser,Username,IdClient,CaseType,Count,LoanType) values (" & FileNum & ",'" & Shortclient & "','" & Now() & "','" & State & "','" & typetext & "'," & StaffID & ",'" & GetFullName() & "'," & ClientIDn & ",'" & CaseTypeIDn & "'," & 1 & ",'" & cbxLoanTypetext & "')"

DoCmd.RunSQL NewCaseType
DoCmd.SetWarnings False

        
Call Restart1CompletionUpdate(FileNum)

MsgBox "File " & FileNum & " has been added to the Restart Queue. Restart Atty fee should be added from Restart Fee Approval button under Fee Requests tab"

'rstFCdetails.Close

'Forms!wizRestartFCdetails1.Refresh
'Forms!wizRestartFCdetails1.Requery
If Forms!wizRestartFCdetails1.Dirty Then
Forms!wizRestartFCdetails1.Dirty = False
End If

DoCmd.Close acForm, "wizrestartcaselist1"
DoCmd.Close acForm, "wizrestartfcdetails1"
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "Journal"


Exit Sub

Exit_Proc:
'DoCmd.Close

'Forms!wizRestartFCdetails1.Refresh
'Forms!wizRestartFCdetails1.Requery

If Forms!wizRestartFCdetails1.Dirty Then
Forms!wizRestartFCdetails1.Dirty = False
End If


DoCmd.Close acForm, "wizrestartfcdetails1"
DoCmd.Close acForm, "wizrestartcaselist1"
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "Journal"

End Sub
Private Sub cmdCancel_Click()
If MsgBox("Are you sure you want to cancel without a journal note entered?", vbYesNo) = vbNo Then
Exit Sub
End If
DoCmd.Close
DoCmd.Close acForm, "wizrestartcaselist1"
DoCmd.Close acForm, "Docswindow"
DoCmd.Close acForm, "journal"
End Sub

Private Sub cmdRejected_Click()

Call RestartRSICompletionUpdateToPutInReviewQ(FileNumber, 7)

MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it is rejected manually."
DoCmd.SetWarnings False
strinfo = "File rejected manualy"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

DoCmd.Close
DoCmd.Close acForm, "wizrestartcaselist1"
DoCmd.Close acForm, "Docswindow"
DoCmd.Close acForm, "journal"

End Sub

Private Sub cmdReturn_Click()
Dim JnlNote As String
On Error GoTo Err_cmdClose_Click

If CurrentProject.AllForms("queRSIReview").IsLoaded Then
    Dim rstqueue As Recordset
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartRSIreviewDateIn = Null
    !RestartRSIreviewUser = GetStaffID
    !RestartRSIreviewReason = Null
    .Update
    End With
    Set rstqueue = Nothing
    DoCmd.Close acForm, "wizrestartcaselist1"
    DoCmd.Close acForm, "Docswindow"
    DoCmd.Close acForm, "journal"
    
    DoCmd.SetWarnings False
    strinfo = "File removed from Restart RSI Mgr"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    DoCmd.Close acForm, "wizrestartfcdetails1"
Else

    JnlNote = "This file has been checked for conflicts via Restart wizard"
'    Dim lrs As Recordset
'    Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'        lrs.AddNew
'        lrs![FileNumber] = FileNumber
'        lrs![JournalDate] = Now
'        lrs![Who] = GetFullName()
'        lrs![Info] = JnlNote & vbCrLf
'        lrs![Color] = 1
'        lrs.Update
'        lrs.Close
DoCmd.SetWarnings False
strinfo = JnlNote & vbCrLf
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
    DoCmd.Close
    DoCmd.Close acForm, "wizrestartcaselist1"


End If

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdSetDisposition_Click()
Dim cost As Currency, Update As Boolean
'On Error GoTo Err_cmdSetDisposition_Click

If Not CurrentProject.AllForms("WizRestartCaseList1").IsLoaded Then
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


'If IsNull(Disposition) And PrivSetDisposition Then
    
    Call SetDispositionWizard(0)
    DoCmd.OpenForm "Journal New Entry", , , , , , FileNumber
    Forms![Journal New Entry]!Info = "The Reason of Adding dispsition is : "
    
    
    If Sale > Date Then     ' if the sale is in the future then try to remove it from the shared calendar
        If Not IsNull(Disposition) And Not IsNull(SaleCalendarEntryID) Then
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
  
'End If

'Milestone Billing for Referral Fee

Dim InvPct As Double
If Disposition = 1 Or Disposition = 2 Then
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
            InvPct = DLookup("MDSalepct", "clientlist", "clientid=" & Forms![Case List]!ClientID)
            If InvPct < 1 Then
            AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due at sale of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
            Else
                AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
            End If
        End If
ElseIf State = "VA" Then
    Select Case LoanType
    Case 4
    FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=177"))
    Case 5
    FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=263"))
    Case Else
    FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
    End Select
       
        If FeeAmount > 0 Then
            InvPct = DLookup("VASalepct", "clientlist", "clientid=" & Forms![Case List]!ClientID)
            If InvPct < 1 Then
            AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due at sale of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
            Else
                AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
            End If
        End If
End If

End If

'DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_cmdSetDisposition_Click:
    Exit Sub

Err_cmdSetDisposition_Click:
    MsgBox Err.Description
    Resume Exit_cmdSetDisposition_Click
End Sub

Private Sub CommEdit_Click()
Dim ctrl As Control
For Each ctrl In Me.sfrmNames.Form.Controls

If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then ' TypeOf ctrl Is CommandButton Then
'If Not ctrl.Locked) Then
ctrl.Locked = False
ctrl.Enabled = True
'Else
'ctrl.Locked = True
End If
'End If
Next
With Me.sfrmNames.Form

.AllowAdditions = True
.AllowEdits = True
.AllowDeletions = True
.cmdCopyClient.Enabled = True
.cmdCopy.Enabled = True
.cmdTenant.Enabled = True
.cmdMERS.Enabled = True
.cmdEnterSSN.Enabled = True
.cmdNoNotice.Enabled = True
.cmdPrintNotice.Enabled = True
.cmdPrintLabel.Enabled = True
.cbxNotice.Enabled = True
.cmdDelete.Enabled = True
.cmdNoNotice.Enabled = True
.cbxNotice.Enabled = True
.cbxNotice.Locked = False




End With
'Exit Sub
End Sub

Private Sub Form_Current()
Dim FC As Recordset, rstEV As Recordset, EVFileNumbers As String

'    Me.AllowEdits = True
'MsgBox ("asdf")
'    sfrmNames.Form.AllowEdits = True
'    sfrmNames.Form.AllowAdditions = True
'    sfrmNames.Form.AllowDeletions = True
'    sfrmNames!cmdCopyClient.Enabled = True
'    sfrmNames!cmdCopy.Enabled = True
'    sfrmNames!cmdTenant.Enabled = True
'    sfrmNames!cmdDelete.Enabled = True
'    sfrmNames!cmdNoNotice.Enabled = True
'
'    cmdCalcPerDiem.Enabled = True
'    cmdPurchaserInvestor.Enabled = True
'
'    Detail.BackColor = -2147483633
'
'Else


    
'End If

      'Referral = Date

'
'Me.Caption = IIf(CaseTypeID = 8, "Monitor ", "") & "Foreclosure File " & Me![FileNumber] & " " & [PrimaryDefName]

If IsNull(Disposition) Then
    lblDisposition.Visible = False
    With SalePrice
        .Enabled = True
        .Locked = False
        .BackStyle = 1
    End With
    If Not IsNull(Notices) Then
        With Notices
            .Enabled = False
            .Locked = True
            .BackStyle = 0
        End With

        With UpdatedNotices
            .Enabled = True
            .Locked = False
        End With
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

    With SalePrice
        .Enabled = False
        .Locked = True
        .BackStyle = 0
    End With
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
    With [567]
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
    With ServiceSent
        .Enabled = False
        .Locked = True
        .BackStyle = 0
    End With
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
        
    End If

End If

If SaleCalendarEntryID = "X" Then
    lblSharedCal1.Caption = "Shared Calendar must be updated manually"
    lblSharedCal2.Caption = "Shared Calendar must be updated manually"
    lblSharedCal1.ForeColor = vbRed
    lblSharedCal2.ForeColor = vbRed
Else

    lblSharedCal1.Caption = "Shared Calendar updates are automatic"
    lblSharedCal2.Caption = "Shared Calendar updates are automatic"
    lblSharedCal1.ForeColor = 10040115
    lblSharedCal2.ForeColor = 10040115
End If

If IsNull(LoanNumber) Then
  LoanNumber.Locked = False
  LoanNumber.BackStyle = 1
  'Call SetObjectAttributes(LoanNumber, True)
Else  ' this allows for copying
  LoanNumber.Locked = True
  LoanNumber.BackStyle = 0
  'Call SetObjectAttributes(LoanNumber, False)
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


txtEvictionBroker = Null
txtEvictionFileNum = Null

End Sub

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub

Private Sub Form_Load()
If Not EditFormRSI Then

    Me.AllowEdits = False
    sfrmNames.Form.AllowEdits = False
    sfrmNames.Form.AllowAdditions = False
    sfrmNames.Form.AllowDeletions = False
    sfrmNames!cmdCopyClient.Enabled = False
    sfrmNames!cmdCopy.Enabled = False
    sfrmNames!cmdTenant.Enabled = False
    sfrmNames!cmdDelete.Enabled = False
    sfrmNames!cmdNoNotice.Enabled = False
    cmdCalcPerDiem.Enabled = False
    cmdPurchaserInvestor.Enabled = False
    Detail.BackColor = 8421631
    
    
Else
    Me.AllowEdits = True
    sfrmNames.Form.AllowEdits = True
    sfrmNames.Form.AllowAdditions = True
    sfrmNames.Form.AllowDeletions = True
    sfrmNames!cmdCopyClient.Enabled = True
    sfrmNames!cmdCopy.Enabled = True
    sfrmNames!cmdTenant.Enabled = True
    sfrmNames!cmdDelete.Enabled = True
    sfrmNames!cmdNoNotice.Enabled = True
    cmdCalcPerDiem.Enabled = True
    cmdPurchaserInvestor.Enabled = True
    Detail.BackColor = -2147483633
    AcceptRestart.Visible = False
    cmdAddtoQueue.Visible = True
   
End If
End Sub


Private Sub SetDispositionWizard(DispositionID As Long)
Dim StatusText As String, FeeAmount As Currency, cost As Currency, Jurisdiction As Long, Update As Boolean, ctr As Integer, rstNames As Recordset

'If DispositionID = 0 Then
  '  If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
  '  SelectedDispositionID = 0
    DoCmd.OpenForm "SetDisposition", , , , , acDialog, "FC"
'Else
    Disposition = DispositionID
'End If

If SelectedDispositionID > 0 Then   ' if it was actually set

    Dim disComplete As Boolean
    disComplete = DLookup("[Completed]", "FCDisposition", "[ID] = " & SelectedDispositionID)
    If (disComplete = True And (IsNull(Me!BidAmount) And IsNull(Me!BidReceived))) Then
      MsgBox "Cannot enter sales results until bid received and bid amount is completed.", vbCritical, "Set Disposition"
      Exit Sub
    End If

    'cmdClose.SetFocus
    'cmdSetDisposition.Enabled = False ' don't allow any changes
    'sfrmFCDIL!DILSentRecord.Enabled = False
    
    Disposition = SelectedDispositionID
    Disposition.Requery
    If MsgBox("Is the disposition date today?", vbYesNo, "Disposition Date Entry") = vbYes Then
    DispositionDate = Date
    Else
    DispositionDate = InputBox("Please enter disposition date", "Disposition Date Entry")
    End If
    Call DispositionWizardDate_AfterUpdate
    
    If StaffID = 0 Then Call GetLoginName
    DispositionStaffID = StaffID
'    DoCmd.RunCommand acCmdSaveRecord
    
    DispositionDesc.Requery
    DispositionInitials.Requery
    Jurisdiction = Forms![wizRestartCaseList1]!JurisdictionID
    
    
    
'VARJINIA POST SALE  SA 11/24/2014
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
  

   
'MARYLAND POST SALE FEES/COSTS
If (State = "MD" Or State = "DC") And disComplete = True Then
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
  
  
  
  
'MARYLAND POST SALE


If Disposition = 1 Then

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
If Disposition = 1 Or Disposition = 2 Then

  If Disposition = 2 Then
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
    If Disposition < 3 And (DispositionDate - TitleThru) > 30 Then
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
If Not IsNull(SalePrice) And Disposition = 2 Then

'Rider Bond calc
    If (SalePrice - 25000) <= 150000 Then
    FeeAmount = (SalePrice - 25000) * 0.00365
    Else
    FeeAmount = (150000 * 0.00365) + (SalePrice - 25000 - 150000) * 0.003
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
If Disposition > 2 And Not IsNull(Docket) Then
AddInvoiceItem FileNumber, "FC-Motion", "Motion to Dismiss Filing Fee", 15, 187, False, False, False, True
End If

    StatusText = Nz(DLookup("StatusInfo", "FCDisposition", "ID=" & Disposition))
    If StatusText <> "" Then AddStatus FileNumber, Now(), StatusText
    
    If Disposition = 6 Then     ' Bankruptcy
        If MsgBox("Do you want to change this file to Bankruptcy?", vbYesNo + vbDefaultButton2) = vbYes Then
            CaseTypeID = 2
        End If
    End If
    If CaseTypeID = 8 And Nz(SaleCompleted) = 0 Then    ' Monitor Sale cancelled
        If MsgBox("Do you want to change this file to Foreclosure?", vbYesNo + vbDefaultButton2) = vbYes Then
            CaseTypeID = 1
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

    If Disposition < 3 Then 'Buyin and 3rd party fees/costs
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
If Disposition = 2 Then
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

'  Call Visuals
End If

End Sub

Private Sub DispositionWizardDate_AfterUpdate()
Dim Reason As Integer

Select Case Disposition

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

Forms![wizRestartCaseList1]!BillCase = True
Forms![wizRestartCaseList1]!BillCaseUpdateUser = GetStaffID()
Forms![wizRestartCaseList1]!BillCaseUpdateDate = Date
Forms![wizRestartCaseList1]![BillCaseUpdateReasonID] = Reason
'Forms![wizRestartCaseList1]!lblBilling.Visible = True
Forms![wizRestartCaseList1].SetFocus
'DoCmd.RunCommand acCmdSaveRecord
'Forms![ForeclosureDetails].SetFocus

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
