VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TitleResolutionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim varFeeAmt As Variant

Private Sub cmdAddTR_Click()
  MsgBox "Adding Title Resolution under construction.", vbCritical, "Add Title Resolution"
  Exit Sub
  
End Sub

Private Sub ComplaintFiled_AfterUpdate()
AddStatus FileNumber, ComplaintFiled, "Complaint Filed"
End Sub

Private Sub ComplaintFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ComplaintFiled)
End Sub

Private Sub ComplaintFiled_DblClick(Cancel As Integer)
ComplaintFiled = Date
AddStatus FileNumber, ComplaintFiled, "Complaint Filed"
End Sub

Private Sub Current_AfterUpdate()
Current.Locked = Current.Value
End Sub



Private Sub Visuals()
Dim X As Boolean, lt As Long


X = (UCase$(Nz(State)) = "MD" Or UCase$(Nz(State)) = "VA")
cmdAudit.Enabled = X
AssessedValue.Enabled = (UCase$(Nz(State)) = "VA")


If IsNull(LoanType) Then
    lt = 0
Else
    lt = LoanType
End If
FHALoanNumber.Enabled = (lt = 2 Or lt = 3)    ' enable for VA or HUD

GroundRentAmount.Enabled = IIf(optLeasehold = 1, -1, 0) ' Changed from optLeasehold because of DC project Ticket 866 SA 06/3

GroundRentPayable.Enabled = IIf(optLeasehold = 1, -1, 0) ' Changed from optLeasehold because of DC project Ticket 866 SA 06/3




End Sub


Private Sub Form_Current()
Dim tr As Recordset

If FileReadOnly Or EditDispute Then
    Me.AllowEdits = False
    cmdAddTR.Enabled = False
    cmdAudit.Enabled = False
    cmdPrint.Enabled = False
    sfrmNamesTR.Form.AllowEdits = False
    sfrmNamesTR.Form.AllowAdditions = False
    sfrmNamesTR.Form.AllowDeletions = False
    cmdCalcPerDiem.Enabled = False
    sfrmStatus.Form.AllowEdits = False
    sfrmStatus.Form.AllowAdditions = False
    sfrmStatus.Form.AllowDeletions = False
 Else
    Me.AllowEdits = True
    cmdAddTR.Enabled = True
    cmdAudit.Enabled = True
    cmdPrint.Enabled = True
    sfrmNamesTR.Form.AllowEdits = True
    sfrmNamesTR.Form.AllowAdditions = True
    sfrmNamesTR.Form.AllowDeletions = True
    cmdCalcPerDiem.Enabled = True
    sfrmStatus.Form.AllowEdits = True
    sfrmStatus.Form.AllowAdditions = True
    sfrmStatus.Form.AllowDeletions = True
    Detail.BackColor = -2147483633
    lblShowCurrent.BackColor = -2147483633
    
    If (IsNull(Disposition) And PrivSetDisposition) Then
      cmdSetDisposition.Enabled = True
 
    End If
    
    
End If

If Me.NewRecord Then    ' fill in info from previous TR, if any
    Set tr = CurrentDb.OpenRecordset("SELECT * FROM TRDetails WHERE FileNumber = " & FileNumber & " AND Current = True", dbOpenDynaset, dbSeeChanges)
    If Not tr.EOF Then
        NewTR = Date
        Referral = Date
        PrimaryFirstName = tr("PrimaryFirstName")
        PrimaryLastName = tr("PrimaryLastName")
        SecondaryFirstName = tr("SecondaryFirstName")
        SecondaryLastName = tr("SecondaryLastName")
        PropertyAddress = tr("PropertyAddress")
        City = tr("City")
        State = tr("State")
        ZipCode = tr("ZipCode")
        TaxID = tr("TaxID")
        optLeasehold = tr("Leasehold")
        GroundRentAmount = tr("GroundRentAmount")
        GroundRentPayable = tr("GroundRentPayable")
        LegalDescription = tr("LegalDescription")
        Comment = tr("Comment")
        DOT = tr("DOT")
        DOTdate = tr("DOTdate")
        OriginalTrustee = tr("OriginalTrustee")
        OriginalBeneficiary = tr("OriginalBeneficiary")
        Liber = tr("Liber")
        Folio = tr("Folio")
        OriginalMortgagors = tr("OriginalMortgagors")
        OriginalPBal = tr("OriginalPBal")
        RemainingPBal = tr("RemainingPBal")
        LoanNumber = tr("LoanNumber")
        LoanType = tr("LoanType")
        LienPosition = tr("LienPosition")
        FHALoanNumber = tr("FHALoanNumber")
        AbstractorCaseNumber = tr("AbstractorCaseNumber")
        CourtCaseNumber = tr("CourtCaseNumber")
        FairDebtAmount = tr("FairDebtAmount")
        
        TitleReviewTo = tr("TitleReviewTo")
        TitleReviewOf = tr("TitleReviewOf")
        TitleReviewFax = tr("TitleReviewFax")
        TitleReviewNameOf = tr("TitleReviewNameOf")
        TitleReviewLiens = tr("TitleReviewLiens")
        TitleReviewJudgments = tr("TitleReviewJudgments")
        TitleReviewTaxes = tr("TitleReviewTaxes")
        TitleReviewStatus = tr("TitleReviewStatus")
        TitleClaimSent = tr("TitleClaimSent")
        TitleClaimResolved = tr("TitleClaimResolved")
    
       
        DOTrecorded = tr!DOTrecorded
       
       
        NewTR = Now()
        If StaffID = 0 Then Call GetLoginName
        NewTRBy = StaffID
        Do While Not tr.EOF     ' make all previously current records not current
            tr.Edit
            tr("Current") = False
            tr.Update
            tr.MoveNext
        Loop
    End If
    tr.Close
    Me!Current = True           ' and make this record current
End If

Current.Locked = Current.Value
Me.Caption = "Title Resolution " & Me![FileNumber] & " " & [PrimaryDefName]
Call Visuals




If Not IsNull(Disposition) Then
    Disposition.Locked = True
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

Private Sub cmdPrint_Click()

On Error GoTo Err_cmdPrint_Click

MsgBox "Title Resolution Print under construction.", vbCritical, "Title Resolution Print"
Exit Sub



If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm "TitleResolutionPrint", , , "[CaseList].[FileNumber]=" & Me![FileNumber]

Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrint_Click
    
End Sub

Private Sub cmdPurchaserInvestor_Click()

On Error GoTo Err_cmdPurchaserInvestor_Click
Me!Purchaser = Investor
Me!PurchaserAddress = InvestorAddress

Exit_cmdPurchaserInvestor_Click:
    Exit Sub

Err_cmdPurchaserInvestor_Click:
    MsgBox Err.Description
    Resume Exit_cmdPurchaserInvestor_Click
    
End Sub

Private Sub Frame218_AfterUpdate()
Call Visuals
End Sub



Private Sub Form_Open(Cancel As Integer)

  If IsNull(LoanNumber) Then
    LoanNumber.Locked = False
    LoanNumber.BackStyle = 1
    'Call SetObjectAttributes(LoanNumber, True)
  Else  ' this allows for copying
    LoanNumber.Locked = True
    LoanNumber.BackStyle = 0
    'Call SetObjectAttributes(LoanNumber, False)
  End If
  

End Sub

Private Sub InitNegNotSucceed_AfterUpdate()
AddStatus FileNumber, InitNegNotSucceed, "Initial Negotiations Unsuccessful"
End Sub

Private Sub InitNegNotSucceed_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(InitNegNotSucceed)
End Sub

Private Sub InitNegNotSucceed_DblClick(Cancel As Integer)
InitNegNotSucceed = Date
AddStatus FileNumber, InitNegNotSucceed, "Initial Negotiations Unsuccessful"
End Sub

Private Sub InterestRate_AfterUpdate()
If InterestRate > 1 Then InterestRate = InterestRate / 100#
End Sub

Private Sub IntroLetterToTCSent_AfterUpdate()
AddStatus FileNumber, IntroLetterToTCSent, "Intro Letter to TC Sent"
End Sub

Private Sub IntroLetterToTCSent_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(IntroLetterToTCSent)
End Sub

Private Sub IntroLetterToTCSent_DblClick(Cancel As Integer)
IntroLetterToTCSent = Date
AddStatus FileNumber, IntroLetterToTCSent, "Intro Letter to TC Sent"
End Sub

Private Sub LoanType_AfterUpdate()
Call Visuals
End Sub

Private Sub optLeasehold_Click()
Call Visuals
End Sub

Private Sub Referral_AfterUpdate()
AddStatus FileNumber, Referral, "Referral Received to TC"
End Sub

Private Sub Referral_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Referral)
End Sub

Private Sub Referral_DblClick(Cancel As Integer)
Referral = Date
AddStatus FileNumber, Referral, "Referral Received to TC"
End Sub


Private Sub ShowCurrent_Click()
If ShowCurrent Then
    Me.Filter = "TRDetails.FileNumber = " & Me![FileNumber] & "AND Current = True"
Else
    Me.Filter = "TRDetails.FileNumber = " & Me![FileNumber]
End If
End Sub

Private Sub State_AfterUpdate()
Call Visuals
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

Private Sub cmdAddFC_Click()
On Error GoTo Err_cmdAddFC_Click

MsgBox "New Title Resolution under construction.", vbCritical, "New Title Resolution"
Exit Sub


' Code lifted from Foreclosure - needs to be customized for Title Resolution
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

MsgBox "Title Resolution audit under construction. ", vbCritical, "Title Resolution Audit"
Exit Sub

' code lifted from foreclosures - needs to be updated for Title Resolution

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
PerDiem = RemainingPBal * InterestRate / 365

Exit_cmdCalcPerDiem_Click:
    Exit Sub

Err_cmdCalcPerDiem_Click:
    MsgBox Err.Description
    Resume Exit_cmdCalcPerDiem_Click
    
End Sub

Private Sub SetDisposition(DispositionID As Long)
Dim StatusText As String, FeeAmount As Currency


If DispositionID = 0 Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    SelectedDispositionID = 0
    DoCmd.OpenForm "SetDisposition", , , , , acDialog, "TR"
Else
    Disposition = DispositionID
End If

If SelectedDispositionID > 0 Then   ' if it was actually set


    cmdClose.SetFocus
    cmdSetDisposition.Enabled = False ' don't allow any changes
    
    Disposition = SelectedDispositionID
    Disposition.Requery
    DispositionDate = Date
    
    If StaffID = 0 Then Call GetLoginName
    DispositionStaffID = StaffID
    DoCmd.RunCommand acCmdSaveRecord
    
    DispositionDesc.Requery
    DispositionInitials.Requery
    
    StatusText = Nz(DLookup("StatusInfo", "TRDisposition", "ID=" & Disposition))
    If StatusText <> "" Then AddStatus FileNumber, Now(), StatusText
    
    
    
    Call Visuals
End If

End Sub

Private Sub cmdSetDisposition_Click()

On Error GoTo Err_cmdSetDisposition_Click

If IsNull(Disposition) And PrivSetDisposition Then

' check to ensure there's a charge record before proceeding
  varFeeAmt = DLookup("[Amount]", "[Fees]", "[State] = '" & Me.State & "' and [ClientID] = " & Forms![Case List]!ClientID & " and [FeeType] = 'TR-COMP'")
  If (IsNull(varFeeAmt)) Then
    MsgBox "Ensure a fee is in the fee table for this client, state and type of Title Resolution Complete before continuing.", vbCritical, "Missing Fees"
    Exit Sub
  End If



    Call SetDisposition(0)
End If

Exit_cmdSetDisposition_Click:
    Exit Sub

Err_cmdSetDisposition_Click:
    MsgBox Err.Description
    Resume Exit_cmdSetDisposition_Click
    
End Sub

Private Sub TitleClaimResolved_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleClaimResolved)
End Sub

Private Sub TitleClaimSent_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleClaimSent)
End Sub

Private Sub TitleClearForDIL_B_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleClearForDIL_B)
End Sub

Private Sub TitleReviewToClient_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleReviewToClient)
End Sub
