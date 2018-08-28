VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditPropertyDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
Me.Undo
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdupdate_Click()
NameJournal = ""
Call makePropertyJournaltext
Dim rstFCdetails As Recordset
Dim strinfo As String
Dim strSQLJournal As String
'With Forms!ForeclosureDetails
'
'   If Not IsNull(Forms!ForeclosureDetails.PrimaryFirstName) Then Forms!ForeclosureDetails.PrimaryFirstName = txtPrimaryFirstName
'     .PrimaryLastName = txtPrimaryLastName
'
'     .SecondaryFirstName = txtSecondaryFirstName
'     .SecondaryLastName = txtSecondaryLastName
'
'     .PropertyAddress = txtPropertyAddress
'     .City = txtCity
'     .State = txtState
'     .ZipCode = txtZipCode
'
'     .CourtCaseNumber = txtCourtCaseNumber
'     .TaxID = txtTaxID
'
'End With





'Call makePropertyJournaltext

    DoCmd.SetWarnings False
    strinfo = NameJournal
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EditPropertyDetails!txtFileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    NameJournal = vbNullString
    

'Dim rstJnl As Recordset
'
'Set rstJnl = CurrentDb.OpenRecordset("select * from journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = txtFileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Property Details changed AFTER title review completed"
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing

'MsgBox "Property Details have been updated"
DoCmd.Close acForm, Me.Name
Forms!foreclosuredetails.Requery
Forms!Journal.Requery
End Sub


Public Function makePropertyJournaltext()
NameJournal = ""

If Nz(PrimaryFirstName) <> Nz(PrimaryFirstName.OldValue) Then
    If IsNull(PrimaryFirstName) And Not IsNull(PrimaryFirstName.OldValue) Then NameJournal = NameJournal + "Removed Primary First Name " & PrimaryFirstName.OldValue & ". "
    If IsNull(PrimaryFirstName.OldValue) And Not IsNull(PrimaryFirstName) Then NameJournal = NameJournal + "Added Primary First Name:" & PrimaryFirstName & ". "
    If Not IsNull(PrimaryFirstName.OldValue) And Not IsNull(PrimaryFirstName) Then NameJournal = NameJournal + "Edit Primary First Name from " & PrimaryFirstName.OldValue & " To " & PrimaryFirstName & ". "
End If

If Nz(PrimaryLastName) <> Nz(PrimaryLastName.OldValue) Then
    If IsNull(PrimaryLastName) And Not IsNull(PrimaryLastName.OldValue) Then NameJournal = NameJournal + "Removed Primary Last Name: " & PrimaryLastName.OldValue & ". "
    If IsNull(PrimaryLastName.OldValue) And Not IsNull(PrimaryLastName) Then NameJournal = NameJournal + "Added Primary Last Name: " & PrimaryLastName & ". "
    If Not IsNull(PrimaryLastName.OldValue) And Not IsNull(PrimaryLastName) Then NameJournal = NameJournal + "Edit Primary Last Name from " & PrimaryLastName.OldValue & " To " & PrimaryLastName & ". "
End If

If Nz(SecondaryFirstName) <> Nz(SecondaryFirstName.OldValue) Then
    If IsNull(SecondaryFirstName) And Not IsNull(SecondaryFirstName.OldValue) Then NameJournal = NameJournal + "Removed Secondary First Name " & SecondaryFirstName.OldValue & ". "
    If IsNull(SecondaryFirstName.OldValue) And Not IsNull(SecondaryFirstName) Then NameJournal = NameJournal + "Added Secondary First Name: " & SecondaryFirstName & ". "
    If Not IsNull(SecondaryFirstName.OldValue) And Not IsNull(SecondaryFirstName) Then NameJournal = NameJournal + "Edit Secondary First Name from " & SecondaryFirstName.OldValue & " To " & SecondaryFirstName & ". "
End If

If Nz(SecondaryLastName) <> Nz(SecondaryLastName.OldValue) Then
    If IsNull(SecondaryLastName) And Not IsNull(SecondaryLastName.OldValue) Then NameJournal = NameJournal + "Removed Secondary Last Name : " & SecondaryLastName.OldValue & ". "
    If IsNull(SecondaryLastName.OldValue) And Not IsNull(SecondaryLastName) Then NameJournal = NameJournal + "Added Secondary Lst Name: " & SecondaryLastName & ". "
    If Not IsNull(SecondaryLastName.OldValue) And Not IsNull(SecondaryLastName) Then NameJournal = NameJournal + "Edit Secondary Last Name " & SecondaryLastName.OldValue & " To " & SecondaryLastName & ". "
End If


If Nz(PropertyAddress) <> Nz(PropertyAddress.OldValue) Then
    If IsNull(PropertyAddress) And Not IsNull(PropertyAddress.OldValue) Then NameJournal = NameJournal + "Removed  Property Address: " & PropertyAddress.OldValue & ". "
    If IsNull(PropertyAddress.OldValue) And Not IsNull(PropertyAddress) Then NameJournal = NameJournal + "Added Property Address: " & PropertyAddress & ". "
    If Not IsNull(PropertyAddress.OldValue) And Not IsNull(PropertyAddress) Then NameJournal = NameJournal + "Edit Property Address " & PropertyAddress.OldValue & " To " & PropertyAddress & ". "
End If

If Nz(Apt) <> Nz(Apt.OldValue) Then
    If IsNull(Apt) And Not IsNull(Apt.OldValue) Then NameJournal = NameJournal + "Removed  Apt/Suit Number : " & Apt.OldValue & ". "
    If IsNull(Apt.OldValue) And Not IsNull(Apt) Then NameJournal = NameJournal + "Added Apt/Suit Number: " & Apt & ". "
    If Not IsNull(Apt.OldValue) And Not IsNull(Apt) Then NameJournal = NameJournal + "Edit Apt/Suit Number " & Apt.OldValue & " To " & Apt & ". "
End If

If Nz(City) <> Nz(City.OldValue) Then
    If IsNull(City) And Not IsNull(City.OldValue) Then NameJournal = NameJournal + "Removed City: " & City.OldValue & ". "
    If IsNull(City.OldValue) And Not IsNull(City) Then NameJournal = NameJournal + "Added City: " & City & ". "
    If Not IsNull(City.OldValue) And Not IsNull(City) Then NameJournal = NameJournal + "Edit City " & City.OldValue & " To " & City & ". "
End If

If Nz(State) <> Nz(State.OldValue) Then
    If IsNull(State) And Not IsNull(State.OldValue) Then NameJournal = NameJournal + "Removed State: " & State.OldValue & ". "
    If IsNull(State.OldValue) And Not IsNull(State) Then NameJournal = NameJournal + "Added State: " & State & ". "
    If Not IsNull(State.OldValue) And Not IsNull(State) Then NameJournal = NameJournal + "Edit State " & State.OldValue & " To " & State & ". "
End If

If Nz(ZipCode) <> Nz(ZipCode.OldValue) Then
    If IsNull(ZipCode) And Not IsNull(ZipCode.OldValue) Then NameJournal = NameJournal + "Removed Zip Code: " & ZipCode.OldValue & ". "
    If IsNull(ZipCode.OldValue) And Not IsNull(ZipCode) Then NameJournal = NameJournal + "Added Zip Code: " & ZipCode & ". "
    If Not IsNull(ZipCode.OldValue) And Not IsNull(ZipCode) Then NameJournal = NameJournal + "Edit Zip Code " & ZipCode.OldValue & " To " & ZipCode & ". "
End If


If Nz(CourtCaseNumber) <> Nz(CourtCaseNumber.OldValue) Then

    If IsNull(CourtCaseNumber) And Not IsNull(CourtCaseNumber.OldValue) Then NameJournal = NameJournal + "Removed Court Case Number: " & CourtCaseNumber.OldValue & ". "
    If Nz(CourtCaseNumber.OldValue) = "" And Nz(CourtCaseNumber) <> "" Then NameJournal = NameJournal + "Added Court Case Number: " & CourtCaseNumber & ". "
    If Nz(CourtCaseNumber.OldValue) <> "" And Nz(CourtCaseNumber) <> "" Then NameJournal = NameJournal + "Edit Court Case Number " & CourtCaseNumber.OldValue & " To " & CourtCaseNumber & ". "

End If
    

If Nz(TaxID) <> Nz(TaxID.OldValue) Then
    If IsNull(TaxID) And Not IsNull(TaxID.OldValue) Then NameJournal = NameJournal + "Removed  TAX ID #: " & TaxID.OldValue & ". "
    If IsNull(TaxID.OldValue) And Not IsNull(TaxID) Then NameJournal = NameJournal + "Added TAX ID #: " & TaxID & ". "
    If Not IsNull(TaxID.OldValue) And Not IsNull(TaxID) Then NameJournal = NameJournal + "Edit TAX ID # " & TaxID.OldValue & " To " & TaxID & ". "
End If
    
If Nz(GroundRentAmount) <> Nz(GroundRentAmount.OldValue) Then
    If IsNull(GroundRentAmount) And Not IsNull(GroundRentAmount.OldValue) Then NameJournal = NameJournal + "Removed Annual Ground Rent: " & GroundRentAmount.OldValue & ". "
    If IsNull(GroundRentAmount.OldValue) And Not IsNull(GroundRentAmount) Then NameJournal = NameJournal + "Added Annual Ground Rent: " & GroundRentAmount & ". "
    If Not IsNull(GroundRentAmount.OldValue) And Not IsNull(GroundRentAmount) Then NameJournal = NameJournal + "Edit Annual Ground Rent: " & GroundRentAmount.OldValue & " To " & GroundRentAmount & ". "
End If

If Nz(GroundRentPayable) <> Nz(GroundRentPayable.OldValue) Then
    If IsNull(GroundRentPayable) And Not IsNull(GroundRentPayable.OldValue) Then NameJournal = NameJournal + "Removed Payable: " & GroundRentPayable.OldValue & ". "
    If IsNull(GroundRentPayable.OldValue) And Not IsNull(GroundRentPayable) Then NameJournal = NameJournal + "Added Payable: " & GroundRentPayable & ". "
    If Not IsNull(GroundRentPayable.OldValue) And Not IsNull(GroundRentPayable) Then NameJournal = NameJournal + "Edit Payable: " & GroundRentPayable.OldValue & " To " & GroundRentPayable & ". "
End If

If Nz(optLeasehold) <> Nz(optLeasehold.OldValue) Then
    Select Case optLeasehold
        Case 0
            NameJournal = NameJournal + "Edit Ownership to Fee Simple. "
        Case 1
            NameJournal = NameJournal + "Edit Ownership to Leasehold. "
        Case 2
            NameJournal = NameJournal + "Edit Ownership to Co-op. "
            'I dont think anyone is using co-op....could be an issue in the future.  9/24/2014 MC
   'This should always be an edit.
    'If Not IsNull(optLeasehold.OldValue) And Not IsNull(optLeasehold) Then NameJournal = NameJournal + "Edit Payable: " & optLeasehold.OldValue & " To " & optLeasehold & ". "
    End Select
End If

makePropertyJournaltext = NameJournal
End Function



Private Sub CourtCaseNumber_AfterUpdate()
If (Nz(CourtCaseNumber)) = "" Then
MsgBox ("You are not allowed to remove the case court number")
Me.Undo
End If

End Sub

'Private Sub RemoveCaseFiled()
'
'With Forms!ForeclosureDetails
''mdeatinn
'!CourtCaseNumber = ""
'!MedCaseNumber = Null
'!MedRequestDate = Null
'!MedRecDocDate = Null
'!MedDocSentDate = Null
'!MedHearingLocation = Null
'!MedHearingResults = Null
'!MedHearingClientContactID = Null,,,"Rosie Admin"
'
''If Not IsNull(Forms!foreclosuredetails.Form!sfrmLMHearing!txtHearing) And Not IsNull(Forms!foreclosuredetails!LMDispositionDesc) Then
''Dim ss As Boolean
''ss = True
''End If
''If ss Then
''Forms!foreclosuredetails.Form!sfrmLMHearing!txtHearing = Null
''Forms!forclosuredetails!LMDispositionDesc = Null
''ss = False
''End If
'
'
'
'' presale
'!HUDOccLetter = Null
'!DocstoClient = Null
'!DocsBack = Null
'!DocBackMilAff = False
'!DocBackSOD = False
'!DocBackLossMitPrelim = False
'!DocBackLossMitFinal = False
'!DocBackAff7105 = False
'!StatementOfDebtDate = Null
'!StatementOfDebtAmount = Null
'!StatementOfDebtPerDiem = Null
'!SentToDocket = Null
'!Docket = Null
'!LienCert = Null
'!FLMASenttoCourt = Null
'!LossMitFinalDate = Null
'!ServiceSent = Null
'!BorrowerServed = Null
'!ServiceMailed = Null
'!FirstPub = Null
'!IRSNotice = Null
'!Notices = Null
'!UpdatedNotices = Null
'!BidReceived = Null
'!BidAmount = Null
'!Sale = Null
'!SaleTime = Null
'!Deposit = Null
'!SaleSet = Null
'!BondNumber = Null
'!BondAmount = Null
'!BondReturned = Null
'!chMannerofService = False
'!ReviewAdProof = Null
'.Form!sfrmCertOfPubFiled!CertOfPubField = Null
'!SaleCert = Null
'
''Post sale tab
'!Report = Null
'!StatePropReg = Null
'!NiSiEnd = Null
'!SaleRat = Null
'!PropReg = Null
'!DeedtoRec = Null
'!DeedtoTitleCo = 0
'!RecordDeed = Null
'!RecordDeedLiber = Null
'!RecordDeedFolio = Null
'!FinalPkg = Null
'!AuditFile = Null
'!AuditRat = Null
'!Audit2File = Null
'!Audit2Rat = Null
''.Form!sfrmAuditorLetter!txtAuditorLetterReceived = Null
'!SalePrice = Null
'!Purchaser = ""
'!PurchaserAddress = ""
''!SubstitutePurchaser = 0
'!OrderSubsPurch = Null
'If Not IsNull(Forms!ForeclosureDetails!ExceptionsHearing) And Not IsNull(Forms!ForeclosureDetails!cbxSustained) Then
'Forms!ForeclosureDetails!ExceptionsHearing = Null
'Forms!ForeclosureDetails!cbxSustained = Null
'End If
'
'If Not IsNull(Forms!ForeclosureDetails!StatusHearing) And Not IsNull(Forms!ForeclosureDetails!StatusResults) Then
'Forms!ForeclosureDetails!StatusHearing = Null
'Forms!ForeclosureDetails!StatusResults = Null
'End If
'
'!ResellMotion = Null
'!ResellServed = Null
'!ResellShowCauseExpires = Null
'!ResellAnswered = Null
'!ResellGranted = Null
'!Disposition = Null
''!DispositionDesc = Null
'!DispositionInitials = Null
'!CorrectiveDeedSent = False
'!CorrectiveDeedRecorded = False
'!Settled = Null
'!ClientPaid = Null
'!chEviction = False
'!REO = False
'''.Form!sfrmDisbursingSurplusTable!Type = 0
''.Form!sfrmSoftHold!SoftID = Null
'!AmmDocBackSOD = False
'!AmmStatementOfDebtDate = Null
'!AmmStatementOfDebtAmount = Null
'!AmmStatementOfDebtPerDiem = Null
'!MonitorMotionSurplusFiled = Null
'!MonitorOrderSurplus = Null
'!MonitorClientPaid = Null
'
'
'
'End With
''Forms!foreclosuredetails.Requery
'
'End Sub

Private Sub Form_Current()

If Me.State = "VA" Then
    Me.optLeasehold.Enabled = False
    Me.GroundRentAmount.Enabled = False
    Me.GroundRentPayable.Enabled = False
ElseIf Me.State = "DC" Then
    Me.OptionCoop.Visible = True
    Me.Coop.Visible = True
End If
End Sub

Private Sub Form_Deactivate()
If CaseNuUpdate = True Then
Exit Sub
Else
DoCmd.RunCommand acCmdSaveRecord
End If

End Sub

Private Sub optLeasehold_Click()

If Me.optLeasehold = 1 Then 'Leasehold
    Me.GroundRentAmount.Enabled = True
    Me.GroundRentPayable.Enabled = True
Else
    Me.GroundRentAmount.Enabled = False
    Me.GroundRentPayable.Enabled = False
End If
End Sub
