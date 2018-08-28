VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditPropertyDetailsCiv"
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
Dim rstCivdetails As Recordset
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
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EditPropertyDetailsCiv!txtFileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
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
Forms!CivilDetails.Requery
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
    If IsNull(CourtCaseNumber.OldValue) And Not IsNull(CourtCaseNumber) Then NameJournal = NameJournal + "Added Court Case Number: " & CourtCaseNumber & ". "
    If Not IsNull(CourtCaseNumber.OldValue) And Not IsNull(CourtCaseNumber) Then NameJournal = NameJournal + "Edit Court Case Number " & CourtCaseNumber.OldValue & " To " & CourtCaseNumber & ". "
End If

If Nz(TaxID) <> Nz(TaxID.OldValue) Then
    If IsNull(TaxID) And Not IsNull(TaxID.OldValue) Then NameJournal = NameJournal + "Removed  TAX ID #: " & TaxID.OldValue & ". "
    If IsNull(TaxID.OldValue) And Not IsNull(TaxID) Then NameJournal = NameJournal + "Added TAX ID #: " & TaxID & ". "
    If Not IsNull(TaxID.OldValue) And Not IsNull(TaxID) Then NameJournal = NameJournal + "Edit TAX ID # " & TaxID.OldValue & " To " & TaxID & ". "
End If

If Nz(Unit) <> Nz(Unit.OldValue) Then
    If IsNull(Unit) And Not IsNull(Unit.OldValue) Then NameJournal = NameJournal + "Removed  Unit #: " & Unit.OldValue & ". "
    If IsNull(Unit.OldValue) And Not IsNull(Unit) Then NameJournal = NameJournal + "Added Unit #: " & Unit & ". "
    If Not IsNull(Unit.OldValue) And Not IsNull(Unit) Then NameJournal = NameJournal + "Edit Unit # " & Unit.OldValue & " To " & Unit & ". "
End If

makePropertyJournaltext = NameJournal
End Function


