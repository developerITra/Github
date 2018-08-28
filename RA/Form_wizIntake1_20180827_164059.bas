VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_wizIntake1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cbxFileType_AfterUpdate()
'If Nz(cbxFileType) <> 1 Then MsgBox "This wizard can only be used to enter Foreclosure files", vbExclamation
End Sub

Private Sub cbxLoanType_AfterUpdate()
Select Case cbxLoanType
Case 1
lblAgencyNbr.Visible = False
txtAgencyNbr.Visible = False
Exit Sub
Case 2
lblAgencyNbr.Caption = "VA Loan Number:"
Case 3
lblAgencyNbr.Caption = "HUD Loan Number:"
Case 4
lblAgencyNbr.Caption = "FNMA Loan Number:"
Case 5
lblAgencyNbr.Caption = "FHLMC Loan Number:"
End Select

lblAgencyNbr.Visible = True
txtAgencyNbr.Visible = True

End Sub

Private Sub CheckAgain_Click()
Dim LoanNumbertext As String
If MsgBox(txtLoanNumber, vbYesNo, "Is this the right number? ") = vbYes Then
txtLoanNumber = Trim(txtLoanNumber)
txtLoanNumber.Requery

Call txtLoanNumber_AfterUpdate
txtLoanNumber.SetFocus
CheckAgain.Visible = False
Else
MsgBox ("Please Add the correct Loan Number")
txtLoanNumber.SetFocus
CheckAgain.Visible = False
End If



End Sub

Private Sub chLPS_AfterUpdate()
If chLPS = -1 Then
chLenstar = 0
chVendorscape = 0
chOther = 0
chLenstar.Enabled = False
chVendorscape.Enabled = False
chOther.Enabled = False
Else
chLenstar.Enabled = True
chVendorscape.Enabled = True
chOther.Enabled = True
End If
End Sub







Private Sub cmdOK_Click()
Dim rstCase As Recordset, rstFC As Recordset, rstNames As Recordset, rstwiz As Recordset, rstLocks As Recordset, rstDIL As Recordset, rstInternet As Recordset, rstFCtitle As Recordset
Dim FileNum As Long, MissingInfo As String, ctr As Integer, FieldName As String, cost As Currency, InvPct As Double, rstWizMT As Recordset

On Error GoTo Err_cmdOK_Click

'If Nz(cbxFileType) <> "Foreclosure" Then
'    MsgBox "This wizard can only be used to enter Foreclosure files", vbCritical
'    Exit Sub
'End If

If Not IsDate(txtReferralDate) Then MissingInfo = MissingInfo & "Referral Date, "
If IsNull(cbxFileType) Then MissingInfo = MissingInfo & "File Type, "
If IsNull(cbxClient) Then MissingInfo = MissingInfo & "Client, "
If IsNull(txtLoanNumber) Then MissingInfo = MissingInfo & "Loan Number, "
If IsNull(txtSSN1) Then MissingInfo = MissingInfo & "Social Security Number, "
If IsNull(txtFirstName1) Or IsNull(txtLastName1) Then MissingInfo = MissingInfo & "Primary Name, "
If IsNull(txtProjectName) Then MissingInfo = MissingInfo & "Project Name, "
If IsNull(txtPropertyAddress) Then MissingInfo = MissingInfo & "Property Address, "
If IsNull(txtZipCode) Then MissingInfo = MissingInfo & "Zip Code, "
If IsNull(txtCity) Then MissingInfo = MissingInfo & "City, "
If IsNull(cbxJurisdictionID) Then MissingInfo = MissingInfo & "Jurisdiction, "
If IsNull(cbxLoanType) Then MissingInfo = MissingInfo & "Loan Type, "
If IsNull(txtPosition) Then MissingInfo = MissingInfo & "Lien Position, "
If IsNull(txtAgencyNbr) And cbxLoanType <> 1 Then MissingInfo = MissingInfo & "Agency Loan Number, "

If MissingInfo <> "" Then
    MsgBox "You cannot continue because the following information is missing:" & vbNewLine & Left$(MissingInfo, Len(MissingInfo) - 2), vbCritical
    Exit Sub
End If

FileNum = ReserveNextCaseNumber()
If StaffID = 0 Then Call GetLoginName
Forms!wizIntake1.FileNum = FileNum
Set rstCase = CurrentDb.OpenRecordset("CaseList", dbOpenDynaset, dbSeeChanges)
With rstCase
    .AddNew
    !FileNumber = FileNum
    !ReferralDate = txtReferralDate
    !PrimaryDefName = txtProjectName
   If Not IsNull(Forms!wizIntake1!TxtProject) Then rstCase!Project = Forms!wizIntake1!TxtProject
    
   If Forms!wizIntake1.cbxFileType = "Foreclosure" Then
        rstCase!CaseTypeID = 1 ' FC
   Else
  
    rstCase!CaseTypeID = 8
    rstCase!BillCaseUpdateReasonID = 34
    rstCase!BillCase = True
    rstCase!BillCaseUpdateDate = Now()
    rstCase!BillCaseUpdateUser = StaffID

''Linda Code6/9/2015
    Dim FeeAmt As Currency
    FeeAmt = Nz(DLookup("MonitorFee", "ClientList", "ClientID=" & Forms![wizIntake1]!cbxClient))
    AddInvoiceItem FileNum, "FC-MON", "Monitor sale fee", Format$(FeeAmt, "Currency"), 0, True, True, False, False

    
  End If
   
        !ClientID = cbxClient
        !JurisdictionID = cbxJurisdictionID
        !Active = True
        !OnStatusReport = True
        !OpenDate = Now()
        !OpenBy = StaffID
        .Update
        .Close
    End With
AddStatus FileNum, txtReferralDate, "Received referral"

' Create FC Detail record
Call AddDetailRecord(1, FileNum, txtReferralDate)
Set rstFC = CurrentDb.OpenRecordset("SELECT * FROM FCDetails WHERE FileNumber=" & FileNum, dbOpenDynaset, dbSeeChanges)
If rstFC.EOF Then
    MsgBox "Can't find details record.  Contact support for assistance.", vbCritical
    Exit Sub
End If
With rstFC
    .Edit
    !LoanNumber = txtLoanNumber
    !PrimaryFirstName = txtFirstName1
    !PrimaryLastName = txtLastName1
    !SecondaryFirstName = txtFirstName2
    !SecondaryLastName = txtLastName2
    !PropertyAddress = txtPropertyAddress
    ![Fair Debt] = TexApt
    !City = txtCity
    !State = txtState
    !ZipCode = txtZipCode
    !LoanType = cbxLoanType
    !LienPosition = txtPosition
    
If txtAgencyNbr.Visible = True Then
Select Case cbxLoanType
'break out once loan number field for VA is included
Case "2"
!FHALoanNumber = txtAgencyNbr
Case "3"
!FHALoanNumber = txtAgencyNbr
Case "4"
!FNMALoanNumber = txtAgencyNbr
Case "5"
!FHLMCLoanNumber = txtAgencyNbr
End Select
End If
    .Update
    .Close
End With

'create record in new FC DIL table
Set rstDIL = CurrentDb.OpenRecordset("FCDIL", dbOpenDynaset, dbSeeChanges)
With rstDIL
.AddNew
    !FileNumber = FileNum
    If cbxFileType = "Monitor" Then rstDIL!Monitor_Refer_reced = Now()
    .Update
    .Close
End With

'create record in new Internet Sources table
Set rstInternet = CurrentDb.OpenRecordset("InternetSites", dbOpenDynaset, dbSeeChanges)
With rstInternet
.AddNew
    !FileNumber = FileNum
    If chLPS = -1 Then
    !LPS_Desktop = True
    'AddInvoiceItem FileNum, "FC-Oth", "Technology Fee- LPS", 80, 192, False, True, False, True
    End If
    If chLenstar = -1 Then
    !Lenstar = True
    'AddInvoiceItem FileNum, "FC-Oth", "Technology Fee- Lenstar", 55, 193, False, True, False, True
    End If
    If chVendorscape = -1 Then
    !Vendorscape = True
    'AddInvoiceItem FileNum, "FC-Oth", "Technology Fee- LPS", 80, 194, False, True, False, True
    End If
    If chOther = -1 Then
    !Other = True
    End If
    .Update
    .Close
End With

' Create Names records
Set rstNames = CurrentDb.OpenRecordset("Names", dbOpenDynaset, dbSeeChanges)
With rstNames
    .AddNew
    !FileNumber = FileNum
    !First = txtFirstName1
    !Last = txtLastName1
    !ProjName = txtLastName1 & ", " & txtFirstName1
    !SSN = txtSSN1
    !Address = txtPropertyAddress
    !Address2 = TexApt
    !City = txtCity
    !State = txtState
    !Zip = txtZipCode
   If ChActiveDuty1 = True Then rstNames!ActiveDuty = ChActiveDuty1
   If Not IsNull(ActiveDutyAsOf1) Then rstNames!ActiveDutyAsOf = ActiveDutyAsOf1
    
    .Update
    ctr = 1
End With
With rstNames
' Add an All Occupants entry for each FC
        .AddNew
        !FileNumber = FileNum
        !First = "All"
        !Last = "Occupants"
        !Address = txtPropertyAddress
        !Address2 = TexApt
        !City = txtCity
        !State = txtState
        !Zip = txtZipCode
        
        .Update
End With
If Not IsNull(txtLastName2) Then
    With rstNames
        .AddNew
        !FileNumber = FileNum
        !First = txtFirstName2
        !Last = txtLastName2
        !ProjName = txtLastName2 & ", " & txtFirstName2
        !SSN = txtSSN2
        !Address = txtPropertyAddress
        !Address2 = TexApt
        !City = txtCity
        !State = txtState
        !Zip = txtZipCode
        If ChActiveDuty2 = True Then rstNames!ActiveDuty = ChActiveDuty2
        If Not IsNull(ActiveDutyAsOf2) Then rstNames!ActiveDutyAsOf = ActiveDutyAsOf2
        .Update
        ctr = 2
    End With
End If

If Not IsNull(txtLastName3) Then
    With rstNames
        .AddNew
        !FileNumber = FileNum
        !First = txtFirstName3
        !Last = txtLastName3
        !ProjName = txtLastName3 & ", " & txtFirstName3
        !SSN = txtSSN3
        !Address = txtPropertyAddress
        !Address2 = TexApt
        !City = txtCity
        !State = txtState
        !Zip = txtZipCode
         If ChActiveDuty3 = True Then rstNames!ActiveDuty = ChActiveDuty3
         If Not IsNull(ActiveDutyAsOf3) Then rstNames!ActiveDutyAsOf = ActiveDutyAsOf3
        
        .Update
        ctr = 3
    End With
End If

If Not IsNull(txtLastName4) Then
    With rstNames
        .AddNew
        !FileNumber = FileNum
        !First = txtFirstName4
        !Last = txtLastName4
        !ProjName = txtLastName4 & ", " & txtFirstName4
        !SSN = txtSSN4
        !Address = txtPropertyAddress
        !Address2 = TexApt
        !City = txtCity
        !State = txtState
        !Zip = txtZipCode
        If ChActiveDuty4 = True Then rstNames!ActiveDuty = ChActiveDuty4
        If Not IsNull(ActiveDutyAsOf4) Then rstNames!ActiveDutyAsOf = ActiveDutyAsOf4
        
        .Update
        ctr = 4
    End With
End If
rstNames.Close

'Add SCRA fee
If cbxClient = 97 Then
cost = ctr * DLookup("ivalue", "db", "ID=" & 32)

'Added stage and putunder the fee. 2/5/15
AddInvoiceItem FileNum, "FC-DOD", "DOD Search - New FC Referral", cost, 0, True, True, False, False
End If

'Add Record to Queue Table
Set rstwiz = CurrentDb.OpenRecordset("WizardQueueStats", dbOpenDynaset, dbSeeChanges)

With rstwiz
    .AddNew
    !FileNumber = FileNum
    !RSIcomplete = Now()
    !RSIuser = StaffID
    .Update
    .Close
End With

Set rstwiz = CurrentDb.OpenRecordset("WizardSupportTwo", dbOpenDynaset, dbSeeChanges)

With rstwiz
    .AddNew
    !FileNumber = FileNum
    !Current = True
    !Count = 1
    .Update
    .Close
End With


AddToList (FileNum) 'Add to list of opencase

'Adding to ValuemRI  05/20
DoCmd.SetWarnings False
Dim Shortclient As String
Dim cbxLoanTypetext As String
Shortclient = DLookup("ShortClientName", "ClientList", "ClientID=" & cbxClient)
cbxLoanTypetext = DLookup("LoanType", "LoanTypes", "ID=" & cbxLoanType)
Dim NewCaseType As String
NewCaseType = "Insert Into ValumeRSI (FileNumber,ShortClientName,Completiondate,State,Type,IdUser,Username,IdClient,CaseType,Count,LoanType) values (" & FileNum & ",'" & Shortclient & "','" & Now() & "','" & txtState & "'," & """New Case""" & "," & StaffID & ",'" & GetFullName() & "'," & cbxClient & "," & """Foreclosure""" & "," & 1 & ",'" & cbxLoanTypetext & "')"
DoCmd.RunSQL NewCaseType
DoCmd.SetWarnings False


'Add Record to FCtitle Table
Set rstFCtitle = CurrentDb.OpenRecordset("FCtitle", dbOpenDynaset, dbSeeChanges)
With rstFCtitle
    .AddNew
    !FileNumber = FileNum
    .Update
    .Close
End With

'Add Record to Locks Table
Set rstLocks = CurrentDb.OpenRecordset("Locks", dbOpenDynaset, dbSeeChanges)
With rstLocks
    .AddNew
    !FileNumber = FileNum
    !StaffID = 0
    .Update
    .Close
End With


Select Case txtState
  Case "VA"
   Select Case cbxLoanType
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
      InvPct = DLookup("VAReferralpct", "clientlist", "clientid=" & cbxClient)
    Else
      InvPct = 1
    End If
    If FeeAmount > 0 Then
      InvPct = DLookup("VAReferralpct", "clientlist", "clientid=" & cbxClient)
      If InvPct < 1 Then
        AddInvoiceItem FileNum, "FC-REF", "Attorney FC Fee- " & Format(InvPct, "percent") & " due when Referral received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      Else
        AddInvoiceItem FileNum, "FC-REF", "Attorney FC fee - up to $2,100.00", InvPct * FeeAmount, 0, True, True, False, False
      End If
    End If
  Case "MD"
    If cbxClient = 531 Then
      AddInvoiceItem FileNum, "FC-MD-M&T", "Usage costs", 100, 0, False, True, False, True
    End If
    Select Case cbxLoanType
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
      InvPct = DLookup("MDReferralpct", "clientlist", "clientid=" & cbxClient)
    Else
      InvPct = 1
    End If
    If FeeAmount > 0 Then
      If InvPct < 1 Then
        AddInvoiceItem FileNum, "FC-REF", "Attorney FC Fee- " & Format(InvPct, "percent") & " due when Referral received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      Else
        AddInvoiceItem FileNum, "FC-REF", "Attorney FC fee - up to $2,100.00", InvPct * FeeAmount, 0, True, True, False, False
      End If
    End If
    Case "DC"
      Select Case cbxLoanType
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
        InvPct = DLookup("DCReferralpct", "clientlist", "clientid=" & cbxClient)
      Else
        InvPct = 1
      End If
      If FeeAmount > 0 Then
        If InvPct < 1 Then
          AddInvoiceItem FileNum, "FC-REF", "Attorney FC Fee- " & Format(InvPct, "percent") & " due when Referral received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
        Else
          AddInvoiceItem FileNum, "FC-REF", "Attorney FC fee - up to $2,100.00", InvPct * FeeAmount, 0, True, True, False, False
        End If
      End If
  End Select



' Add Responsibility (1 = Intake)
Call AddFileResponsibilityHistory(FileNum, 1, StaffID)

MsgBox "New file # " & FileNum & " has been entered", vbInformation
'Open case list to docs tab and open journal window

'Create Journal Entry for electronic file and manual entry


'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  lrs.AddNew
'
'  lrs![FileNumber] = FileNum
'  lrs![JournalDate] = Now
'  lrs![Who] = GetFullName()
'  lrs![Info] = "File is electronic file. " & vbCrLf
'  lrs![Color] = 1
'  lrs.Update
strinfo = "File is electronic file. " & vbCrLf
strinfo = Replace(strinfo, "'", "''")
DoCmd.SetWarnings False
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNum & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
If Not IsNull(txtJournal) Then
'  lrs.AddNew
'  lrs![FileNumber] = FileNum
'  lrs![JournalDate] = Now
'  lrs![Who] = GetFullName()
'  lrs![Info] = txtJournal & vbCrLf
'  lrs![Color] = 1
'  lrs.Update
  strinfo = txtJournal & vbCrLf
strinfo = Replace(strinfo, "'", "''")
DoCmd.SetWarnings False
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNum & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
  End If
  
  If (confilictName = True) Or (conflictAddress = True) Then
'  lrs.AddNew
'  lrs![FileNumber] = FileNum
'  lrs![JournalDate] = Now
'  lrs![Who] = GetFullName()
'  lrs![Info] = "File was created with conflicts present." & vbCrLf
'  lrs![Color] = 1
'  lrs.Update
'  lrs.Close
strinfo = "File was created with conflicts present." & vbCrLf
strinfo = Replace(strinfo, "'", "''")
DoCmd.SetWarnings False
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNum & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
  'Set lrs = Nothing
    If confilictName Then confilictName = False
    If conflictAddress Then conflictAddress = False
  End If
    
  
  
Call SelectDocsTab(FileNum)
'Check conflict

DoCmd.SetWarnings False
DoCmd.OpenQuery "ConflictUpdate"
DoCmd.OpenQuery "ConnflictUpdateName"
DoCmd.OpenQuery "ConflictWizardUpdate"

    
    

DoCmd.SetWarnings True

DoCmd.Close acForm, Me.Name
'Call ClearForm

MsgBox "File has been created"


Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub ClearForm()
Dim C As Control

cmdConflicts.Enabled = False
Call ShowConflictList(False)
Call SetConflictVisual(txtLoanNumber, False)
For Each C In Me.Controls
    If C.ControlType = acComboBox Or C.ControlType = acTextBox Then C = Null
Next
txtReferralDate = Now()
txtReferralDate.SetFocus

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





Private Sub ComMon_Click()
If cbxFileType.Value = "Foreclosure" Then
cbxFileType.Value = "Monitor"
'ComMon.Caption = "Monitor"
ComMon.Caption = "Foreclosure"
Else
cbxFileType.Value = "Foreclosure"
ComMon.Caption = "Monitor"
End If

End Sub

Private Sub Form_Current()
cbxFileType.Value = "Foreclosure"
TexPropertyAddress = "99999999"
NameFirst = "99999999"
NameSecond = "99999999"
End Sub

Private Sub Form_Open(Cancel As Integer)
txtReferralDate = Now()
End Sub

Private Sub SetConflictVisual(tb As TextBox, Conflict As Boolean)
If Conflict Then
    tb.FontBold = True
    tb.ForeColor = vbRed
    cmdConflicts.Enabled = True
Else
    tb.FontBold = False
    tb.ForeColor = vbBlack
End If
End Sub




Private Sub lstConflicts_DblClick(Cancel As Integer)
Dim WarningLevel As Integer

Dim F As Form
Dim FormClosed As Boolean

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "Wizards", "wizIntake1"  ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed







DoCmd.OpenForm "wizRestartCaseList1", , , "FileNumber = " & lstConflicts
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & lstConflicts
DoCmd.OpenForm "Journal", , , "FileNumber = " & lstConflicts
Forms!wizRestartCaseList1.SetFocus
WarningLevel = Nz(DMax("Warning", "Journal", "FileNumber=" & lstConflicts))
With Forms!wizRestartCaseList1
Select Case WarningLevel
    Case 50
        .imgWarning.Picture = dbLocation & "dollar.emf"
        .imgWarning.Visible = True
    Case 100
        .imgWarning.Picture = dbLocation & "papertray.emf"
        .imgWarning.Visible = True
    Case 200
        .imgWarning.Picture = dbLocation & "house.emf"
        .imgWarning.Visible = True
    Case 300
        .imgWarning.Picture = dbLocation & "caution.bmp"
        .imgWarning.Visible = True
    Case 400
        .imgWarning.Picture = dbLocation & "stop.emf"
        .imgWarning.Visible = True
    Case Else
        .imgWarning.Visible = False
End Select
End With
End Sub

Private Sub txtFirstName1_AfterUpdate()


Dim StFirst As String
If (InStr([txtFirstName1], " ")) > 0 Then
    NameFirst = Left([txtFirstName1], InStr([txtFirstName1], " ") - 1)
    Else
    NameFirst = txtFirstName1
End If
If IsNull(txtLastName1) Then
    Call SetConflictVisual(txtFirstName1, DCount("FileNumber", "Names", "First Like """ & NameFirst & "*"" ") > 0)
    Else
    Dim conf As Boolean
    conf = False
    conf = DCount("FileNumber", "Names", "Last Like """ & NameSecond & "*"" " & " AND " & "First Like """ & NameFirst & "*"" ")
        If conf Then
        Call SetConflictVisual(txtLastName1, True)
        Call SetConflictVisual(txtFirstName1, True)
        Else
        Call SetConflictVisual(txtLastName1, False)
        Call SetConflictVisual(txtFirstName1, False)
        End If
End If

If (txtFirstName1) <> "" Then
If NameSecond = "99999999" Then
NameSecond = Null
End If
Else
If txtLastName1 <> "" Then
NameFirst = Null
Else
NameFirst = "99999999"
NameSecond = "99999999"

End If
End If

End Sub

Private Sub txtLastName1_AfterUpdate()



txtProjectName = txtLastName1 & IIf(IsNull(txtFirstName1), "", ", " & txtFirstName1)
If (InStr([txtLastName1], " ")) > 0 Then
    NameSecond = Left([txtLastName1], InStr([txtLastName1], " ") - 1)
    Else
    NameSecond = txtLastName1
End If
If IsNull(txtFirstName1) Then
    Call SetConflictVisual(txtLastName1, DCount("FileNumber", "Names", "Last Like """ & NameSecond & "*"" ") > 0)
    Else
    Dim conf As Boolean
    conf = False
    conf = DCount("FileNumber", "Names", "Last Like """ & NameSecond & "*"" " & " AND " & "First Like """ & NameFirst & "*"" ")
    If conf Then
        Call SetConflictVisual(txtLastName1, True)
        Call SetConflictVisual(txtFirstName1, True)
        Else
        Call SetConflictVisual(txtLastName1, False)
        Call SetConflictVisual(txtFirstName1, False)
    End If
End If

If (txtLastName1) <> "" Then
If NameFirst = "99999999" Then
NameFirst = Null
End If
Else
If txtFirstName1 <> "" Then
NameSecond = Null
Else
NameSecond = "99999999"
NameFirst = "99999999"
End If
End If

If (txtFirstName1.ForeColor = vbRed And txtLastName1.ForeColor = vbRed) Then
confilictName = True
End If




End Sub

Private Sub txtLoanNumber_AfterUpdate()

If Not IsNull(txtLoanNumber) Then
    Call SetConflictVisual(txtLoanNumber, (DCount("FileNumber", "FCDetails", "LoanNumber='" & txtLoanNumber & "'") > 0))
 If IsNull(NameFirst) Then NameFirst = "99999"
End If
 If txtLoanNumber.ForeColor <> vbRed Then
 CheckAgain.Visible = True
 
 End If
End Sub
Private Sub txtSSN1_AfterUpdate()

If Not IsNull(txtSSN1) Then
    Call SetConflictVisual(txtSSN1, (DCount("FileNumber", "Names", "SSN=txtSSN1") > 0))
End If

End Sub
Private Sub txtSSN2_AfterUpdate()

If Not IsNull(txtSSN2) Then
    Call SetConflictVisual(txtSSN2, (DCount("FileNumber", "Names", "SSN=txtSSN2") > 0))
End If

End Sub
Private Sub txtSSN3_AfterUpdate()

If Not IsNull(txtSSN3) Then
    Call SetConflictVisual(txtSSN3, (DCount("FileNumber", "Names", "SSN=txtSSN3") > 0))
End If

End Sub
Private Sub txtSSN4_AfterUpdate()

If Not IsNull(txtSSN4) Then
    Call SetConflictVisual(txtSSN4, (DCount("FileNumber", "Names", "SSN=txtSSN4") > 0))
End If

End Sub

Private Sub txtPropertyAddress_AfterUpdate()

If Not IsNull(txtPropertyAddress) Then
 TexPropertyAddress = txtPropertyAddress
 Call SetConflictVisual(txtPropertyAddress, (DCount("FileNumber", "qrySearchAddress", "PropertyAddress like """ & TexPropertyAddress & "*""") > 0))
 Else
 TexPropertyAddress = "999999"
End If
End Sub

Private Sub txtZipCode_AfterUpdate()
Dim rstZip As Recordset, JurisdictionID As Variant

If Not IsNull(txtZipCode) Then
    Set rstZip = CurrentDb.OpenRecordset("SELECT * FROM ZipCodes WHERE ZipCode = '" & Left$(txtZipCode, 5) & "' and Preferred = 'Yes'", dbOpenSnapshot)
    If Not rstZip.EOF Then
        txtCity = StrConv(rstZip!City, vbProperCase)
        txtState = rstZip!State
        cbxJurisdictionID = Null
        JurisdictionID = DLookup("JurisdictionID", "JurisdictionList", "Jurisdiction Like '" & rstZip!County & "*'" & " And State like '" & txtState & "'")
              
        If Not IsNull(JurisdictionID) Then
            If txtState = DLookup("State", "JurisdictionList", "JurisdictionID=" & JurisdictionID) Then
                cbxJurisdictionID = JurisdictionID
            End If
        End If
    Else
        txtCity = Null
        txtState = Null
        cbxJurisdictionID = Null
        MsgBox "CAUTION: unknown Zip Code", vbExclamation
    End If
    rstZip.Close
End If

If (txtPropertyAddress.ForeColor = vbRed And txtZipCode.ForeColor = vbRed And txtZipCode.ForeColor = vbRed) Then
conflictAddress = True
End If

End Sub

Private Sub cmdConflicts_Click()

On Error GoTo Err_cmdConflicts_Click

Call ShowConflictList(Not lstConflicts.Visible)

Exit_cmdConflicts_Click:
    Exit Sub

Err_cmdConflicts_Click:
    MsgBox Err.Description
    Resume Exit_cmdConflicts_Click
    
End Sub

Private Sub ShowConflictList(Show As Boolean)
Dim rstConflicts As Recordset

If Show Then
    DoCmd.MoveSize , , 8 * 1440         ' change width to 8 inches, leave position and height alone
    lstConflicts.Left = 4.75 * 1440
    lstConflicts.Width = 3 * 1440
    lstConflicts.Top = 0.1 * 1440
    lstConflicts.Height = 3.5 * 1440
    lstConflicts.Visible = True
    cmdConflicts.Caption = "Hide Conflicts"
    lstConflicts.RowSource = "qryConflictOutput"
    'lstConflicts.RowSource = "SELECT DISTINCT FileNumber, LoanNumber AS Conflict FROM FCdetails WHERE LoanNumber='" & txtLoanNumber & "' " & _
                                "UNION ALL SELECT DISTINCT FileNumber, PropertyAddress AS Conflict FROM FCdetails WHERE PropertyAddress=""" & txtPropertyAddress & """"
Else
    lstConflicts.Visible = False
    lstConflicts.Left = 0
    lstConflicts.Width = 100
    lstConflicts.Top = 0
    lstConflicts.Height = 100
    DoCmd.MoveSize , , 6 * 1440         ' restore width
    cmdConflicts.Caption = "Show Conflicts"
End If
End Sub


