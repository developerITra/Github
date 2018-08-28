VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmNamesUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cbxNotice_AfterUpdate()

 NoticeTypechange = True


'Dim n As Integer
'If Forms.Parent.Name = "ForeclosureDetails" Then
'MsgBox ("as")
'
'
'If Not IsNull(Forms!ForeclosureDetails!Notices) Then
' If Me.NewRecord Then
'    If IsNull(Forms!ForeclosureDetails!UpdatedNotices) Then
'         n = 1
'        Forms!ForeclosureDetails!UpdatedNotices.Locked = False
'        Forms!ForeclosureDetails!UpdatedNotices.Enabled = True
'        Forms!ForeclosureDetails!UpdatedNotices.BackStyle = 1
'        Forms!ForeclosureDetails!TimesUpdatedNotice = n
'        DoCmd.GoToRecord , , acNewRec
'        Else
'        Forms!ForeclosureDetails!UpdatedNotices.Locked = False
'        Forms!ForeclosureDetails!UpdatedNotices.Enabled = True
'        Forms!ForeclosureDetails!UpdatedNotices.BackStyle = 1
'
'         n = n + 1
'        Forms!ForeclosureDetails!TimesUpdatedNotice = Forms!ForeclosureDetails!TimesUpdatedNotice + n
'        DoCmd.GoToRecord , , acNewRec
'    End If
' End If
'End If
'End If
End Sub

Private Sub ChServed_AfterUpdate()
       If (Forms!foreclosuredetails!State = "DC") Then
    Debug.Print (DCount("Served", "Names", "Filenumber=" & FileNumber & " and Served=1"))
        If (Nz(DCount("Served", "Names", "Filenumber=" & FileNumber & " and Served=1"), -1) = 0) Then
            Dim InvPct As Double
            Dim cbxClient As Integer
            cbxClient = Nz(DLookup("ClientID", "Caselist", "FileNumber=" & FileNumber))
            Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
                Case 1 'Conventional
                    FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=" & cbxClient))
                Case 2 'VA or Veteran's Affairs
                    FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
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
                InvPct = DLookup("DCServiceCompPct", "clientlist", "clientid=" & cbxClient)
            Else
                InvPct = 1
            End If
            If FeeAmount > 0 Then
                If InvPct <= 1 Then
                    AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when borrower served of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
                Else
                    'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee", InvPct * FeeAmount, 0, True, True, False, False
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdNoNotice_Click()

On Error GoTo Err_cmdNoNotice_Click
cbxNotice = Null
NoticeType = 0


Exit_cmdNoNotice_Click:
    Exit Sub

Err_cmdNoNotice_Click:
    MsgBox Err.Description
    Resume Exit_cmdNoNotice_Click
    
End Sub

Private Sub cmdCopy_Click()

On Error GoTo Err_cmdCopy_Click

'Select Case Forms.Parent.Name
If IsLoadedF("ForeclosureDetails") = True Then
    'Case "ForeclosureDetails"
        [Address] = Forms!foreclosuredetails!PropertyAddress
        [Address2] = Forms!foreclosuredetails![Fair Debt]
        [City] = Forms!foreclosuredetails!City
        [State] = Forms!foreclosuredetails!State
        [Zip] = Forms!foreclosuredetails!ZipCode
        
End If

If IsLoadedF("EvictionDetails") = True Then
'    Case "EvictionDetails"
        [Address] = Forms.EvictionDetails!sfrmForeclosure.Form!PropertyAddress
        [Address2] = Forms.EvictionDetails!sfrmForeclosure.Form!Apt
        [City] = Forms.EvictionDetails!sfrmForeclosure.Form!City
        [State] = Forms.EvictionDetails!sfrmForeclosure.Form!State
        [Zip] = Forms.EvictionDetails!sfrmForeclosure.Form!ZipCode
End If

If IsLoadedF("wizReferralII") = True Then

'    Case "wizReferralII"
'        'Added a case to deal with auto populating address info into RSII wizard.
'        'Left MsgBox lines in for future tests
'        'Adress is grabbed from parent form wizReferralII and populated into sfrmNames
'        'This code is also called when Occupant button is pushed.
'        'Patrick J. Fee 240-401-6820 8/3/11.
'        'MsgBox (Me.Parent.Name)
'        'MsgBox (Me.Name)
        [Address] = Forms.wizreferralII.PropertyAddress
        [Address2] = Forms.wizreferralII.Apt
        [City] = Forms.wizreferralII.City
        [State] = Forms.wizreferralII.State
        [Zip] = Forms.wizreferralII.ZipCode
End If

If IsLoadedF("BankruptcyDetails") = True Then
'sfrmPropAddr

'    Case Else
'        [Address] = Form.Parent.PropertyAddress
'        [City] = Form.Parent.City
'        [State] = Form.Parent.State
'        [Zip] = Form.Parent.ZipCode
       [Address] = Forms!BankruptcyDetails!sfrmPropAddr.Form!PropertyAddress
       [Address2] = Forms!BankruptcyDetails!sfrmPropAddr.Form!Apt
       [City] = Forms!BankruptcyDetails!sfrmPropAddr.Form!City
       [State] = Forms!BankruptcyDetails!sfrmPropAddr.Form!State
       [Zip] = Forms!BankruptcyDetails!sfrmPropAddr.Form!ZipCode
'End Select
End If

If IsLoadedF("wizNOI") = True Then

        [Address] = Forms.wizNOI.PropertyAddress
        [Address2] = Forms.wizNOI.Apt
        [City] = Forms.wizNOI.City
        [State] = Forms.wizNOI.State
        [Zip] = Forms.wizNOI.ZipCode
End If

If IsLoadedF("wizFairDebt") = True Then

        [Address] = Forms.wizfairdebt.PropertyAddress
        [Address2] = Forms.wizfairdebt.Apt
        [City] = Forms.wizfairdebt.City
        [State] = Forms.wizfairdebt.State
        [Zip] = Forms.wizfairdebt.ZipCode
End If

If IsLoadedF("WizDemand") = True Then

        [Address] = Forms.WizDemand.PropertyAddress
        [Address2] = Forms.WizDemand.Apt
        [City] = Forms.WizDemand.City
        [State] = Forms.WizDemand.State
        [Zip] = Forms.WizDemand.ZipCode
End If



Exit_cmdCopy_Click:
    Exit Sub

Err_cmdCopy_Click:
    MsgBox Err.Description
    Resume Exit_cmdCopy_Click
    
End Sub

Private Sub cmdDelete_Click()

On Error GoTo Err_cmdDelete_Click

If (vbYes = MsgBox("Do you want to delete this name?", vbYesNo, "Confirm Delete")) Then
  DoCmd.RunCommand acCmdDeleteRecord
End If


Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
    
End Sub

Private Sub cmdCopyClient_Click()
Dim C As Recordset
Dim d As Recordset
Dim Nform As String


On Error GoTo Err_cmdCopyClient_Click
If IsLoadedF("Case List") = False Then

        If IsLoadedF("WizDemand") = True Then
        Nform = [Forms]!WizDemand![FileNumber]
        ElseIf IsLoadedF("wizreferralII") = True Then Nform = [Forms]!wizreferralII![FileNumber]
        ElseIf IsLoadedF("wizNOI") = True Then Nform = [Forms]!wizNOI![FileNumber]
        ElseIf IsLoadedF("wizSAI") = True Then Nform = [Forms]!wizSAI![FileNumber]
        ElseIf IsLoadedF("wizFairDebt") = True Then Nform = [Forms]!wizfairdebt![FileNumber]
                End If
    Set d = CurrentDb.OpenRecordset("SELECT ClientList.* FROM ClientList INNER JOIN CaseList ON ClientList.ClientID = CaseList.ClientID WHERE CaseList.FileNumber=" & Nform, dbOpenSnapshot)
    If d.EOF Then
    MsgBox "Client information not found", vbInformation
    Else
    [Company] = d("LongClientName")
        [First] = d("ContactFirstName")
        [Last] = d("ContactLastName")
        [Address] = d("StreetAddress")
        [Address2] = d("StreetAddr2")
        [City] = d("City")
        [State] = d("State")
        [Zip] = d("ZipCode")
        cbxNotice = 6
    End If
    d.Close



Else


    Set C = CurrentDb.OpenRecordset("SELECT ClientList.* FROM ClientList INNER JOIN CaseList ON ClientList.ClientID = CaseList.ClientID WHERE CaseList.FileNumber=" & [Forms]![Case List]![FileNumber], dbOpenSnapshot)
    If C.EOF Then
        MsgBox "Client information not found", vbInformation
    Else
        [Company] = C("LongClientName")
        [First] = C("ContactFirstName")
        [Last] = C("ContactLastName")
        [Address] = C("StreetAddress")
        [Address2] = C("StreetAddr2")
        [City] = C("City")
        [State] = C("State")
        [Zip] = C("ZipCode")
        cbxNotice = 6
    End If
    C.Close
End If

Exit_cmdCopyClient_Click:
    Exit Sub

Err_cmdCopyClient_Click:
    MsgBox Err.Description
    Resume Exit_cmdCopyClient_Click
    
End Sub

Private Sub cmdPrintNotice_Click()
Dim rstFC As Recordset, rstNames As Recordset

On Error GoTo Err_cmdPrintNotice_Click


'Check to make sure the notice was sent to all parties once already
Set rstFC = CurrentDb.OpenRecordset("select * from fcdetails where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If Not IsNull(rstFC!Notices) Then
'Capture print output destination
DoCmd.OpenForm "PrintOptions", acNormal, , , , acDialog
'Pass through existing DoReport procedure
DoReport "Notice " & rstFC!State & " Name", Forms!foreclosuredetails!PrintOutput, FileNumber, , "ID=" & Me!ID
'DoReport "Notice to Occupant", Forms!foreclosuredetails!PrintOutput, FileNumber
'Update sent notice date in Names table based on user decision
If MsgBox("Do you want to update the Notices Sent date to the current date?", vbYesNo) = vbYes Then
    Set rstNames = CurrentDb.OpenRecordset("select * from Names where ID=" & ID, dbOpenDynaset, dbSeeChanges)
        With rstNames
        .Edit
        !SaleNoticeSent = Date
        .Update
        .Close
        End With
    'Add status to status table using AddStatus function
    AddStatus FileNumber, Date, "Notice of Sale sent to " & First & " " & Last
    
   AddInvoiceItem FileNumber, "FC-NOT", "Additional Sale Notice - Certified Postage", (Nz(DLookup("Value", "StandardCharges", "ID=" & 8))), 76, False, False, False, True
   AddInvoiceItem FileNumber, "FC-NOT", "Additional Sale Notice - First Class Postage", (Nz(DLookup("Value", "StandardCharges", "ID=" & 1))), 76, False, False, False, True

    Forms!foreclosuredetails!UpdatedNotices = Date
    End If
Else: MsgBox "You must send the notices out to all parties before using the individual print option", vbCritical
End If
rstFC.Close
Me.Requery
Exit_cmdPrintNotice_Click:
    Exit Sub

Err_cmdPrintNotice_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintNotice_Click

End Sub

Private Sub cmdTenant_Click()

On Error GoTo Err_cmdTenant_Click

[Company] = "All Occupants"
Call cmdCopy_Click
cbxNotice = 4

Exit_cmdTenant_Click:
    Exit Sub

Err_cmdTenant_Click:
    MsgBox Err.Description
    Resume Exit_cmdTenant_Click
    
End Sub







Private Sub ComUndo_Click()
Me.Undo
DoCmd.Close
End Sub

Private Sub ComUpdate_Click()
Call makeJournaltext
Dim JournalClose As Boolean
If Me.NewRecord Then 'Add New Name
    If IsNull(Me.Company) And IsNull(Me.First) And IsNull(Me.Last) And _
    IsNull(Me.AKA) Then
    Me.Undo
    DoCmd.Close acForm, "sfrmNamesUpdate", acSaveNo
    Else
    'Case changes on Zip
        If ZipChange = True Then
        FetchZipCodeCityState Zip, Me.City, Me.State
        End If


                If NoticeTypechange = True Then
                Dim n As Integer
                    If IsLoadedF("ForeclosureDetails") = True Then
                   
                   
                        If Not IsNull(Forms!foreclosuredetails!Notices) Then
                            If Me.NewRecord Then
'                                If IsNull(Forms!ForeclosureDetails!UpdatedNotices) Then 'I stoped SA  (change logic)
'                                     n = 1
'                                    Forms!ForeclosureDetails!UpdatedNotices.Locked = False
'                                    Forms!ForeclosureDetails!UpdatedNotices.Enabled = True
'                                    Forms!ForeclosureDetails!UpdatedNotices.BackStyle = 1
'                                    Forms!ForeclosureDetails!TimesUpdatedNotice = n
'
'                                 Else
                                    Forms!foreclosuredetails!UpdatedNotices = Null
                                    n = Forms!foreclosuredetails!TimesUpdatedNotice
                                    Forms!foreclosuredetails!UpdatedNotices.Locked = False
                                    Forms!foreclosuredetails!UpdatedNotices.Enabled = True
                                    Forms!foreclosuredetails!UpdatedNotices.BackStyle = 1
                                    n = n + 1
                                    Forms!foreclosuredetails!TimesUpdatedNotice = n
                        
                               ' End If
                            End If
                        End If
                            
                    End If
                   '--
                   If IsLoadedF("wizNOI") = True Then
                   
                        If Not IsNull(Forms!wizNOI!Notices) Then
                            If Me.NewRecord Then
'                                If IsNull(Forms!wizNOI!UpdatedNotices) Then
'                                     n = 1
'                                    Forms!wizNOI!UpdatedNotices.Locked = False
'                                    Forms!wizNOI!UpdatedNotices.Enabled = True
'                                    Forms!wizNOI!UpdatedNotices.BackStyle = 1
'                                    Forms!wizNOI!TimesUpdatedNotice = n
'
'                                 Else
                                    Forms!wizNOI!UpdatedNotices = Null
                                    n = Forms!wizNOI!TimesUpdatedNotice
                                    Forms!wizNOI!UpdatedNotices.Locked = False
                                    Forms!wizNOI!UpdatedNotices.Enabled = True
                                    Forms!wizNOI!UpdatedNotices.BackStyle = 1
                                    n = n + 1
                                    Forms!wizNOI!TimesUpdatedNotice = n
                        
                               ' End If
                            End If
                        End If
                            
                    End If
                    
                    '--
                    
                     If IsLoadedF("wizFairDebt") = True Then
                   
                   
                        If Not IsNull(Forms!wizfairdebt!Notices) Then
                            If Me.NewRecord Then
'                                If IsNull(Forms!wizFairDebt!UpdatedNotices) Then
'                                     n = 1
'                                    Forms!wizFairDebt!UpdatedNotices.Locked = False
'                                    Forms!wizFairDebt!UpdatedNotices.Enabled = True
'                                    Forms!wizFairDebt!UpdatedNotices.BackStyle = 1
'                                    Forms!wizFairDebt!TimesUpdatedNotice = n
'
'                                 Else
                                    Forms!wizfairdebt!UpdatedNotices = Null
                                    n = Forms!wizfairdebt!TimesUpdatedNotice
                                    Forms!wizfairdebt!UpdatedNotices.Locked = False
                                    Forms!wizfairdebt!UpdatedNotices.Enabled = True
                                    Forms!wizfairdebt!UpdatedNotices.BackStyle = 1
                                    n = n + 1
                                    Forms!wizfairdebt!TimesUpdatedNotice = n
                        
                               ' End If
                            End If
                        End If
                            
                    End If
                    
                    '--
                    If IsLoadedF("WizDemand") = True Then
                    
                        If Not IsNull(Forms!WizDemand!Notices) Then
                                If Me.NewRecord Then
    '                                If IsNull(Forms!wizNOI!UpdatedNotices) Then
    '                                     n = 1
    '                                    Forms!wizNOI!UpdatedNotices.Locked = False
    '                                    Forms!wizNOI!UpdatedNotices.Enabled = True
    '                                    Forms!wizNOI!UpdatedNotices.BackStyle = 1
    '                                    Forms!wizNOI!TimesUpdatedNotice = n
    '
    '                                 Else
                                        Forms!WizDemand!UpdatedNotices = Null
                                        n = Forms!WizDemand!TimesUpdatedNotice
                                        Forms!WizDemand!UpdatedNotices.Locked = False
                                        Forms!WizDemand!UpdatedNotices.Enabled = True
                                        Forms!WizDemand!UpdatedNotices.BackStyle = 1
                                        n = n + 1
                                        Forms!WizDemand!TimesUpdatedNotice = n
                            
                                   ' End If
                                End If
                            End If
                            
                    End If
                    
                    
                    
                    
                    
                End If
        
                Forms!Journal.cmdNewJournalEntry_Click
GiveTime:
                Wait 1
                
                    If IsLoadedF("Journal New Entry") = True Then
                        GoTo GiveTime
                    Else
                                If UpdateName Then
                                    Me.Undo
                                    DoCmd.Close acForm, "sfrmNamesUpdate", acSaveNo
                                    UpdateName = False
                                Else
                            
                                    DoCmd.SetWarnings False
                                    strinfo = NameJournal
                                    strinfo = Replace(strinfo, "'", "''")
                                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!sfrmNamesUpdate!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                                    DoCmd.RunSQL strSQLJournal
                                    DoCmd.SetWarnings True
                                    NameJournal = vbNullString
                            
                                    DoCmd.Close acForm, "sfrmNamesUpdate", acSaveYes
                                    
                                    
                                    If IsLoadedF("ForeclosureDetails") = True Then Forms!foreclosuredetails!sfrmNames.Requery
                                    If IsLoadedF("BankruptcyDetails") = True Then Forms!BankruptcyDetails!sfrmNames.Requery
                                    If IsLoadedF("EvictionDetails") = True Then Forms!EvictionDetails!sfrmNames.Requery
                                    If IsLoadedF("wizReferralII") = True Then Forms!wizreferralII!sfrmNames.Requery
                                    If IsLoadedF("wizNOI") = True Then Forms!wizNOI!sfrmNames.Requery
                                    If IsLoadedF("wizFairDebt") = True Then Forms!wizfairdebt!sfrmNames.Requery
                                    If IsLoadedF("WizDemand") = True Then Forms!WizDemand!sfrmNames.Requery
                                    If IsLoadedF("wizSAI") = True Then Forms!wizSAI!sfrmNames.Requery
                                    
                                
                                End If
            
                    End If
        
            End If

    Forms!Journal.Requery


Else

    If IsNull(Me.Company) And IsNull(Me.First) And IsNull(Me.Last) And _
        IsNull(Me.AKA) Then
        Me.Undo
        MsgBox ("You can not remove all data")
        Exit Sub
    Else
    
        If ZipChange = True Then
        FetchZipCodeCityState Zip, Me.City, Me.State
        End If
    
    
        If NoticeTypechange = True Then
            Dim m As Integer
            If IsLoadedF("ForeclosureDetails") = True Then
             
               If Not IsNull(Forms!foreclosuredetails!Notices) Then
                    If Me.NewRecord Then
                        If IsNull(Forms!foreclosuredetails!UpdatedNotices) Then
                m = 1
                Forms!foreclosuredetails!UpdatedNotices.Locked = False
                Forms!foreclosuredetails!UpdatedNotices.Enabled = True
                Forms!foreclosuredetails!UpdatedNotices.BackStyle = 1
                Forms!foreclosuredetails!TimesUpdatedNotice = m
                DoCmd.GoToRecord , , acNewRec
                Else
                Forms!foreclosuredetails!UpdatedNotices.Locked = False
                Forms!foreclosuredetails!UpdatedNotices.Enabled = True
                Forms!foreclosuredetails!UpdatedNotices.BackStyle = 1
        
                 m = m + 1
                Forms!foreclosuredetails!TimesUpdatedNotice = Forms!foreclosuredetails!TimesUpdatedNotice + m
                DoCmd.GoToRecord , , acNewRec
                        End If
            
            
            
            
                    End If
                End If
            End If
        End If
    
    
    '9/3/14
GiveTimeE:
    Wait 1
    
    If IsLoadedF("Journal New Entry") = True Then
    
    GoTo GiveTimeE
    Else
    
    
    If UpdateName Then
    Me.Undo
    DoCmd.Close acForm, "sfrmNamesUpdate", acSaveNo
    UpdateName = False
    Else
    
        
        DoCmd.SetWarnings False
        strinfo = NameJournal
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!sfrmNamesUpdate!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        NameJournal = vbNullString
    
        
    DoCmd.Close acForm, "sfrmNamesUpdate", acSaveYes
    
    
    If IsLoadedF("ForeclosureDetails") = True Then Forms!foreclosuredetails!sfrmNames.Requery
    If IsLoadedF("BankruptcyDetails") = True Then Forms!BankruptcyDetails!sfrmNames.Requery
    If IsLoadedF("EvictionDetails") = True Then Forms!EvictionDetails!sfrmNames.Requery
    If IsLoadedF("wizReferralII") = True Then Forms!wizreferralII!sfrmNames.Requery
    If IsLoadedF("wizNOI") = True Then Forms!wizNOI!sfrmNames.Requery
    If IsLoadedF("wizFairDebt") = True Then Forms!wizfairdebt!sfrmNames.Requery
    If IsLoadedF("WizDemand") = True Then Forms!WizDemand!sfrmNames.Requery
    If IsLoadedF("wizSAI") = True Then Forms!wizSAI!sfrmNames.Requery
    
    
    End If
    
    End If
    
    End If
    
    Forms!Journal.Requery
End If
If NoticeTypechange Then NoticeTypechange = False

End Sub





Private Sub Form_Current()
'wizFairDebt
'wizReferralII
'DeceasedChange = False
'WizDemand
'wizSAI
If ZipChange Then ZipChange = False
If NoticeTypechange Then NoticeTypechange = False
'SSNChange = False

If IsLoadedF("ForeclosureDetails") = True Then Me.FileNumber = Forms!foreclosuredetails!FileNumber
If IsLoadedF("BankruptcyDetails") = True Then Me.FileNumber = Forms!BankruptcyDetails!FileNumber
If IsLoadedF("EvictionDetails") = True Then Me.FileNumber = Forms!EvictionDetails!FileNumber
If IsLoadedF("wizReferralII") = True Then Me.FileNumber = Forms!wizreferralII!FileNumber
If IsLoadedF("wizNOI") = True Then Me.FileNumber = Forms!wizNOI!FileNumber
If IsLoadedF("wizFairDebt") = True Then Me.FileNumber = Forms!wizfairdebt!FileNumber
If IsLoadedF("WizDemand") = True Then Me.FileNumber = Forms!WizDemand!FileNumber
If IsLoadedF("wizSAI") = True Then Me.FileNumber = Forms!wizSAI!FileNumber


If (DLookup("Casetypeid", "Caselist", "Filenumber = " & Me.FileNumber)) = 8 Then



Dim ctrl As Control
For Each ctrl In Me.Controls

If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is Label Or TypeOf ctrl Is CommandButton Or TypeOf ctrl Is ComboBox Then  ' TypeOf ctrl Is CommandButton Then
ctrl.Visible = False
End If

Next

Me.Label22.Visible = True
Me.Company.Visible = True
Me.Label23.Visible = True
Me.First.Visible = True
Me.Last.Visible = True
Me.AKA.Visible = True
Me.Address.Visible = True
Me.Address2.Visible = True
Me.City.Visible = True
Me.State.Visible = True
Me.Zip.Visible = True
Me.cmdPrintLabel.Visible = True
Me.Label3.Visible = True
Me.Label4.Visible = True
Me.Label85.Visible = True
Me.ComUpdate.Visible = True
Me.ComUndo.Visible = True


End If



 
End Sub

'Private Sub Form_Open(Cancel As Integer)
'If Not CheckNameEdit() Then
'Dim ctrl As Control
'For Each ctrl In Me.Controls
'
'If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then ' TypeOf ctrl Is CommandButton Then
''If Not ctrl.Locked) Then
'ctrl.Locked = True
''Else
''ctrl.Locked = True
'End If
''End If
'Next
'
'cmdCopyClient.Enabled = False
'cmdCopy.Enabled = False
'cmdTenant.Enabled = False
'cmdMERS.Enabled = False
'cmdEnterSSN.Enabled = False
'cmdNoNotice.Enabled = False
'cmdPrintNotice.Enabled = False
'cmdPrintLabel.Enabled = True
'cbxNotice.Enabled = False
'cmdDelete.Enabled = False
'
'
'
'
'Exit Sub
'Else
'
'cmdCopy.Enabled = Not (Me.Parent.Name = "CollectionDetails")
'cmdTenant.Enabled = Not (Me.Parent.Name = "CollectionDetails")
'End If
'End Sub

Private Sub cmdMERS_Click()

On Error GoTo Err_cmdMERS_Click
Company = "MERS, Inc."
Address = "1818 Library Street, Suite 300"
City = "Reston"
State = "VA"
Zip = "20190"
cbxNotice = 2   ' Jr. Lienholder

Exit_cmdMERS_Click:
    Exit Sub

Err_cmdMERS_Click:
    MsgBox Err.Description
    Resume Exit_cmdMERS_Click
    
End Sub

Private Sub cmdPrintLabel_Click()
Dim rstLabelData As Recordset, sql As String

On Error GoTo Err_cmdPrintLabel_Click

sql = "SELECT CaseList.FileNumber, CaseList.PrimaryDefName, ClientList.ShortClientName, Names.Company, Names.Deceased, Names.Last, Names.First, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip " & _
        "FROM ClientList INNER JOIN (CaseList INNER JOIN Names ON CaseList.FileNumber = Names.FileNumber) ON ClientList.ClientID = CaseList.ClientID " & _
        "WHERE ID=" & Me!ID
Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
Do While Not rstLabelData.EOF
    Call StartLabel
    Print #6, FormatName(rstLabelData!Company, IIf(rstLabelData!Deceased = True, "Estate of " & rstLabelData!First, rstLabelData!First), rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
    Print #6, "|FONTSIZE 8"
    Print #6, "|BOTTOM"
    Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
    Call FinishLabel
    rstLabelData.MoveNext
Loop
rstLabelData.Close

Exit_cmdPrintLabel_Click:
    Exit Sub

Err_cmdPrintLabel_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintLabel_Click
    
End Sub

Private Sub cmdEnterSSN_Click()

On Error GoTo Err_cmdEnterSSN_Click
If Not IsNull(ID) Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    
    
    If (Not IsNull([SSN])) Then
      MsgBox "SSN has already been entered.  Cannot be updated.", vbCritical, "SSN"
      Exit Sub
    End If
    
    DoCmd.OpenForm "EnterSSN", , , "[ID]=" & Me![ID], , acDialog
    DoCmd.Requery
    Else
    If Me.NewRecord Then DoCmd.OpenForm "EnterSSN", , , , acFormAdd, acDialog
    
End If

Exit_cmdEnterSSN_Click:
    Exit Sub

Err_cmdEnterSSN_Click:
    MsgBox Err.Description
    Resume Exit_cmdEnterSSN_Click
    
End Sub



Private Sub Form_Deactivate()
'Me.Undo

'DoCmd.Close acForm, "sfrmNamesUpdate", acSaveNo

End Sub




Private Sub Last_AfterUpdate()


If Owner = True Then
ProjName = Last & ", " & First
End If

If Mortgagor = True Then
ProjName = Last & ", " & First
End If

If Noteholder = True Then
ProjName = Last & ", " & First
End If


End Sub

Private Sub Zip_AfterUpdate()
'FetchZipCodeCityState Zip, Me.City, Me.State


ZipChange = True

End Sub
'Private Sub cmdPrintNotice_click()
'If DCount("[ID]", "Names", "Nz([NoticeType]) = 0 AND [FileNumber]=" & [Forms]![case list]![FileNumber]) = 0 Then
'        Call DoReport("Notice " & Me!State, PrintTo)
'        If Me!State = "MD" Then
'            Call DoReport("Notice MD County Attorney", PrintTo)
'            Call DoReport("Notice MD All Occupants", PrintTo)
'        End If
'        If MsgBox("Update Notices Sent = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
'        'insert name recordset and copy
'            Forms!foreclosuredetails!Notices = Now()
'            AddStatus [CaseList.FileNumber], Now(), "Sent notices"
'        End If
'    Else
'        Call MsgBox("One or more Send Notice Options are missing", vbCritical)
'    End If
'
'End Sub

Public Function CheckNameEdit()
Dim R1 As String
Dim R2 As String
Dim RC1 As String
Dim RC2 As String
Dim IsComplete As Boolean
Dim IsFormOpen As Boolean
Dim R As Recordset
IsComplete = False

If IsLoadedF("wizReferralII") = True Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
End If

If IsLoadedF("WizDemand") = True Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
End If

If IsLoadedF("wizFairDebt") = True Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
End If

If IsLoadedF("wizNOI") = True Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
End If

If IsLoadedF("wizRestartFCdetails1") = True Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
End If


If IsLoadedF("ForeclosureDetails") = True Then
If (Forms!foreclosuredetails!WizardSource) = "Restart" Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
Else

If IsNull(Forms!foreclosuredetails!WizardSource) Then

 Set R = CurrentDb.OpenRecordset("Select * From WizardQueueStats WHERE FileNumber = " & FileNumber & " And Current = True", dbOpenDynaset, dbSeeChanges)
    If Not IsNull(R!RSIcomplete) Or Not IsNull(R!RSIcomplete) Or Not IsNull(R!RestartRSIComplete) Or Not IsNull(R!RestartComplete) Then
    IsComplete = True
    CheckNameEdit = IsComplete
    Exit Function

End If
End If
End If
End If


End Function
Public Function makeJournaltext()
NameJournal = ""
If Nz(Company) <> Nz(Company.OldValue) Then
    If IsNull(Company) And Not IsNull(Company.OldValue) Then NameJournal = NameJournal + "Removed Company Name " & Company.OldValue & ". "
    If IsNull(Company.OldValue) And Not IsNull(Company) Then NameJournal = NameJournal + "Added Company Name:" & Company & ". "
    If Not IsNull(Company.OldValue) And Not IsNull(Company) Then NameJournal = NameJournal + "Edit Company Name from " & Company.OldValue & " To " & Company & ". "
End If

If Nz(First) <> Nz(First.OldValue) Then
    If IsNull(First) And Not IsNull(First.OldValue) Then NameJournal = NameJournal + "Removed First Name: " & First.OldValue & ". "
    If IsNull(First.OldValue) And Not IsNull(First) Then NameJournal = NameJournal + "Added First Name: " & First & ". "
    If Not IsNull(First.OldValue) And Not IsNull(First) Then NameJournal = NameJournal + "Edit First Name from " & First.OldValue & " To " & First & ". "
End If

If Nz(Last) <> Nz(Last.OldValue) Then
    If IsNull(Last) And Not IsNull(Last.OldValue) Then NameJournal = NameJournal + "Removed Last Name " & Last.OldValue & ". "
    If IsNull(Last.OldValue) And Not IsNull(Last) Then NameJournal = NameJournal + "Added Last Name: " & Last & ". "
    If Not IsNull(Last.OldValue) And Not IsNull(Last) Then NameJournal = NameJournal + "Edit Last Name from " & Last.OldValue & " To " & Last & ". "
End If

If Nz(AKA) <> Nz(AKA.OldValue) Then
    If IsNull(AKA) And Not IsNull(AKA.OldValue) Then NameJournal = NameJournal + "Removed AKA: " & AKA.OldValue & ". "
    If IsNull(AKA.OldValue) And Not IsNull(AKA) Then NameJournal = NameJournal + "Added AKA: " & AKA & ". "
    If Not IsNull(AKA.OldValue) And Not IsNull(AKA) Then NameJournal = NameJournal + "Edit AKA " & AKA.OldValue & " To " & AKA & ". "
End If


If Nz(Address) <> Nz(Address.OldValue) Then
    If IsNull(Address) And Not IsNull(Address.OldValue) Then NameJournal = NameJournal + "Removed Address: " & Address.OldValue & ". "
    If IsNull(Address.OldValue) And Not IsNull(Address) Then NameJournal = NameJournal + "Added Address: " & Address & ". "
    If Not IsNull(Address.OldValue) And Not IsNull(Address) Then NameJournal = NameJournal + "Edit Address " & Address.OldValue & " To " & Address & ". "
End If

If Nz(Address2) <> Nz(Address2.OldValue) Then
    If IsNull(Address2) And Not IsNull(Address2.OldValue) Then NameJournal = NameJournal + "Removed Address: " & Address2.OldValue & ". "
    If IsNull(Address2.OldValue) And Not IsNull(Address2) Then NameJournal = NameJournal + "Added Address: " & Address2 & ". "
    If Not IsNull(Address2.OldValue) And Not IsNull(Address2) Then NameJournal = NameJournal + "Edit Address " & Address2.OldValue & " To " & Address2 & ". "
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


If Nz(Zip) <> Nz(Zip.OldValue) Then
    If IsNull(Zip) And Not IsNull(Zip.OldValue) Then NameJournal = NameJournal + "Removed Zip: " & Zip.OldValue & ". "
    If IsNull(Zip.OldValue) And Not IsNull(Zip) Then NameJournal = NameJournal + "Added Zip: " & Zip & ". "
    If Not IsNull(Zip.OldValue) And Not IsNull(Zip) Then NameJournal = NameJournal + "Edit Zip " & Zip.OldValue & " To " & Zip & ". "
End If

If Nz(cbxNotice) <> Nz(cbxNotice.OldValue) Then
    If (IsNull(cbxNotice) Or cbxNotice = 0) And Not IsNull(cbxNotice.OldValue) Then NameJournal = NameJournal + "Removed Send Notice: " & getNoticeSend(cbxNotice.OldValue) & ". "
    If IsNull(cbxNotice.OldValue) And Not IsNull(cbxNotice) Then NameJournal = NameJournal + "Added Send Notice: " & getNoticeSend(cbxNotice) & ". "
    If (Not IsNull(cbxNotice.OldValue) And Not IsNull(cbxNotice) And cbxNotice <> 0) Then NameJournal = NameJournal + "Edit Send Notice from " & getNoticeSend(cbxNotice.OldValue) & " To " & getNoticeSend(cbxNotice) & ". "
End If

If Nz(PR) <> Nz(PR.OldValue) Then
    If IsNull(PR) And Not IsNull(PR.OldValue) Then NameJournal = NameJournal + "Removed PR: " & PR.OldValue & ". "
    If IsNull(PR.OldValue) And Not IsNull(PR) Then NameJournal = NameJournal + "Added PR: " & PR & ". "
    If Not IsNull(PR.OldValue) And Not IsNull(PR) Then NameJournal = NameJournal + "Edit PR " & PR.OldValue & " To " & PR & ". "
End If


If Nz(Phone) <> Nz(Phone.OldValue) Then
    If IsNull(Phone) And Not IsNull(Phone.OldValue) Then NameJournal = NameJournal + "Removed Phone " & Phone.OldValue & ". "
    If IsNull(Phone.OldValue) And Not IsNull(Phone) Then NameJournal = NameJournal + "Added Phone: " & Phone & ". "
    If Not IsNull(Phone.OldValue) And Not IsNull(Phone) Then NameJournal = NameJournal + "Edit Phone from " & Phone.OldValue & " To " & Phone & ". "
End If

If Nz(SaleNoticeSent) <> Nz(SaleNoticeSent.OldValue) Then
    If IsNull(SaleNoticeSent) And Not IsNull(SaleNoticeSent.OldValue) Then NameJournal = NameJournal + "Removed Notice Sent " & SaleNoticeSent.OldValue & ". "
    If IsNull(SaleNoticeSent.OldValue) And Not IsNull(SaleNoticeSent) Then NameJournal = NameJournal + "Added Notice Sent: " & SaleNoticeSent & ". "
    If Not IsNull(SaleNoticeSent.OldValue) And Not IsNull(SaleNoticeSent) Then NameJournal = NameJournal + "Edit Notice Sent from " & SaleNoticeSent.OldValue & " To " & SaleNoticeSent & ". "
End If

If Nz(ActiveDutyAsOf) <> Nz(ActiveDutyAsOf.OldValue) Then
    If IsNull(ActiveDutyAsOf) And Not IsNull(ActiveDutyAsOf.OldValue) Then NameJournal = NameJournal + "Removed Active Duty As Of: " & ActiveDutyAsOf.OldValue & ". "
    If IsNull(ActiveDutyAsOf.OldValue) And Not IsNull(ActiveDutyAsOf) Then NameJournal = NameJournal + "Added Active Duty As Of: " & ActiveDutyAsOf & ". "
    If Not IsNull(ActiveDutyAsOf.OldValue) And Not IsNull(ActiveDutyAsOf) Then NameJournal = NameJournal + "Edit Active Duty As Of " & ActiveDutyAsOf.OldValue & " To " & ActiveDutyAsOf & ". "
End If

If SSNChange Then
    NameJournal = NameJournal + "Added SSN. "
    SSNChange = False
End If
      

If Nz(Owner) <> Nz(Owner.OldValue) Then
    If Owner <> -1 And Owner.OldValue = -1 Then NameJournal = NameJournal + "Unchecked Owner. "
    If (Owner.OldValue <> -1 Or IsNull(Owner.OldValue)) And Owner = -1 Then NameJournal = NameJournal + "Checked Owner. "
End If

If Nz(Mortgagor) <> Nz(Mortgagor.OldValue) Then
    If Mortgagor <> -1 And Mortgagor.OldValue = -1 Then NameJournal = NameJournal + "Unchecked Mortgagor. "
    If (Mortgagor.OldValue <> -1 Or IsNull(Mortgagor.OldValue)) And Mortgagor = -1 Then NameJournal = NameJournal + "Checked Mortgagor. "
End If

If Nz(Noteholder) <> Nz(Noteholder.OldValue) Then
    If Noteholder <> -1 And Noteholder.OldValue = -1 Then NameJournal = NameJournal + "Unchecked Borrower. "
    If (Noteholder.OldValue <> -1 Or IsNull(Noteholder.OldValue)) And Noteholder = -1 Then NameJournal = NameJournal + "Checked Borrower. "
End If

If Nz(FairDebt) <> Nz(FairDebt.OldValue) Then
    If FairDebt <> -1 And FairDebt.OldValue = -1 Then NameJournal = NameJournal + "Unchecked Fair Debt. "
    If (FairDebt.OldValue <> -1 Or IsNull(FairDebt.OldValue)) And FairDebt = -1 Then NameJournal = NameJournal + "Checked Fair Debt. "
End If

If Nz(BKDebtor) <> Nz(BKDebtor.OldValue) Then
    If BKDebtor <> -1 And BKDebtor.OldValue = -1 Then NameJournal = NameJournal + "Unchecked BK Debtor. "
    If (BKDebtor.OldValue <> -1 Or IsNull(BKDebtor.OldValue)) And Not IsNull(BKDebtor) Then NameJournal = NameJournal + "Checked BK Debtor. "
End If

If Nz(BKCoDebtor) <> Nz(BKCoDebtor.OldValue) Then
    If BKCoDebtor <> -1 And BKCoDebtor.OldValue = -1 Then NameJournal = NameJournal + "Unchecked BK CoDebtor. "
    If (BKCoDebtor.OldValue <> -1 Or IsNull(BKCoDebtor.OldValue)) And BKCoDebtor = -1 Then NameJournal = NameJournal + "Checked BK CoDebtor. "
End If

If Nz(COLDebtor) <> Nz(COLDebtor.OldValue) Then
    If COLDebtor <> -1 And COLDebtor.OldValue = -1 Then NameJournal = NameJournal + "Unchecked COL Debtor. "
    If (COLDebtor.OldValue <> -1 Or IsNull(COLDebtor.OldValue)) And Not IsNull(COLDebtor) Then NameJournal = NameJournal + "Checked COL Debtor. "
End If

If Nz(COS) <> Nz(COS.OldValue) Then
    If COS <> -1 And COS.OldValue = -1 Then NameJournal = NameJournal + "Unchecked COS. "
    If (COS.OldValue <> -1 Or IsNull(COS.OldValue)) And Not IsNull(COS) Then NameJournal = NameJournal + "Checked COS. "
End If

If Nz(chActiveDuty) <> Nz(chActiveDuty.OldValue) Then
    If chActiveDuty <> -1 And chActiveDuty.OldValue = -1 Then NameJournal = NameJournal + "Unchecked Active Duty. "
    If (chActiveDuty.OldValue <> -1 Or IsNull(chActiveDuty.OldValue)) And Not IsNull(chActiveDuty) Then NameJournal = NameJournal + "Checked Active Duty. "
End If

If Nz(EV) <> Nz(EV.OldValue) Then
    If EV <> -1 And EV.OldValue = -1 Then NameJournal = NameJournal + "Unchecked EV. "
    If (EV.OldValue <> -1 Or IsNull(EV.OldValue)) And Not IsNull(EV) Then NameJournal = NameJournal + "Checked EV. "
End If

If Nz(Deceased) <> Nz(Deceased.OldValue) Then
    If Deceased <> -1 And Deceased.OldValue = -1 Then NameJournal = NameJournal + "Unchecked Deceased. "
    If (Deceased.OldValue <> -1 Or IsNull(Deceased.OldValue)) And Not IsNull(Deceased) Then NameJournal = NameJournal + "Checked Deceased. "
End If

If Nz(Tenant) <> Nz(Tenant.OldValue) Then
    If Tenant <> -1 And Tenant.OldValue = -1 Then NameJournal = NameJournal + "Unchecked Tenant. "
    If (Tenant.OldValue <> -1 Or IsNull(Tenant.OldValue)) And Not IsNull(Tenant) Then NameJournal = NameJournal + "Checked Tenant. "
End If

If Nz(ChDefendant) <> Nz(ChDefendant.OldValue) Then
    If ChDefendant <> -1 And ChDefendant.OldValue = -1 Then NameJournal = NameJournal + "Unchecked Defendant. "
    If (ChDefendant.OldValue <> -1 Or IsNull(ChDefendant.OldValue)) And Not IsNull(ChDefendant) Then NameJournal = NameJournal + "Checked Defendant. "
End If

If Nz(ChServed) <> Nz(ChServed.OldValue) Then
    If ChServed <> -1 And ChServed.OldValue = -1 Then NameJournal = NameJournal + "Unchecked Served. "
    If (ChServed.OldValue <> -1 Or IsNull(ChServed.OldValue)) And Not IsNull(ChServed) Then NameJournal = NameJournal + "Checked Served. "
End If

If Nz(ChNOI) <> Nz(ChNOI.OldValue) Then
    If ChNOI <> -1 And ChNOI.OldValue = -1 Then NameJournal = NameJournal + "Unchecked NOI. "
    If (ChNOI.OldValue <> -1 Or IsNull(ChNOI.OldValue)) And Not IsNull(ChNOI) Then NameJournal = NameJournal + "Checked NOI. "
End If

If Nz(Demand) <> Nz(Demand.OldValue) Then
    If Demand <> -1 And Demand.OldValue = -1 Then NameJournal = NameJournal + "Unchecked Demand. "
    If (Demand.OldValue <> -1 Or IsNull(Demand.OldValue)) And Not IsNull(Demand) Then NameJournal = NameJournal + "Checked Demand. "
End If
End Function

Public Function getNoticeSend(A As Long)
getNoticeSend = DLookup("NoticeType", "Noticetypes", "ID= " & A)
End Function



