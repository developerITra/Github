VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit





Private Sub cbxNotice_AfterUpdate()
Dim n As Integer
If Me.Parent.Name = "ForeclosureDetails" Then
        
If Not IsNull(Forms!foreclosuredetails!Notices) Then
 If Me.NewRecord Then
    If IsNull(Forms!foreclosuredetails!UpdatedNotices) Then
         n = 1
        Forms!foreclosuredetails!UpdatedNotices.Locked = False
        Forms!foreclosuredetails!UpdatedNotices.Enabled = True
        Forms!foreclosuredetails!UpdatedNotices.BackStyle = 1
        Forms!foreclosuredetails!TimesUpdatedNotice = n
        DoCmd.GoToRecord , , acNewRec
        Else
        Forms!foreclosuredetails!UpdatedNotices.Locked = False
        Forms!foreclosuredetails!UpdatedNotices.Enabled = True
        Forms!foreclosuredetails!UpdatedNotices.BackStyle = 1

         n = n + 1
        Forms!foreclosuredetails!TimesUpdatedNotice = Forms!foreclosuredetails!TimesUpdatedNotice + n
        DoCmd.GoToRecord , , acNewRec
    End If
 End If
End If
End If
End Sub

Private Sub cmdNoNotice_Click()

On Error GoTo Err_cmdNoNotice_Click
cbxNotice = Null

Exit_cmdNoNotice_Click:
    Exit Sub

Err_cmdNoNotice_Click:
    MsgBox Err.Description
    Resume Exit_cmdNoNotice_Click
    
End Sub

Private Sub cmdCopy_Click()

On Error GoTo Err_cmdCopy_Click
If IsLoadedF("ForeclosureDetails") = True Then


        [Address] = Forms!foreclosuredetails!PropertyAddress
        [Address2] = Forms!foreclosuredetails![Fair Debt]
        [City] = Forms!foreclosuredetails!City
        [State] = Forms!foreclosuredetails!State
        [Zip] = Forms!foreclosuredetails!ZipCode
End If

If IsLoadedF("EvictionDetails") = True Then
   
        [Address] = Me.Parent!sfrmForeclosure!PropertyAddress
        [Address2] = Forms!foreclosuredetails!Apt
        [City] = Me.Parent!sfrmForeclosure!City
        [State] = Me.Parent!sfrmForeclosure!State
        [Zip] = Me.Parent!sfrmForeclosure!ZipCode
 End If
 
If IsLoadedF("wizReferralII") = True Then
  
        'Added a case to deal with auto populating address info into RSII wizard.
        'Left MsgBox lines in for future tests
        'Adress is grabbed from parent form wizReferralII and populated into sfrmNames
        'This code is also called when Occupant button is pushed.
        'Patrick J. Fee 240-401-6820 8/3/11.
        'MsgBox (Me.Parent.Name)
        'MsgBox (Me.Name)
        [Address] = Me.Parent.PropertyAddress
        [Address2] = Me.Parent.Apt
        [City] = Me.Parent.City
        [State] = Me.Parent.State
        [Zip] = Me.Parent.ZipCode
End If

'
'
'    Case Else
'        [Address] = Me.Parent.PropertyAddress
'        [City] = Me.Parent.City
'        [State] = Me.Parent.State
'        [Zip] = Me.Parent.ZipCode
''        [Address] = Me.Parent!sfrmPropAddr!PropertyAddress
''        [City] = Me.Parent!sfrmPropAddr!City
''        [State] = Me.Parent!sfrmPropAddr!State
''        [Zip] = Me.Parent!sfrmPropAddr!ZipCode
'End Select

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





Private Sub Deceased_BeforeUpdate(Cancel As Integer)
If Deceased = True Then
    PR.Enabled = True
    PR.Locked = False
    PR.BackStyle = 1
End If
End Sub

Private Sub Deceased_Click()
Deceased = True
End Sub


'Private Sub Form_BeforeUpdate(Cancel As Integer)

'If CheckNameEdit() = False Then
'Me.AllowEdits = False
'Me.AllowDeletions = False
'Me.AllowAdditions = False
'MsgBox ("good")
'
'Cancel = 0
'
'Dim ctlC As Control
'        ' For each control.
'        For Each ctlC In Me.Controls
'            If ctlC.ControlType = acTextBox Then
'                ' Restore Old Value.
'                ctlC.Value = ctlC.OldValue
'            End If
'        Next ctlC
'
'
'Exit Sub
'
'End If


'    If MsgBox("Changes have been made to this record." _
'        & vbCrLf & vbCrLf & "Do you want to save these changes?" _
'        , vbYesNo, "Changes Made...") = vbYes Then
'            DoCmd.Save
'        Else
'            DoCmd.RunCommand acCmdUndo
'    End If

'End Sub

Private Sub Form_Current()


If (DLookup("Casetypeid", "Caselist", "Filenumber = " & FileNumber)) = 8 Then



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
End If

Exit_cmdEnterSSN_Click:
    Exit Sub

Err_cmdEnterSSN_Click:
    MsgBox Err.Description
    Resume Exit_cmdEnterSSN_Click
    
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
FetchZipCodeCityState Zip, Me.City, Me.State
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

