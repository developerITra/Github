VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Title Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cbxAbstractor_AfterUpdate()
Abstractor = cbxAbstractor
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close acForm, "Print Title Order"

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdComplete_Click()
Dim strSearchType As String
Dim statusMsg As String
Dim NfileNumber As Long
Forms!foreclosuredetails.Requery

Select Case Order
    Case 1
       statusMsg = "Ordered full title search by Client"
       strSearchType = "Full"
    Case 2
       statusMsg = "Ordered title rundown by Client from " & Format$(RundownDate, "m/d/yyyy")
       strSearchType = "Update"
    Case 3
       statusMsg = "Ordered title rundown by Client from present owner"
       strSearchType = "Rundown"
    Case 4
       statusMsg = "Ordered 2 owner search by Client"
       strSearchType = "2 Owner"

End Select

Select Case Weekday(Date)
    Case vbMonday, vbTuesday, vbWednesday     ' due Wednesday, Thursday, Friday
        TitleDue = Format$(DateAdd("d", 2, Date), "mmmm d, yyyy")
    Case vbThursday     ' due Monday
        TitleDue = Format$(DateAdd("d", 4, Date), "mmmm d, yyyy")
    Case vbFriday       ' due Tuesday
        TitleDue = Format$(DateAdd("d", 4, Date), "mmmm d, yyyy")
    Case vbSaturday     ' due Tuesday
        TitleDue = Format$(DateAdd("d", 3, Date), "mmmm d, yyyy")
    Case vbSunday       ' due Tuesday
        TitleDue = Format$(DateAdd("d", 2, Date), "mmmm d, yyyy")
End Select
DateRequired = Format$(TitleDue, "mmmm d, yyyy")

Forms![foreclosuredetails]!TitleOrder = Now()
Forms![foreclosuredetails]!TitleDue = DateRequired
If Not IsNull(Forms![foreclosuredetails]!TitleBack) Then Forms![foreclosuredetails]!TitleBack = Null
If Not IsNull(Forms![foreclosuredetails]!TitleThru) Then Forms![foreclosuredetails]!TitleThru = Null
If Not IsNull(Forms![foreclosuredetails]!TitleReviewToClient) Then Forms![Print Title Order]!TitleReviewToClient = Null
Forms![foreclosuredetails]!TitleSearchType = strSearchType

AddStatus FileNumber, Now(), statusMsg

'Dim lrs As Recordset
'
'            Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'            lrs.AddNew
'            lrs![FileNumber] = FileNumber
'            lrs![JournalDate] = Now
'            lrs![Who] = GetFullName()
'            lrs![Info] = statusMsg
'            lrs.Update
'            lrs.Close
            
DoCmd.SetWarnings False
strinfo = statusMsg
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "')"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
'If IsNull(rstqueue!TitleCompleted) Then
With rstqueue
.Edit
If Not IsNull(rstqueue!TitleOutDateInQueue) Then rstqueue!TitleOutDateInQueue = Null
If Not IsNull(rstqueue!TitleRevewDateInQueue) Then rstqueue!TitleRevewDateInQueue = Null
!TitleCompleted = Now
!TitleCompUser = GetStaffID
!TitleLastEditDate = Now
!TitleLastEditUser = GetStaffID
!TitelMissingReson = False
.Update
End With
'End If
Set rstqueue = Nothing

DoCmd.SetWarnings False
'Dim DeletsSQl As Recordset
'DoCmd.RunSQL "delete * from SCRAQueueFiles"
DoCmd.RunSQL "DELETE * FROM TitleOrderFinal WHERE File=" & FileNumber

 'DoCmd.RunSQL "Delete * FROM  TitleOrderFinal  Where File =" & FileNumber
 
'If Not DeletsSQl.EOF Then
'DeletsSQl.Delete
'End If
'Set DeletsSQl = Nothing
DoCmd.SetWarnings True

If IsLoadedF("queTitelOrder") Then
NfileNumber = Forms!queTitelOrder!lstFiles.Column(0)


    If CurrentProject.AllForms("queTitelOrder").IsLoaded = True And FileNumber = NfileNumber Then
    
    Dim t As Recordset
    Set t = CurrentDb.OpenRecordset("TitleOrderHistory", dbOpenDynaset, dbSeeChanges)
    t.AddNew
    t!CaseFile = FileNumber
    t!Client = Forms!queTitelOrder!lstFiles.Column(2)
    t!Name = GetFullName()
    t!ClientOrder = Now
    t!ClientOrderC = 1
    t!DataDisiminated = True
    t!Stage = Forms!queTitelOrder!lstFiles.Column(6)
    t!Days = Forms!queTitelOrder!lstFiles.Column(8)
    t!dateOfStage = Forms!queTitelOrder!lstFiles.Column(7)
    t.Update
    Set t = Nothing
    
    Else
    
    Dim s As Recordset
    Set s = CurrentDb.OpenRecordset("TitleOrderHistory", dbOpenDynaset, dbSeeChanges)
    s.AddNew
    s!CaseFile = FileNumber
    s!Client = ClientShortName(Forms![Case List]!ClientID)
    s!Name = GetFullName()
    s!ClientOrder = Now
    s!ClientOrderC = 1
    s!DataDisiminated = True
    s!Stage = "Hard Order"
    s!Days = 0
    s!UpdateFromDM = 0
    s!dateOfStage = Date
    s.Update
    
    
    Set s = Nothing
    End If
Else
Dim X As Recordset
Set X = CurrentDb.OpenRecordset("TitleOrderHistory", dbOpenDynaset, dbSeeChanges)
X.AddNew
X!CaseFile = FileNumber
X!Client = ClientShortName(Forms![Case List]!ClientID)
X!Name = GetFullName()
X!ClientOrder = Now
X!ClientOrderC = 1
X!DataDisiminated = True
X!Stage = "Hard Order"
X!Days = 0
X!UpdateFromDM = 0
X!dateOfStage = Date
X.Update
Set X = Nothing
End If

DoCmd.Close acForm, "Print title Order"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"



End Sub

Private Sub Commcancel_Click()
DoCmd.Close acForm, "Print Title Order"

End Sub

Private Sub Form_Current()
 If ClientOrder = True Then
    Dim ctrl As Control
    For Each ctrl In Forms![Print Title Order].Controls
    If TypeOf ctrl Is Label Or TypeOf ctrl Is CheckBox Or TypeOf ctrl Is CommandButton Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Then
    If ctrl.Name = "label4" Or ctrl.Name = "label6" Or ctrl.Name = "RundownDate" Or ctrl.Name = "Label22" Or ctrl.Name = "Label36" Or ctrl.Name = "cmdcomplete" Or ctrl.Name = "Commcancel" Then
    ctrl.Visible = True
    Else
    ctrl.Visible = False
    End If
    End If
    Next
      ClientOrder = False
End If
    
 
 


 
 
 
''Dim rstTitleOrders As Recordset, ClientID As Long, JurisdictionID As Long, Conventional As Boolean, Agency As Boolean, LoanType As Integer
'
''Abstractor = ...
'
''ClientID = DLookup("clientid", "caseList", "filenumber=" & FileNumber)
''JurisdictionID = DLookup("jurisdictionid", "caselist", "filenumber=" & FileNumber)
''LoanType = DLookup("loantype", "fcdetails", "filenumber=" & FileNumber & " and current=" & True)
'
'Select Case LoanType
'Case 4, 5
'Agency = DLookup("AcerAgency", "clientlist", "clientid=" & ClientID)
'Case 1
'Conventional = DLookup("AcerConventional", "clientlist", "clientid=" & ClientID)
'End Select
'
'If Agency = True Or Conventional = True Then
'Set rstTitleOrders = CurrentDb.OpenRecordset("SELECT * FROM TitleOrders WHERE Filenumber=" & FileNumber & " AND abstractor=89", dbOpenDynaset, dbSeeChanges)
'If Not rstTitleOrders.EOF Then
'Abstractor = 89 'Acer update
'Else
'If IsNull([TitleThru]) And IsNull([TitleDue]) Then
'Abstractor = 89 'New Acer Order
'End If
'End If
'End If
'
'If Nz(Abstractor, 0) <> 89 Then
'If IsNull(DLookup("clientabstrator", "clientlist", "filenumber=" & ClientID)) Then
'Abstractor = DLookup("abstractor", "jurisdictionlist", "jurisdictionid=" & JurisdictionID)
'Else
'Abstractor = DLookup("clientabstrator", "clientlist", "filenumber=" & ClientID)
'End If
'End If
'
''txtAbstractor = DLookup("Name", "Abstractors", "ID=" & Abstractor)
'If DLookup("TitleAcer", "clientlist", "clientid=" & ClientID) = True Then
'cbxAbstractor.RowSource = "select id, Name from abstractors where id=" & Abstractor & " or id=89"
'cbxAbstractor = Abstractor
'
'If Abstractor = 89 Then
'Option21.Enabled = False
'Option3.Enabled = False
'cmdPrint.Enabled = False
'cmdView.Enabled = False
'End If
'Else
'cbxAbstractor.RowSource = "select id, Name from abstractors where id=" & Abstractor
'cbxAbstractor = Abstractor
'
'End If
'
'
'
'Select Case Weekday(Date)
'    Case vbMonday, vbTuesday, vbWednesday     ' due Wednesday, Thursday, Friday
'        TitleDue = Format$(DateAdd("d", 2, Date), "mmmm d, yyyy")
'    Case vbThursday     ' due Monday
'        TitleDue = Format$(DateAdd("d", 4, Date), "mmmm d, yyyy")
'    Case vbFriday       ' due Tuesday
'        TitleDue = Format$(DateAdd("d", 4, Date), "mmmm d, yyyy")
'    Case vbSaturday     ' due Tuesday
'        TitleDue = Format$(DateAdd("d", 3, Date), "mmmm d, yyyy")
'    Case vbSunday       ' due Tuesday
'        TitleDue = Format$(DateAdd("d", 2, Date), "mmmm d, yyyy")
'End Select
'DateRequired = Format$(TitleDue, "mmmm d, yyyy")
'
''Deactivated per Diane request with Acer ordering
'If Not IsNull([TitleThru]) Then
'  RundownDate = [TitleThru]
'  Order.DefaultValue = 2
'  Else
'  Order.DefaultValue = 4
'End If
'txtLiber = [Liber]
'txtFolio = [Folio]
'' SSN = MortgagorNames(0, 12)
'
'Select Case Nz([TitleSearchType])
'  Case "Full"
'    Order = 1
'  Case "Update"
'    Order = 2
'  Case "Rundown"
'    Order = 3
'  Case "2 Owner", ""
'    Order = 4
'End Select
'
'If Not IsNull([TitleThru]) Then
' ' RundownDate = [TitleThru]
'  Order = 2
'  Else
'  Order = 4
'End If

'If CurrentProject.AllForms("ForeclosureDetails").IsLoaded Then
'If Not IsNull(Forms!ForeclosureDetails!Disposition) Then
'MsgBox "Caution:  This file has a disposition.  Please check with your supervisor before ordering title", vbExclamation
'End If
'End If


End Sub
Private Sub cmdPrint_Click()

Dim statusMsg As String

On Error GoTo Err_cmdOK_Click

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

If Forms![Print Title Order]!Order = 2 Then
    If IsNull([TitleThru]) Then
    MsgBox "Must have good through date to order update."
    Exit Sub
    Else
    End If
    End If


'If SLS file, title must be ordered differently
If ClientID = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation
'Call DoReport("Title Order", acViewNormal)

Dim strSearchType As String

Select Case Order
    Case 1
       statusMsg = "Ordered full title search"
       strSearchType = "Full"
    Case 2
       statusMsg = "Ordered title rundown from " & Format$(RundownDate, "m/d/yyyy")
       strSearchType = "Update"
    Case 3
       statusMsg = "Ordered title rundown from present owner"
       strSearchType = "Rundown"
    Case 4
       statusMsg = "Ordered 2 owner search"
       strSearchType = "2 Owner"
    
End Select
'Linda 12_17_14

Forms![foreclosuredetails]!TitleSearchType = strSearchType

Call DoReport("Title Order", acViewNormal)

If chForeclosure Then
    If MsgBox("Update Title Ordered = " & Format$(Date, "m/d/yyyy") & vbNewLine & "and clear Title Received and Title Through" & vbNewLine & "and add to status?", vbYesNo) = vbYes Then
        Forms![Print Title Order]!TitleOrder = Now()
        Forms![Print Title Order]!TitleDue = Format$(DateRequired, "m/d/yyyy")
        If Not IsNull(Forms![Print Title Order]!TitleBack) Then Forms![Print Title Order]!TitleBack = Null
        If Not IsNull(Forms![Print Title Order]!TitleThru) Then Forms![Print Title Order]!TitleThru = Null
        If Not IsNull(Forms![Print Title Order]!TitleReviewToClient) Then Forms![Print Title Order]!TitleReviewToClient = Null
        AddStatus FileNumber, Now(), statusMsg
    End If
End If

Forms![Print Title Order]!TitleSearchType = strSearchType
cmdCancel.Caption = "Close"

If IsLoadedF("wizReferralII") = True Then
If RSII = True Then
Forms!wizreferralII.Requery
End If
End If

Exit_cmdOK_Click:
DoCmd.Close acForm, "Print Title Order"
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    

End Sub

Private Sub cmdView_Click()

Dim statusMsg As String

On Error GoTo Err_cmdOK_Click

'If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

If Forms![Print Title Order]!Order = 2 Then
    If IsNull([TitleThru]) Then
    MsgBox "Must have good through date to order update."
    Exit Sub
    Else
    End If
    End If


'If SLS file, title must be ordered differently
If ClientID = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation

'Forms![ForeclosureDetails]!TitleSearchType = strSearchType

'Call DoReport("Title Order", acPreview)

Dim strSearchType As String

Select Case Order
    Case 1
       statusMsg = "Ordered full title search"
       strSearchType = "Full"
    Case 2
       statusMsg = "Ordered title rundown from " & Format$(RundownDate, "m/d/yyyy")
       strSearchType = "Update"
    Case 3
       statusMsg = "Ordered title rundown from present owner"
       strSearchType = "Rundown"
    Case 4
       statusMsg = "Ordered 2 owner search"
       strSearchType = "2 Owner"
    
End Select
'Linda 12_17_14
Forms![foreclosuredetails]!TitleSearchType = strSearchType

Call DoReport("Title Order", acPreview)

If chForeclosure Then
    If MsgBox("Update Title Ordered = " & Format$(Date, "m/d/yyyy") & vbNewLine & "and clear Title Received and Title Through" & vbNewLine & "and add to status?", vbYesNo) = vbYes Then
        Forms![Print Title Order]!TitleOrder = Now()
        Forms![Print Title Order]!TitleDue = DateRequired
        If Not IsNull(Forms![Print Title Order]!TitleBack) Then Forms![Print Title Order]!TitleBack = Null
        If Not IsNull(Forms![Print Title Order]!TitleThru) Then Forms![Print Title Order]!TitleThru = Null
        If Not IsNull(Forms![Print Title Order]!TitleReviewToClient) Then Forms![Print Title Order]!TitleReviewToClient = Null
        AddStatus FileNumber, Now(), statusMsg
    End If
End If



Forms![Print Title Order]!TitleSearchType = strSearchType
cmdCancel.Caption = "Close"

If IsLoadedF("wizReferralII") = True Then
If RSII = True Then
Forms!wizreferralII.Requery
End If
End If

Exit_cmdOK_Click:
DoCmd.Close acForm, "Print Title Order"
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    

End Sub

Private Sub cmdAcrobat_Click()

Dim statusMsg As String, rstAcer As Recordset, rstTitleOrders As Recordset, rstFCdetails As Recordset

On Error GoTo Err_cmdOK_Click

'If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

If Forms![Print Title Order]!Order = 2 Then
    If IsNull([TitleThru]) Then
    MsgBox "Must have good through date to order update."
    
    DoCmd.Close acForm, "Print Title Order"
    Exit Sub
    Else
    End If
    End If


'If SLS file, title must be ordered differently
If ClientID = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation

'Linda 12_17_14
Dim strSearchType As String

Select Case Order
    Case 1
       statusMsg = "Ordered full title search"
       strSearchType = "Full"
    Case 2
       statusMsg = "Ordered title rundown from " & Format$(RundownDate, "m/d/yyyy")
       strSearchType = "Update"
    Case 3
       statusMsg = "Ordered title rundown from present owner"
       strSearchType = "Rundown"
    Case 4
       statusMsg = "Ordered 2 owner search"
       strSearchType = "2 Owner"

End Select

Forms![foreclosuredetails]!TitleSearchType = strSearchType

Call DoReport("Title Order", -2)

'
'If chForeclosure Then
'    If MsgBox("Update Title Ordered = " & Format$(Date, "m/d/yyyy") & vbNewLine & "and clear Title Received and Title Through" & vbNewLine & "and add to status?", vbYesNo) = vbYes Then
'        Forms![Print Title Order]!TitleOrder = Now()
'        Forms![Print Title Order]!TitleDue = DateRequired
'        Forms![Print Title Order]!TitleBack = Null
'        Forms![Print Title Order]!TitleThru = Null
'        Forms![Print Title Order]!TitleReviewToClient = Null
'        AddStatus FileNumber, Now(), statusMsg
'
'
'
'If Abstractor = 89 Then 'Acer
'Dim Last As String, First As String, Address As String, City As String, State As String
'
'Set rstFCdetails = CurrentDb.OpenRecordset("SELECT * FROM fcdetails WHERE FileNumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
'With rstFCdetails
'Last = !PrimaryLastName
'First = !PrimaryFirstName
'Address = !PropertyAddress
'City = !City
'State = !State
'
'.Close
'End With
'
'Set rstAcer = CurrentDb.OpenRecordset("SELECT * FROM RA_Title_Ordered", dbOpenDynaset, dbSeeChanges)
'
'With rstAcer
'.AddNew
'!RA_Number = FileNumber
'!Last_Name = Last
'!First_Name = First
'!Address = Address
'!City = City
'!State = State
'!DueDate = DateRequired
'!Jurisdiction = DLookup("jurisdictionid", "caseList", "filenumber=" & FileNumber)
'!Client = DLookup("clientid", "caseList", "filenumber=" & FileNumber)
'!RequestDate = Now
'.Update
'.Close
'End With
'End If
'
'Set rstTitleOrders = CurrentDb.OpenRecordset("SELECT * FROM TitleOrders", dbOpenDynaset, dbSeeChanges)
'
'With rstTitleOrders
'.AddNew
'!FileNumber = FileNumber
'!Abstractor = Abstractor
'!DateOrdered = Now
'!OrderedBy = GetStaffID
'.Update
'.Close
'End With
'
'End If
'End If

'Forms![Print Title Order]!TitleSearchType = strSearchType
'cmdCancel.Caption = "Close"

If IsLoadedF("wizReferralII") = True Then
If RSII = True Then
Forms!wizreferralII.Requery
End If
End If

'Forms!DocsWindow!cmdAddDoc.SetFocus = True


Exit_cmdOK_Click:
'DoCmd.Close acForm, "Print Title Order"
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Command63_Click()
On Error GoTo Err_Command63_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command63_Click:
    Exit Sub

Err_Command63_Click:
    MsgBox Err.Description
    Resume Exit_Command63_Click
    
End Sub
