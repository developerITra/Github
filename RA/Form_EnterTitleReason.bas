VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterTitleReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btn1_AfterUpdate()
If btn1 = True Then
btn2 = False
btn3 = False
btn5 = False
btn4 = False
End If
End Sub

Private Sub btn2_AfterUpdate()
If btn2 = True Then
btn1 = False
btn3 = False
btn5 = False
btn4 = False
End If
End Sub

Private Sub btn3_AfterUpdate()
If btn3 = True Then
btn1 = False
btn2 = False
btn5 = False
btn4 = False
End If
End Sub

Private Sub btn4_AfterUpdate()
If btn4 = True Then
Other.Enabled = True
btn1 = False
btn2 = False
btn3 = False
btn5 = False
Else
Other.Enabled = False
End If
End Sub

Private Sub btn5_AfterUpdate()
If btn5 = True Then
btn1 = False
btn2 = False
btn4 = False
btn3 = False
End If
End Sub

Private Sub cmdClientOrder_Click()
ClientOrder = True
'If CurrentProject.AllForms("queTitelOrder").IsLoaded = True Then
'If Forms!queTitelOrder!lstFiles.Column(10) = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation ' by Diane request 09/12/14
'DoCmd.OpenForm "Print Title Order" ', , , "Caselist.FileNumber=" & FileNumber, , , acViewNormal
'Forms![Print Title Order]!FileNumber = FileNumber
'Else
'If Forms![Case List]!ClientID = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation
DoCmd.OpenForm "Print Title Order"

Call FullTitleOrederPrintFormWithData

DoCmd.Close acForm, "EnterTitleReason"
End Sub

Private Sub cmdComplete_Click()
Dim ctr As Integer, rstwizqueue As Recordset, Other As String, rstdocs As Recordset, JrlTxt As String, rstFCdetailsCurrent As Recordset, rstFCdetailsPrior As Recordset
Dim s As Recordset
Dim lrs As Recordset
Dim t As Recordset
Dim Jrs As Recordset
Dim AFileNumber As Long

DoCmd.Hourglass True
Dim rstqueue As Recordset
ctr = 0
If btn1 = True Then
ctr = ctr + 1
End If
If btn2 = True Then
ctr = ctr + 1
End If
If btn3 = True Then
ctr = ctr + 1
End If
If btn5 = True Then
ctr = ctr + 1
End If
If btn4 = True Then
ctr = ctr + 1
End If

If ctr = 0 Then
MsgBox "Please select a reason", vbCritical
Exit Sub
End If

If btn1 = True Then
JrlTxt = " Client Orders Own? "
    Set rstdocs = CurrentDb.OpenRecordset("TitleDocumentMissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Client Orders Own"
    !DocNeededby = GetStaffID
    .Update
    End With
    
  
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !TitleLastEditDate = Now
    !TitleLastEditUser = GetStaffID
    !TitelMissingReson = True
    .Update
    End With
    Set rstqueue = Nothing
      
End If

If btn2 = True Then
    If JrlTxt = "" Then
    JrlTxt = " Title is within 30 days "
    Else: JrlTxt = JrlTxt & ", Title is within 30 days"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("TitleDocumentMissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Title is within 30 days"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !TitleLastEditDate = Now
    !TitleLastEditUser = GetStaffID
    !TitelMissingReson = True
    .Update
    End With
    Set rstqueue = Nothing
  End If
    
    If btn3 = True Then
    If JrlTxt = "" Then
    JrlTxt = " Service Transferred "
    Else: JrlTxt = JrlTxt & ", Service Transferred"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("TitleDocumentMissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Service Transferred"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !TitleLastEditDate = Now
    !TitleLastEditUser = GetStaffID
    !TitelMissingReson = True
    .Update
    End With
    Set rstqueue = Nothing
    End If
    
If btn5 = True Then

    If JrlTxt = "" Then
    JrlTxt = " Prior order not yet received "
    Else: JrlTxt = JrlTxt & ", Prior order not yet received"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("TitleDocumentMissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Prior order not yet received"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !TitleLastEditDate = Now
    !TitleLastEditUser = GetStaffID
    !TitelMissingReson = True
    .Update
    End With
    Set rstqueue = Nothing
    
End If
If btn4 = True Then
If JrlTxt = "" Then
    JrlTxt = Me!Other
    Else: JrlTxt = JrlTxt & ", " & Me!Other
    End If
    Set rstdocs = CurrentDb.OpenRecordset("TitleDocumentMissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Other"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !TitleLastEditDate = Now
    !TitleLastEditUser = GetStaffID
    !TitelMissingReson = True
    .Update
    End With
    Set rstqueue = Nothing
    
End If

If CurrentProject.AllForms("queTitelOrder").IsLoaded = True Then
    AFileNumber = Forms!queTitelOrder!lstFiles.Column(0)
    If FileNumber = AFileNumber Then
       
    Set t = CurrentDb.OpenRecordset("TitleOrderHistory", dbOpenDynaset, dbSeeChanges)
    t.AddNew
    t!CaseFile = FileNumber
    t!Client = Forms!queTitelOrder!lstFiles.Column(2)
    t!Name = GetFullName()
    t!TitleNotNeeded = Now
    t!TitleNotNeededC = 1
    t!DataDisiminated = True
    t!Stage = Forms!queTitelOrder!lstFiles.Column(6)
    t!Days = Forms!queTitelOrder!lstFiles.Column(8)
    t!UpdateFromDM = 0
    t!dateOfStage = Forms!queTitelOrder!lstFiles.Column(7)
    t.Update
    Set t = Nothing
    
    Else
    
    
    Set s = CurrentDb.OpenRecordset("TitleOrderHistory", dbOpenDynaset, dbSeeChanges)
    s.AddNew
    s!CaseFile = FileNumber
    s!Client = ClientShortName(Forms![Case List]!ClientID)
    s!Name = GetFullName()
    s!TitleNotNeeded = Now
    s!TitleNotNeededC = 1
    s!DataDisiminated = True
    s!Stage = "Hard Order"
    s!Days = 0
    s!UpdateFromDM = 0
    s!dateOfStage = Now()
    s.Update
    Set s = Nothing
    End If

Else
Set s = CurrentDb.OpenRecordset("TitleOrderHistory", dbOpenDynaset, dbSeeChanges)
s.AddNew
s!CaseFile = FileNumber
s!Client = ClientShortName(Forms![Case List]!ClientID)
s!Name = GetFullName()
s!TitleNotNeeded = Now
s!TitleNotNeededC = 1
s!DataDisiminated = True
s!Stage = "Hard Order"
s!Days = 0
s!UpdateFromDM = 0
s!dateOfStage = Now()
s.Update
Set s = Nothing

End If



If CurrentProject.AllForms("queTitelOrder").IsLoaded = True Then
  AFileNumber = Forms!queTitelOrder!lstFiles.Column(0)
  If FileNumber = AFileNumber Then
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  lrs.AddNew
'  lrs![FileNumber] = FileNumber
'  lrs![JournalDate] = Now
'  lrs![Who] = GetFullName()
'  lrs![Info] = "Title did not order because :  " & JrlTxt & " - For " & Forms!queTitelOrder!lstFiles.Column(6) & vbCrLf
'  lrs![Color] = 1
' ' lrs![Warning] = 100
'  lrs.Update
'  'imgWarning.Picture = dbLocation & "papertray.emf"
'' imgWarning.Visible = True
'    lrs.Close
'    Set lrs = Nothing
    DoCmd.SetWarnings False
    strinfo = "Title did not order because :  " & JrlTxt & " - For " & Forms!queTitelOrder!lstFiles.Column(6) & vbCrLf
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Else
'  Set Jrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  Jrs.AddNew
'  Jrs![FileNumber] = FileNumber
'  Jrs![JournalDate] = Now
'  Jrs![Who] = GetFullName()
'  Jrs![Info] = "Title did not order because :  " & JrlTxt & " - For " & " Hard Order" & vbCrLf
'  Jrs![Color] = 1
'  Jrs.Update
'  Jrs.Close
'  Set Jrs = Nothing
  
    DoCmd.SetWarnings False
    strinfo = "Title did not order because :  " & JrlTxt & " - For " & " Hard Order" & vbCrLf
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
  End If
Else
'Set Jrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'Jrs.AddNew
'Jrs![FileNumber] = FileNumber
'Jrs![JournalDate] = Now
'Jrs![Who] = GetFullName()
'Jrs![Info] = "Title did not order because :  " & JrlTxt & " - For " & " Hard Order" & vbCrLf
'Jrs![Color] = 1
'Jrs.Update
'Jrs.Close
'Set Jrs = Nothing

    DoCmd.SetWarnings False
    strinfo = "Title did not order because :  " & JrlTxt & " - For " & " Hard Order" & vbCrLf
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
End If

DoCmd.SetWarnings False
'Dim DeletsSQl As Recordset
''heresarabd
DoCmd.RunSQL "DELETE * FROM TitleOrderFinal WHERE File=" & FileNumber

'Set DeletsSQl = CurrentDb.OpenRecordset("Select * FROM  TitleOrderFinal  Where File =" & FileNumber, dbOpenDynaset, dbSeeChanges)
'DeletsSQl.FindFirst
'DeletsSQl.Delete
'DeletsSQl.Update
'Set DeletsSQl = Nothing


DoCmd.SetWarnings True


If CurrentProject.AllForms("queTitelOrder").IsLoaded = True Then
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryTitle_All", dbOpenDynaset, dbSeeChanges)
If rstqueue.EOF Then
    cntr = 0
    Else
    rstqueue.MoveLast
    cntr = rstqueue.RecordCount
End If
Set rstqueue = Nothing

'
'
'
'Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryTitle_All", dbOpenDynaset, dbSeeChanges)
'Do Until rstqueue.EOF
'cntr = cntr + 1
'rstqueue.MoveNext
'Loop
'Forms!queTitelOrder!QueueCount = cntr
'Set rstqueue = Nothing
End If
'DoCmd.Hourglass False

'MsgBox "   Completed    "
'DoCmd.Hourglass True
If CurrentProject.AllForms("queTitelOrder").IsLoaded = True Then
Forms!queTitelOrder!lstFiles.Requery
Forms!queTitelOrder.Requery

Forms!queTitelOrder!lstFiles = Null


End If

DoCmd.Hourglass False
Call ReleaseFile(FileNumber)
'Call RestartWaitingCompletionUpdate(FileNumber)
'DoCmd.Close acForm, Me.Name
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"


DoCmd.Close


End Sub
Private Sub Command22_Click()
'btn1.Visible = True
btn2.Visible = True
btn3.Visible = True
btn4.Visible = True
btn5.Visible = True
Other.Visible = True
'Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
cmdComplete.Visible = True

    
End Sub

Private Sub cmdTitleW_Click()
'If Me.Dirty Then
'DoCmd.RunCommand acCmdSaveRecord
'End If
'If Forms![Case List]!ClientID = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation
'If CurrentProject.AllForms("queTitelOrder").IsLoaded = True Then
'If Forms!queTitelOrder!lstFiles.Column(10) = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation  ' removed by Diane request 09/12/14
'DoCmd.OpenForm "Print Title Order" ', , , "Caselist.FileNumber=" & FileNumber, , , acViewNormal
'Forms![Print Title Order]!FileNumber = FileNumber
'Else
'If Forms![Case List]!ClientID = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation
DoCmd.OpenForm "Print Title Order"

Call FullTitleOrederPrintFormWithData

DoCmd.Close acForm, "EnterTitleReason"

  
End Sub




Private Sub Form_Current()
'Dim T As Recordset

'Set T = CurrentDB.OpenRecordset ("Select * From FCdetails Where FileNubmer="

If Not IsNull(Forms!foreclosuredetails!TitleOrder) And IsNull(Forms!foreclosuredetails!TitleBack) Then
AcerNote.Visible = True
End If

If IsLoaded("queTitelOrder") = True Then
If Forms!queTitelOrder!lstFiles.Column(6) = "Last Post Sale                                    " Then
    If LastPostSale = False Then Command22.Enabled = False
End If
End If

End Sub


Private Sub FullTitleOrederPrintFormWithData()


Forms![Print Title Order]!FileNumber = Forms![foreclosuredetails]!FileNumber
Forms![Print Title Order]!TitleOrder = Forms![foreclosuredetails]!TitleOrder
Forms![Print Title Order]!TitleThru = Forms![foreclosuredetails]!TitleThru
Forms![Print Title Order]!TitleDue = Forms![foreclosuredetails]!TitleDue
Forms![Print Title Order]!TitleSearchType = Forms![foreclosuredetails]!TitleSearchType
Forms![Print Title Order]!DOTdate = Forms![foreclosuredetails]!DOTdate
Forms![Print Title Order]!ClientID = DLookup("clientid", "caseList", "filenumber=" & FileNumber)
Forms![Print Title Order]!JurisdictionID = Forms![foreclosuredetails]!JurisdictionID
Forms![Print Title Order]!LoanType = DLookup("loantype", "fcdetails", "filenumber=" & FileNumber & " and current=" & True)



Dim rstTitleOrders As Recordset, ClientID As Long, JurisdictionID As Long, Conventional As Boolean, Agency As Boolean, LoanType As Integer
Dim Abstractor As Integer
Dim TitleDue As Date
Dim DateRequired As Date


'Abstractor = ...

ClientID = DLookup("clientid", "caseList", "filenumber=" & FileNumber)
JurisdictionID = DLookup("jurisdictionid", "caselist", "filenumber=" & FileNumber)
LoanType = DLookup("loantype", "fcdetails", "filenumber=" & FileNumber & " and current=" & True)

Select Case LoanType
Case 4, 5
Agency = DLookup("AcerAgency", "clientlist", "clientid=" & ClientID)
Case 1
Conventional = DLookup("AcerConventional", "clientlist", "clientid=" & ClientID)
End Select

If Agency = True Or Conventional = True Then
Set rstTitleOrders = CurrentDb.OpenRecordset("SELECT * FROM TitleOrders WHERE Filenumber=" & FileNumber & " AND abstractor=89", dbOpenDynaset, dbSeeChanges)
    If Not rstTitleOrders.EOF Then
    Abstractor = 89 'Acer update
    Else
        If IsNull(Forms![foreclosuredetails]![TitleThru]) And IsNull(Forms![foreclosuredetails]![TitleDue]) Then
        Abstractor = 89 'New Acer Order
        End If
    End If
End If

If Nz(Abstractor, 0) <> 89 Then
    If IsNull(DLookup("clientabstrator", "clientlist", "filenumber=" & ClientID)) Then
    Abstractor = DLookup("abstractor", "jurisdictionlist", "jurisdictionid=" & JurisdictionID)
    Else
    Abstractor = DLookup("clientabstrator", "clientlist", "filenumber=" & ClientID)
    End If
End If
Forms![Print Title Order]!Abstractor = Abstractor
    If DLookup("TitleAcer", "clientlist", "clientid=" & ClientID) = True Then
    Forms![Print Title Order]!cbxAbstractor.RowSource = "select id, Name from abstractors where id=" & Abstractor & " or id=89"
    Forms![Print Title Order]!cbxAbstractor = Abstractor
    
        If Abstractor = 89 Then
            Forms![Print Title Order]!Order.Locked = False
'           If Forms![Print Title Order]!Option21.Enabled = True Then Forms![Print Title Order]!Option21.Enabled = False
'           ' If Forms![Print Title Order]!Option3.Enabled = True Then Forms![Print Title Order]!Option3.Enabled = False
'                If Forms![Print Title Order]!cmdPrint.Enabled = True Then Forms![Print Title Order]!cmdPrint.Enabled = False
'                    If Forms![Print Title Order]!cmdView.Enabled = True Then Forms![Print Title Order]!cmdView.Enabled = False
        End If
    Else
    Forms![Print Title Order]!cbxAbstractor.RowSource = "select id, Name from abstractors where id=" & Abstractor
    Forms![Print Title Order]!cbxAbstractor = Abstractor
    
    End If



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
Forms![Print Title Order]!DateRequired = Format$(TitleDue, "mmmm d, yyyy")

'Deactivated per Diane request with Acer ordering
If Not IsNull(Forms![foreclosuredetails]!TitleThru) Then
  Forms![Print Title Order]!RundownDate = Forms![foreclosuredetails]!TitleThru
  Forms![Print Title Order]!Order.DefaultValue = 2
  Else
  Forms![Print Title Order]!Order.DefaultValue = 4
End If
Forms![Print Title Order]!txtLiber = Forms![foreclosuredetails]![Liber]
Forms![Print Title Order]!txtFolio = Forms![foreclosuredetails]![Folio]
' SSN = MortgagorNames(0, 12)

Select Case Nz(Forms![foreclosuredetails]![TitleSearchType])
  Case "Full"
    Forms![Print Title Order]!Order = 1
  Case "Update"
    Forms![Print Title Order]!Order = 2
  Case "Rundown"
    Forms![Print Title Order]!Order = 3
  Case "2 Owner", ""
    Forms![Print Title Order]!Order = 4
End Select

If Not IsNull(Forms![foreclosuredetails]![TitleThru]) Then
 ' RundownDate = [TitleThru]
  Forms![Print Title Order]!Order = 2
  Else
  Forms![Print Title Order]!Order = 4
End If


'End If
Forms![Print Title Order]!RSII = True
End Sub

