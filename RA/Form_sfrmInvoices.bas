VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCapture_Click()
Dim NewFilespec As String

If Not IsNull(Filespec) Then
    If MsgBox("There is already an image associated with this invoice.  Do you want to replace it?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
End If
NewFilespec = CaptureImage()
If NewFilespec <> "" Then Filespec = NewFilespec

End Sub



Private Sub ComFilter_Click()
If ComFilter.Caption = "All Invoices" Then

    EscFilter = True
    ComFilter.Caption = "Escro Invoice"
    Me.FilterOn = False
    Me.Filter = ""
Else
    EscFilter = False
    ComFilter.Caption = "All Invoices"
    Me.Filter = "Invoiceid = " & Nz(Forms!QueueAccounESC.lstFiles.Column(17))
    Me.FilterOn = True
End If
    
'
'Me.Filter = ""
'Me.FilterOn = False
'Me.Filter = ""
'ComFilter.Caption = "Escro Invoice"
'Me.Form.Requery
'Else
'Me.Filter = "Invoiceid = " & Nz(Forms!QueueAccounESC.lstFiles.Column(17))
'Me.FilterOn = True
'ComFilter.Caption = "All Invoices"
'End If
End Sub

Private Sub DatePaid_DblClick(Cancel As Integer)
DatePaid = Date
PaidAmount.SetFocus
End Sub

Private Function CaptureImage() As String
Dim InputFilespec As String, fileextension As String, Path As String, FileName As String, newfilename As String, i As Integer, Prompt As String, DocType As String
Const GroupCode = "I"
Const GroupDelimiter = ";"

On Error GoTo Err_cmdCapture_Click

CaptureImage = ""

InputFilespec = OpenFile(Me.Parent)
If InputFilespec = "" Then Exit Function

For i = Len(InputFilespec) To 0 Step -1
    If Asc(Mid$(InputFilespec, i, 1)) <> 0 Then Exit For
Next i
If i = 0 Then
    MsgBox "Invalid file specification: " & InputFilespec, vbCritical
    Exit Function
End If
InputFilespec = Left$(InputFilespec, i)

For i = Len(InputFilespec) To 0 Step -1
    If Mid$(InputFilespec, i, 1) = "." Then Exit For
Next i
If i = 0 Then
    MsgBox "Invalid file specification: " & InputFilespec, vbCritical
    Exit Function
End If
fileextension = Mid$(InputFilespec, i)

For i = Len(InputFilespec) To 0 Step -1
    If Mid$(InputFilespec, i, 1) = "\" Then Exit For
Next i
If i = 0 Then
    MsgBox "Invalid file specification: " & InputFilespec, vbCritical
    Exit Function
End If

Path = Left$(InputFilespec, i)
FileName = Mid$(InputFilespec, i + 1)
newfilename = GroupDelimiter & GroupCode & GroupDelimiter & Format$(Now(), "yyyymmdd hhnnss") & fileextension

Do While Dir$(DocLocation & DocBucket(Me.Parent!FileNumber) & "\" & Me.Parent!FileNumber & "\" & newfilename) <> ""
    Wait 2
    newfilename = GroupDelimiter & GroupCode & GroupDelimiter & Format$(Now(), "yyyymmdd hhnnss") & fileextension
    Exit Do
Loop

CaptureImage = newfilename  ' save in the record
FileCopy InputFilespec, DocLocation & DocBucket(Me.Parent!FileNumber) & "\" & Me.Parent!FileNumber & "\" & newfilename
If MsgBox("New document " & newfilename & " accepted.  OK to delete " & InputFilespec & "?", vbQuestion + vbYesNo) = vbYes Then Kill InputFilespec

Exit_cmdCapture_Click:
    Exit Function

Err_cmdCapture_Click:
    If Err.Number = 76 Then     ' path not found
        MkDir DocLocation & DocBucket(Me.Parent!FileNumber) & "\" & Me.Parent!FileNumber & "\"
        Resume
    Else
        CaptureImage = ""
        MsgBox Err.Description
        Resume Exit_cmdCapture_Click
    End If
End Function

Private Sub cmdView_Click()

On Error GoTo Err_cmdView_Click

If IsNull(Filespec) Then
    MsgBox "There is no image associated with this invoice", vbInformation
    Exit Sub
End If
StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & Filespec

Exit_cmdView_Click:
    Exit Sub

Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click
    
End Sub

Private Sub cmdAdd_Click()
Dim rstInv As Recordset, rstDoc As Recordset, Filespec As String, InvoiceNumber As String

On Error GoTo Err_cmdAdd_Click

If StaffID = 0 Then Call GetLoginName

Filespec = CaptureImage()
If Filespec = "" Then
    MsgBox "You must capture the image in order to create an invoice.", vbCritical
    Exit Sub
End If
    
InvoiceNumber = InputBox$("Enter Invoice Number:")
If InvoiceNumber = "" Then Exit Sub

Set rstInv = CurrentDb.OpenRecordset("Invoices", dbOpenDynaset, dbSeeChanges)
With rstInv
    .AddNew
    !FileNumber = Me.Parent!FileNumber
  
    Select Case Forms![Case List]!SCRAID
    Case "AccPSAdvanced"
     rstInv!InvoiceType = 711
    Case "AccLitig"
     rstInv!InvoiceType = 710
    Case Else
     rstInv!InvoiceType = 0
    End Select

    !InvoiceNumber = InvoiceNumber
    !AdditionalInvoiceNeeded = 0
    !Filespec = Filespec
    !DateSent = Now()
    !CreatedBy = StaffID
    .Update
    .Close
End With
Me.Requery

'Commented by JAE 10-30-2014 'Document Speed'
'Set rstDoc = CurrentDb.OpenRecordset("DocIndex", dbOpenDynaset, dbSeeChanges)
'rstDoc.AddNew
'rstDoc!FileNumber = FileNumber
'rstDoc!DocTitleID = 0
'rstDoc!DocGroup = "I"
'rstDoc!StaffID = GetStaffID()
'rstDoc!DateStamp = Now()
'rstDoc!Filespec = Filespec
'rstDoc!Notes = Filespec
'rstDoc.Update
'rstDoc.Close

DoCmd.SetWarnings False
Dim strSQLValues As String: strSQLValues = ""
Dim strSQL As String: strSQL = ""
strSQL = ""
strSQLValues = FileNumber & "," & 0 & ",'" & "I" & "'," & GetStaffID() & ",'" & Now() & "','" & Replace(Filespec, "'", "''") & "','" & Replace(Filespec, "'", "''") & "'"
'Debug.Print strSQLValues
strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
'Debug.Print strSQL
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings True

Open dbLocation & "Invoice.log" For Append As #1
Print #1, Format$(Now(), "mm/dd/yyyy hh:nn am/pm") & "|" & GetLoginName() & "|" & Me.Parent!FileNumber & "|" & Filespec
Close #1


Select Case Forms![Case List]!SCRAID
    Case "AccPSAdvanced"
    Case "AccLitig"
    Case Else
    DoCmd.OpenForm "UnbilledEvents"
    Forms!unbilledevents!FileNumber = Forms![Case List]!FileNumber
    Forms!unbilledevents!lstBillingReasons.RowSource = "SELECT BillingReasonsFCarchive.ID, BillingReasonsFC.Reason, BillingReasonsFCarchive.Date AS [Date] FROM BillingReasonsFCarchive INNER JOIN BillingReasonsFC ON BillingReasonsFCarchive.billingreasonid=BillingReasonsFC.ID WHERE (((BillingReasonsFCarchive.FileNumber)=" & Forms!unbilledevents!FileNumber & " AND Invoiced is null));"
End Select


Exit_cmdAdd_Click:
    Exit Sub

Err_cmdAdd_Click:
    MsgBox Err.Description
    Resume Exit_cmdAdd_Click
    
End Sub

Private Sub Form_Current()

InvoiceType.Locked = (Nz(InvoiceType) <> 0)
DateSent.Locked = Not (IsNull(DateSent) Or DateSent = Date)
InvoiceAmount.Locked = Not IsNull(InvoiceAmount)
DatePaid.Locked = Not (PrivReceivePayments And (IsNull(DatePaid) Or DatePaid = Date))
PaidAmount.Locked = Not (PrivReceivePayments And IsNull(PaidAmount))
AdditionalInvoiceNeeded.Locked = Not PrivAccounting


UpdateInvoiceTypeList


If Forms![Case List]!SCRAID = "AccEsc" Then
Me.ComFilter.Visible = True
If EscFilter = False Then
'And Me.ComFilter.Caption = "Invoices" Then
Me.Filter = "Invoiceid = " & Nz(Forms!QueueAccounESC.lstFiles.Column(17))
Me.FilterOn = True
End If
End If
End Sub

Private Sub UpdateInvoiceTypeList()

If (NewRecord) Then
  InvoiceType.RowSource = "SELECT InvoiceTypes.ID, InvoiceTypes.InvoiceType FROM InvoiceTypes WHERE active = true ORDER BY InvoiceTypes.ID"
Else
  InvoiceType.RowSource = "SELECT InvoiceTypes.ID, InvoiceTypes.InvoiceType FROM InvoiceTypes ORDER BY InvoiceTypes.ID"
End If

Me.InvoiceType.Requery

End Sub

Private Sub Form_Open(Cancel As Integer)
EscFilter = False
Me.AllowEdits = PrivAccounting Or PrivReceivePayments
cmdAdd.Enabled = PrivAccounting
cmdPrepareInvoice.Enabled = PrivAccounting
cmdCapture.Enabled = PrivAccounting
If PrivWriteOff Then WriteOffAmount.Locked = False


If FileReadOnly Or EditDispute Then

    Dim ctl As Control
    Dim lngI As Long
    Dim bSkip As Boolean

    For Each ctl In Form.Controls
    Select Case ctl.ControlType
    Case acTextBox, acComboBox, acListBox, acOptionGroup, acCheckBox, acSubform, acOptionButton
         bSkip = False
    
            If Not bSkip Then ctl.Locked = True
            
            
    Case acCommandButton

            If Not bSkip Then ctl.Enabled = False
         
       
    End Select
    Next
End If




End Sub





Private Sub PaidAmount_AfterUpdate()

If StaffID = 0 Then Call GetLoginName
ReceivedBy = StaffID
If IsNull(DatePaid) Then DatePaid = Date

Dim intDisp As Variant
Dim intRPAmtRecClient As Variant

If ([InvoiceType] = 101 Or [InvoiceType] = 102) Then  ' Foreclosure Invoice
 intDisp = DLookup("[Disposition]", "[FCDetails]", "[FileNumber] = " & Me.FileNumber & " and Current = true")

 If (intDisp = 1) Then  ' disposition = "Buy-In"
   GetAmount ("How much is for real property tax?")
   If (FeeAmount > 0) Then
     intRPAmtRecClient = FeeAmount + Nz(DLookup("[RPAmtReceivedClient]", "[FCDetails]", "[FileNumber] = " & Me.FileNumber & " and Current = true"), 0)

     DoCmd.SetWarnings False  'update amt received from client
     DoCmd.RunSQL ("UPDATE FCDetails set RPAmtReceivedClient = " & intRPAmtRecClient & " WHERE [FileNumber] = " & Me.FileNumber & " and Current = true")
     DoCmd.SetWarnings True

   End If
 End If
End If

If Not IsNull(BTnumber) And PaidAmount >= 0 Then
Call cmd_EmailBT_Click(46, Me.FileNumber, " The BT paid for the file number: " & Me.FileNumber & "  , BT Check number: " & Trim(Forms![Case List]!sfrmInvoices.Form!BTChecNumber) & "  and BT number: " & Trim(Forms![Case List]!sfrmInvoices.Form!BTnumber) & ". The amount is: $" & Format(Forms![Case List]!sfrmInvoices.Form!PaidAmount, "###,##0.00"), " BT paid " & Me.FileNumber & " - " & ClientShortName(Forms![Case List]!ClientID) & ", " & Forms![Case List]!PrimaryDefName)
End If

'ESC 1 step proceduer SA 4/5/15

Dim clientShor As String
Dim strSQL As String
Dim DateShow As Date

clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)

DoCmd.SetWarnings False

Select Case Forms![Case List]![CaseType]


    Case "Foreclosure"
         intDisp = DLookup("[Disposition]", "[FCDetails]", "[FileNumber] = " & Me.FileNumber & " and Current = true")
           Select Case intDisp
           
                    Case 1, 2
                    
                    DateShow = DLookup("[AuditRat]", "[FCdetails]", "[FileNumber] = " & Me.FileNumber & " and Current = true")
                    If IsNull(DateShow) Then DateShow = DatePaid
               
                    strSQL = "Insert into Accou_EscQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, StaffID, StaffName, CaseType, Disposition, " & _
                    " InvoiceID,DateInvoicedPaid,InvoiceType,InvoiceNumber,BTnumber,InvoiceAmount,PaidAmount,DateShouldShowUp,WhoInvoiced,DatePaid) Values (" & FileNumber & _
                    ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & clientShor & " '," & Forms![Case List]!ClientID & " ,#" & Now() & "# ," & GetStaffID() & _
                    ", '" & GetFullName() & "', '" & Forms![Case List]![CaseType] & "', " & intDisp & " , InvoiceID, #" & DatePaid & "#, '" & InvoiceType & "', '" & InvoiceNumber & "', '" & BTnumber & _
                    "', " & InvoiceAmount & ", " & PaidAmount & ", #" & DateShow & "#,'" & CreatedByName & "', #" & DatePaid & "# )"
                    
                    DoCmd.RunSQL strSQL
                    
                    Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 16, 26, 27, 28, 29, 30, 31, 32

                    DateShow = Date + 30
                    
                    strSQL = "Insert into Accou_EscQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, StaffID, StaffName, CaseType, Disposition, " & _
                    " InvoiceID,DateInvoicedPaid,InvoiceType,InvoiceNumber,BTnumber,InvoiceAmount,PaidAmount,DateShouldShowUp,WhoInvoiced,DatePaid) Values (" & FileNumber & _
                    ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & clientShor & " '," & Forms![Case List]!ClientID & " ,#" & Now() & "# ," & GetStaffID() & _
                    ", '" & GetFullName() & "','" & Forms![Case List]![CaseType] & "', " & intDisp & " , InvoiceID, #" & DatePaid & "#, '" & InvoiceType & "', '" & InvoiceNumber & "', '" & BTnumber & _
                    "', " & InvoiceAmount & ", " & PaidAmount & ", #" & DateShow & "#,'" & CreatedByName & "', #" & DatePaid & "# )"
                    
                    DoCmd.RunSQL strSQL
                    
                    Case Else
                    
                    DateShow = Date + 30
                    
                    strSQL = "Insert into Accou_EscQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, StaffID, StaffName, CaseType, " & _
                    " InvoiceID,DateInvoicedPaid,InvoiceType,InvoiceNumber,BTnumber,InvoiceAmount,PaidAmount,DateShouldShowUp,WhoInvoiced,DatePaid) Values (" & FileNumber & _
                    ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & clientShor & " '," & Forms![Case List]!ClientID & " ,#" & Now() & "# ," & GetStaffID() & _
                    ", '" & GetFullName() & "','" & Forms![Case List]![CaseType] & "',InvoiceID, #" & DatePaid & "#, '" & InvoiceType & "', '" & InvoiceNumber & "', '" & BTnumber & _
                    "', " & InvoiceAmount & ", " & PaidAmount & ", #" & DateShow & "#,'" & CreatedByName & "', #" & DatePaid & "#)"
                    
                    DoCmd.RunSQL strSQL
                    
                    
            
        End Select
        
    Case Else
    
    DoCmd.SetWarnings False
    DateShow = Date + 14
    
    strSQL = "Insert into Accou_EscQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, StaffID, StaffName, CaseType, " & _
    " InvoiceID,DateInvoicedPaid,InvoiceType,InvoiceNumber,BTnumber,InvoiceAmount,PaidAmount,DateShouldShowUp,WhoInvoiced,DatePaid) Values (" & FileNumber & _
    ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & clientShor & " '," & Forms![Case List]!ClientID & " ,#" & Now() & "# ," & GetStaffID() & _
    ", '" & GetFullName() & "','" & Forms![Case List]![CaseType] & "', " & InvoiceID & ", #" & DatePaid & "#, '" & InvoiceType & "', '" & InvoiceNumber & "', '" & BTnumber & _
    "', " & InvoiceAmount & ", " & PaidAmount & ", #" & DateShow & "#,'" & CreatedByName & "', #" & DatePaid & "#)"
    
    DoCmd.RunSQL strSQL
    
    DoCmd.SetWarnings True
              
    
    
    
End Select

DoCmd.SetWarnings True



End Sub

Private Sub cmdPrepareInvoice_Click()

On Error GoTo Err_cmdPrepareInvoice_Click

Dim i As Integer
Dim strTbl As String
Dim strFilter As String


    Select Case Forms![Case List]!CaseTypeID
      Case 5 'Civil
         strTbl = "NamesCIV"
         strFilter = "Debtor = true"
      Case 10 ' Title Resolution
         strTbl = "NamesTR"
         strFilter = "Debtor = true"
      Case Else
         strTbl = "Names"
         strFilter = "Mortgagor = true"
    End Select

i = DCount("[ID]", strTbl, strFilter & " and FileNumber = " & Forms![Case List]!FileNumber)
If i = 0 Then
  MsgBox "No identifying debtors are listed in access under names and therefore, you cannot bill this file as is.  Please correct data in access and try again.", vbCritical
  Exit Sub
End If

DoCmd.OpenForm "InvoiceCreate", , , "FileNumber=" & Forms![Case List]!FileNumber
 Select Case Forms![Case List]!SCRAID
    Case "AccPSAdvanced"
     Forms!InvoiceCreate!cbxInvType = 711
    Case "AccLitig"
     Forms!InvoiceCreate!cbxInvType = 710
    Case Else
     Forms!InvoiceCreate!cbxInvType = 0
    End Select

'DoCmd.OpenForm "UnbilledEvents"
'Forms!unbilledevents!FileNumber = Forms![case list]!FileNumber
'Forms!unbilledevents!lstBillingReasons.RowSource = "SELECT BillingReasonsFCarchive.ID, BillingReasonsFC.Reason, BillingReasonsFCarchive.Date AS [Date] FROM BillingReasonsFCarchive INNER JOIN BillingReasonsFC ON BillingReasonsFCarchive.billingreasonid=BillingReasonsFC.ID WHERE (((BillingReasonsFCarchive.FileNumber)=" & Forms!unbilledevents!FileNumber & " AND Invoiced is null));"



Exit_cmdPrepareInvoice_Click:
    Exit Sub

Err_cmdPrepareInvoice_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrepareInvoice_Click
    
End Sub

Private Sub cmdViewInvoice_Click()

On Error GoTo Err_cmdViewInvoice_Click

DoCmd.SetWarnings False
    
    DoCmd.RunSQL "DELETE * FROM rptInvoice"
    DoCmd.RunSQL "DELETE * FROM rptInvoiceType"
    DoCmd.OpenQuery ("Append InvoiceItem")
    DoCmd.OpenQuery ("Append Invoice Type")
    DoCmd.OpenQuery ("UpdateInvoiceInfo")

DoCmd.SetWarnings True


If DCount("InvoiceItemID", "InvoiceItems", "InvoiceID='" & InvoiceNumber & "'") = 0 Then
    MsgBox "No Invoice data found for this invoice", vbInformation
Else
    Select Case Forms![Case List]!CaseTypeID
      Case 5 'Civil
         DoCmd.OpenReport "Invoice TR CIV", acPreview, , "Invoices.InvoiceNumber=""" & Me!InvoiceNumber & """"
      Case 10 ' Title Resolution
         DoCmd.OpenReport "Invoice TR CIV", acPreview, , "Invoices.InvoiceNumber=""" & Me!InvoiceNumber & """"
      Case Else
         'DoCmd.OpenReport "Invoice", acPreview, , "Invoices.InvoiceNumber=""" & Me!InvoiceNumber & """"
        DoCmd.OpenReport "Invoice", acPreview, , "InvoiceNumber=""" & Me!InvoiceNumber & """"

    End Select
    
    
End If

Exit_cmdViewInvoice_Click:
    Exit Sub

Err_cmdViewInvoice_Click:
    MsgBox Err.Description
    Resume Exit_cmdViewInvoice_Click
    
End Sub

Private Sub WriteOffAmount_AfterUpdate()
DoCmd.OpenForm "EnterWriteOffReason"
Forms!enterwriteoffreason!FileNumber = FileNumber
End Sub
