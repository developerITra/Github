VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_AdvPostSaleCostPkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdCancel_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM PostSaleAdvancePkg;"
DoCmd.SetWarnings True
DoCmd.Close
End Sub

Private Sub cmdOK_Click()

Me.Requery

Dim rs, rs1, rs2 As Recordset
Dim i As Integer
Dim Des As String
Dim Vendor As String
            
Set rs = CurrentDb.OpenRecordset("PostSaleAdvancePkg", dbOpenDynaset, dbSeeChanges)
If rs.EOF Then
    MsgBox ("There is no data")
Exit Sub
End If
            
            
If Form!sfrmPostSaleAdvanceCost!txtCKRequest.ForeColor = vbRed Then
GoTo G:
Else
   If MsgBox("Are you sure you want to close without requesting checks?", vbQuestion + vbYesNo) = vbYes Then
        GoTo G:
   Else
   Exit Sub
   End If
End If
                
G:
            
If IsNull(txtDesc) Or txtDesc = "" Then
    MsgBox ("Fill in the Approval box is required")
Exit Sub
End If
            
Set rs1 = CurrentDb.OpenRecordset("PostSaleCostAdvancePKG", dbOpenDynaset, dbSeeChanges)
            
        If MsgBox("Are you sure to add above item(s) to View Bill Sheet ?", vbQuestion + vbYesNo) = vbYes Then
            
            Do While Not rs.EOF
                Des = rs!Description
            
                If rs!Description = "Attorney Fee" Or rs!Description = "Property Registration Fee" Then
                    AddInvoiceItem Forms![Case List]!FileNumber, "FC-PSA", Des, rs!Amount, 0, True, True, False, False
                Else
                    AddInvoiceItem Forms![Case List]!FileNumber, "FC-PSA", Des, rs!Amount, 0, False, False, False, True
                End If
            
                Set rs2 = CurrentDb.OpenRecordset("Select max(InvoiceItemID)as [InvoiceitemNo] FROM InvoiceItems where filenumber=" & Forms![Case List]!FileNumber, dbOpenDynaset, dbSeeChanges)
            
            rs1.AddNew
                rs1!FileNumber = Forms![Case List]!FileNumber
                rs1!Description = rs!Description
                rs1!Amount = rs!Amount
                rs1!Timestamp = Now()
                rs1!Vendor = rs!Vendor
                rs1!Fees = 0
                rs1!Username = GetFullName()
                rs1!InvoiceItemID = rs2!InvoiceitemNo
                rs1.Update
                rs.MoveNext
            Loop
            
            rs1.Close
            Set rs1 = Nothing
            rs.Close
            Set rs = Nothing
            
        DoCmd.SetWarnings False
        strinfo = "Post Sale Cost Advance." & txtDesc
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms![Case List]!FileNumber & ",Now,GetFullName(),'" & strinfo & "',2 )"
            
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
                
            
'MsgBox ("Post Sale Advances Cost added to View Bill Sheet")
            
Else
Exit Sub
End If
            


''here sarab
'************************************************** SA Cods 3/22/15 PS Advanced project
Dim newfilename As String
Dim selecteddoctype As Long
Dim fileextension As String
Dim DocDate As Date
Dim strSQL As String
Dim strSQLValues As String
Dim DocIDNo As Long
Dim clientShor As String
Dim rstBillReasons As Recordset

strSQL = ""
strSQLValues = ""
Me.txtDesc.SetFocus
DocDate = Now
selecteddoctype = 1566

Me.Box80.Visible = False
Me.cmdOK.Visible = False
Me.cmdCancel.Visible = False

DoCmd.SetWarnings False

newfilename = "PS Advanced Costs" & " " & Format$(Now(), "yyyymmdd hhnnss")

If Dir$(DocLocation & DocBucket(txtFilenum) & "\" & txtFilenum & "\" & newfilename) <> "" Then
    MsgBox txtFilenum & " already exists.", vbCritical
    Exit Sub
End If

DoCmd.OutputTo acOutputForm, "AdvPostSaleCostPkg", acFormatPDF, DocLocation & DocBucket(txtFilenum) & "\" & txtFilenum & "\" & newfilename & ".pdf", False, "", 0

newfilename = newfilename & ".pdf"

strSQLValues = txtFilenum & "," & selecteddoctype & ",'" & "B" & "'," & GetStaffID() & ",'" & DocDate & "','" & Replace(newfilename, "'", "''") & "','" & Replace(newfilename, "'", "''") & "'"
strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
DoCmd.RunSQL (strSQL)


Me.Box80.Visible = True
Me.cmdOK.Visible = True
Me.cmdCancel.Visible = True


       
clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)
DocIDNo = GetLastDocIDNo(GetStaffID(), selecteddoctype, txtFilenum)

'Dim DismSent As Date
'If Not IsNull(DLookup("[DismissalSent]", "[FCDetails]", "[FileNumber] = " & Me.txtFilenum & " and Current = true")) Then
'    DismSent = DLookup("[DismissalSent]", "[FCDetails]", "[FileNumber] = " & Me.txtFilenum & " and Current = true")
'Else
'    DismSent = ""
'End If

strSQL = "Insert into Accou_PSAdvancedCostsPackageQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, Hold, MangNotic, DocIndexID, DocumentId, StaffID, StaffName) Values (" & txtFilenum & ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & _
        clientShor & " '," & Forms![Case List]!ClientID & ", Now(), '','', " & DocIDNo & ", " & selecteddoctype & ", " & GetStaffID() & ", '" & GetFullName() & "'" & ")"

DoCmd.RunSQL strSQL

'Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & txtFilenum, dbOpenDynaset, dbSeeChanges)
'With rstBillReasons
'.AddNew
'!FileNumber = txtFilenum
'!billingreasonid = 33
'!userid = GetStaffID
'!Date = Date
'.Update
'End With


 


DoCmd.SetWarnings True


'********************************************************* SA



DoCmd.Close acForm, "AdvPostSaleCostPkg"
End Sub

Private Sub Form_Open(Cancel As Integer)

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM PostSaleAdvancePkg;"
DoCmd.SetWarnings True

Me.txtDesc = ""
Form!sfrmPostSaleAdvanceCost.Requery
Me.Requery
End Sub
