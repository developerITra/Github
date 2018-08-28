VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Invoice"
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
    !InvoiceType = 0
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

End Sub

Private Sub Form_Open(Cancel As Integer)
Me.AllowEdits = PrivAccounting Or PrivReceivePayments
cmdCapture.Enabled = PrivAccounting
End Sub

Private Sub PaidAmount_AfterUpdate()

If StaffID = 0 Then Call GetLoginName
ReceivedBy = StaffID
If IsNull(DatePaid) Then DatePaid = Date

End Sub
