VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_wizReferralII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cboSortby_AfterUpdate()
  UpdateDocumentList
End Sub
Private Sub UpdateDocumentList()
Dim GroupName As String

On Error GoTo UpdateDocumentListErr

'Select Case optDocType
'    Case 1
        GroupName = ""
'    Case 2
'        GroupName = "B"
'End Select

lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name],DocIndex.doctitleid AS DocType, DocIndex.Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='" & GroupName & "' AND Filespec IS NOT NULL and DeleteDate is null ORDER BY " & Me.cboSortby
lstDocs.Requery

Exit Sub

UpdateDocumentListErr:
    MsgBox Err.Description, vbCritical
    Exit Sub
    

End Sub

Private Sub cmdAddDoc_Click()

'8/29/14
FileNO = txtFileNumber

    Dim fso
    Dim ss As String
    ss = "SSN"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not (fso.FolderExists(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\")) Then
    fso.CreateFolder (DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\")
    End If
    If Not (fso.FolderExists(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & ss & "\")) Then
    fso.CreateFolder (DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & ss & "\")
    Dim objFSO
    Dim objFolder
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\")
    If objFolder.Attributes = objFolder.Attributes And 2 Then
       objFolder.Attributes = objFolder.Attributes Xor 2
    End If
    End If


Dim Filespec As String, fileextension As String, Path As String, FileName As String, newfilename As String, i As Integer, Prompt As String
Dim GroupCode As String, DocType As String, rstDoc As Recordset, DocDateInput As String, DocDate As Date

On Error GoTo Err_cmdAddDoc_Click

Me.Refresh

Filespec = OpenFile(Me)
If Filespec = "" Then Exit Sub

For i = Len(Filespec) To 0 Step -1
    If Asc(Mid$(Filespec, i, 1)) <> 0 Then Exit For
Next i
If i = 0 Then
    MsgBox "Invalid file specification: " & Filespec, vbCritical
    Exit Sub
End If
Filespec = Left$(Filespec, i)

For i = Len(Filespec) To 0 Step -1
    If Mid$(Filespec, i, 1) = "." Then Exit For
Next i
If i = 0 Then
    MsgBox "Invalid file specification: " & Filespec, vbCritical
    Exit Sub
End If
fileextension = Mid$(Filespec, i)

For i = Len(Filespec) To 0 Step -1
    If Mid$(Filespec, i, 1) = "\" Then Exit For
Next i
If i = 0 Then
    MsgBox "Invalid file specification: " & Filespec, vbCritical
    Exit Sub
End If

Path = Left$(Filespec, i)
FileName = Mid$(Filespec, i + 1)
DoCmd.OpenForm "Select Document Type", , , , , acDialog
If selecteddoctype = 0 Then Exit Sub

GroupCode = Nz(DLookup("GroupCode", "DocumentTitles", "ID=" & selecteddoctype))
'If GroupCode = "" Then
    newfilename = DLookup("Title", "DocumentTitles", "ID=" & selecteddoctype) & " " & Format$(Now(), "yyyymmdd hhnnss") & fileextension
'Else
'    NewFilename = GroupDelimiter & GroupCode & GroupDelimiter & DLookup("Title", "DocumentTitles", "ID=" & SelectedDocType) & " " & Format$(Now(), "yyyymmdd hhnn") & FileExtension
'End If

If Dir$(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & newfilename) <> "" Then
    MsgBox newfilename & " already exists.", vbCritical
    Exit Sub
End If

If PrivDocDate Then
    DocDateInput = InputBox$("Enter scan date:", , Format$(Date, "m/d/yyyy"))
    If DocDateInput = "" Then Exit Sub
    If Not IsDate(DocDateInput) Then
        MsgBox ("Invalid or unrecognized date"), vbCritical
        Exit Sub
    End If
    DocDate = CVDate(DocDateInput)
    If DocDate = Date Then DocDate = Now()  ' if user took default (today) then also store the time
Else
    DocDate = Now()
End If




Select Case selecteddoctype
Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572

FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & ss & "\" & newfilename

'DoCmd.SetWarnings False
'strInfo = "Added SSN Document "
'strInfo = Replace(strInfo, "'", "''")
'strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strInfo & "',1 )"
'DoCmd.RunSQL strSQLJournal
'DoCmd.SetWarnings True

Case Else
FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & newfilename
End Select




'FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & NewFilename




'Commented by JAE 10-30-2014 'Document Speed'

'Set rstDoc = CurrentDb.OpenRecordset("DocIndex", dbOpenDynaset, dbSeeChanges)
'rstDoc.AddNew
'rstDoc!FileNumber = FileNumber
'rstDoc!DocTitleID = SelectedDocType
'rstDoc!DocGroup = GroupCode
'rstDoc!StaffID = GetStaffID()
'rstDoc!DateStamp = DocDate
'rstDoc!Filespec = NewFilename
'rstDoc!Notes = NewFilename
'rstDoc.Update
'rstDoc.Close

DoCmd.SetWarnings False
Dim strSQLValues As String: strSQLValues = ""
Dim strSQL As String: strSQL = ""
strSQL = ""
strSQLValues = FileNumber & "," & selecteddoctype & ",'" & GroupCode & "'," & GetStaffID() & ",'" & DocDate & "','" & Replace(newfilename, "'", "''") & "','" & Replace(newfilename, "'", "''") & "'"
'Debug.Print strSQLValues
strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
'Debug.Print strSQL
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings True

lstDocs.Requery
If MsgBox("New document " & newfilename & " accepted.  OK to delete " & Filespec & "?", vbQuestion + vbYesNo) = vbYes Then Kill Filespec

Exit_cmdAddDoc_Click:
    Exit Sub

Err_cmdAddDoc_Click:
    If Err.Number = 76 Then     ' path not found
        MkDir DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\"
        Resume
    Else
        MsgBox Err.Description
        Resume Exit_cmdAddDoc_Click
    End If

End Sub

Private Sub cmdCalcPerDiem_Click()
On Error GoTo Err_cmdCalcPerDiem_Click
PerDiem = RemainingPBal * InterestRate / 100 / 360 'changed to bond market convention to conform to industry standard AW 10/24

Exit_cmdCalcPerDiem_Click:
    Exit Sub

Err_cmdCalcPerDiem_Click:
    MsgBox Err.Description
    Resume Exit_cmdCalcPerDiem_Click

End Sub

Private Sub CmdEditAddress_Click()
PropertyAddress.Enabled = True
PropertyAddress.Locked = False
PropertyAddress.BackStyle = 1
PropertyAddress.BackColor = -2147483643
Apt.Enabled = True
Apt.Locked = False
Apt.BackStyle = 1
Apt.BackColor = -2147483643



End Sub

Private Sub CmdEditPorject_Click()
PrimaryDefName.Enabled = True
PrimaryDefName.Locked = False
PrimaryDefName.BackStyle = 1
PrimaryDefName.BackColor = -2147483643
End Sub

Private Sub cmdOK_Click()
Dim rstCase As Recordset, rstFC As Recordset, rstNames As Recordset
Dim FileNum As Long, MissingInfo As String, WizardType As String, Exceptions As Long

On Error GoTo Err_cmdOK_Click

'Validation
If IsNull(OriginalPBal) Then MissingInfo = MissingInfo & "Original Principal, "
If IsNull(LPIDate) Then MissingInfo = MissingInfo & "LPI"

If LoanType = 4 Then
If IsNull(FNMALoanNumber) Then MissingInfo = "FNMA Loan Number"
End If
If LoanType = 5 Then
If IsNull(FHLMCLoanNumber) Then MissingInfo = "FHLMC Loan Number"
End If

If MissingInfo <> "" Then
    MsgBox "You cannot continue because the following information is missing:" & vbNewLine & MissingInfo, vbCritical
    Exit Sub
End If



Call AddDefaultTrustees(FileNumber)
Exceptions = CountExceptions

Call RSIICompletionUpdate(FileNumber, Exceptions)

DoCmd.SetWarnings False  ' Add records to ValumeRSII
Dim Shortclient As String
Dim ExceptionNo As Integer

ExceptionNo = DLookup("RSIIexceptions", "wizardqueuestats", "filenumber=" & FileNumber & " AND current=true")
Shortclient = DLookup("ShortClientName", "ClientList", "ClientID=" & ClientID)

Dim NewCaseType As String
NewCaseType = "Insert Into ValumeRSII (FileNumber,ShortClientName,Completiondate,State,IdUser,Username,Count,RSIIexceptions) values (" & FileNumber & ",'" & Shortclient & "'" & ",Now, " & "'" & State & "'," & StaffID & ",'" & GetFullName() & "'," & 1 & "," & ExceptionNo & ")"
DoCmd.RunSQL NewCaseType
DoCmd.SetWarnings False


Call ReleaseFile(FileNumber)
MsgBox "Referral Specialist II Wizard complete", vbInformation

txtFileNumber = Null
txtFileNumber.SetFocus
Call ConfirmationVisible(False)
Call FieldsVisible(False)

Me.RecordSource = ""
DoCmd.Close
Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
If Me.RecordSource <> "" Then
    If Me.Dirty Then Me.Undo
    If Not IsNull(FileNumber) Then ReleaseFile (FileNumber)
End If
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, Me.Name

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdView_Click()


On Error GoTo Err_cmdView_Click

Dim i As Long
For i = 0 To lstDocs.ListCount - 1

    Select Case lstDocs.Column(4, i)
  
    Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572
                If Not PrivSSN Then
                MsgBox (" You are not authorized to open SSN doc")
                Exit Sub
                Else
                If lstDocs.Selected(i) Then StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\" & lstDocs.Column(3, i)
                End If
    Case Else

     If lstDocs.Selected(i) Then StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)
    End Select

  


Next i

Exit_cmdView_Click:
    Exit Sub

Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click

End Sub

Private Sub cmdViewInternetSources_Click()
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm "Internet Sources"
End Sub

Private Sub ComAddName_Click()
DoCmd.OpenForm "sfrmNamesUpdate", , , , acFormAdd
End Sub


Private Sub ComAttachDoc_Click()
On Error GoTo Error_Msg
Dim myMail As Outlook.MailItem
Dim OLK As Object 'Oulook.Application
Dim Atmt As Object 'Attachment
Dim Mensaje As Object 'Outlook.MailItem
Dim Adjuntos As String
Dim Body As String
Dim i As Integer
Dim myAttachments As Outlook.Attachments
Dim AttachmentPath As String

  AttachmentPath = """" & DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3) & """"


Set OLK = CreateObject("Outlook.Application")
Set Mensaje = OLK.ActiveInspector.CurrentItem


    With Mensaje
    
   ' MsgBox "Mail was sent on: " & .SentOn
        If .SentOn < Now() Then
            MsgBox ("Please Reply to current email")
            Exit Sub
        Else
            .Attachments.Add AttachmentPath, olByValue, 1
            .Display
        End If
    
    End With



Error_Msg:
If Err = 91 Then
    MsgBox ("Please open the email you wish to attach the document to.")
    Exit Sub
End If
End Sub

Private Sub ComGroup_Click()

Dim Filespec As String, fileextension As String, Path As String, FileName As String, newfilename As String, i As Integer, Prompt As String
Dim GroupCode As String, DocType As String, rstDoc As Recordset, DocDateInput As String, DocDate As Date, UpdateFlag As Boolean, UpdateCase As String
Dim strSCRAQueueFiles As Recordset
Dim rstqueueCount As Integer
Dim objCAcroPDDocDestination As Acrobat.CAcroPDDoc
Dim objCAcroPDDocSource As Acrobat.CAcroPDDoc
Dim A As Integer
DocDate = Now()


If (IsNull(lstDocs.Column(0))) Then
      MsgBox "Please select a document before continuing.", vbCritical, "Select Document"
      Exit Sub
End If

For i = 0 To lstDocs.ListCount - 1
    
    If lstDocs.Selected(i) Then
        
        Select Case lstDocs.Column(4, i)
        Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572
     
    
               MsgBox (" You selected SSN document, not allowed to be in group ")
               Exit Sub
                
        End Select
    End If
Next i

Set objCAcroPDDocDestination = CreateObject("AcroExch.PDDoc")
Set objCAcroPDDocSource = CreateObject("AcroExch.PDDoc")

DoCmd.OpenForm "Select Document Type group", , , , , acDialog
If selecteddoctype = 0 Then Exit Sub


For i = 0 To lstDocs.ListCount - 1
    
    If lstDocs.Selected(i) Then
        objCAcroPDDocDestination.Open DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)
        A = i
        i = lstDocs.ListCount - 1
    End If

Next i

 
For i = 0 To lstDocs.ListCount - 1

    If lstDocs.Selected(i) Then
     If i = A Then i = i + 1
     

 
 
    objCAcroPDDocSource.Open DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)
     
   

    
     

  If Not objCAcroPDDocDestination.InsertPages(objCAcroPDDocDestination.GetNumPages - 1, objCAcroPDDocSource, 0, objCAcroPDDocSource.GetNumPages, 0) Then
       MsgBox (" Error in selected documents")
       objCAcroPDDocSource.Close
       Exit Sub
  End If
  objCAcroPDDocSource.Close
  End If
  
Next i
GroupCode = Nz(DLookup("GroupCode", "DocumentTitles", "ID=" & selecteddoctype))
newfilename = DLookup("Title", "DocumentTitles", "ID=" & selecteddoctype) & " " & Format$(Now(), "yyyymmdd hhnnss") & ".pdf"
objCAcroPDDocDestination.Save 1, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & newfilename
objCAcroPDDocDestination.Close
Set objCAcroPDDocSource = Nothing
Set objCAcroPDDocDestination = Nothing
'End Sub

DoCmd.SetWarnings False
Dim strSQLValues As String: strSQLValues = ""
strSQL = ""
strSQLValues = FileNumber & "," & selecteddoctype & ",'" & GroupCode & "'," & GetStaffID() & ",'" & DocDate & "','" & Replace(newfilename, "'", "''") & "','" & Replace(newfilename, "'", "''") & "'"
'Debug.Print strSQLValues
strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
'Debug.Print strSQL
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings True


lstDocs.Requery
End Sub



Private Sub CommEdit_Click()
DoCmd.OpenForm "sfrmNamesUpdate", , , WhereCondition:="ID= " & Forms!wizreferralII!sfrmNames!ID


'Dim ctrl As Control
'For Each ctrl In Me.sfrmNames.Form.Controls
'
'If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then ' TypeOf ctrl Is CommandButton Then
''If Not ctrl.Locked) Then
'ctrl.Locked = False
'ctrl.Enabled = True
''Else
''ctrl.Locked = True
'End If
''End If
'Next
'With Me.sfrmNames.Form
'
'.AllowAdditions = True
'.AllowEdits = True
'.AllowDeletions = True
'.cmdCopyClient.Enabled = True
'.cmdCopy.Enabled = True
'.cmdTenant.Enabled = True
'.cmdMERS.Enabled = True
'.cmdEnterSSN.Enabled = True
'.cmdNoNotice.Enabled = True
'.cmdPrintNotice.Enabled = True
'.cmdPrintLabel.Enabled = True
'.cbxNotice.Enabled = True
'.cmdDelete.Enabled = True
'.cmdNoNotice.Enabled = True
'.cbxNotice.Enabled = True
'.cbxNotice.Locked = False
'
'End With
''Exit Sub
End Sub

Private Sub Form_Current()
If Not IsNull(LoanNumber) Then Me.LoanNumber.Locked = True

End Sub

Private Sub Form_Open(Cancel As Integer)
Me.RecordSource = ""

'8/29/14
'Fileno = txtFileNumber
End Sub

Private Sub InterestRate_AfterUpdate()
'If InterestRate > 1 Then InterestRate = InterestRate / 100
End Sub

Private Sub LoanType_AfterUpdate()

'Call Visuals
Dim lt As Integer

If IsNull(LoanType) Then
    lt = 0
Else
    lt = LoanType
End If


'FHALoanNumber.Enabled = (lt = 2 Or lt = 3)    ' enable for VA or HUD
'FNMALoanNumber.Enabled = (lt = 4)
'FHLMCLoanNumber.Enabled = (lt = 5)

Call RSIIFieldsVisible(True)

End Sub

Private Sub LPIDate_AfterUpdate()
txt567 = LPIDate + 180
End Sub

Private Sub lstDocs_DblClick(Cancel As Integer)
'
'This function will open the selected document when double-clicked.
'It does the same thing as the "View/Edit Selected" button below the list of docs.
'Patrick J. Fee 8/2/11 240-401-6820.
'





On Error GoTo Err_lstDocs_DblClick
Dim i As Long
For i = 0 To lstDocs.ListCount - 1

    Select Case lstDocs.Column(4, i)
  
    Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572
                If Not PrivSSN Then
                MsgBox (" You are not authorized to open SSN doc")
                Exit Sub
                Else
                If lstDocs.Selected(i) Then StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\" & lstDocs.Column(3, i)
                End If
    Case Else

     If lstDocs.Selected(i) Then StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)
    End Select

  


Next i

Exit_lstDocs_DblClick:
    Exit Sub

Err_lstDocs_DblClick:
    MsgBox Err.Description
    Resume Exit_lstDocs_DblClick
End Sub

Private Sub OriginalTrustee_AfterUpdate()
Dim t As Recordset, fcdetails As Recordset

If OriginalTrustee = 3 Then
Set t = CurrentDb.OpenRecordset("Trustees", dbOpenDynaset, dbSeeChanges)
t.AddNew
t!FileNumber = FileNumber
t!Trustee = OriginalTrustee
t!Assigned = Now()
Set fcdetails = CurrentDb.OpenRecordset("SELECT SubstituteTrustees FROM fcdetails WHERE FileNumber=" & txtFileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
fcdetails.Edit
fcdetails!SubstituteTrustees = False
fcdetails.Update
End If

End Sub



Private Sub txtFileNumber_AfterUpdate()
Dim rstfiles As Recordset

Set rstfiles = CurrentDb.OpenRecordset("SELECT FileNumber FROM Caselist WHERE FileNumber=" & txtFileNumber, dbOpenSnapshot)
If rstfiles.EOF Then
    MsgBox "No such file number: " & txtFileNumber
    txtFileNumber = Null
    txtFileNumber.SetFocus
Else
    If LockFile(txtFileNumber) Then
        Me.RecordSource = CurrentDb.QueryDefs("qryWizRefSpecII").sql
        Me.Filter = "FileNumber=" & txtFileNumber
        Me.FilterOn = True
        Call LoanType_AfterUpdate
        Call ConfirmationVisible(True)
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
End If
rstfiles.Close

End Sub

'Private Sub cmdYes_Click()
'
'On Error GoTo Err_cmdYes_Click
'lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name] FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='' AND Filespec IS NOT NULL and DeleteDate is null"
'lstDocs.Requery
'Call FieldsVisible(True)
'
'Exit_cmdYes_Click:
'    Exit Sub
'
'Err_cmdYes_Click:
'    MsgBox Err.Description
'    Resume Exit_cmdYes_Click
'
'End Sub
'
'Private Sub cmdNo_Click()
'
'On Error GoTo Err_cmdNo_Click
'
'If Me.Dirty Then Me.Undo
'Me.RecordSource = ""
'txtFileNumber = Null
'txtFileNumber.SetFocus
'Call ConfirmationVisible(False)
'Call FieldsVisible(False)
'
'Exit_cmdNo_Click:
'    Exit Sub
'
'Err_cmdNo_Click:
'    MsgBox Err.Description
'    Resume Exit_cmdNo_Click
'
'End Sub

Private Sub ConfirmationVisible(SetVisible As Boolean)

PrimaryDefName.Visible = SetVisible
PropertyAddress.Visible = SetVisible
'txtJurisdiction.Visible = SetVisible
'LongClientName.Visible = SetVisible
'cmdYes.Enabled = SetVisible
'cmdNo.Enabled = SetVisible

End Sub

Private Sub FieldsVisible(SetVisible As Boolean)

tabWiz.Visible = SetVisible
cmdOK.Enabled = SetVisible
AssessedValue.Enabled = (UCase$(Nz(State)) = "VA")

If SetVisible Then
    DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
Else
    DoCmd.Close acForm, "Journal"
End If

End Sub

Private Sub cmdTitle_Click()
On Error GoTo Err_cmdTitle_Click

If Me.Dirty Then
DoCmd.RunCommand acCmdSaveRecord
End If
If ClientID = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation
DoCmd.OpenForm "Print Title Order", , , "Caselist.FileNumber=" & FileNumber, , , acViewNormal
Forms![Print Title Order]!RSII = True
Exit_cmdTitle_Click:
    Exit Sub

Err_cmdTitle_Click:
    MsgBox Err.Description
    Resume Exit_cmdTitle_Click
    
End Sub
Private Function CountExceptions()
'Counts exceptions to populate Exception field in RSII volume report

Dim ctr As Integer, ctra As Integer, rstNames As Recordset

If IsNull(DOTdate) Then
ctr = ctr + 1
End If

If IsNull(Liber) Then
ctr = ctr + 1
End If

If IsNull(Folio) Then
ctr = ctr + 1
End If

If IsNull(DOTrecorded) Then
ctr = ctr + 1
End If

If IsNull(OriginalPBal) Then
ctr = ctr + 1
End If

If IsNull(RemainingPBal) Then
ctr = ctr + 1
End If

If IsNull(LoanType) Then
ctr = ctr + 2 'account for missing loan number as well if loan type not selected
Else
'Check for loan number exception based on loan type
Select Case LoanType
Case "FHLMC"
If IsNull(FHLMCLoanNumber) Then
ctr = ctr + 1
End If
Case "FNMA"
If IsNull(FNMALoanNumber) Then
ctr = ctr + 1
End If
Case "VA"
If IsNull(FHALoanNumber) Then
ctr = ctr + 1
End If
End Select
End If



If IsNull(InterestRate) Then
ctr = ctr + 1
End If

If IsNull(FairDebtAmount) Then
ctr = ctr + 1
End If

If IsNull(TaxID) Then
ctr = ctr + 1
End If

If IsNull(PerDiem) Then
ctr = ctr + 1
End If

If State = "MD" Then
If IsNull(AssessedValue) Then
ctr = ctr + 1
End If
End If

'Look at names subform
Set rstNames = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber=" & FileNumber, dbOpenSnapshot)


Do While Not rstNames.EOF
        If rstNames!NoticeType >= 1 Then
        ctra = ctra + 1
        End If
        rstNames.MoveNext
    Loop


If rstNames.RecordCount > ctra Then
ctr = ctr + (rstNames.RecordCount - ctra)
End If

CountExceptions = ctr


End Function
