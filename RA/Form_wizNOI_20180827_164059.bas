VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_wizNOI"
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
Private Sub cmdAcrobat_Click()

If Forms!wizNOI!ClientSentNOI <> "C" Then
    If IsNull(Forms!wizNOI!FairDebtDate) Then
    MsgBox ("There is no Fairdebt date")
    Exit Sub
    End If
End If

'Open client specific attachments
Dim rstClient As Recordset
Set rstClient = CurrentDb.OpenRecordset("select * from clientlist where clientid=" & ClientID, dbOpenDynaset, dbSeeChanges)
With rstClient
Select Case LoanType
Case 5 'Freddie
Call StartDoc(TemplatePath & "Freddie Mac- HAMP-HUD Package.pdf")
If !NOIFreddie1 = True Or !NOIFreddie2 = True Or !NOIFreddie3 = True Or !NOIFreddie4 = True Or !NOIFreddie5 = True Then
Call StartDoc(TemplatePath & !NOIdoc1)
Else
MsgBox "At least one client specific attachment must be selected in the Clients menu to proceed.  See your manager.", vbCritical
Exit Sub
End If
If !NOIFreddie2 = True Then Call StartDoc(TemplatePath & !NOIdoc2)
If !NOIFreddie3 = True Then Call StartDoc(TemplatePath & !NOIdoc3)
If !NOIFreddie4 = True Then Call StartDoc(TemplatePath & !NOIdoc4)
If !NOIFreddie5 = True Then Call StartDoc(TemplatePath & !NOIdoc5)
Case 4 'Fannie
If !NOIFannie1 = True Or !NOIFannie2 = True Or !NOIFannie3 = True Or !NOIFannie4 = True Or !NOIFannie5 = True Then
Call StartDoc(TemplatePath & !NOIdoc1)
Else
MsgBox "At least one client specific attachment must be selected in the Clients menu to proceed.  See your manager.", vbCritical
Exit Sub
End If
If !NOIFannie2 = True Then Call StartDoc(TemplatePath & !NOIdoc2)
If !NOIFannie3 = True Then Call StartDoc(TemplatePath & !NOIdoc3)
If !NOIFannie4 = True Then Call StartDoc(TemplatePath & !NOIdoc4)
If !NOIFannie5 = True Then Call StartDoc(TemplatePath & !NOIdoc5)
Case 2 Or 3 'HUD
If !NOIHUD1 = True Or !NOIHUD2 = True Or !NOIHUD3 = True Or !NOIHUD4 = True Or !NOIHUD5 = True Then
Call StartDoc(TemplatePath & !NOIdoc1)
Else
MsgBox "At least one client specific attachment must be selected in the Clients menu to proceed.  See your manager.", vbCritical
Exit Sub
End If
If !NOIHUD2 = True Then Call StartDoc(TemplatePath & !NOIdoc2)
If !NOIHUD3 = True Then Call StartDoc(TemplatePath & !NOIdoc3)
If !NOIHUD4 = True Then Call StartDoc(TemplatePath & !NOIdoc4)
If !NOIHUD5 = True Then Call StartDoc(TemplatePath & !NOIdoc5)
Case 1 ' Conventional
If !NOIConv1 = True Or !NOIConv2 = True Or !NOIConv3 = True Or !NOIConv4 = True Or !NOIConv5 = True Then
Call StartDoc(TemplatePath & !NOIdoc1)
Else
MsgBox "At least one client specific attachment must be selected in the Clients menu to proceed.  See your manager.", vbCritical
Exit Sub
End If
If !NOIConv2 = True Then Call StartDoc(TemplatePath & !NOIdoc2)
If !NOIConv3 = True Then Call StartDoc(TemplatePath & !NOIdoc3)
If !NOIConv4 = True Then Call StartDoc(TemplatePath & !NOIdoc4)
If !NOIConv5 = True Then Call StartDoc(TemplatePath & !NOIdoc5)
End Select
End With

Call DoReport("45 Day Notice WizTest", -2)
cmdPrint.Enabled = True
cmdPrint.SetFocus
cmdAcrobat.Enabled = False
Call StartDoc(TemplatePath & "\CounselingResources.pdf")
Call StartDoc(TemplatePath & "\LossMitApp.pdf")

End Sub

Private Sub cmdClientSentNOI_Click()
Dim NOIdate As Date
If MsgBox("Are you sure the Client mailed the NOI? ", vbYesNo) = vbYes Then
'NOIdate = InputBox("Please enter the date the Client mailed the NOI", "NOI Wizard") 'stopped as should not filled NOI befoe Atty apporove SA NOI project 2/23/15
'If Not IsNull(NOIdate) Then
'NOI = NOIdate
'End If

        Dim rstNOI As Recordset
                    Set rstNOI = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where current = true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
                    With rstNOI
                    .Edit
                    !NOIClientSent = True
                    .Update
                    .Close
                    End With
        
        
        Me.Requery
        Dim rstFCdetails As Recordset
                    Set rstFCdetails = CurrentDb.OpenRecordset("Select * FROM FCdetails where current = true and  filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
                    With rstFCdetails
                    .Edit
                    !ClientSentNOI = "C"
                   '   !NOI = NOIdate 'stopped as should not filled NOI befoe Atty apporove SA NOI project 2/23/15
                    .Update
                    .Close
                    End With
        If Forms!wizNOI!ComLable.Visible = True Then Forms!wizNOI!ComLable.Visible = False
        
End If
Me.Requery
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


Select Case selecteddoctype
'Sub GeneralMissingDoc(FileNumber As Integer, DocTitleNO As Integer, Demaind As Boolean, FD As Boolean, Intake As Boolean, NOI As Boolean, Dockting As Boolean)

Case 1549
    Call GeneralMissingDoc(FileNumber, 1549, True, False, False, True, False)
   
Case 988
     Call GeneralMissingDoc(FileNumber, 988, False, False, False, True, False)
Case 1553
    Call GeneralMissingDoc(FileNumber, 1553, False, False, False, True, False)
End Select



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



Private Sub cmdOK_Click()
Dim rstCase As Recordset, rstFC As Recordset, rstNames As Recordset, rstdocs As Recordset
Dim FileNum As Long, MissingInfo As String, WizardType As String, Exceptions As Long
Dim NOIdate As Date
Dim qtypstge As Double
Dim rstsql As String
Dim strSQLJournal As String

WizardType = "NOI"
On Error GoTo Err_cmdOK_Click

If Forms!queNOIdocs!lstFiles.Column(8) <> "1. Approved" Then
MsgBox (" The file did not approved by Atty")
Exit Sub
End If

'-------- added 2/20/15
    'If Forms!wizNOI!ClientSentNOI <> "C" Or IsNull(Forms!wizNOI!ClientSentNOI) Or Forms!wizNOI!ClientSentNOI = "" Then
    Call GeneralMissingDoc(FileNumber, 0, False, False, True, False, False)
        
        If Forms!wizNOI!ClientSentNOI = "C" Then
        Else
            FeeAmount = Nz(DLookup("FeeNOI", "ClientList", "ClientID=" & ClientID))
            
                If FeeAmount = 0 Or IsNull(FeeAmount) Then
                    MsgBox ("See Operations Manager for fee for this client, Cannot complete NOI wizard until fee is entered")
                Exit Sub
                End If
                
                If FeeAmount > 0 Then
                    AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI", FeeAmount, 0, True, True, False, False
                'Else
                    'AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI", 1, 0, True, True, False, False 'set unknown fee as $1, per Diane
                End If
                
            FeeAmount = Nz(DLookup("FeeNOIPostage", "ClientList", "ClientID=" & ClientID))
                    'Removed 2/6 to replace with postage entry when doc uploaded
                    If FeeAmount > 0 Then
                        qtypstge = DCount("[FileNumber]", "[NOTCnt]", "FileNumber=" & [FileNumber])
                        AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI Postage", (qtypstge * FeeAmount), 76, False, False, False, True
                    'Else
                    'AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI Postage", 1, 76, False, False, False, True
                    End If

    End If

'-------

If Forms!wizNOI!ClientSentNOI = "C" Then

    NOIdate = InputBox("Please enter the date the Client mailed the NOI", "NOI Wizard") 'stopped as should not filled NOI befoe Atty apporove SA NOI project 2/23/15
    If Not IsNull(NOIdate) Then
        NOI = NOIdate
    Else
    NOI = Date
    End If


Else

NOI = Date


End If

'Add invocing


            'FeeAmount = Nz(DLookup("FeeNOI", "ClientList", "ClientID=" & ClientID))
                'If FeeAmount > 0 Then
                    'AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI", FeeAmount, 0, True, True, False, False
                'Else
                    'AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI", 1, 0, True, True, False, False 'set unknown fee as $1, per Diane
                'End If
                
            'FeeAmount = Nz(DLookup("FeeNOIPostage", "ClientList", "ClientID=" & ClientID))
                    'Removed 2/6 to replace with postage entry when doc uploaded
                    'If FeeAmount > 0 Then
                    'qtypstge = DCount("[FileNumber]", "[qry45days]", "FileNumber=" & [FileNumber])
                    'AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI Postage", (qtypstge * FeeAmount), 76, False, False, False, True
                    'Else
                    'AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI Postage", 1, 76, False, False, False, True
                    'End If





'Call AddFileResponsibilityHistory(FileNum, ?, StaffID)
Call NOICompletionUpdate(WizardType, FileNumber)

    
    DoCmd.SetWarnings False
    rstsql = "Insert into ValumeNOI (CaseFile, Client, Name, NOISentBy, NOIComplete, NOICompleteC ) values (Forms!wizNOI!FileNumber, ClientShortName(forms!wizNOI!ClientID),Getfullname(),'" & Forms!wizNOI!ClientSentNOI & "',Now(),1) "
    DoCmd.RunSQL rstsql
    
    
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!wizNOI!FileNumber & ",Now,GetFullName(),'" & " 45 days Notice Completed " & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True


           
         
        


Call ReleaseFile(FileNumber)
AddStatus FileNumber, Now(), "45 Day Notice Completed"




txtFileNumber = Null
txtFileNumber.SetFocus
Call ConfirmationVisible(False)
Call FieldsVisible(False)

Set rstdocs = CurrentDb.OpenRecordset("select * from documentmissing where filenbr=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstdocs
Do Until .EOF
.Delete
.MoveNext
Loop
'.Update
End With
rstdocs.Close



Dim lrs As Recordset
Set lrs = CurrentDb.OpenRecordset("select * from journal where FileNumber=" & FileNumber & " AND warning = 100", dbOpenDynaset, dbSeeChanges)
With lrs
'.Edit
Do Until .EOF
.Edit
![Warning] = 0
.Update
.MoveNext

Loop
'.Update
End With
lrs.Close



'MsgBox "NOI Wizard complete", vbInformation
'Me.RecordSource = ""
DoCmd.Close acForm, Me.Name


Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub
Private Sub cmdOKdocsmsng_Click()
Dim rstCase As Recordset, rstFC As Recordset, rstNames As Recordset
Dim FileNum As Long, MissingInfo As String

On Error GoTo Err_cmdOK_Click


Call ReleaseFile(FileNumber)

DoCmd.OpenForm "EnterNOIDocs"
Forms!EnterNOIDocs!FileNumber = Forms!wizNOI!FileNumber


Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
'If Me.RecordSource <> "" Then 'to save the changes , as per Diane request SA 08/23/14


   ' If Me.Dirty Then Me.Undo
    If Not IsNull(FileNumber) Then ReleaseFile (FileNumber)
'End If
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, Me.Name

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdPreView_Click()


On Error GoTo Err_cmdpreview_Click


Dim statusMsg As String, FileNum As Long, MissingInfo As String, rstClientAttach As Recordset

If Not IsNull(txtDisposition) Then
MsgBox "You cannot print an NOI for this file because it has a Disposition", vbCritical
Exit Sub
End If

FileNum = FileNumber

If IsNull(LoanNumber) Then MissingInfo = MissingInfo & "Loan Number, "
If IsNull(LastPaymentDated) Then MissingInfo = MissingInfo & "Last Payment Received, "
If IsNull(LastPaymentApplied) Then MissingInfo = MissingInfo & "Last Payment Applied, "
If IsNull(AmountOwedNOI) Then MissingInfo = MissingInfo & "Amount to Cure Default, "
If IsNull(DateOfDefault) Then MissingInfo = MissingInfo & "Date of Default, "
If IsNull(SecuredParty) Then MissingInfo = MissingInfo & "Secured Party, "
If IsNull(SecuredPartyPhone) Then MissingInfo = MissingInfo & "Secured Party Phone, "
If IsNull(MortgageLender) Then MissingInfo = MissingInfo & "Mortgage Lender, "
If IsNull(MortgageLenderLicense) Then MissingInfo = MissingInfo & "Mortgage Lender License, "
If IsNull(Option339) And IsNull(Option341) Then MissingInfo = MissingInfo & "Type of Default, "
If IsNull(DLookup("loanmodagent", "clientlist", "clientid=" & ClientID)) Then MissingInfo = MissingInfo & "Loan Mod Agent (In client screen)"

If MissingInfo <> "" Then
    MsgBox "You cannot continue because the following information is missing:" & vbNewLine & Left$(MissingInfo, Len(MissingInfo) - 2), vbCritical
    Exit Sub
End If

'client loss mit package validation
If LoanType = 5 Then
Set rstClientAttach = CurrentDb.OpenRecordset("select * from clientlist where clientid=" & ClientID, dbOpenDynaset, dbSeeChanges)
If rstClientAttach.RecordCount = 0 Then
MsgBox "Client's Loss Mit attachments are required for this file, please see your manager", vbCritical
Exit Sub
End If
End If

Call DoReport("45 Day Notice WizTest", acPreview)
cmdAcrobat.Enabled = True
cmdAcrobat.SetFocus
cmdPreview.Enabled = False
    
'cmdCancel.Caption = "Close"

Exit_cmdpreview_Click:
DoCmd.Close acForm, "45 Day Notice"
    Exit Sub

Err_cmdpreview_Click:
    MsgBox Err.Description
    Resume Exit_cmdpreview_Click
End Sub



Private Sub ComAddName_Click()
DoCmd.OpenForm "sfrmNamesUpdate", , , , acFormAdd

End Sub

Private Sub Command194_Click()
Me!btnUndo.Enabled = True
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

Private Sub ComLable_Click()
Dim sql As String
Dim i As Integer
Dim rstLabelData As Recordset
Dim cntr As Integer
Dim matter As String
Dim sql2 As String
Dim rstLabelData2 As Recordset
Dim J As Integer
Dim cntr2 As Integer

If Forms!wizNOI!ClientSentNOI <> "C" Then
    If IsNull(Forms!wizNOI!FairDebtDate) Then
    MsgBox ("There is no Fairdebt date")
    Exit Sub
    End If
End If


sql = "SELECT qry45DaysNew.FileNumber, qry45DaysNew.Names_Company, qry45DaysNew.Names_First, qry45DaysNew.Names_Last, qry45DaysNew.Names_Address, qry45DaysNew.Names_Address2, qry45DaysNew.Names_City, qry45DaysNew.Names_State, qry45DaysNew.Names_Zip, qry45DaysNew.PrimaryDefName, qry45DaysNew.Deceased, ClientList.ShortClientName FROM ClientList INNER JOIN qry45DaysNew ON ClientList.ClientID = qry45DaysNew.ClientID where qry45DaysNew.filenumber=" & Forms![wizNOI]!FileNumber


        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            For i = 1 To 4
                Call StartLabel
                Print #6, FormatName(rstLabelData!Names_Company, IIf(rstLabelData!Deceased = True, "Estate of " & rstLabelData!Names_First, rstLabelData!Names_First), rstLabelData!Names_Last, "", rstLabelData!Names_Address, rstLabelData!Names_Address2, rstLabelData!Names_City, rstLabelData!Names_State, rstLabelData!Names_Zip)
                Print #6, "|FONTSIZE 8"
                Print #6, "|BOTTOM"
                Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
                matter = rstLabelData!PrimaryDefName
                Call FinishLabel
            Next i
'            rstLabelData.MoveNext
'            cntr = cntr + 1
'        Loop
'        rstLabelData.Close
'        cntr = cntr * 2
   
   
   
   sql2 = "SELECT ClientList.ShortClientName, CaseList.PrimaryDefName, CaseList.FileNumber, ClientList.ClientNameAsInvestor, ClientList.StreetAddress, ClientList.StreetAddr2, ClientList.City, ClientList.state, ClientList.ZipCode" & _
        " FROM (ClientList INNER JOIN CaseList ON ClientList.ClientID = CaseList.ClientID) INNER JOIN FCdetails ON CaseList.FileNumber = FCdetails.FileNumber " & _
        " WHERE (((CaseList.FileNumber)=" & [Forms]![wizNOI]![FileNumber] & ") AND ((FCdetails.Current)=True));"

    Set rstLabelData2 = CurrentDb.OpenRecordset(sql2, dbOpenSnapshot)
        Do While Not rstLabelData2.EOF
            For J = 1 To 2
                Call StartLabel
                Print #6, FormatName("", "", rstLabelData2!ClientNameAsInvestor, " ", rstLabelData2!StreetAddress, rstLabelData2!StreetAddr2, rstLabelData2!City, rstLabelData2!State, rstLabelData2!ZipCode)
                Print #6, "|FONTSIZE 8"
                Print #6, "|BOTTOM"
                Print #6, rstLabelData2!FileNumber & " / " & rstLabelData2!ShortClientName & " / " & rstLabelData2!PrimaryDefName
                matter = rstLabelData2!PrimaryDefName
                Call FinishLabel
            Next J
            rstLabelData2.MoveNext
            cntr2 = cntr2 + 1
        Loop
        rstLabelData2.Close
        cntr2 = cntr2 * 2
        
        
    rstLabelData.MoveNext
            cntr = cntr + 1
        Loop
        rstLabelData.Close
        cntr = cntr * 2

End Sub

Private Sub CommEdit_Click()
DoCmd.OpenForm "sfrmNamesUpdate", , , WhereCondition:="ID= " & Forms!wizNOI!sfrmNames!ID

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


Private Sub ComTestNOI_Click()
On Error GoTo Err_cmdpreview_Click

'If Forms!wizNOI!ClientSentNOI <> "C" Then
    If IsNull(Forms!wizNOI!FairDebtDate) Then
    MsgBox ("There is no Fairdebt date")
    Exit Sub
    End If
'End If

Dim statusMsg As String, FileNum As Long, MissingInfo As String, rstClientAttach As Recordset

If Not IsNull(txtDisposition) Then
MsgBox "You cannot print an NOI for this file because it has a Disposition", vbCritical
Exit Sub
End If

FileNum = FileNumber

If IsNull(LoanNumber) Then MissingInfo = MissingInfo & "Loan Number, "
If IsNull(LastPaymentDated) Then MissingInfo = MissingInfo & "Last Payment Received, "
If IsNull(LastPaymentApplied) Then MissingInfo = MissingInfo & "Last Payment Applied, "
If IsNull(AmountOwedNOI) Then MissingInfo = MissingInfo & "Amount to Cure Default, "
If IsNull(DateOfDefault) Then MissingInfo = MissingInfo & "Date of Default, "
If IsNull(SecuredParty) Then MissingInfo = MissingInfo & "Secured Party, "
If IsNull(SecuredPartyPhone) Then MissingInfo = MissingInfo & "Secured Party Phone, "
If IsNull(MortgageLender) Then MissingInfo = MissingInfo & "Mortgage Lender, "
If IsNull(MortgageLenderLicense) Then MissingInfo = MissingInfo & "Mortgage Lender License, "
If IsNull(Option339) And IsNull(Option341) Then MissingInfo = MissingInfo & "Type of Default, "
If IsNull(DLookup("loanmodagent", "clientlist", "clientid=" & ClientID)) Then MissingInfo = MissingInfo & "Loan Mod Agent (In client screen)"

If MissingInfo <> "" Then
    MsgBox "You cannot continue because the following information is missing:" & vbNewLine & Left$(MissingInfo, Len(MissingInfo) - 2), vbCritical
    Exit Sub
End If

'client loss mit package validation
If LoanType = 5 Then
Set rstClientAttach = CurrentDb.OpenRecordset("select * from clientlist where clientid=" & ClientID, dbOpenDynaset, dbSeeChanges)
If rstClientAttach.RecordCount = 0 Then
MsgBox "Client's Loss Mit attachments are required for this file, please see your manager", vbCritical
Exit Sub
End If
End If

Call DoReport("45 Day Notice WizTest", acPreview)
cmdAcrobat.Enabled = True
cmdAcrobat.SetFocus
cmdPreview.Enabled = False
    
'cmdCancel.Caption = "Close"

Exit_cmdpreview_Click:
DoCmd.Close acForm, "45 Day Notice"
    Exit Sub

Err_cmdpreview_Click:
    MsgBox Err.Description
    Resume Exit_cmdpreview_Click
End Sub

Private Sub Form_Current()
If PrivNewNOIFDDemaind Then
NewFairDebt.Visible = True
NewDemand.Visible = True
New45Notice.Visible = True
End If



'Dim CheckID As Integer
'CheckID = (DLookup("ID", "Staff", "ID =" & GetStaffID()))
'If CheckID = 1 Or CheckID = 557 Or CheckID = 103 Or CheckID = 455 Then ComTestNOI.Visible = True

End Sub



Private Sub Form_Open(Cancel As Integer)
Me.RecordSource = ""
End Sub

Private Sub cmdPrint_Click()

If Forms!wizNOI!ClientSentNOI <> "C" Then
    If IsNull(Forms!wizNOI!FairDebtDate) Then
    MsgBox ("There is no Fairdebt date")
    Exit Sub
    End If
End If

Dim MortLndrTxt As String, MortOrigTxt As String, MortLndrLic As String, MortOrigLic As String, qtypstge As Double, sql As String, i As Integer
Dim statusMsg As String, JnlNote As String, FileNum As Long, MissingInfo As String, rstNames As Recordset, rstwiz As Recordset, rstLabelData As Recordset
Dim matter As String, cntr As Integer

MortOrigTxt = "The Mortgage Originator is missing, "
MortOrigLic = "The Mortgage Originator License is missing"

FileNum = FileNumber
'On Error GoTo Err_cmdPrint_Click

If Not IsNull(txtDisposition) Then
MsgBox "You cannot print an NOI for this file because it has a Disposition", vbCritical
Exit Sub
End If

If IsNull(LoanNumber) Then MissingInfo = MissingInfo & "Loan Number, "
If IsNull(LastPaymentDated) Then MissingInfo = MissingInfo & "Last Payment Received, "
If IsNull(LastPaymentApplied) Then MissingInfo = MissingInfo & "Last Payment Applied, "
If IsNull(AmountOwedNOI) Then MissingInfo = MissingInfo & "Amount to Cure Default, "
If IsNull(DateOfDefault) Then MissingInfo = MissingInfo & "Date of Default, "
If IsNull(SecuredParty) Then MissingInfo = MissingInfo & "Project Name, "
If IsNull(SecuredPartyPhone) Then MissingInfo = MissingInfo & "Property Address, "
If IsNull(MortgageLender) Then MissingInfo = MissingInfo & "Mortgage Lender, "
If IsNull(MortgageLenderLicense) Then MissingInfo = MissingInfo & "Mortgage Lender License, "
If IsNull(Option339) And IsNull(Option341) Then MissingInfo = MissingInfo & "Type of Default"

If MissingInfo <> "" Then
    MsgBox "You cannot continue because the following information is missing:" & vbNewLine & Left$(MissingInfo, Len(MissingInfo) - 2), vbCritical
    Exit Sub
End If

If IsNull(MortgageOriginator) Then
JnlNote = JnlNote & MortOrigTxt
End If
If IsNull(MortgageOriginatorLicense) Then
JnlNote = JnlNote & MortOrigLic
End If

Dim lrs As Recordset
If JnlNote <> "" Then
'            Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'            lrs.AddNew
'            lrs![FileNumber] = FileNum
'            lrs![JournalDate] = Now
'            lrs![Who] = GetFullName()
'            lrs![Info] = JnlNote & vbCrLf
'            lrs![Color] = 1
'            lrs.Update
'            lrs.Close
            
            DoCmd.SetWarnings False
            strinfo = JnlNote & vbCrLf
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNum & ",Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            End If

Call DoReport("45 Day Notice WizTest", acViewNormal)
Call DoReport("45 Day Notice WizTest", acViewNormal)
Call StartDoc(TemplatePath & "\CounselingResources.pdf")
Call StartDoc(TemplatePath & "\LossMitApp.pdf")

'Open client specific attachments
Dim rstClient As Recordset
Set rstClient = CurrentDb.OpenRecordset("select * from clientlist where clientid=" & ClientID, dbOpenDynaset, dbSeeChanges)
With rstClient
Select Case LoanType
Case 5 'Freddie
Call StartDoc(TemplatePath & "Freddie Mac- HAMP-HUD Package.pdf")
If !NOIFreddie1 = True Then Call StartDoc(TemplatePath & !NOIdoc1)
If !NOIFreddie2 = True Then Call StartDoc(TemplatePath & !NOIdoc2)
If !NOIFreddie3 = True Then Call StartDoc(TemplatePath & !NOIdoc3)
If !NOIFreddie4 = True Then Call StartDoc(TemplatePath & !NOIdoc4)
If !NOIFreddie5 = True Then Call StartDoc(TemplatePath & !NOIdoc5)
Case 4 'Fannie
If !NOIFannie1 = True Then Call StartDoc(TemplatePath & !NOIdoc1)
If !NOIFannie2 = True Then Call StartDoc(TemplatePath & !NOIdoc2)
If !NOIFannie3 = True Then Call StartDoc(TemplatePath & !NOIdoc3)
If !NOIFannie4 = True Then Call StartDoc(TemplatePath & !NOIdoc4)
If !NOIFannie5 = True Then Call StartDoc(TemplatePath & !NOIdoc5)
Case 2 Or 3 'HUD
If !NOIHUD1 = True Then Call StartDoc(TemplatePath & !NOIdoc1)
If !NOIHUD2 = True Then Call StartDoc(TemplatePath & !NOIdoc2)
If !NOIHUD3 = True Then Call StartDoc(TemplatePath & !NOIdoc3)
If !NOIHUD4 = True Then Call StartDoc(TemplatePath & !NOIdoc4)
If !NOIHUD5 = True Then Call StartDoc(TemplatePath & !NOIdoc5)
Case 1 ' Conventional
If !NOIConv1 = True Then Call StartDoc(TemplatePath & !NOIdoc1)
If !NOIConv2 = True Then Call StartDoc(TemplatePath & !NOIdoc2)
If !NOIConv3 = True Then Call StartDoc(TemplatePath & !NOIdoc3)
If !NOIConv4 = True Then Call StartDoc(TemplatePath & !NOIdoc4)
If !NOIConv5 = True Then Call StartDoc(TemplatePath & !NOIdoc5)
End Select
End With

''Print labels, 4 per address Stopped by SA on 12/31 as we added new Lebal icon
'    '    sql = "SELECT Names.Company, Names.First, Names.Last, Names.AKA, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip, CaseList.FileNumber, ClientList.FairDebt, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (ClientList RIGHT JOIN (CaseList RIGHT JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID) LEFT JOIN FCdetails ON CaseList.FileNumber=FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & Forms![wizNOI]!FileNumber & ") And  ((FCdetails.Current)=True));"
'        sql = "SELECT qry45days.FileNumber, qry45Days.Names_Company, qry45Days.Names_First, qry45Days.Names_Last, qry45Days.Names_Address, qry45Days.Names_Address2, qry45Days.Names_City, qry45Days.Names_State, qry45Days.Names_Zip, qry45Days.PrimaryDefName, qry45Days.Deceased, ClientList.ShortClientName FROM ClientList INNER JOIN qry45Days ON ClientList.ClientID = qry45Days.ClientID where qry45days.filenumber=" & Forms![wizNOI]!FileNumber
'        'sql = "SELECT Names.Company, Names.First, Names.Last, Names.AKA, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip, Names.Deceased, CaseList.FileNumber, ClientList.FairDebt, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (ClientList RIGHT JOIN (CaseList RIGHT JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID) LEFT JOIN FCdetails ON CaseList.FileNumber=FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & Forms![wizFairDebt]!FileNumber & ") And ((Names.FairDebt)=True) And ((FCdetails.Current)=True));"
'        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
'        Do While Not rstLabelData.EOF
'            For i = 1 To 4
'                Call StartLabel
'                Print #6, FormatName(rstLabelData!Names_Company, IIf(rstLabelData!Deceased = True, "Estate of " & rstLabelData!Names_First, rstLabelData!Names_First), rstLabelData!Names_Last, "", rstLabelData!Names_Address, rstLabelData!Names_Address2, rstLabelData!Names_City, rstLabelData!Names_State, rstLabelData!Names_Zip)
'                Print #6, "|FONTSIZE 8"
'                Print #6, "|BOTTOM"
'                Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
'                matter = rstLabelData!PrimaryDefName
'                Call FinishLabel
'            Next i
'            rstLabelData.MoveNext
'            cntr = cntr + 1
'        Loop
'        rstLabelData.Close
'        cntr = cntr * 2

  'a from a to b stopped with changes on NOI proceduer SA 2/16/14
'    If MsgBox("Update 45 Day Notice = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
'        NOI = Date
'        AddStatus FileNum, Now(), "45 Day Notice sent"
'
'            FeeAmount = Nz(DLookup("FeeNOI", "ClientList", "ClientID=" & ClientID))
'            If FeeAmount > 0 Then
'                AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI", FeeAmount, 0, True, True, False, False
'            Else
'                AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI", 1, 0, True, True, False, False 'set unknown fee as $1, per Diane
'            End If
'            FeeAmount = Nz(DLookup("FeeNOIPostage", "ClientList", "ClientID=" & ClientID))
'            'Removed 2/6 to replace with postage entry when doc uploaded
'            If FeeAmount > 0 Then
'            qtypstge = DCount("[FileNumber]", "[qry45days]", "FileNumber=" & [FileNumber])
'            AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI Postage", (qtypstge * FeeAmount), 76, False, False, False, True
'            Else
'            AddInvoiceItem FileNumber, "FC-NOI", "45 Day NOI Postage", 1, 76, False, False, False, True
' b           End If
                        
    'Update Journal with missing info note
    If JnlNote <> "" Then
    '2/11/14
        DoCmd.SetWarnings False
        strinfo = JnlNote
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizNOI!FileNum,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True

    End If
      
'    End If
    
    Set rstwiz = CurrentDb.OpenRecordset("select * from wizardqueuestats where filenumber=" & FileNum, dbOpenDynaset, dbSeeChanges)
    rstwiz.Edit
    rstwiz![NOIBulkUpload] = False
    rstwiz.Update
    rstwiz.Close
    
cmdCancel.Caption = "Close"

Exit_cmdPrint_Click:
DoCmd.Close acForm, "45 Day Notice"
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrint_Click
    

End Sub



Private Sub lstDocs_DblClick(Cancel As Integer)
'
'This function will open the selected document when double-clicked.
'It does the same thing as the "View/Edit Selected" button below the list of docs.
'Patrick J. Fee 8/2/11 240-401-6820.
'
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

Private Sub ConfirmationVisible(SetVisible As Boolean)


'cmdYes.Enabled = SetVisible
'cmdNo.Enabled = SetVisible

End Sub

Private Sub FieldsVisible(SetVisible As Boolean)

tabWiz.Visible = SetVisible
cmdOK.Enabled = SetVisible
'AssessedValue.Enabled = (UCase$(Nz(State)) = "VA")

If SetVisible Then
    DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
Else
    DoCmd.Close acForm, "Journal"
End If

End Sub

Private Sub cmdTitle_Click()
On Error GoTo Err_cmdTitle_Click

If ClientID = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation
DoCmd.OpenForm "Print Title Order", , , "Caselist.FileNumber=" & FileNumber, , acDialog, acViewNormal

Exit_cmdTitle_Click:
    Exit Sub

Err_cmdTitle_Click:
    MsgBox Err.Description
    Resume Exit_cmdTitle_Click
    
End Sub
Private Sub cmdView_Click()

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

Private Sub New45Notice_Click()
If (IsNull(Disposition) Or (Disposition = 1 Or Disposition = 2)) Then

    If Not IsNull([NOI]) Or ([ClientSentNOI] = "C") Then
        '45 Day Notice sent
        If MsgBox(" You are about to remove dates ? ", vbOKCancel) = vbOK Then
        AddStatus FileNumber, Now(), "Removed 45 Days Notice (" & [NOI] & ") by " & GetFullName
    
        DoCmd.SetWarnings False
        strinfo = "Removed 45 Days Notice (" & [NOI] & ") by " & GetFullName
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizNOI!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
    
       
        Dim rstNOI As Recordset
        Set rstNOI = CurrentDb.OpenRecordset("Select * From WizardQueueStats Where FileNumber = " & FileNumber & "And Current = True", dbOpenDynaset, dbSeeChanges)
            If Not rstNOI.EOF Then
                With rstNOI
                .Edit
                'If IsNull(!FairDebtComplete) Then !FairDebtComplete = #1/2/2012#
                !NOIcomplete = Null
                !DateInWaiitingQueueNOI = Null
                !DateInQueueNOI = Null
                '!Add45 = "45day"
                .Update
                End With
                Else
                MsgBox ("There is no Currrent Wizard Record for this File, Please Contact the IT")
            End If
        Set rstNOI = Nothing
        Forms!wizNOI![NOI] = Null
        
                  
            If Forms!wizNOI!txtClientSentNOI = "C" Then
                AddStatus FileNumber, Now(), "Removed C Of NOI by " & GetFullName
                
                DoCmd.SetWarnings False
                strinfo = "Removed C Of NOI by " & GetFullName
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizNOI!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery
            
                Forms!wizNOI!txtClientSentNOI = ""
            End If
            
                Forms!wizNOI.Requery
        Else
        Exit Sub
        End If
        
    End If
    
Else
MsgBox ("The File has dispsotion not buy in or 3rd party, proceduer canceld")
Exit Sub
End If

End Sub

Private Sub NewDemand_Click()
If (IsNull(Disposition) Or (Disposition = 1 Or Disposition = 2)) Then
    If Not IsNull([AccelerationIssued]) Or Not IsNull([AccelerationLetter]) Or ([ClientSentAcceleration] = "C") Then
    If MsgBox(" You are about to remove dates ? ", vbOKCancel) = vbOK Then
        AddStatus FileNumber, Now(), "Removed Demand Issued (" & [AccelerationIssued] & ") by " & GetFullName
    
        DoCmd.SetWarnings False
        strinfo = "Removed Demand Issued (" & [AccelerationIssued] & ") by " & GetFullName
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizNOI!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
    
        'Put file in Demand queue
        Dim rstAccelerationIssued As Recordset
        Set rstAccelerationIssued = CurrentDb.OpenRecordset("Select * From WizardQueueStats Where FileNumber = " & FileNumber & "And Current = True", dbOpenDynaset, dbSeeChanges)
            If Not rstAccelerationIssued.EOF Then
                With rstAccelerationIssued
                .Edit
                .Edit
                    !DemandComplete = Null
                    !DemandWaiting = Null
                    !DemandQueue = Null
                .Update
                End With
                Else
                MsgBox ("There is no Currrent Wizard Record for this File, Please Contact the IT")
            End If
        Set rstAccelerationIssued = Nothing
        Forms!wizNOI![AccelerationIssued] = Null
        
            If Not IsNull(Forms!wizNOI![AccelerationLetter]) Then
                AddStatus FileNumber, Now(), "Removed Demand Expires date (" & [AccelerationLetter] & ") by " & GetFullName
                
                DoCmd.SetWarnings False
                strinfo = "Removed Demand Expires date (" & [AccelerationLetter] & ") by " & GetFullName
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizNOI!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery
                
                Forms!wizNOI!AccelerationLetter = Null
            End If
            
            If Forms!wizNOI!txtClientSentAcceleration = "C" Then
                AddStatus FileNumber, Now(), ":  Removed C from the Demand Field by " & GetFullName
                
                DoCmd.SetWarnings False
                strinfo = ":  Removed C from the Demand Field by " & GetFullName
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizNOI!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery
            
                Forms!wizNOI!txtClientSentAcceleration = ""
            End If
            
            Forms!wizNOI.Requery
         Else
         Exit Sub
         End If
         
    End If

Else
MsgBox ("The File has dispsotion not buy in or 3rd party, proceduer canceld")
Exit Sub
End If


End Sub

'
Private Sub NewFairDebt_Click()
If (IsNull(Disposition) Or (Disposition = 1 Or Disposition = 2)) Then

    If Not IsNull([FairDebtDate]) Then
       If MsgBox(" You are about to remove dates ? ", vbOKCancel) = vbOK Then
        AddStatus FileNumber, Now(), "Removed Fair Debt (" & [FairDebtDate] & ") by " & GetFullName
    
        DoCmd.SetWarnings False
        strinfo = "Removed Fair Debt (" & [FairDebtDate] & ") by " & GetFullName
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizNOI!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
    
        'Put file in FairDebt queue
        Dim rstWFairDebt As Recordset
        Set rstWFairDebt = CurrentDb.OpenRecordset("Select * From WizardQueueStats Where FileNumber = " & FileNumber & "And Current = True", dbOpenDynaset, dbSeeChanges)
            If Not rstWFairDebt.EOF Then
                With rstWFairDebt
                .Edit
                 ' If IsNull(!RSIIcomplete) Then !RSIIcomplete = #1/1/2011#
                If Not IsNull(!FairDebtComplete) Then !FairDebtComplete = Null
                If Not IsNull(!FairDebtWaiting) Then !FairDebtWaiting = Null
               ' If Not IsNull(!NOIcomplete) Then !NOIcomplete = Null
               ' !DateInQueueNOI = Null
                
                !AddFair = "Fair"
              '  If Not IsNull(!RestartQueue) Then !FairDebtRestart = #1/1/2011#
                .Update
                End With
                Else
                MsgBox ("There is no Currrent Wizard Record for this File, Please Contact the IT")
            End If
        Set rstWFairDebt = Nothing
    
        Forms!wizNOI![FairDebtDate] = Null
        Forms!wizNOI.Requery
        Else
        Exit Sub
        End If
        
        
    
    End If
Else
MsgBox ("The File has dispsotion not buy in or 3rd party, proceduer canceld")
Exit Sub
End If

End Sub

Private Sub SentAtty_Click()
Dim rstqueue As Recordset, rstdocs As Recordset, cntr As Integer
Dim strSQLJournal As String
Dim rstsql As String

If Forms!wizNOI!ClientSentNOI <> "C" Then
    If IsNull(Forms!wizNOI!FairDebtDate) Then
    MsgBox ("There is no Fairdebt date")
    Exit Sub
    End If
End If

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!DateSentAttyNOI = Now
!AttyMilestone1_5 = Null
!AttyMilestone1_5Reject = False
!AttyMilestoneMgr1_5 = Null
If IsNull(rstqueue!DateInWaiitingQueueNOI) Then rstqueue!DateInWaiitingQueueNOI = Now

.Update
End With

Set rstqueue = Nothing

    DoCmd.SetWarnings False
    rstsql = "Insert into ValumeNOI (CaseFile, Client, Name, NOISentBy, NOIAttyReview, NOIAttyReviewC ) values (Forms!wizNOI!FileNumber, ClientShortName(forms!wizNOI!ClientID),Getfullname(),'" & Forms!wizNOI!ClientSentNOI & "',Now(),1) "
    DoCmd.RunSQL rstsql
    
    
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!wizNOI!FileNumber & ",Now,GetFullName(),'" & " 45 days Notice Sent To Atty For Review " & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True




Call ReleaseFile(FileNumber)



MsgBox "NOI Sent to Atty Wizard complete", vbInformation




Call ConfirmationVisible(False)
Call FieldsVisible(False)


'Me.RecordSource = ""
DoCmd.Close acForm, Me.Name

'    If IsLoadedF("wizNOI") = True Then
'
'       Forms!wizNOI!lstFiles.Requery
'        Forms!wizNOI.Requery
'
'
'        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueNOI", dbOpenDynaset)
'        If rstqueue.EOF Then
'            cntr = 0
'            Else
'            rstqueue.MoveLast
'            cntr = rstqueue.RecordCount
'        End If
'        Forms!wizNOI!QueueCount = cntr
'        Set rstqueue = Nothing
'    End If

End Sub
