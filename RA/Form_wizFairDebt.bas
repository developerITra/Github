VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_wizFairDebt"
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

lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], [doctitleid] AS DocType , Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='" & GroupName & "' AND Filespec IS NOT NULL and DeleteDate is null ORDER BY " & Me.cboSortby
lstDocs.Requery



Exit Sub

UpdateDocumentListErr:
    MsgBox Err.Description, vbCritical
    Exit Sub
    


End Sub

Private Sub cmdAcrobat_Click() '****** -2 = PDFS!  At this point they werent using PRintTo  :( :( **********
                                '***** BE SURE TO MAKE CHANGES IN EACH APPROPIATE SECTION         **********
If Me!State = "VA" Then

    If DLookup("clientid", "caselist", "filenumber=" & FileNumber) = 97 Then
        Call DoReport("Fair Debt Letter VA Wiz", -2)
        Call DoReport("Loss Mitigation Solicitation Letter wiz JP", -2)
    
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 385 Then
        Call DoReport("Loss Mitigation Solicitation Letter Wiz NSTAR", -2)
        Call DoReport("Fair Debt Letter VA Wiz", -2)
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 602 Then
        Call DoReport("Loss Mitigation Solicitation Letter Wiz", -2)
        Call DoReport("Fair Debt Letter VA Wiz", -2)
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 567 Then   'Champion
        Call DoReport("Fair Debt Letter VA Wiz", -2)
        Call DoReport("Loss Mitigation Solicitation Letter Champion", -2)
    Else
        If LoanType = 5 And State = "VA" Or LoanType = 4 And State = "VA" Then
            Call DoReport("Loss Mitigation Solicitation Letter Wiz", -2)
            Call DoReport("Fair Debt Letter VA Wiz", -2)
        Else
            Call DoReport("Fair Debt Letter VA Wiz", -2)
        End If
    End If
Else
    If DLookup("clientid", "caselist", "filenumber=" & FileNumber) = 97 Then
        Call DoReport("Fair Debt Letter VA Wiz", -2) 'make VA and Md the same tickt 863 05/24
        Call DoReport("Fair Debt Letter VA Wiz", -2)
        Call DoReport("Loss Mitigation Solicitation Letter wiz JP", -2)
    
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 385 Then
        Call DoReport("Loss Mitigation Solicitation Letter Wiz NSTAR", -2)
       ' Call DoReport("Fair Debt Letter Wiz", -2)
        Call DoReport("Fair Debt Letter VA Wiz", -2)
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 602 Then
        Call DoReport("Loss Mitigation Solicitation Letter Wiz", -2)
       ' Call DoReport("Fair Debt Letter Wiz", -2)
        Call DoReport("Fair Debt Letter VA Wiz", -2)
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 567 Then   'Champion
        Call DoReport("Fair Debt Letter VA Wiz", -2)
        Call DoReport("Loss Mitigation Solicitation Letter Champion", -2)
    Else
        'Call DoReport("Fair Debt Letter Wiz", -2)
        Call DoReport("Fair Debt Letter VA Wiz", -2)
    End If

End If

End Sub

Private Sub cmdCalcPerDiem_Click()

On Error GoTo Err_cmdCalcPerDiem_Click
PerDiem = RemainingPBal * InterestRate / 100 / 365

Exit_cmdCalcPerDiem_Click:
    Exit Sub

Err_cmdCalcPerDiem_Click:
    MsgBox Err.Description
    Resume Exit_cmdCalcPerDiem_Click
    
End Sub

Private Sub cmdOKdocsmsng_Click()
Dim rstCase As Recordset, rstFC As Recordset, rstNames As Recordset
Dim FileNum As Long, MissingInfo As String

On Error GoTo Err_cmdOK_Click


Call ReleaseFile(FileNumber)


FileNum = txtFileNumber

'Me.RecordSource = ""
'DoCmd.Close acForm, Me.Name
DoCmd.OpenForm "EnterFairDebtReason"
Forms!EnterFairDebtReason!FileNumber = FileNum

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
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

Case 1371 ', 4, 1517
    Call GeneralMissingDoc(FileNumber, 1371, False, True, True, False, False)
Case 4
    Call GeneralMissingDoc(FileNumber, 4, False, True, True, False, False)
    Call GeneralMissingDoc(FileNumber, 4, False, False, False, False, False, , True)
Case 1517
    Call GeneralMissingDoc(FileNumber, 1517, False, True, True, False, False)
    Call GeneralMissingDoc(FileNumber, 1517, False, False, False, False, False, , True)
Case 1522
    Call GeneralMissingDoc(FileNumber, 1522, False, True, True, False, False)

Case 1450
        Call GeneralMissingDoc(FileNumber, 1450, False, True, True, False, True)

Case 1493
    Call GeneralMissingDoc(FileNumber, 1493, False, False, False, False, False, , True)
    If Me.ClientID <> 385 And Me.ClientID <> 446 Then
        Call GeneralMissingDoc(FileNumber, 1493, False, True, True, False, False, 1)
    End If
Case 1523
    Call GeneralMissingDoc(FileNumber, 1523, False, False, False, False, False, , True)
    If Me.ClientID <> 385 And Me.ClientID <> 446 Then
        Call GeneralMissingDoc(FileNumber, 1523, False, True, True, False, False, 1)
    End If

Case 1569
 If Me.ClientID = 385 Then 'Nationstar
        Call GeneralMissingDoc(FileNumber, 1569, False, True, False, False, False, 385)
 End If

Case 1571
 If Me.ClientID = 385 Then 'Nationstar
        Call GeneralMissingDoc(FileNumber, 1571, False, True, False, False, False, 385)
 End If
 
 Case 1570
 If Me.ClientID = 446 Then 'BOA
        Call GeneralMissingDoc(FileNumber, 1570, False, True, False, False, False, 446)
 End If
 
 Case 1572
 If Me.ClientID = 446 Then 'BOA
        Call GeneralMissingDoc(FileNumber, 1572, False, True, False, False, False, 446)
 End If
 
 Case 1554

 Call GeneralMissingDoc(FileNumber, 1554, False, True, False, False, False)

 
End Select

'If me.ClientID <> 385 and Me.ClientID <> 446


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
Private Sub cmdView_Click()
Dim i As Integer
For i = 0 To lstDocs.ListCount - 1
    
If lstDocs.Selected(i) Then

Select Case lstDocs.Column(4, i)
     
        Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572
    
'            If lstDocs.Column(4, i) = 1516 Then
                If Not PrivSSN Then
                MsgBox (" You are not authorized to open SSN doc")
                Exit Sub
                End If
        End Select


Select Case lstDocs.Column(4, i)
      
        Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572
        StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\" & lstDocs.Column(3, i)
        Case Else
        StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)
        End Select
End If
Next i

Exit_cmdView_Click:
    Exit Sub

Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click

End Sub

Private Sub cmdOK_Click()
Dim rstCase As Recordset, rstFC As Recordset, rstNames As Recordset, rstdocs As Recordset
Dim FileNum As Long, MissingInfo As String, WizardType As String, Exceptions As Long
Dim rstsql As String
MsgBox (Forms!wizfairdebt!cmdOK.Caption)

On Error GoTo Err_cmdOK_Click
If Forms!wizfairdebt!cmdOK.Caption <> "Sent to Atty" Then
   ' If FairDebt = "" Then
        If MsgBox("Update Fair debt Sent ? ", vbYesNo) = vbYes Then
    
            If MsgBox("Fair Debt Sent today ? = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
                   Forms!wizfairdebt!FairDebtDate = Now()
                Else
                 Dim fairdate As Date
AF:                fairdate = InputBox("Input Fair debt in past date with format m/d/yyyy ")
                    If fairdate > Now() Then
                        MsgBox ("You can not put date in futuer")
                        GoTo AF
                    End If
                Forms!wizfairdebt!FairDebtDate = fairdate
            End If
            
             
            
                    
                    
                    
                    DoCmd.SetWarnings False
                    rstsql = "Insert into ValumeFD (CaseFile, Client, Name, FDComplete, FDCompleteC,state ) values (Forms!wizFairDebt!FileNumber, ClientShortName(forms!wizFairDebt!ClientID),Getfullname(),Now(),1, Forms!wizFairDebt!State) "
                    DoCmd.RunSQL rstsql
                    DoCmd.SetWarnings True
                                        
                    'DoCmd.RunCommand acCmdSaveRecord
                    AddStatus FileNumber, Now(), "Fair Debt Letter sent"
        
                    Dim fairdebtCnt As Integer
                    Dim qtypstge As Integer
                    Dim JnlNote As String
                    fairdebtCnt = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and FairDebt = true")
                        If (fairdebtCnt > 0) Then
                                'FeeAmount = Nz(DLookup("FairDebtPostage", "ClientList", "ClientID=" & ClientID), 0)
                                '2-5-15
                                FeeAmount = Nz(DLookup("Value", "StandardCharges", "ID=" & 1))


                                    'If FeeAmount > 0 Then
                                        qtypstge = DCount("[FileNumber]", "[qryFairDebt]", "FileNumber=" & [FileNumber])
                                        AddInvoiceItem FileNumber, "FC-FairDebt", "Fair Debt Postage", (qtypstge * FeeAmount), 76, False, False, False, True
                                    'Else
                                        'AddInvoiceItem FileNumber, "FC-FairDebt", "Fair Debt Postage", 1, 76, False, False, False, True
                                    'End If
            
                        End If
        
                        If LoanType = 5 And State = "VA" Then
            
            
                            Forms!wizfairdebt!LossMitSolicitationDate = Now()
                            AddStatus FileNumber, Now(), "Loss Mitigation Solicitation Letter sent"
                            JnlNote = "Fair Debt & Loss Mit Package sent"
            
                               Dim lossMitSolCnt As Integer
                               lossMitSolCnt = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and FairDebt = true")
                                If (lossMitSolCnt > 0) Then
                                End If
                         Else
            
                        End If
                    
                    
                 JnlNote = "Fair Debt notice sent"
            'Dim lrs As Recordset
            '
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
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
        
        End If
        
   
    
    
    
    
   
End If

'Call AddFileResponsibilityHistory(FileNum, ?, StaffID)
Call FairDebtCompletionUpdate(WizardType, FileNumber)
Call RemoveDocMissFairDebt(FileNumber)
DoCmd.SetWarnings False
Call ReleaseFile(FileNumber)
DoCmd.SetWarnings True
MsgBox "Fair Debt Wizard complete", vbInformation

txtFileNumber = Null
txtFileNumber.SetFocus
Call ConfirmationVisible(False)
Call FieldsVisible(False)


'Me.RecordSource = ""
DoCmd.Close acForm, Me.Name

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

Private Sub cmdPreView_Click()
On Error GoTo Err_cmdpreview_Click

Dim statusMsg As String, FileNum As Long, MissingInfo As String
FileNum = FileNumber

If Not IsNull(txtDisposition) Then
MsgBox "You cannot print a Fair Debt letter for this file because it has a Disposition", vbCritical
Exit Sub
End If

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
FileNum = FileNumber
On Error GoTo Err_cmdpreview_Click
If IsNull(FairDebtAmount) Then
MsgBox "You must enter a Fair Debt Amount before printing the Fair Debt Letter", vbCritical
Exit Sub
End If

If Me!State = "VA" Then
    If DLookup("clientid", "caselist", "filenumber=" & FileNumber) = 97 Then
        Call DoReport("Fair Debt Letter VA Wiz", acPreview)
        Call DoReport("Loss Mitigation Solicitation Letter wiz JP", acPreview)
    
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 385 Then
  '       Call DoReport("Loss Mitigation Solicitation Letter Wiz", acPreview)
         Call DoReport("Fair Debt Letter VA Wiz", acPreview)
        Call DoReport("Loss Mitigation Solicitation Letter Wiz NSTAR", acPreview)
        
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 602 Then   'Mei 10_23/15 open 2 docs for BOKF
        Call DoReport("Fair Debt Letter VA Wiz", acPreview)
        Call DoReport("Loss Mitigation Solicitation Letter Wiz", acPreview)
        
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 567 Then   'Champion
        Call DoReport("Fair Debt Letter VA Wiz", acPreview)
        Call DoReport("Loss Mitigation Solicitation Letter Champion", acPreview)
    
    Else
        If LoanType = 5 And State = "VA" Or LoanType = 4 And State = "VA" Then
            Call DoReport("Loss Mitigation Solicitation Letter Wiz", acPreview)
        End If
        Call DoReport("Fair Debt Letter VA Wiz", acPreview)
    End If
        
Else
    If DLookup("clientid", "caselist", "filenumber=" & FileNumber) = 97 Then
        'Call DoReport("Fair Debt Letter Wiz", acPreview) ' stop as we should make one format ticket no.863
        Call DoReport("Fair Debt Letter VA Wiz", acPreview)
        Call DoReport("Loss Mitigation Solicitation Letter wiz JP", acPreview)
    
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 385 Then
        Call DoReport("Loss Mitigation Solicitation Letter Wiz NSTAR", acPreview)
        'Call DoReport("Fair Debt Letter Wiz", acPreview)
        Call DoReport("Fair Debt Letter VA Wiz", acPreview)
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 602 Then   'Mei 10_23/15 open 2 docs for BOKF
        Call DoReport("Fair Debt Letter VA Wiz", acPreview)
        Call DoReport("Loss Mitigation Solicitation Letter wiz", acPreview)
    ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 567 Then   'Champion
        Call DoReport("Fair Debt Letter VA Wiz", acPreview)
        Call DoReport("Loss Mitigation Solicitation Letter Champion", acPreview)
    Else
       ' Call DoReport("Fair Debt Letter Wiz", acPreview)
        Call DoReport("Fair Debt Letter VA Wiz", acPreview)
    End If

        Forms!wizfairdebt!LossMitSolicitationDate = Now()
        AddStatus FileNum, Now(), "Loss Mitigation Solicitation Letter sent"
End If

Exit Sub
Exit_cmdpreview_Click:
DoCmd.Close acForm, "WizFairDebt"
    Exit Sub

Err_cmdpreview_Click:
    MsgBox Err.Description
    Resume Exit_cmdpreview_Click
    

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

Private Sub ComLable_Click()
Dim sql As String
Dim i As Integer
Dim rstLabelData As Recordset
Dim cntr As Integer
Dim matter As String


sql = "SELECT qryFairDebtLabel.FileNumber, qryFairDebtLabel.Names_Company, qryFairDebtLabel.Names_First, qryFairDebtLabel.Names_Last, qryFairDebtLabel.Names_Address, qryFairDebtLabel.Names_Address2, " & _
"qryFairDebtLabel.Names_City, qryFairDebtLabel.Names_State, qryFairDebtLabel.Names_Zip, qryFairDebtLabel.PrimaryDefName, qryFairDebtLabel.Deceased, ClientList.ShortClientName FROM ClientList INNER JOIN qryFairDebtLabel ON ClientList.ClientID = qryFairDebtLabel.ClientID where qryFairDebtLabel.filenumber=" & Forms![wizfairdebt]!FileNumber


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
            rstLabelData.MoveNext
            cntr = cntr + 1
        Loop
        rstLabelData.Close
        cntr = cntr * 2
   
   
   
  
End Sub

Private Sub CommEdit_Click()

DoCmd.OpenForm "sfrmNamesUpdate", , , WhereCondition:="ID= " & Forms!wizfairdebt!sfrmNames!ID

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



Private Sub comSentAtty_Click()

If CurrentProject.AllForms("queFairDebtWaiting").IsLoaded = True Then
    If Forms!queFairDebtWaiting!lstFiles.Column(11) = 0 Or Forms!queFairDebtWaiting!lstFiles.Column(12) = 0 Then
    MsgBox ("Connot send to attorney until Mising Figures or Needs Military Confirm is removed from waiting")
    Exit Sub
    End If
End If

    
    

Dim rstqueue As Recordset
Dim cntr As Integer
Dim rstsql As String


 Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
 '   If IsNull(rstqueue!FairDebtAttyReview) Then
    With rstqueue
    .Edit
 
    !FairDebtUser = GetStaffID
 '   !FairDebtReason = Null
    !AttyMilestone1 = Null
    !AttyMilestone1Reject = False
    !FairDebtAttyReview = Now()
    !FairDebtWaiting = Now()
    !FairDebtComplete = Null
    !AddFair = ""
    .Update
    End With
Set rstqueue = Nothing

                    DoCmd.SetWarnings False
                    rstsql = "Insert into ValumeFD (CaseFile, Client, Name, FDAttyReview, FDAttyReviewC,state) values (Forms!wizFairDebt!FileNumber, ClientShortName(forms!wizFairDebt!ClientID),Getfullname(),Now(),1, Forms!wizFairDebt!State) "
                    DoCmd.RunSQL rstsql
                    DoCmd.SetWarnings True

Call ReleaseFile(FileNumber)

    DoCmd.SetWarnings False
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!wizfairdebt!FileNumber & ",Now,GetFullName(),'" & " Fair debt Sent To Atty For Review " & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

MsgBox "Fair Debt Sent to Atty Wizard complete", vbInformation



If CurrentProject.AllForms("queFairDebt").IsLoaded = True Then
    Forms!queFairDebt!lstFiles.Requery
    Forms!queFairDebt.Requery
    
    
    
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuefairdebt", dbOpenDynaset, dbSeeChanges)
    Do Until rstqueue.EOF
    cntr = cntr + 1
    rstqueue.MoveNext
    Loop
    Forms!queFairDebt!QueueCount = cntr
    Set rstqueue = Nothing
End If
If CurrentProject.AllForms("queFairDebtwaiting").IsLoaded = True Then
    Forms!queFairDebtWaiting!lstFiles.Requery
    Forms!queFairDebtWaiting.Requery
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuefairdebtwaiting", dbOpenDynaset, dbSeeChanges)
    Do Until rstqueue.EOF
    cntr = cntr + 1
    rstqueue.MoveNext
    Loop
    Forms!queFairDebtWaiting!QueueCount = cntr
    Set rstqueue = Nothing
End If

Call ConfirmationVisible(False)
Call FieldsVisible(False)


'Me.RecordSource = ""
DoCmd.Close acForm, Me.Name


End Sub

Private Sub Form_Current()
If PrivNewNOIFDDemaind Then
NewFairDebt.Visible = True
NewDemand.Visible = True
New45Notice.Visible = True
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Me.RecordSource = ""
End Sub

Private Sub cmdPrint_Click()
Dim qtypstge As Double, cntr As Integer, i As Integer, sql As String, matter As String, rstLabelData As Recordset
Dim statusMsg As String, JnlNote As String, FileNum As Long, MissingInfo As String, rstNames As Recordset, rstwiz As Recordset
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
FileNum = FileNumber
On Error GoTo Err_cmdOK_Click
If Not IsNull(txtDisposition) Then
MsgBox "You cannot print a Fair Debt letter for this file because it has a Disposition", vbCritical
Exit Sub
End If
If IsNull(FairDebtAmount) Then
MsgBox "You must enter a Fair Debt Amount before printing the Fair Debt Letter", vbCritical
Exit Sub
End If
If LoanType = 4 And IsNull(FNMALoanNumber) Then
MsgBox "You must enter a FNMA loan number before printing the Fair Debt Letter", vbCritical
Exit Sub
End If
If LoanType = 5 And IsNull(FHLMCLoanNumber) Then
MsgBox "You must enter a FHLMC loan number before printing the Fair Debt Letter", vbCritical
Exit Sub
End If

If Me!State = "VA" Then
        If DLookup("clientid", "caselist", "filenumber=" & FileNumber) = 97 Then
            Call DoReport("Fair Debt Letter VA Wiz", acViewNormal)
            Call DoReport("Loss Mitigation Solicitation Letter wiz JP", acViewNormal)
        
        ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 385 Then
            'Call DoReport("Loss Mitigation Solicitation Letter Wiz", acViewNormal)
            Call DoReport("Fair Debt Letter VA Wiz", acViewNormal)
            Call DoReport("Loss Mitigation Solicitation Letter Wiz NSTAR", acViewNormal)
        ElseIf DLookup("clientid", "caselist", "filenumber=" & FileNumber) = 602 Then
            Call DoReport("Fair Debt Letter VA Wiz", acViewNormal)
            Call DoReport("Loss Mitigation Solicitation Letter wiz", acViewNormal)
        ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 567 Then   'Champion
            Call DoReport("Fair Debt Letter VA Wiz", acViewNormal)
            Call DoReport("Loss Mitigation Solicitation Letter Champion", acViewNormal)
        Else
            Call DoReport("Fair Debt Letter VA Wiz", acViewNormal)
        End If
Else
        If DLookup("clientid", "caselist", "filenumber=" & FileNumber) = 97 Then
            Call DoReport("Fair Debt Letter VA Wiz", acPreview)
           ' Call DoReport("Fair Debt Letter Wiz", acViewNormal) ticket 861
            Call DoReport("Loss Mitigation Solicitation Letter wiz JP", acViewNormal)
        
        ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 385 Then
            Call DoReport("Loss Mitigation Solicitation Letter Wiz NSTAR", acViewNormal)
           ' Call DoReport("Fair Debt Letter Wiz", acViewNormal) 'Make VA and MD the same t icket no.863 05/24/2014 sa
            Call DoReport("Fair Debt Letter VA Wiz", acViewNormal)
        ElseIf DLookup("clientid", "caselist", "filenumber=" & FileNumber) = 602 Then
            Call DoReport("Fair Debt Letter VA Wiz", acViewNormal)
            Call DoReport("Loss Mitigation Solicitation Letter wiz", acViewNormal)
        ElseIf DLookup("clientid", "Caselist", "filenumber=" & FileNumber) = 567 Then   'Champion
            Call DoReport("Fair Debt Letter VA Wiz", acViewNormal)
            Call DoReport("Loss Mitigation Solicitation Letter Champion", acViewNormal)
        Else
            Call DoReport("Fair Debt Letter VA Wiz", acViewNormal)
        End If
End If

If LoanType = 5 And State = "VA" Or LoanType = 4 And State = "VA" Then
    If DLookup("clientid", "caselist", "filenumber=" & FileNumber) <> 97 Then
    Call DoReport("Loss Mitigation Solicitation Letter Wiz", acViewNormal)
    End If
End If

'Call StartDoc(TemplatePath & "Borrowers_Assistance_Form_Chase.pdf")

   'Sarab stopped today as per Daine 2/215
'If MsgBox("Update Fair Debt Sent = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
'            Forms!wizFairDebt!FairDebtDate = Now()
'            'DoCmd.RunCommand acCmdSaveRecord
'            AddStatus FileNum, Now(), "Fair Debt Letter sent"
'
'            Dim fairdebtCnt As Integer
'            fairdebtCnt = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and FairDebt = true")
'            If (fairdebtCnt > 0) Then
'                    FeeAmount = Nz(DLookup("FairDebtPostage", "ClientList", "ClientID=" & ClientID), 0)
'                        If FeeAmount > 0 Then
'                        qtypstge = DCount("[FileNumber]", "[qryFairDebt]", "FileNumber=" & [FileNumber])
'                        AddInvoiceItem FileNumber, "FC-FairDebt", "Fair Debt Postage", (qtypstge * FeeAmount), 76, False, False, False, True
'                        Else
'                        AddInvoiceItem FileNumber, "FC-FairDebt", "Fair Debt Postage", 1, 76, False, False, False, True
'                        End If
'
'            End If
'
'            If LoanType = 5 And State = "VA" Then
'
'
'                Forms!wizFairDebt!LossMitSolicitationDate = Now()
'                AddStatus FileNum, Now(), "Loss Mitigation Solicitation Letter sent"
'                JnlNote = "Fair Debt & Loss Mit Package sent"
'
'                   Dim lossMitSolCnt As Integer
'                   lossMitSolCnt = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and FairDebt = true")
'                    If (lossMitSolCnt > 0) Then
'                    End If
'           Else
'
'            End If
'         JnlNote = "Fair Debt notice sent"
'    'Dim lrs As Recordset
'    '
'    '            Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'    '            lrs.AddNew
'    '            lrs![FileNumber] = FileNum
'    '            lrs![JournalDate] = Now
'    '            lrs![Who] = GetFullName()
'    '            lrs![Info] = JnlNote & vbCrLf
'    '            lrs![Color] = 1
'    '            lrs.Update
'    '            lrs.Close
'
'        DoCmd.SetWarnings False
'        strinfo = JnlNote & vbCrLf
'        strinfo = Replace(strinfo, "'", "''")
'        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNum & ",Now,GetFullName(),'" & strinfo & "',1 )"
'        DoCmd.RunSQL strSQLJournal
'        DoCmd.SetWarnings True
'
'End If

'Print labels, 1 per address and none to client
        sql = "SELECT Names.Company, Names.First, Names.Last, Names.AKA, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip, Names.Deceased, CaseList.FileNumber, ClientList.FairDebt, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (ClientList RIGHT JOIN (CaseList RIGHT JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID) LEFT JOIN FCdetails ON CaseList.FileNumber=FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & Forms![wizfairdebt]!FileNumber & ") And ((Names.FairDebt)=True) And ((FCdetails.Current)=True));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
  
                Call StartLabel
                Print #6, FormatName(rstLabelData!Company, IIf(rstLabelData!Deceased = True, "Estate of " & rstLabelData!First, rstLabelData!First), rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
                Print #6, "|FONTSIZE 8"
                Print #6, "|BOTTOM"
                Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
                matter = rstLabelData!PrimaryDefName
                Call FinishLabel
 
            rstLabelData.MoveNext
        
        Loop
        rstLabelData.Close
        
cmdCancel.Caption = "Close"
Exit Sub

Exit_cmdOK_Click:
DoCmd.Close acForm, "WizFairDebt"
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    

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

Private Sub lstDocs_DblClick(Cancel As Integer)
Dim i As Integer
For i = 0 To lstDocs.ListCount - 1
    
If lstDocs.Selected(i) Then

Select Case lstDocs.Column(4, i)
        Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572
    
'            If lstDocs.Column(4, i) = 1516 Then
                If Not PrivSSN Then
                MsgBox (" You are not authorized to open SSN doc")
                Exit Sub
                End If
        End Select


Select Case lstDocs.Column(4, i)
        Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572
        StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\" & lstDocs.Column(3, i)
        Case Else
        StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)
        End Select
End If
Next i
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
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizFairDebt!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
    
        'Put file in FairDebt queue
        Dim rstNOI As Recordset
        Set rstNOI = CurrentDb.OpenRecordset("Select * From WizardQueueStats Where FileNumber = " & FileNumber & "And Current = True", dbOpenDynaset, dbSeeChanges)
            If Not rstNOI.EOF Then
                With rstNOI
                .Edit
               ' If IsNull(!FairDebtComplete) Then !FairDebtComplete = #1/2/2012#
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
        Forms!wizfairdebt![NOI] = Null
        
                  
            If Forms!wizfairdebt!txtClientSentNOI = "C" Then
                AddStatus FileNumber, Now(), "Removed C Of NOI by " & GetFullName
                
                DoCmd.SetWarnings False
                strinfo = "Removed C Of NOI by " & GetFullName
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizFairDebt!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery
            
                Forms!wizfairdebt!txtClientSentNOI = ""
            End If
            
                Forms!wizfairdebt.Requery
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
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizFairDebt!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery
        
            'Put file in Demand queue
            Dim rstAccelerationIssued As Recordset
            Set rstAccelerationIssued = CurrentDb.OpenRecordset("Select * From WizardQueueStats Where FileNumber = " & FileNumber & "And Current = True", dbOpenDynaset, dbSeeChanges)
                If Not rstAccelerationIssued.EOF Then
                    With rstAccelerationIssued
                   
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
            Forms!wizfairdebt![AccelerationIssued] = Null
            
                If Not IsNull(Forms!wizfairdebt![AccelerationLetter]) Then
                    AddStatus FileNumber, Now(), "Removed Demand Expires date (" & [AccelerationLetter] & ") by " & GetFullName
                    
                    DoCmd.SetWarnings False
                    strinfo = "Removed Demand Expires date (" & [AccelerationLetter] & ") by " & GetFullName
                    strinfo = Replace(strinfo, "'", "''")
                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizFairDebt!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                    DoCmd.RunSQL strSQLJournal
                    DoCmd.SetWarnings True
                    Forms!Journal.Requery
                    
                    Forms!wizfairdebt!AccelerationLetter = Null
                End If
                
                If Forms!wizfairdebt!txtClientSentAcceleration = "C" Then
                    AddStatus FileNumber, Now(), ":  Removed C from the Demand Field by " & GetFullName
                    
                    DoCmd.SetWarnings False
                    strinfo = ":  Removed C from the Demand Field by " & GetFullName
                    strinfo = Replace(strinfo, "'", "''")
                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!wizFairDebt!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                    DoCmd.RunSQL strSQLJournal
                    DoCmd.SetWarnings True
                    Forms!Journal.Requery
                
                    Forms!wizfairdebt!txtClientSentAcceleration = ""
                End If
                
                Forms!wizfairdebt.Requery
                Else
                Exit Sub
                End If
                
                
        End If
        
Else
MsgBox ("The File has dispsotion not buy in or 3rd party, proceduer canceld")
Exit Sub
End If


End Sub

Private Sub NewFairDebt_Click()
If (IsNull(Disposition) Or (Disposition = 1 Or Disposition = 2)) Then

        If Not IsNull([FairDebtDate]) Then
        If MsgBox(" You are about to remove dates ? ", vbOKCancel) = vbOK Then
            AddStatus FileNumber, Now(), "Removed Fair Debt (" & [FairDebtDate] & ") by " & GetFullName
        
            DoCmd.SetWarnings False
            strinfo = "Removed Fair Debt (" & [FairDebtDate] & ") by " & GetFullName
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!WizFairDebt!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
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
               '     If Not IsNull(!NOIcomplete) Then !NOIcomplete = Null
                    !AddFair = "Fair"
                  '  If Not IsNull(!RestartQueue) Then !FairDebtRestart = #1/1/2011#
                    .Update
                    End With
                    Else
                    MsgBox ("There is no Currrent Wizard Record for this File, Please Contact the IT")
                End If
            Set rstWFairDebt = Nothing
        
            Forms!wizfairdebt![FairDebtDate] = Null
            Forms!wizfairdebt.Requery
        
          
        Else
        Exit Sub
        End If
        
    
    End If

Else
MsgBox ("The File has dispsotion not buy in or 3rd party, proceduer canceld")
Exit Sub
End If


End Sub
