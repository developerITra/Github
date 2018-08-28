VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DocsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Sub Form_Current()
If PrivSSN Then SpecialDoc.Visible = True
Call UpdateDocumentList

End Sub
Private Sub cmdAddDoc_Click()

'8/29/14
FileNO = FileNumber

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
'If SelectedDocType <> (1511 Or 1513 Or 1514 Or 1515 Or 1516 Or 1517 Or 1518 Or 1519 Or 1520 Or 1521 Or 1522) Then ' not ssn doc

'If PrivSSN Then
FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\" & newfilename

DoCmd.SetWarnings False
strinfo = "Added SSN Document "
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
Case Else
FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & newfilename
End Select

'Else
'MsgBox (" You are not authorized to add SSN ")
'Exit Sub
'End If
'End If 'this is for ssn


'FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & NewFilename old one

'Commented by JAE 'Document Speed'
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



Call UpdateDocumentList
If MsgBox("New document " & newfilename & " accepted.  OK to delete " & Filespec & "?", vbQuestion + vbYesNo) = vbYes Then Kill Filespec


If selecteddoctype = 205 Then
If Forms!foreclosuredetails!WizardSource = "Title" Then
Forms!foreclosuredetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!cmdcloserestart.Visible = True

'Forms!foreclosuredetails!cmdcloserestart.Visible = False
End If
End If




Select Case selecteddoctype
Dim rstFCdetails As Recordset
Case 1   'If title or title update, update fc records

   Call GeneralMissingDoc(FileNumber, 1, False, False, True, False, False)
   Call GeneralMissingDoc(FileNumber, 1, False, False, False, False, False, , True)
    
    

AddStatus FileNumber, Date, "Received title"
If CaseTypeID = 1 Then
    If ClientID <> 328 Or ClientID = 328 Then 'SLS does own title
    
        'change made on 2_27_15
    DoCmd.OpenForm "GetTitleSearchFee", , , , , acDialog, "Enter Title Search costs, zero if none|FC-TC1|Title Search|Abstractor"

    
    'DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter Title Search costs, zero if none|FC-TC1|Title Search|Abstractor"
        'If JurisdictionID = 4 Or JurisdictionID = 18 Then 'PG and Balt City judgment search
        'DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "PG County or Balt City File:  Enter Judgment Search costs, zero if none|FC-TC1|Judgment Search|Title"
        'End If
          If Forms!foreclosuredetails!WizardSource = "TitleOut" Then
                If MsgBox(" Do you need to Upload a Bill-Title document ? ", vbYesNo) = vbNo Then
                Forms!foreclosuredetails!cmdWizComplete.Visible = True
                Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Upload completed"
                Forms!foreclosuredetails!cmdWizComplete.SetFocus
                Forms!foreclosuredetails!cmdWaiting.Visible = False
                'Forms!foreclosuredetails!cmdcloserestart.Visible = False
                Else
                BillTitle = True
                'DoCmd.OpenForm "Select Document Type"
                'Forms![Select Document Type].Visible = False
                cmdAddDoc.SetFocus
                End If
    
            End If
    End If
End If
    
Case 591   'If title or title update, update fc records
    
        Call GeneralMissingDoc(FileNumber, 591, False, False, False, False, False, , True)
        Call GeneralMissingDoc(FileNumber, 591, False, False, True, False, False)
    AddStatus FileNumber, Date, "Received title"
If CaseTypeID = 1 Then
    If ClientID <> 328 Or ClientID = 328 Then 'SLS does own title
    
        'change made on 2_27_15
    DoCmd.OpenForm "GetTitleSearchFee", , , , , acDialog, "Enter Title Search costs, zero if none|FC-TC1|Title Search|Abstractor"

    
    
    'DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter Title Search costs, zero if none|FC-TC1|Title Search|Abstractor"
        'If JurisdictionID = 4 Or JurisdictionID = 18 Then 'PG and Balt City judgment search
        'DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "PG County or Balt City File:  Enter Judgment Search costs, zero if none|FC-TC1|Judgment Search|Title"
        'End If
            If Forms!foreclosuredetails!WizardSource = "TitleOut" Then
                If MsgBox(" Do you need to Upload a Bill-Title Update document ? ", vbYesNo) = vbNo Then
                Forms!foreclosuredetails!cmdWizComplete.Visible = True
                Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Update Upload completed"
                Forms!foreclosuredetails!cmdWizComplete.SetFocus
                Forms!foreclosuredetails!cmdWaiting.Visible = False
                Forms!foreclosuredetails!cmdcloserestart.Visible = False
                Else
                BillTitleUpdate = True
                'DoCmd.OpenForm "Select Document Type"
                'Forms![Select Document Type].Visible = False
                cmdAddDoc.SetFocus
                End If

            End If
    End If
End If



Case 16
If BillTitle = True Then
Forms!foreclosuredetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Upload completed"
Forms!foreclosuredetails!cmdWizComplete.SetFocus
Forms!foreclosuredetails!cmdWaiting.Visible = False
'Forms!foreclosuredetails!cmdcloserestart.Visible = False
End If

Case 1360
If BillTitleUpdate = True Then
Forms!foreclosuredetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Update Upload completed"
Forms!foreclosuredetails!cmdWizComplete.SetFocus
Forms!foreclosuredetails!cmdWaiting.Visible = False
Forms!foreclosuredetails!cmdcloserestart.Visible = False
End If



Case 1549 ' These for Demand missing docs SA 07/24/14
DoCmd.SetWarnings False
    'Dim strSQL As String
    'Demand
    strSQL = "UPDATE DemandDocsNeeded SET " & " DocReceived = #" & Now() & "# AND docreceivedby = " & GetStaffID & _
    " WHERE FileNumber = " & FileNumber & " AND DocName = ('" & "Missing Figures" & "')"
    DoCmd.RunSQL strSQL
    strSQL = ""
    'NOI
    strSQL = "UPDATE DocumentMissing SET " & " DocRecd = -1 , DocRecdBy = " & GetStaffID & _
    " WHERE FileNbr = " & FileNumber & " AND DocName = ('" & "Rfigs" & "')"
    DoCmd.RunSQL strSQL
    strSQL = ""
    
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "Reinstatement Figures Document Uploaded" & "',1 )"
    DoCmd.RunSQL strSQLJournal
    strSQLJournal = ""
    
    
    
    Dim rstdocs As Recordset
    Dim rstqueue As Recordset
    
    Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DemandDocsNeeded where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
    If rstdocs.EOF Then
       
        strSQL = "UPDATE wizardqueuestats SET DemandDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                DoCmd.RunSQL strSQL
                strSQL = ""
        Set rstqueue = Nothing
        Set rstdocs = Nothing
    End If
        
    Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DocumentMissing where FileNbr=" & FileNumber & " AND Not (DocRecd)", dbOpenDynaset, dbSeeChanges)
        If rstdocs.EOF Then
           
            strSQL = "UPDATE wizardqueuestats SET DocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                    DoCmd.RunSQL strSQL
                    strSQL = ""
            Set rstqueue = Nothing
            Set rstdocs = Nothing
    End If
DoCmd.SetWarnings True


Case 988 ' These for NOI missing docs SA 08/24/14
DoCmd.SetWarnings False
    
    ' NOI
    strSQL = "UPDATE DocumentMissing SET " & " DocRecd = -1 , DocRecdBy = " & GetStaffID & _
    " WHERE FileNbr = " & FileNumber & " AND DocName = ('" & "Payment Dates" & "')"
    DoCmd.RunSQL strSQL
    strSQL = ""
     
    strSQL = "UPDATE DocumentMissing SET " & " DocRecd = -1 , DocRecdBy = " & GetStaffID & _
    " WHERE FileNbr = " & FileNumber & " AND DocName = ('" & "Default Dates" & "')"
    DoCmd.RunSQL strSQL
    strSQL = ""
    
    
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "Payment History Document Uploaded" & "',1 )"
    DoCmd.RunSQL strSQLJournal
    strSQLJournal = ""
   
    Forms!Journal.Requery
    
    Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DocumentMissing where FileNbr=" & FileNumber & " AND Not (DocRecd)", dbOpenDynaset, dbSeeChanges)
        If rstdocs.EOF Then
           strSQL = "UPDATE wizardqueuestats SET DocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                    DoCmd.RunSQL strSQL
                    strSQL = ""
            Set rstqueue = Nothing
            Set rstdocs = Nothing
    End If
DoCmd.SetWarnings True

Case 1553 ' These for NOI missing docs SA 08/24/14
DoCmd.SetWarnings False
    
    ' NOI
    strSQL = "UPDATE DocumentMissing SET " & " DocRecd = -1 , DocRecdBy = " & GetStaffID & _
    " WHERE FileNbr = " & FileNumber & " AND DocName = ('" & "Client Sent NOI Copy" & "')"
    DoCmd.RunSQL strSQL
    strSQL = ""
    
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "Client Sent NOI Document Uploaded" & "',1 )"
    DoCmd.RunSQL strSQLJournal
    strSQLJournal = ""
   
    Forms!Journal.Requery
    
    Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DocumentMissing where FileNbr=" & FileNumber & " AND Not (DocRecd)", dbOpenDynaset, dbSeeChanges)
        If rstdocs.EOF Then
           strSQL = "UPDATE wizardqueuestats SET DocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                    DoCmd.RunSQL strSQL
                    strSQL = ""
            Set rstqueue = Nothing
            Set rstdocs = Nothing
    End If
DoCmd.SetWarnings True


Case 1550 ' Demand missing docs SA 07/24/14
    DoCmd.SetWarnings False
    strSQL = "UPDATE DemandDocsNeeded SET " & " DocReceived = #" & Now() & "# AND docreceivedby = " & GetStaffID & _
    " WHERE FileNumber = " & FileNumber & " AND DocName = ('" & "Need Fee Approval" & "')"
    DoCmd.RunSQL strSQL
    strSQL = ""
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "Demand Fee Approval Dcument Uploaded" & "',1 )"
    DoCmd.RunSQL strSQLJournal
    strSQLJournal = ""
    DoCmd.SetWarnings True




End Select






Exit_cmdAddDoc_Click:
    Exit Sub

Err_cmdAddDoc_Click:
'    If Err.Number = 76 Then     ' path not found
'
'        If SelectedDocType <> (1511 Or 1513 Or 1514 Or 1515 Or 1516 Or 1517 Or 1518 Or 1519 Or 1520 Or 1521 Or 1522) Then
'         MkDir DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\"
'        Else
'         'If PrivSSN Then
'            MkDir DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\"
'            Dim objFSO
'            Dim objFolder
'            Set objFSO = CreateObject("Scripting.FileSystemObject")
'            Set objFolder = objFSO.GetFolder(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN")
'                If objFolder.Attributes = objFolder.Attributes And 2 Then
'                   objFolder.Attributes = objFolder.Attributes Xor 2
'                End If
'
'           DoCmd.SetWarnings False
'            strInfo = "Added SSN Document "
'            strInfo = Replace(strInfo, "'", "''")
'            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strInfo & "',1 )"
'            DoCmd.RunSQL strSQLJournal
'            DoCmd.SetWarnings True
'
'          'Else
'         ' MsgBox ("You are not authorized to Add SSN Doc")
'          'Exit Sub
'         ' End If
'
'         End If
'
'     '   MkDir DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\"
'        Resume
'    Else
        MsgBox Err.Description
        Resume Exit_cmdAddDoc_Click
'    End If
End Sub
Private Sub lstDocs_DblClick(Cancel As Integer)
Call cmdView_Click
End Sub
Private Sub cmdAddDocRequest_Click()
On Error GoTo Err_cmdAddDocRequest_Click

DoCmd.OpenForm "Add Document Request"

Exit_cmdAddDocRequest_Click:
  Exit Sub
  
Err_cmdAddDocRequest_Click:
  MsgBox Err.Description
  Resume Exit_cmdAddDocRequest_Click
  
End Sub

Private Sub cmdAll_Click()
Dim i As Long

On Error GoTo Err_cmdAll_Click

For i = 0 To lstDocs.ListCount - 1
    lstDocs.Selected(i) = True
Next i

Exit_cmdAll_Click:
    Exit Sub

Err_cmdAll_Click:
    MsgBox Err.Description
    Resume Exit_cmdAll_Click
    
End Sub

Private Sub cmdInvert_Click()
Dim i As Long

On Error GoTo Err_cmdInvert_Click

For i = 0 To lstDocs.ListCount - 1
    If lstDocs.Selected(i) Then
        lstDocs.Selected(i) = False
    Else
        lstDocs.Selected(i) = True
    End If
Next i

Exit_cmdInvert_Click:
    Exit Sub

Err_cmdInvert_Click:
    MsgBox Err.Description
    Resume Exit_cmdInvert_Click
    
End Sub

Private Sub cmdView_Click()
Dim i As Long

On Error GoTo Err_cmdView_Click

For i = 0 To lstDocs.ListCount - 1

        Select Case lstDocs.Column(4, i)
       
        Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572
        'If lstDocs.Column(4, i) = (1511 Or 1513 Or 1514 Or 1515 Or 1516 Or 1517 Or 1518 Or 1519 Or 1520 Or 1521 Or 1522) Then
        If PrivSSN Then
                If lstDocs.Selected(i) Then StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\" & lstDocs.Column(3, i)
                Else
                MsgBox ("You are not authirized to open SSN")
                Exit Sub
                End If
        Case Else
        If lstDocs.Selected(i) Then StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)
       ' End If
       End Select
       
        
Next i

Exit_cmdView_Click:
    Exit Sub

Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click
    
        
    
    
End Sub
Private Sub UpdateDocumentList()
Dim GroupName As String

On Error GoTo UpdateDocumentListErr

Select Case optDocType
    Case 1
        GroupName = ""
    Case 2
        GroupName = "B"
End Select

lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name],DocTitleID FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='" & GroupName & "' AND Filespec IS NOT NULL and DeleteDate is null ORDER BY " & Me.cboSortby
lstDocs.Requery

Exit Sub

UpdateDocumentListErr:
    MsgBox Err.Description, vbCritical
    Exit Sub
    
End Sub
Private Sub cboSortby_AfterUpdate()
  UpdateDocumentList
  
End Sub
Private Sub cmdViewFolder_Click()

On Error GoTo Err_cmdViewFolder_Click

Shell "Explorer """ & DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\""", vbNormalFocus

Exit_cmdViewFolder_Click:
    Exit Sub

Err_cmdViewFolder_Click:
    MsgBox Err.Description
    Resume Exit_cmdViewFolder_Click
    
End Sub
Private Sub cmdDeleteDoc_Click()
On Error GoTo Err_cmdDeleteDoc_Click

If (IsNull(lstDocs.Column(0))) Then
  MsgBox "Please select a document before continuing.", vbCritical, "Select Document"
  Exit Sub
End If

Dim ls_LoginName As String
ls_LoginName = GetLoginName()

DoCmd.SetWarnings False
DoCmd.RunSQL ("UPDATE DocIndex set DeleteDate = Now(), DeleteStaff = '" & ls_LoginName & "' WHERE DocID = " & lstDocs.Column(0))

DoCmd.SetWarnings True

Call UpdateDocumentList

Exit_cmdDeleteDoc_Click:
  Exit Sub
  
Err_cmdDeleteDoc_Click:
  MsgBox Err.Description
  Resume Exit_cmdDeleteDoc_Click
  
End Sub
Private Sub Command111_Click()
On Error GoTo Err_Command111_Click


    DoCmd.Close

Exit_Command111_Click:
    Exit Sub

Err_Command111_Click:
    MsgBox Err.Description
    Resume Exit_Command111_Click
    
End Sub

Private Sub SpecialDoc_Click()
Shell "Explorer """ & DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\""", vbNormalFocus
End Sub
