VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_WizDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub AccelerationExpires_AfterUpdate()
If Not IsNull([AccelerationExpires]) And IsNull(AccelerationExpires.OldValue) Then
CmdContinue.Visible = True
cmdClientSentAccel.Caption = "Cancel"
End If
End Sub

Private Sub AccelerationIssued_AfterUpdate()
If Not IsNull([AccelerationIssued]) And IsNull(AccelerationIssued.OldValue) Then
CmdContinue.Visible = True
cmdClientSentAccel.Caption = "Cancel"
End If

End Sub




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

Private Sub CmdContinue_Click()
If MsgBox("Are you sure the Client mailed the Demand? ", vbYesNo) = vbYes Then
Forms!WizDemand!AccelerationIssued.Locked = False
Forms!WizDemand!AccelerationExpires.Locked = False

    ClientSentDemand = True
    Call cmdAddDoc_Click
    Call ClientSentDemandProceduer
    Call DemandCompletionUpdate("", FileNumber)
    
'    DoCmd.SetWarnings False
'    Dim rstsql As String
'    rstsql = "Insert InTo ValumeDemand (CaseFile, Client, Name, DemandWaiting, DemandWaitingC,Demandcompleted, DemandcompletedC,SentByClient ) Values ( FileNumber, ClientShortName(forms!wizdemand!ClientID),Getfullname(),null,0,now(),1, 'Yes')"
'    DoCmd.RunSQL rstsql
'    DoCmd.SetWarnings True
    
    
    'Call ReleaseFile(FileNumber)
    'MsgBox "Demand Letter Wizard complete", vbInformation
    Forms!WizDemand.Requery
    
   ' DoCmd.Close acForm, "WizDemand"
  '  DoCmd.Close acForm, "Journal"

End If



End Sub

Private Sub cmdOKdocsmsng_Click()
Dim rstCase As Recordset, rstFC As Recordset, rstNames As Recordset
Dim FileNum As Long, MissingInfo As String

On Error GoTo Err_cmdOK_Click

Call ReleaseFile(FileNumber)

FileNum = txtFileNumber

'Me.RecordSource = ""
'DoCmd.Close acForm, Me.Name
DoCmd.OpenForm "EnterDemandReason"
Forms!EnterDemandReason!FileNumber = FileNum

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
If ClientSentDemand Then
 selecteddoctype = 124
    ClientSentDemand = False
 GoTo CSDP 'Client Sent Demand Process
End If
    DoCmd.OpenForm "Select Document Type", , , , , acDialog
    If selecteddoctype = 0 Then Exit Sub
CSDP:
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


'Update Queue missing Document

Select Case selecteddoctype
'Sub GeneralMissingDoc(FileNumber As Integer, DocTitleNO As Integer, Demaind As Boolean, FD As Boolean, Intake As Boolean, NOI As Boolean, Dockting As Boolean)

Case 1550

    Call GeneralMissingDoc(FileNumber, 1550, True, False, False, False, False)
Case 1549
    Call GeneralMissingDoc(FileNumber, 1549, True, False, False, True, False)
   
Case 124
        Dim costamt As Currency
        costamt = InputBox("Please enter Demand postage")
        AddInvoiceItem FileNumber, "FC-Demand", "Demand postage", costamt, 76, False, True, False, False
       Call GeneralMissingDoc(FileNumber, 124, True, False, False, False, False)
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

Private Sub cmdOKd_Click()
Dim rstCase As Recordset, rstFC As Recordset, rstNames As Recordset, rstdocs As Recordset
Dim FileNum As Long, MissingInfo As String, WizardType As String, Exceptions As Long
Dim rstsql As String
Dim strinfo As String
Dim strSQLJournal  As String
Dim sql As String, matter As String
Dim wzdQue As Recordset
Dim rstLabelData As Recordset
Dim i As Integer, cntr As Integer


If Forms!WizDemand!ClientSentAcceleration = "C" Then
    If IsNull(AccelerationIssued) Or IsNull(AccelerationLetter) Then
    MsgBox "The date the demand letter was sent needs to be entered before completing this wizard.", vbCritical
    Exit Sub
    Else
    
    DoCmd.SetWarnings False
    rstsql = "Insert InTo ValumeDemand (CaseFile, Client, Name, DemandWaiting, DemandWaitingC,Demandcompleted, DemandcompletedC,SentByClient ) Values ( FileNumber, ClientShortName(forms!wizdemand!ClientID),Getfullname(),null,0,now(),1, 'Yes')"
    DoCmd.RunSQL rstsql
    DoCmd.SetWarnings True
                Set wzdQue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where current = true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
                With wzdQue
                .Edit
                !DateSentAttyDemand = Null
                !AttyMilestone1_25 = Null
                !AttyMilestone1_25Reject = False
                !AttyMilestoneMgr1_25 = Null
                !DemandWaiting = Null
                !DemandComplete = Now
                !DemandUser = GetStaffID
                !DemandQueue = Null
                .Update
                .Close
                End With
   
    
    AddStatus FileNumber, Date, "Demand Sent By Client Completed"

    DoCmd.SetWarnings False
    strinfo = "Demand Sent By Client Completed"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

    Call DemandCompletionUpdate("", FileNumber)
    Call ReleaseFile(FileNumber)
    DoCmd.Close acForm, "Journal"
    DoCmd.Close acForm, Me.Name
    
   
    MsgBox "Demand Letter Wizard complete", vbInformation
    Exit Sub
    End If
Else 'RA

If IsNull(Forms!WizDemand!FairDebtDate) Then
MsgBox ("There is no FairDebt Date, cannot complete Demand")
Exit Sub
Else

 'Started invocing
        
        Dim ClientID As Integer, qtypstge As Integer, LoanType As Integer
                LoanType = Forms!WizDemand!LoanType
                ClientID = DLookup("clientid", "caselist", "filenumber=" & FileNumber)
                    'Select Case LoanType
                        'Case 4
                        'FeeAmount = Nz(DLookup("FeeAcceleration", "ClientList", "ClientID=177"))
                        'Case 5
                        'FeeAmount = Nz(DLookup("FeeAcceleration", "ClientList", "ClientID=263"))
                        'Case Else
                        FeeAmount = Nz(DLookup("FeeAcceleration", "ClientList", "ClientID=" & ClientID))
                    'End Select
                        
                        
                
                'Discretionary invoicing for demand letters (ability to override)
                        
                'If MsgBox("Do you want to override the standard fee of $" & FeeAmount & " for this client?", vbYesNo) = vbYes Then ' as per Diane request on 06/26 sA
                'FeeAmount = InputBox("Please enter fee, then rememeber to note the journal")
                'MsgBox "Please upload fee approval to documents"
                'End If
                        
                If FeeAmount = 0 Or IsNull(FeeAmount) Then
                    MsgBox ("See Operations Manager for fee for this client, Cannot complete Demand wizard until fee is entered")
                    Dim rs As Recordset
                    Set rs = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)

                        With rs
                        .Edit
                        !DateSentAttyDemand = Null
                        !DemandWaiting = Now
                        !AttyMilestone1_25 = Null
                        !AttyMilestone1_25Reject = False
                        !AttyMilestoneMgr1_25 = Null
                        !DemandComplete = Null

                        .Update
                        End With
                        rs.Close
                        Set rs = Nothing
                Exit Sub
                End If
                
         
                
                'If FeeAmount > 0 Then
                    AddInvoiceItem FileNumber, "FC-Acc", "Acceleration Letter", FeeAmount, 0, True, True, False, False
                'else
                    'AddInvoiceItem FileNumber, "FC-Acc", "Acceleration Letter", 1, 0, True, True, False, False 'set unknown fee as $1, per Diane
                'End If
               
        '        FeeAmount = Nz(DLookup("AccelerationPostage", "ClientList", "ClientID=" & ClientID))
        '        If FeeAmount > 0 Then
        '        qtyPstge = DCount("[FileNumber]", "[qryFairDebt]", "FileNumber=" & [FileNumber])
        '        AddInvoiceItem FileNumber, "FC-Acceleration", "Acceleration Letter mailed", (qtyPstge * FeeAmount), 76, False, True, False, True
        '        Else
        '        AddInvoiceItem FileNumber, "FC-Acceleration", "Acceleration Letter mailed", 1, 76, False, True, False, True
        '        End If
                Dim iDemand As Integer
                iDemand = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and Demand = true")
                 If (iDemand > 0) Then
                    FeeAmount = Nz(DLookup("AccelerationPostage", "ClientList", "ClientID=" & ClientID), 0)
                    'If FeeAmount > 0 Then
                    'qtypstge = DCount("[FileNumber]", "[qryFairDebt]", "FileNumber=" & [FileNumber])
                    'AddInvoiceItem FileNumber, "FC-ACC", "Acceleration Letter Postage", (qtypstge * FeeAmount), 76, False, False, False, True
                    AddInvoiceItem FileNumber, "FC-ACC", "Acceleration Letter Postage", (iDemand * FeeAmount), 76, False, False, False, True

                    'Else
                    'AddInvoiceItem FileNumber, "FC-ACC", "Acceleration Letter Postage", 1, 76, False, False, False, True
                    'End If
                 End If
            'end invocing
On Error GoTo Err_cmdOK_Click


    If MsgBox("Would you like to Update the Demand letter date? ", vbYesNo) = vbYes Then
    AccelerationIssued = Date
    
    
    
        DoCmd.SetWarnings False
       
        rstsql = "Insert InTo ValumeDemand (CaseFile, Client, Name, DemandWaiting, DemandWaitingC,Demandcompleted, DemandcompletedC,SentByClient ) Values ( FileNumber, ClientShortName(forms!wizdemand!ClientID),Getfullname(),null,0,now(),1, 'No')"
        DoCmd.RunSQL rstsql
        DoCmd.SetWarnings True
        
       
            
    
    
    AccelerationExpires = DateAdd("d", Date, InputBox("Please enter the number of days expired"))
    
    
    End If
    
    If IsNull(AccelerationIssued) Then
    MsgBox "The date the demand letter was sent needs to be entered before completing this wizard.", vbCritical
    Exit Sub
    End If
    
    
    
     
                    Set wzdQue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where current = true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
                    With wzdQue
                    .Edit
                    !DateSentAttyDemand = Null
                    !AttyMilestone1_25 = Null
                    !AttyMilestone1_25Reject = False
                    !AttyMilestoneMgr1_25 = Null
                    !DemandWaiting = Null
                    !DemandComplete = Now
                    !DemandUser = GetStaffID
                    !DemandQueue = Null
                    .Update
                    .Close
                    End With
       
        
        AddStatus FileNumber, Date, "Demand Sent By RA Completed"
    
        DoCmd.SetWarnings False
        strinfo = "Demand Sent by RA Completed"
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
    
    Call DemandCompletionUpdate("", FileNumber)
    
    Call ReleaseFile(FileNumber)
    MsgBox "Demand Letter Wizard complete", vbInformation
    
    
    
    
   ' txtFileNumber = Null
    'txtFileNumber.SetFocus
  '  Call ConfirmationVisible(False)
   ' Call FieldsVisible(False)
    
    DoCmd.Close acForm, "Journal"
    DoCmd.Close acForm, Me.Name
    
Exit_cmdOK_Click:
        Exit Sub
    
Err_cmdOK_Click:
        MsgBox Err.Description
        Resume Exit_cmdOK_Click
        
End If
End If

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

Private Sub ComAddName_Click()
DoCmd.OpenForm "sfrmNamesUpdate", , , , acFormAdd

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
Dim sql As String, matter As String
Dim rstLabelData As Recordset
Dim i As Integer, cntr As Integer

'Print labels, 4 per address
        sql = "SELECT qryDemand.FileNumber, qryDemand.Names_Company, qryDemand.Names_First, qryDemand.Names_Last, qryDemand.Names_Address, qryDemand.Names_Address2, qryDemand.Names_City, qryDemand.Names_State, qryDemand.Names_Zip, qryDemand.PrimaryDefName, qryDemand.Deceased, ClientList.ShortClientName"
        sql = sql + " FROM ClientList INNER JOIN qryDemand ON ClientList.ClientID = qryDemand.ClientID"
        sql = sql + " where qryDemand.filenumber=" & Forms![WizDemand]!txtFileNumber
       
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

Private Sub Command254_Click()
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

Private Sub CommEdit_Click()
DoCmd.OpenForm "sfrmNamesUpdate", , , WhereCondition:="ID= " & Forms!WizDemand!sfrmNames!ID

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
If PrivNewNOIFDDemaind Then
NewFairDebt.Visible = True
NewDemand.Visible = True
New45Notice.Visible = True
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
Me.RecordSource = ""
'Forms!WizDemand.lstDocs.Requery

End Sub

Private Sub cmdPrint_Click()
DoCmd.OpenForm "Print Demand", , , "fileNumber=" & Forms!WizDemand!FileNumber & " and current=true", , , acViewNormal & "|FC"
Forms![Print Demand]!Option64.Enabled = True
Forms![Print Demand]!optDocType = 3

End Sub
Private Sub cmdPreView_Click()
If IsNull([FairDebtDate]) Then
    MsgBox ("Cannot process demand untill Fair Debt Sent. Put on Waiting Q")
    Exit Sub
Else

    DoCmd.OpenForm "Print Payoff", , , "fileNumber=" & Forms!WizDemand!FileNumber & " and current=true", , , acPreview & "|FC"
    Forms![Print Payoff]!Option64.Enabled = True
    Forms![Print Payoff]!optDocType = 3
End If
End Sub


Private Sub ConfirmationVisible(SetVisible As Boolean)


'cmdYes.Enabled = SetVisible
'cmdNo.Enabled = SetVisible

End Sub

Private Sub FieldsVisible(SetVisible As Boolean)

tabWiz.Visible = SetVisible
cmdOKd.Enabled = SetVisible
'AssessedValue.Enabled = (UCase$(Nz(State)) = "VA")

If SetVisible Then
    DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
Else
    DoCmd.Close acForm, "Journal"
End If

End Sub
Private Sub cmdClientSentAccel_Click()
'If Not IsNull([AccelerationIssued]) And Not IsNull([AccelerationExpires]) Then
'    If MsgBox("Are you sure you want to continue with Client Sends Demand proceduer?", vbYesNo) = vbYes Then
'
'
'
'    Dim wzdQue As Recordset
'                Set wzdQue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where current = true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
'                With wzdQue
'                .Edit
'                !ClientSentAcceleration = True
'                .Update
'                .Close
'                End With
'
'    Me.Requery
'    Dim rstFCdetails As Recordset
'                Set rstFCdetails = CurrentDb.OpenRecordset("Select * FROM FCdetails where current = true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
'                With rstFCdetails
'                .Edit
'                !ClientSentAcceleration = "C"
'                .Update
'                .Close
'                End With
'    AddStatus FileNumber, Date, "Acceleration Letter noted as sent by client"
'
'
'    DoCmd.SetWarnings False
'    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!WizDemand!txtFileNumber,Now(),GetFullName(),'" & "Demand sent by Client" & " '," & 1 & ")"
'    DoCmd.RunSQL strSQLJournal
'    DoCmd.SetWarnings True
'
'    Call cmdOKd_Click
'
'    cmdOK.Enabled = False
'    Else
'    Me.Undo
'    AccelerationIssued.Locked = True
'    AccelerationExpires.Locked = True
'    Exit Sub
'    End If
'Else
'    MsgBox "Please enter the dates that the Acceleration letter was issued and expires", , "Demand Wizard"
If cmdClientSentAccel.Caption = "Cancel" Then
AccelerationIssued = AccelerationIssued.OldValue
AccelerationExpires = AccelerationExpires.OldValue
CmdContinue.Visible = False
cmdClientSentAccel.Caption = "Press to Update Demand Date if Sent by Client"
Else
    AccelerationIssued.Locked = False
    AccelerationExpires.Locked = False
End If

End Sub

Private Sub ClientSentDemandProceduer()
Dim wzdQue As Recordset
Dim rstFCdetails As Recordset
                Set wzdQue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where current = true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
                With wzdQue
                .Edit
                !ClientSentAcceleration = True
                .Update
                .Close
                End With

'Me.Requery
    
                Set rstFCdetails = CurrentDb.OpenRecordset("Select * FROM FCdetails where current = true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
                With rstFCdetails
                .Edit
                !ClientSentAcceleration = "C"
                .Update
                .Close
                End With
    AddStatus FileNumber, Date, "Acceleration Letter noted as sent by client"


    DoCmd.SetWarnings False
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!WizDemand!txtFileNumber,Now(),GetFullName(),'" & "Demand sent by Client" & " '," & 1 & ")"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    
    
'
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
    
    If MsgBox(" You are about to remove dates ? ", vbOKCancel) = vbOK Then
    
        '45 Day Notice sent
        AddStatus FileNumber, Now(), "Removed 45 Days Notice (" & [NOI] & ") by " & GetFullName
    
        DoCmd.SetWarnings False
        strinfo = "Removed 45 Days Notice (" & [NOI] & ") by " & GetFullName
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!WizDemand!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
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
        Forms!WizDemand![NOI] = Null
        
                  
            If Forms!WizDemand!txtClientSentNOI = "C" Then
                AddStatus FileNumber, Now(), "Removed C Of NOI by " & GetFullName
                
                DoCmd.SetWarnings False
                strinfo = "Removed C Of NOI by " & GetFullName
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!WizDemand!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery
            
                Forms!WizDemand!txtClientSentNOI = ""
            End If
            
                Forms!WizDemand.Requery
    
    
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
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!WizDemand!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
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
                Forms!WizDemand![AccelerationIssued] = Null
                
                    If Not IsNull(Forms!WizDemand![AccelerationLetter]) Then
                        AddStatus FileNumber, Now(), "Removed Demand Expires date (" & [AccelerationLetter] & ") by " & GetFullName
                        
                        DoCmd.SetWarnings False
                        strinfo = "Removed Demand Expires date (" & [AccelerationLetter] & ") by " & GetFullName
                        strinfo = Replace(strinfo, "'", "''")
                        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!WizDemand!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                        DoCmd.RunSQL strSQLJournal
                        DoCmd.SetWarnings True
                        Forms!Journal.Requery
                        
                        Forms!WizDemand!AccelerationLetter = Null
                    End If
                    
                    If Forms!WizDemand!txtClientSentAcceleration = "C" Then
                        AddStatus FileNumber, Now(), ":  Removed C from the Demand Field by " & GetFullName
                        
                        DoCmd.SetWarnings False
                        strinfo = ":  Removed C from the Demand Field by " & GetFullName
                        strinfo = Replace(strinfo, "'", "''")
                        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!WizDemand!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
                        DoCmd.RunSQL strSQLJournal
                        DoCmd.SetWarnings True
                        Forms!Journal.Requery
                    
                        Forms!WizDemand!txtClientSentAcceleration = ""
                    End If
                    
                    Forms!WizDemand.Requery
               
           
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
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!WizDemand!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
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
        
            Forms!WizDemand![FairDebtDate] = Null
            Forms!WizDemand.Requery
         
       
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

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!DateSentAttyDemand = Now
 If IsNull(rstqueue!DemandWaiting) Then rstqueue!DemandWaiting = Now
!AttyMilestone1_25 = Null
!AttyMilestone1_25Reject = False
!AttyMilestoneMgr1_25 = Null

.Update
End With

Set rstqueue = Nothing

    DoCmd.SetWarnings False
  
    rstsql = "Insert InTo ValumeDemand (CaseFile, Client, Name, DemandAttyReview, DemandAttyReviewC,SentByClient ) Values ( FileNumber, ClientShortName(forms!wizdemand!ClientID),Getfullname(),now(),1, 'No')"
    DoCmd.RunSQL rstsql
    

    
    
    
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!WizDemand!FileNumber & ",Now,GetFullName(),'" & " Demand Sent To Atty For Review " & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True




Call ReleaseFile(FileNumber)



MsgBox "Demand Sent to Atty Wizard complete", vbInformation


Call ConfirmationVisible(False)
Call FieldsVisible(False)



'Me.RecordSource = ""
DoCmd.Close acForm, Me.Name

End Sub
