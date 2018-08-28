VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Case List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const GroupDelimiter = ";"

Private Sub Active_AfterUpdate()

If (Forms![Case List]!Active = 0) Then

      Dim Ac, rs1, rs2  As Recordset
          If (Me.State = "MD" Or Me.State = "VA" Or Me.State = "DC") And Me.CaseType = "Foreclosure" Then

              Set rs1 = CurrentDb.OpenRecordset("SELECT * FROM LMHearings WHERE FileNumber =" & Forms![Case List].FileNumber, dbOpenSnapshot)

            Set Ac = CurrentDb.OpenRecordset("SELECT * FROM FCdetails WHERE FileNumber =" & Forms![Case List].FileNumber & " And current=True", dbOpenSnapshot)
                If (Ac![Disposition] = "1" Or Ac![Disposition] = "2") Then
                    If IsNull(Ac![AuditFile]) Or IsNull(Ac![AuditRat]) Then
                         MsgBox ("You cannot close this file because either the audit is not filed or ratified. See Post-Sale dept.")
                         Me.Active = True
                         Exit Sub
                    End If
                Else
                  'edited on 4_15_15
                   If IsNull(Ac![Disposition]) Then

                        MsgBox ("You cannot close this file because it has no disposition. ")
                        Me.Active = True
                        Exit Sub

                   ElseIf Not IsNull(Ac!ExceptionsHearing) Then
                        MsgBox ("You cannot close this file because hearing does not have a hearing status on. ")
                        Me.Active = True
                        Exit Sub
                   ElseIf (Not IsNull(Ac!ExceptionsHearing) And (IsNull(Ac!ExceptionsStatus) Or Ac!ExceptionsStatus = "Continue")) Then
                        MsgBox ("You cannot close this file because hearing does not have a hearing status on. ")
                        Me.Active = True
                        Exit Sub
                   ElseIf ((Not IsNull(Ac!StatusHearing) And Ac!StatusHearing >= Date) And (IsNull(Ac!StatusResults) Or Ac!StatusResults = "Continue")) Then
                        MsgBox ("You cannot close this file because Status hearing does not have a status Results on. ")
                        Me.Active = True
                        Exit Sub
                   'ElseIf (Not rs1.EOF) Then
                        'MsgBox ("You cannot close this file because Mediation hearing does not have a status Results on. ")
                        'Me.Active = True
                   'ElseIf Ac!LMDisposition = 4 Then
                        'MsgBox ("You cannot close this file because either no Mediation disposition or as Continued.")
                        'Me.Active = True

          '         Else

         '               Active.Value = 0
                   End If
                End If

                Ac.Close
                Set Ac = Nothing

                rs1.Close
                Set rs1 = Nothing
            End If
  'added DC MEDIATION check on 6/25/15

        If Me.State = "DC" And Me.CaseType = "Foreclosure" Then

            Dim checkmark As Boolean

            checkmark = False
            Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM LMHearings_DC WHERE FileNumber =" & Forms![Case List].FileNumber, dbOpenSnapshot)

            Do Until rs2.EOF

            If Not IsNull(rs2!Hearing) And IsNull(rs2!CondactedTypeID) Then
                checkmark = True
            End If
            rs2.MoveNext

            Loop

            If checkmark = True Then
                MsgBox ("File can not be closed because DC hearing disposition is missing. ")
                Active = True
                Exit Sub
                
           ' Else
          '      Active = False
            End If
            rs2.Close
            Set rs2 = Nothing
        
    '    Exit Sub
        End If






    If DCount("*", "qryNeedToInvoiceBK", "FileNumber=" & FileNumber) > 0 Then
        MsgBox "This file cannot be closed because an invoice needs to be done.  See the Accounting department for assistance.", vbCritical
        Exit Sub
    End If
  '  If IsNull(CloseDate) Then
        Select Case MsgBox("Do you really want to close this file?", vbQuestion + vbYesNoCancel)
            Case vbYes
                OnStatusReport = False
                ' OnStatusReport.Enabled = False
                CloseDate = Now
                ClosedBy = GetLoginName()
                ClosedNumber = ReserveNextClosedNumber()
                Call AddStatus(FileNumber, Now(), "File closed")
            Case vbNo, vbCancel
                Active = True
                Exit Sub
                
        End Select
   ' End If
    
End If

                                    

End Sub

Private Sub Box226_Click()
MsgBox (CurrentProject.Name)
End Sub

Private Sub btn_RestartFeeApproval_Click()
DoCmd.OpenForm "GetRestarApprovalFee", , , , , acDialog, "FC-REF"
End Sub

Private Sub btnEditProject_Click()
DoCmd.OpenForm "EditProjectName"
Forms!EditProjectName.txtFileNumber = FileNumber
Forms!EditProjectName.txtProjectName = PrimaryDefName
If Dirty Then DoCmd.RunCommand acCmdSaveRecord


End Sub

Private Sub CaseTypeID_AfterUpdate()
Call UpdateCaption
End Sub





Private Sub cboSortby_AfterUpdate()
  UpdateDocumentList
  
End Sub

Private Sub cbxDetails_Change()
If Not IsNull(cbxDetails) Then
    If cbxDetails <> 8 Then
 
        Call Details(cbxDetails)
        Else
        MonitorChoose = True
        Call Details(cbxDetails)
    
    End If
End If

End Sub

Private Sub ChOffset_Click()
If ChOffset Then
LabOffset.Visible = True
Call AddStatus(FileNumber, Now(), "Check Offset")
Else
LabOffset.Visible = False
Call AddStatus(FileNumber, Now(), "Unchecked Offset")
End If


End Sub

Private Sub ClientID_AfterUpdate()
If CaseType.Value = "Foreclosure" Then
    If Forms![Case List].[ClientID] = 97 Then
        AIF.Value = False
        AIF.Enabled = False
    ElseIf Forms![Case List].[ClientID] = 446 Then
        AIF.Value = False
        AIF.Enabled = False
    ElseIf Forms![Case List].[ClientID] = 404 Then  'Bogman
        AIF.Value = False
        AIF.Enabled = False
    End If
End If
UpdateAbstractor
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_cmdAdd_Click

  DoCmd.OpenForm "Add Check Request"

Exit_cmdAdd_Click:
  Exit Sub
  
Err_cmdAdd_Click:
  MsgBox Err.Description
  Resume Exit_cmdAdd_Click
  
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

Private Sub CmdAttDoc_Click()
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

Private Sub cmdDeleteDoc_Click()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else

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
End If
  
End Sub

Private Sub cmdImportEmail_Click()
On Error GoTo Err_cmdImportEmail_Click


MsgBox "Emails for the last 14 days will be displayed.", vbOKOnly, "Emails"
DoCmd.OpenForm "frmEmailImport"

Exit_cmdImportEmail_Click:
  Exit Sub
  
Err_cmdImportEmail_Click:
  MsgBox Err.Description
  Resume Exit_cmdImportEmail_Click
  
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

Private Sub cmdJudCost_Click()
DoCmd.OpenForm "GetJudgmentCosts", , , , , acDialog, "FC-JUD"
End Sub

Private Sub cmdNewBillSheet_Click()
On Error GoTo Err_cmdNewBillSheet_Click

If Forms![Case List]!CaseTypeID = 2 Then
    DoCmd.OpenForm "BK_Filling_History"
Else
  DoCmd.OpenReport "Bill Sheet New", acViewPreview
End If
'DoCmd.SetWarnings False

'DoCmd.OpenQuery ("MK_BillSheetFees")
'DoCmd.OpenQuery ("MK_BillSheetCosts")

'DoCmd.SetWarnings True
  
'DoCmd.OpenForm ("Bill Sheet")
Exit_cmdNewBillSheet_Click:
  Exit Sub
  
Err_cmdNewBillSheet_Click:
  MsgBox Err.Description
  Resume Exit_cmdNewBillSheet_Click
End Sub

Private Sub cmdNonStandardFlatFee_Click()
DoCmd.OpenForm "GetNonHourlyFlatFee", , , , , acDialog, "FC-OTH"
End Sub

Private Sub cmdPostSaleCost_Click()
'AdvPostSaleCostPkg
DoCmd.OpenForm "AdvPostSaleCostPkg"

End Sub

Private Sub cmdSkipTrace_Click()
DoCmd.OpenForm "GetSkipTraceCostApproval", , , , , acDialog, "FC-SKP"

End Sub

Private Sub cmdTimeEntry_Click()
DoCmd.OpenForm "GetHours", , , , , acDialog, "FC-OTH"


End Sub

Private Sub cmdView_Click()
Dim i As Long

On Error GoTo Err_cmdView_Click

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
        
        If SCRAID = "Service" And lstDocs.Column(4, i) = 97 Then
        Forms!foreclosuredetails!cmdWizComplete.Enabled = True
        'Open PDF of service packet and make copy
        Dim PDFapp As cacroapp
        Dim CurrentPDF As CAcroAVDoc
        Dim MergedPDF As CAcroPDDoc
        Dim rstDoc As Recordset
        Dim strProofFileName As String: strProofFileName = ""
        Set PDFapp = CreateObject("acroexch.App")
        Set MergedPDF = CreateObject("acroexch.pddoc")
        'Added by Josh for Proof File Name Issue
        strProofFileName = "Proof of Service " & Format$(Now(), "yyyymmdd hhnnss")
        If MergedPDF.Open(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)) Then
        End If
        If MergedPDF.Save(PDSaveFull, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & strProofFileName & ".pdf") = False Then
        MsgBox "Cannot save"
        End If
        MergedPDF.Close
        Set CurrentPDF = CreateObject("acroexch.avdoc")
        If CurrentPDF.Open(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & strProofFileName & ".pdf", "Proof of Service") Then
        If CurrentPDF.BringToFront = True Then
        End If
        End If
        'Commented by JAE 'Document Speed'
        'Set rstDoc = CurrentDb.OpenRecordset("DocIndex", dbOpenDynaset, dbSeeChanges)
        'rstDoc.AddNew
        'rstDoc!FileNumber = FileNumber
        'rstDoc!DocTitleID = 812 'Proof of Service
        'rstDoc!DocGroup = ""
        'rstDoc!StaffID = GetStaffID()
        'rstDoc!DateStamp = Date
        'rstDoc!Filespec = "Proof of Service " & Format$(Now(), "yyyymmdd hhnn") & ".pdf"
        'rstDoc!Notes = "Proof of Service.pdf"
        'rstDoc.Update
        'rstDoc.Close
        DoCmd.SetWarnings False
        Dim strSQLValues As String: strSQLValues = ""
        Dim strSQL As String: strSQL = ""
        strSQLValues = FileNumber & "," & "812" & ",'" & "" & "'," & GetStaffID() & ",'" & Date & "','" & strProofFileName & ".pdf" & "','" & "Proof of Service.pdf" & "'"
        'Debug.Print strSQLValues
        strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
        'Debug.Print strSQL
        DoCmd.RunSQL (strSQL)
        DoCmd.SetWarnings True
        Call UpdateDocumentList
        Else
        
        'If lstDocs.Column(4, i) = 1516 Then
        Select Case lstDocs.Column(4, i)
        Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572
        StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\" & lstDocs.Column(3, i)
        Case Else
        StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)
        End Select
        
        
        'End If
        End If
    End If
Next i

Exit_cmdView_Click:
    Exit Sub

Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click
    
End Sub

 Sub cmdDetails_Click()
'Removed Private clause to call from FC print close event
Call Details(CaseTypeID)
End Sub

Private Sub Details(CaseType As Long)
Dim stDocName As String
Dim stLinkCriteria As String
Dim Details As Recordset

On Error GoTo Err_cmdDetails_Click
If Not FileReadOnly Then
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
End If
Select Case CaseType
    Case 1, 11   ' Foreclosure  or Pending and Monitor
    
        'added on 6/5/15
        'If Me.CaseTypeID = 1 And cbxDetails.Column(0) = 8 Then MonitorChoose = True
        
        stDocName = "ForeclosureDetails"
        stLinkCriteria = "[FileNumber]=" & Me![FileNumber] & " AND Current = True"
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM FCDetails WHERE FileNumber = " & Me!FileNumber & " and current=true", dbOpenSnapshot)
        'New code to check for Current Record
        If Details.RecordCount = 0 Then
        MsgBox "There is not a Foreclosure record marked as Current for this file, please have your manager update via Data Manager", vbCritical
        Exit Sub
        End If
        
     Case 8    '  Monitor
     'added on 6/5/15
        If Me.CaseTypeID = 1 And cbxDetails.Column(0) = 8 Then MonitorChoose = True
     'If IsNull(cbxDetails) Then MonitorChoose = True
    
        stDocName = "ForeclosureDetails"
        stLinkCriteria = "[FileNumber]=" & Me![FileNumber] & " AND Current = True"
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM FCDetails WHERE FileNumber = " & Me!FileNumber & " and current=true", dbOpenSnapshot)
        'New code to check for Current Record
        If Details.RecordCount = 0 Then
        MsgBox "There is not a Foreclosure record marked as Current for this file, please have your manager update via Data Manager", vbCritical
        Exit Sub
        End If
    
    Case 2      ' Bankruptcy
        stDocName = "BankruptcyDetails"
        stLinkCriteria = "[FileNumber]=" & Me![FileNumber] & " AND Current = True"
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM BKDetails WHERE FileNumber = " & Me!FileNumber, dbOpenSnapshot)
    Case 4      ' Collection
        stDocName = "CollectionDetails"
        stLinkCriteria = "[FileNumber]=" & Me![FileNumber]
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM COLDetails WHERE FileNumber = " & Me!FileNumber, dbOpenSnapshot)
    Case 5
        stDocName = "CivilDetails"
        stLinkCriteria = "[FileNumber]=" & Me![FileNumber]
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM CIVDetails WHERE FileNumber = " & Me!FileNumber, dbOpenSnapshot)
    Case 7      ' Eviction
        stDocName = "EvictionDetails"
        stLinkCriteria = "[FileNumber]=" & Me![FileNumber] & " AND Current = True"
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM EVDetails WHERE FileNumber = " & Me!FileNumber, dbOpenSnapshot)
    Case 9      ' REO
        stDocName = "REODetails"
        stLinkCriteria = "[FileNumber]=" & Me![FileNumber]
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM REODetails WHERE FileNumber = " & Me!FileNumber, dbOpenSnapshot)
    Case 10      ' Title Resolution
        stDocName = "TitleResolutionDetails"
        stLinkCriteria = "[FileNumber]=" & Me![FileNumber]
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM TRDetails WHERE FileNumber = " & Me!FileNumber, dbOpenSnapshot)
    
       
    Case Else
        MsgBox "Details not available for this case type", vbExclamation
        Exit Sub
        

End Select

If Details.EOF Then
    If IsNull(ReferralDate) Then
        MsgBox "Referral Date is required", vbCritical
        Exit Sub
    End If
    Call AddDetailRecord(CaseType, Me!FileNumber, ReferralDate)
    If CaseType = 2 Or CaseType = 5 Or CaseType = 7 Or CaseType = 8 Or CaseType = 9 Then   ' BK, EV, REO: make sure FC record exists
        Details.Close
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM FCDetails WHERE FileNumber = " & Me!FileNumber, dbOpenSnapshot)
        If Details.EOF Then Call AddDetailRecord(1, Me!FileNumber, ReferralDate)
    End If
End If
Details.Close

If stDocName <> "Monitor" Then

DoCmd.OpenForm stDocName, , , stLinkCriteria

Else

DoCmd.OpenForm "ForeclosureDetails", , , stLinkCriteria

End If


    If Forms![Case List]!SCRAID = "AccLitig" Then
    
            Select Case stDocName
            Case "ForeclosureDetails"
               ' If CaseType <> 8 Then
                
                        Forms!foreclosuredetails!Page96.Visible = False
                        Forms!foreclosuredetails!Trustees.Visible = False
                        Forms!foreclosuredetails!Page412.Visible = False
                        Forms!foreclosuredetails!pgNOI.Visible = False
                        Forms!foreclosuredetails!Page256.Visible = False
                        Forms!foreclosuredetails!pgMediation.Visible = False
                        Forms!foreclosuredetails![Pre-Sale].Visible = False
                        Forms!foreclosuredetails![Post-Sale].Visible = False
                        Forms!foreclosuredetails!pgRealPropTaxes.Visible = False
                        Forms!foreclosuredetails!pageStatus.Visible = False
'            Case "Monitor"
'                  '  Forms!foreclosureDetails!Page96.Visible = False
'                    '    Forms!foreclosureDetails!Trustees.Visible = False
'                        Forms!foreclosureDetails!Page412.Visible = False
'                        Forms!foreclosureDetails!pgNOI.Visible = False
'                      '  Forms!foreclosureDetails!Page256.Visible = False
'                        Forms!foreclosureDetails!pgMediation.Visible = False
'                     '   Forms!foreclosureDetails![Pre-Sale].Visible = False
'                    '    Forms!foreclosureDetails![Post-Sale].Visible = False
'                        Forms!foreclosureDetails!pgRealPropTaxes.Visible = False
'                        Forms!foreclosureDetails!pageStatus.Visible = False
''                End If
                
            
            Case "BankruptcyDetails"
                    Forms!BankruptcyDetails!Page80.Visible = False
                    Forms!BankruptcyDetails!Page37.Visible = False
                    Forms!BankruptcyDetails!Page45.Visible = False
                    Forms!BankruptcyDetails!Page254.Visible = False
                    Forms!BankruptcyDetails!Page199.Visible = False
                    Forms!BankruptcyDetails!pgPlan.Visible = False
                    Forms!BankruptcyDetails!pgMisc.Visible = False
                    Forms!BankruptcyDetails!Page108.Visible = False
                    Forms!BankruptcyDetails!Exhibits.Visible = False
                    Forms!BankruptcyDetails!pageStatus.Visible = False
                    Forms!BankruptcyDetails!pgBOA.Visible = False
            Case "CollectionDetails"
                    Forms!CollectionDetails!Page64.Visible = False
                    Forms!CollectionDetails!Page88.Visible = False
                    Forms!CollectionDetails!Page34.Visible = False
                    Forms!CollectionDetails!Page45.Visible = False
            Case "CivilDetails"
                    Forms!CivilDetails!Timeline.Visible = False
                    Forms!CivilDetails!Page96.Visible = False
                    Forms!CivilDetails!pageStatus.Visible = False
            Case "EvictionDetails"
                    Forms!EvictionDetails!Page96.Visible = False
                    Forms!EvictionDetails!Page227.Visible = False
                    Forms!EvictionDetails!pageCFK.Visible = False
                    Forms!EvictionDetails!Page300.Visible = False
                    Forms!EvictionDetails!Page309.Visible = False
                    Forms!EvictionDetails!Comments.Visible = False
                    Forms!EvictionDetails!pageStatus.Visible = False
            Case "REODetails"
                    Forms!REODetails!Page45.Visible = False
                    Forms!REODetails!pageStatus.Visible = False
            Case "TitleResolutionDetails"
                    Forms!TitleResolutionDetails!Page96.Visible = False
                    Forms!TitleResolutionDetails!Page195.Visible = False
                    Forms!TitleResolutionDetails!Page256.Visible = False
                    Forms!TitleResolutionDetails![Title Clearance].Visible = False
                    Forms!TitleResolutionDetails!pageStatus.Visible = False
            
            End Select
    End If





Exit_cmdDetails_Click:
    Exit Sub

Err_cmdDetails_Click:
    MsgBox Err.Description
    Resume Exit_cmdDetails_Click
    
End Sub

Private Sub cmdPrint_Click()
' 3/11/10 Disabled this button because FC might be in Fair Debt dispute.
'           We could check, but does anyone really use this button any more?
Exit Sub

On Error GoTo Err_cmdPrint_Click
Call cmdDetails_Click       ' need the details form open

Select Case CaseTypeID
    ' Might be in Fair Debt dispute.
    'Case 1, 8   ' Foreclosure
    '    DoCmd.OpenForm "ForeclosurePrint", , , "[CaseList].[FileNumber]=" & Me![CaseList.FileNumber]
    
    ' Cannot jump to Bankruptcy print because there might be more than 1 current record
    'Case 2  ' Bankruptcy
    '    DoCmd.OpenForm "BankruptcyPrint", , , "[CaseList].[FileNumber]=" & Me![CaseList.FileNumber]
    
    Case 4      ' Collection
        DoCmd.OpenForm "CollectionPrint", , , "[CaseList].[FileNumber]=" & Me![CaseList.FileNumber]
    
    Case 7      ' Eviction
        DoCmd.OpenForm "EvictionPrint", , , "[CaseList].[FileNumber]=" & Me![CaseList.FileNumber]
    
    Case Else
        MsgBox "Printing is not available for this case type", vbExclamation
End Select

Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrint_Click
    
End Sub

Private Sub cmdGoToFile_Click()

On Error GoTo Err_cmdGoToFile_Click

If (CheckOpenJournalEntry) Then
  DoCmd.OpenForm "Select File"
End If

Exit_cmdGoToFile_Click:
    Exit Sub

Err_cmdGoToFile_Click:
    MsgBox Err.Description
    Resume Exit_cmdGoToFile_Click
    
End Sub

Private Sub cmdSearch_Click()

On Error GoTo Err_cmdSearch_Click

If (CheckOpenJournalEntry) Then
  DoCmd.OpenForm "Search"
End If

Exit_cmdSearch_Click:
    Exit Sub

Err_cmdSearch_Click:
    MsgBox Err.Description
    Resume Exit_cmdSearch_Click
    
End Sub

Private Sub cmdViewBillSheet_Click()
On Error GoTo Err_cmdViewBillSheet_Click

  DoCmd.OpenReport "Bill Sheet", acViewPreview

Exit_cmdViewBillSheet_Click:
  Exit Sub
  
Err_cmdViewBillSheet_Click:
  MsgBox Err.Description
  Resume Exit_cmdViewBillSheet_Click
End Sub

Private Sub cmdViewInternetSources_Click()
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm "Internet Sources", , , "[FileNumber]=" & Me![FileNumber]
End Sub





Private Sub CmdWizLiti_Click()
DoCmd.OpenForm "EnterLitigationBillingDetails"
End Sub

Private Sub CmdWizPS_Click()
DoCmd.OpenForm "EnterPSAdvancedCostDetails"
End Sub

Private Sub ComEditInvestor_Click()
'DoCmd.OpenForm "EditInvestor", , , WhereCondition:="FileNumber= " & Forms![Case list]!FileNumber

DoCmd.OpenForm "EditInvestor"
Forms!EditInvestor.FileNumber = Me.FileNumber
Forms!EditInvestor.Investor = Me.Investor
Forms!EditInvestor.InvestorAddr = Me.InvestorAddress
Forms!EditInvestor.AIF = Me.AIF
If Dirty Then DoCmd.RunCommand acCmdSaveRecord

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

Call UpdateDocumentList
lstDocs.Requery

End Sub

Private Sub ComWizESC_Click()
If Not WizESC Then
DoCmd.OpenForm "Audit_ESC", , , "InvoiceID = " & Forms![Case List].sfrmInvoices.Form.InvoiceID
Else

Dim intDisp As Variant
Dim intRPAmtRecClient As Variant
Dim clientShor As String
Dim strSQL As String
Dim DateShow As Date

clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)

DoCmd.SetWarnings False

intDisp = DLookup("[Disposition]", "[FCDetails]", "[FileNumber] = " & Me.FileNumber & " and Current = true")
           Select Case intDisp
           
                                        
                    Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 16, 26, 27, 28, 29, 30, 31, 32, 1, 2

                    DateShow = Date
                    
                    strSQL = "Insert into Accou_EscQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, StaffID, StaffName, CaseType, Disposition, " & _
                    " InvoiceID,DateInvoicedPaid,InvoiceType,InvoiceNumber,BTnumber,InvoiceAmount,PaidAmount,DateShouldShowUp,WhoInvoiced) Values (" & FileNumber & _
                    ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & clientShor & " '," & Forms![Case List]!ClientID & " ,#" & Now() & "# ," & GetStaffID() & _
                    ", '" & GetFullName() & "','" & Forms![Case List]![CaseType] & "', " & intDisp & " , " & Forms![Case List].sfrmInvoices.Form.InvoiceID & ", #" & _
                     Forms![Case List].sfrmInvoices.Form.DatePaid & "#, '" & Forms![Case List].sfrmInvoices.Form.InvoiceType & "', '" & Forms![Case List].sfrmInvoices.Form.InvoiceNumber & "', '" & Forms![Case List].sfrmInvoices.Form.BTnumber & _
                    "', " & Forms![Case List].sfrmInvoices.Form.InvoiceAmount & ", " & Forms![Case List].sfrmInvoices.Form.PaidAmount & ", #" & DateShow & "#,'" & Forms![Case List].sfrmInvoices.Form.CreatedByName & "' )"
                    
                    DoCmd.RunSQL strSQL
                    
                    Case Else
                    
                    DateShow = Date
                    
                    strSQL = "Insert into Accou_EscQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, StaffID, StaffName, CaseType, " & _
                    " InvoiceID,DateInvoicedPaid,InvoiceType,InvoiceNumber,BTnumber,InvoiceAmount,PaidAmount,DateShouldShowUp,WhoInvoiced) Values (" & FileNumber & _
                    ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & clientShor & " '," & Forms![Case List]!ClientID & " ,#" & Now() & "# ," & GetStaffID() & _
                    ", '" & GetFullName() & "','" & Forms![Case List]![CaseType] & "'," & Forms![Case List].sfrmInvoices.Form.InvoiceID & ", #" & Forms![Case List].sfrmInvoices.Form.DatePaid & "#, '" & _
                    Forms![Case List].sfrmInvoices.Form.InvoiceType & "', '" & Forms![Case List].sfrmInvoices.Form.InvoiceNumber & "', '" & Forms![Case List].sfrmInvoices.Form.BTnumber & _
                    "', " & Forms![Case List].sfrmInvoices.Form.InvoiceAmount & ", " & Forms![Case List].sfrmInvoices.Form.PaidAmount & ", #" & DateShow & "#,'" & Forms![Case List].sfrmInvoices.Form.CreatedByName & "' )"
                    
                    DoCmd.RunSQL strSQL
                    
  DoCmd.SetWarnings True
  End Select
  DoCmd.OpenForm "Audit_ESC", , , "InvoiceID = " & Forms![Case List].sfrmInvoices.Form.InvoiceID

End If


End Sub

Private Sub Form_Close()
If WizESC Then WizESC = False
DoCmd.Close acForm, "Journal"
Call ReleaseFile(FileNumber)
If EMailStatus = 1 Then MsgBox "Reminder: EMail is still active", vbInformation
EditDispute = False

End Sub

Private Sub Form_Current()
If privProject Then Me.ProcessProject.Locked = False
If privProject And (Forms![Case List]!Project = "BH" Or Forms![Case List]!Project = "MWC") And Forms![Case List]!ProcessProject = False Then BHproject = True

If CheckStaffConflict(Forms![Case List]!FileNumber) Then cmdDetails.Enabled = False

   



If ChOffset Then LabOffset.Visible = True
Dim WarningLevel As Integer
If PrivSSN Then SpecialDocSSN.Visible = True
If FileReadOnly Or EditDispute Then
  '  Me.AllowEdits = False
    cmdPrint.Enabled = False
    cmdChangeToFC.Enabled = False
    cmdChangeToBK.Enabled = False
    cmdChangeToEV.Enabled = False
    cmdChangetoPND.Enabled = False
    cmdChangeToREO.Enabled = False
    btnEditProject.Enabled = False
    cmdImportEmail.Enabled = False
    cmdViewScan.Enabled = False
    cmdFileLabel.Enabled = False
    cmdInvClient.Enabled = False
    tglEMail.Enabled = False
    Detail.BackColor = ReadOnlyColor
Else
    Me.AllowEdits = True
    cmdPrint.Enabled = (CaseTypeID <> 2)                        ' BK must print from detail screen
    cmdChangeToFC.Enabled = (CaseTypeID = 2 Or CaseTypeID = 11)                    ' BK or PND -> FC
    cmdChangeToBK.Enabled = (CaseTypeID = 1 Or CaseTypeID = 7)  ' FC -> BK  or  EV -> BK
    cmdChangeToEV.Enabled = (CaseTypeID = 2)                    ' BK -> EV
    cmdChangetoPND.Enabled = (CaseTypeID = 2)
    'cmdChangeToREO.Enabled = (CaseTypeID = 1 Or CaseType = 7)   ' FC -> REO  or  EV -> REO
        
    cmdInvClient.Enabled = True
    tglEMail.Enabled = DLookup("iValue", "DB", "Name='EMailPDF'")
    Detail.BackColor = -2147483633
End If

If Not PrivPND Then cmdChangetoPND.Enabled = False
If Not PrivFC Then cmdChangeToFC.Enabled = False
If Not IsNull(ServicerRelease) Then lblServicer.Visible = True
If Not IsNull(ServicerRelease) Then lblServicer.Visible = True
If BillCase = True Then lblBilling.Visible = True

cmdDeleteDoc.Enabled = CBool(DLookup("PrivDeleteDocs", "Staff", "name = '" & GetLoginName() & "'"))   'PrivDeleteDocs
cmdViewFolder.Enabled = CBool(DLookup("PrivDeleteDocs", "Staff", "name = '" & GetLoginName() & "'"))


Call UpdateCaption
If Not CheckStaffConflict(Forms![Case List]!FileNumber) Then

    cmdDetails.SetFocus
    DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
    WarningLevel = Nz(DMax("Warning", "Journal", "FileNumber=" & FileNumber))
    Select Case WarningLevel
        Case 50
            imgWarning.Picture = dbLocation & "dollar.emf"
            imgWarning.Visible = True
        Case 100
            imgWarning.Picture = dbLocation & "papertray.emf"
            imgWarning.Visible = True
        Case 200
            imgWarning.Picture = dbLocation & "house.emf"
            imgWarning.Visible = True
        Case 300
            imgWarning.Picture = dbLocation & "caution.bmp"
            imgWarning.Visible = True
        Case 400
            imgWarning.Picture = dbLocation & "stop.emf"
            imgWarning.Visible = True
        Case Else
            imgWarning.Visible = False
    End Select
End If


If Nz(DCount("Warning", "Journal", "FileNumber=" & FileNumber & " and warning=50")) Then
imgAcctg.Picture = dbLocation & "dollar.emf"
imgAcctg.Visible = True
End If

cmdViewScan.Enabled = (Dir$(ClosedScanLocation & FileNumber & "*.pdf") <> "")
Call UpdateDocumentList
If IsNull(JurisdictionID) Then MsgBox "Jurisdiction is missing!", vbExclamation

If DCount("*", "CIVDetails", "FCFileNumber=" & FileNumber) > 0 Then
    Detail.BackColor = vbYellow
    MsgBox "CAUTION! Litigation in progress, see an attorney!", vbExclamation
End If


'Dim FC As Recordset
'Dim ss As Integer
'If CaseType = 1 Then
'If IsNull(ss = DLookup("Disposition", "FCdetails", "FileNumber= Forms![Case List].[FileNumber] And Current = True")) Then
'Set FC = CurrentDb.OpenRecordset("Select ID, Disposition FROM FCDisposition WHERE ID = " & ss, dbOpenDynaset, dbSeeChanges)
'Des = "Disposition is: " & FC!Disposition
'End If

lstBillingReasons.RowSource = "SELECT BillingReasonsFCarchive.ID, BillingReasonsFC.Reason, BillingReasonsFCarchive.date FROM BillingReasonsFCarchive INNER JOIN BillingReasonsFC ON BillingReasonsFCarchive.billingreasonID = BillingReasonsFC.ID WHERE (((BillingReasonsFCarchive.Invoiced) Is Null) AND((BillingReasonsFCarchive.FileNumber)=" & FileNumber & "));"
lstBillingReasons.Requery

Dim ss As Integer
Dim FC As Recordset

Select Case (Forms![Case List].[CaseType])

 Case "Foreclosure"
    If IsNull(ss = DLookup("Disposition", "FCdetails", "FileNumber= Forms![Case List].[FileNumber] And Current = True")) Then
     Des.Value = ""
     cmdChangeToBK.Enabled = False
     cmdChangeToFC.Enabled = False
     cmdChangeToEV.Enabled = False
     cmdChangeToREO.Enabled = False
     cmdChangetoPND.Enabled = False
     
    Else
    ss = DLookup("Disposition", "FCdetails", "FileNumber= Forms![Case List].[FileNumber] And Current = True")
     If IsNull([ss]) Then
      Des.Value = " "
      Else
     Set FC = CurrentDb.OpenRecordset("Select ID, Disposition FROM FCDisposition WHERE ID = " & ss, dbOpenDynaset, dbSeeChanges)
      Des = "Disposition is: " & FC!Disposition
      FC.Close
    End If
    End If
    Case "Bankruptcy"
    Des.Value = " "
        End Select
        
    If CaseType.Value = "Foreclosure" Then
    If Forms![Case List].[ClientID] = 97 Then
        AIF.Value = False
        AIF.Enabled = False
    ElseIf Forms![Case List].[ClientID] = 531 Then
        AIF.Value = False
        AIF.Enabled = False
    ElseIf Forms![Case List].[ClientID] = 404 Then  'Bogman
        AIF.Value = False
        AIF.Enabled = False
    Else
    If Forms![Case List].[ClientID] = 446 Then
    AIF.Value = False
    AIF.Enabled = False
'    Else
'    AIF.Value = True
'    AIF.Enabled = True
    End If
    End If
    End If
    
    
    '11/3/14
                        
If Me.CaseType = "Foreclosure" Then
    Me.JurisdictionID.RowSource = "SELECT DISTINCTROW JurisdictionID, Jurisdiction & ' , ' & State AS Expr1 FROM JurisdictionList where JurisdictionID not in(61, 189) ORDER BY Jurisdiction & ' , ' & State;"
Else
    Me.JurisdictionID.RowSource = "SELECT DISTINCTROW JurisdictionID, Jurisdiction & ' , ' & State AS Expr1 FROM JurisdictionList ORDER BY Jurisdiction & ' , ' & State;"
End If



End Sub

Private Sub cmdClose_Click()
If WizESC Then WizESC = False
On Error GoTo Err_cmdClose_Click

If (CheckOpenJournalEntry) And Not CurrentProject.AllForms("foreclosuredetails").IsLoaded And Not CurrentProject.AllForms("bankruptcydetails").IsLoaded And Not CurrentProject.AllForms("evictiondetails").IsLoaded Then
  DoCmd.Close
  Else
  'Removed by JE 07-14-2014
  'Dim rstLocks As Recordset
  'Set rstLocks = CurrentDb.OpenRecordset("select * from locksarchive", dbOpenDynaset, dbSeeChanges)
  'With rstLocks
  '.AddNew
  '!FileNumber = FileNumber
  '!StaffID = GetStaffID()
  '!Type = "X"
  '.Update
  'End With
  'Added by JE 07-14-2014
  Dim str_SQL As String
  str_SQL = "INSERT INTO LocksArchive(FileNumber,StaffID,[TimeStamp],[Type]) VALUES (" & FileNumber & "," & GetStaffID() & ",'" & Now() & "','X')"
  'Debug.Print str_SQL
  RunSQL (str_SQL)
  
  MsgBox "Please close the Details screen first", vbCritical
  Exit Sub
End If

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub



Private Sub UpdateCaption()
Me.Caption = FileNumber & " " & PrimaryDefName & " " & CaseType 'CaseTypeID.Column(1)
End Sub


Private Sub UpdateDocumentList()
Dim GroupName As String

On Error GoTo UpdateDocumentListErr

Select Case optDocType
    Case 1
        GroupName = ""
        lstDocs.ColumnCount = 5
        lstDocs.ColumnWidths = "0 in; 0.4 in; 0.75 in; 2 in; 0 in"
        
    Case 2
        GroupName = "B"
        lstDocs.ColumnCount = 6
        lstDocs.ColumnWidths = "0 in; 0.4 in; 0.75 in; 3 in; 0 in ;0.3 in "
End Select

lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], [doctitleid] AS DocType , Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='" & GroupName & "' AND Filespec IS NOT NULL and DeleteDate is null ORDER BY " & Me.cboSortby
lstDocs.Requery

Exit Sub

UpdateDocumentListErr:
    MsgBox Err.Description, vbCritical
    Exit Sub
    
End Sub




Private Sub Form_Open(Cancel As Integer)
BHproject = False
Dim s As Recordset
Dim t As Recordset
Dim ss As Integer

'MsgBox (FileNumber)
Forms![Case List]!lstDocs.RowSource = "Select SELECT DocIndex.DocID, Staff.Initials, Format(Datestamp,'mm/dd/yyyy') AS [Date Entered], DocIndex.Filespec AS [File Name], DocIndex.doctitleid AS DocType, DocIndex.Hold, DocIndex.FileNumber FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID = Staff.ID WHERE (((DocIndex.Filespec) Is Not Null) AND ((DocIndex.DocGroup)='B') AND ((DocIndex.DeleteDate) Is Null) And (( DocIndex.FileNumber) = " & FileNumber & ")));"

Active.Locked = Not PrivCloseFiles
OnStatusReport.Locked = Not PrivCloseFiles
'Me.Page91.Enabled = PrivBillingEdits
cmdDeleteDoc.Enabled = CBool(DLookup("PrivDeleteDocs", "Staff", "name = '" & GetLoginName() & "'"))   'PrivDeleteDocs
cmdViewFolder.Enabled = CBool(DLookup("PrivDeleteDocs", "Staff", "name = '" & GetLoginName() & "'"))

cmdDeleteDoc.Enabled = PrivDeleteDocs
tglEMail.Enabled = DLookup("iValue", "DB", "Name='EMailPDF'")

If EMailStatus = 1 Then
    tglEMail = True
    MsgBox "Reminder: EMail is still active", vbInformation
End If

'8/25/14
If PrivitLimitedView = True Or CheckStaffConflict(FileNumber) = True Then
Me.Page120.Visible = False
Me.Page97.Visible = False
Me.pageAccounting.Visible = False
Me.Page91.Visible = False
Me.pageCheckRequest.Visible = False
Me.pgDocRequest.Visible = False
Me.pgConflicts.Visible = False
End If

If FileReadOnly Or EditDispute Then

    Dim ctl As Control
    Dim lngI As Long
    Dim bSkip As Boolean

    For Each ctl In Form.Controls
    Select Case ctl.ControlType
    Case acTextBox, acListBox, acOptionGroup, acCheckBox, acSubform, acOptionButton
         bSkip = False
            If ctl.Name = "lstDocs" Then bSkip = True
            If Not bSkip Then ctl.Locked = True
            
            
    Case acCommandButton
        bSkip = False
            If ctl.Name = "cbxDetails" Then bSkip = True
            If ctl.Name = "cmdDetails" Then bSkip = True
            If ctl.Name = "cmdGoToFile" Then bSkip = True
            If ctl.Name = "cmdSearch" Then bSkip = True
            If ctl.Name = "cmdClose" Then bSkip = True
            If ctl.Name = "cmdSelectFile" Then bSkip = True
            If Not bSkip Then ctl.Enabled = False
         
    Case acComboBox
        bSkip = False
        If ctl.Name = "cbxDetails" Then bSkip = True
        If Not bSkip Then ctl.Locked = True
    
    
    End Select
    Next
End If

'11/3/14
                        
If Me.CaseType = "Foreclosure" Then
    Me.JurisdictionID.RowSource = "SELECT DISTINCTROW JurisdictionID, Jurisdiction & ' , ' & State AS Expr1 FROM JurisdictionList where JurisdictionID not in(61, 189) ORDER BY Jurisdiction & ' , ' & State;"
Else
    Me.JurisdictionID.RowSource = "SELECT DISTINCTROW JurisdictionID, Jurisdiction & ' , ' & State AS Expr1 FROM JurisdictionList ORDER BY Jurisdiction & ' , ' & State;"
End If


End Sub

Private Sub JurisdictionID_AfterUpdate()
  UpdateAbstractor
  
    If ([CaseTypeID] = 1 Or [CaseTypeID] = 7) Then  ' foreclosure or eviction
  
      If (Not IsNull([JurisdictionID])) Then
        If ([JurisdictionID] > 0) Then
    
          Dim JurisdictionState As String
      
          JurisdictionState = DLookup("[State]", "[JurisdictionList]", "[JurisdictionID] = " & [JurisdictionID])
          DoCmd.SetWarnings False
          DoCmd.RunSQL "update FCdetails set State = '" & JurisdictionState & "' where FileNumber = " & Me.FileNumber
          DoCmd.SetWarnings True
        End If
      End If
    End If
  
  
End Sub

Private Sub lstDocs_DblClick(Cancel As Integer)
Call cmdView_Click
End Sub

Public Sub optDocType_AfterUpdate()
Call UpdateDocumentList
End Sub



Private Sub PrimaryDefName_AfterUpdate()
Call UpdateCaption
End Sub

Private Sub cmdInvClient_Click()
Dim C As Recordset

On Error GoTo Err_cmdInvClient_Click
Set C = CurrentDb.OpenRecordset("SELECT * FROM ClientList WHERE ClientID = " & ClientID, dbOpenSnapshot)

If Not C.EOF Then
    If (C!ClientID = 567 And Forms![Case List]!State = "MD") Then
        Investor = "Nationstar Mortgage LLC d/b/a Champion Mortgage Company of Texas"
        InvestorAddr = C("StreetAddress") & IIf(IsNull(C("StreetAddr2")), "", vbNewLine & C("StreetAddr2")) & _
        vbNewLine & C("City") & ", " & C("State") & " " & C("ZipCode")
    ElseIf (C!ClientID = 567 And Forms![Case List]!State = "VA") Then
        Investor = "Nationstar Mortgage LLC, doing business in the Commonwealth of Virginia as Virginia Nationstar LLC d/b/a Champion Mortgage Company"
        InvestorAddr = C("StreetAddress") & IIf(IsNull(C("StreetAddr2")), "", vbNewLine & C("StreetAddr2")) & _
        vbNewLine & C("City") & ", " & C("State") & " " & C("ZipCode")
    Else
        Investor = C("ClientNameAsInvestor")
        InvestorAddr = C("StreetAddress") & IIf(IsNull(C("StreetAddr2")), "", vbNewLine & C("StreetAddr2")) & _
        vbNewLine & C("City") & ", " & C("State") & " " & C("ZipCode")
    End If
End If


C.Close

Exit_cmdInvClient_Click:
    Exit Sub

Err_cmdInvClient_Click:
    MsgBox Err.Description
    Resume Exit_cmdInvClient_Click
    
End Sub

Private Sub ProcessProject_AfterUpdate()
If Forms![Case List]!ProcessProject = False Then
Call AddStatus(FileNumber, Date, "Un checked Project Process")
Else
Call AddStatus(FileNumber, Date, "Checked Project Process")
End If

End Sub

Private Sub ReferralDate_AfterUpdate()
Call AddStatus(FileNumber, ReferralDate, "Referral Date")
End Sub

Private Sub ReferralDocsReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ReferralDocsReceived = Date
End If
End Sub

Private Sub RestartReceived_AfterUpdate()
If Not IsNull(RestartReceived) Then
    If Active And OnStatusReport Then
        AddStatus FileNumber, RestartReceived, "Restart Received"
    Else
        RestartReceived = Null
        MsgBox "File must be Active and On Status Report in order to Restart", vbCritical
    End If
End If
End Sub

Private Sub RestartReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    RestartReceived = Date
    Call RestartReceived_AfterUpdate
End If

End Sub

Private Sub ServicerRelease_AfterUpdate()
Dim Status As String, rstJnl As Recordset, rstBillReasons As Recordset

If Not IsNull(ServicerRelease) Then servicereffective = InputBox("Please enter the effective date")
BillCase = True
BillCaseUpdateDate = Date
BillCaseUpdateUser = GetStaffID
[BillCaseUpdateReasonID] = 2
lblBilling.Visible = True
lstBillingReasons.Requery

Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstBillReasons
.AddNew
!FileNumber = FileNumber
!billingreasonid = 2
!UserID = GetStaffID
!Date = Date
.Update
End With


Status = "Servicer Release notified on " & ServicerRelease & "; effective " & servicereffective
AddStatus FileNumber, Now(), Status

'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = Status
'!Color = 2
'.Update
'End With
'Set rstJnl = Nothing

DoCmd.SetWarnings False
strinfo = Status
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',2 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

End Sub

Private Sub SpecialDocSSN_Click()

Shell "Explorer """ & DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\""", vbNormalFocus

End Sub



Private Sub tglEMail_Click()
If tglEMail Then        ' click on
    If Not EMailInit() Then tglEMail = False
Else                    ' click off
    If Not EMailEnd() Then tglEMail = True
End If
End Sub

Private Sub cmdChangeToFC_Click()

On Error GoTo Err_cmdChangeToFC_Click

If CaseTypeID = 2 Then
    Dim rsMyRS As Recordset
    Dim sqlString As String
    
    sqlString = "SElect Names.*"
    sqlString = sqlString + " From [Names]"
    sqlString = sqlString + " Where Names.Filenumber =" & Forms![Case List]!FileNumber & ";"

    Set rsMyRS = CurrentDb.OpenRecordset(sqlString, dbOpenDynaset, dbSeeChanges)

    If Not (rsMyRS.BOF And rsMyRS.EOF) Then rsMyRS.MoveFirst
    Do While Not rsMyRS.EOF
        If rsMyRS!NoticeType = 10 Then
            rsMyRS.Edit
            rsMyRS!NoticeType = 16
            rsMyRS.Update
            rsMyRS.MoveNext
        Else
            rsMyRS.MoveNext
        End If
    Loop
End If

If CaseTypeID = 2 Or CaseTypeID = 11 Then  ' BK or PND
Dim rstNames As Recordset, rstBKdetails As Recordset, rstTrustees As Recordset, rstBKAtty As Recordset
Dim AttyID As String
Dim TrusteeID As String, TrusteePhone As String, TrusteeFirst As String, TrusteeLast As String, TrusteeAddress As String, TrusteeAddress2 As String, TrusteeCity As String, TrusteeState As String, TrusteeZip As String

Set rstBKdetails = CurrentDb.OpenRecordset("select * from BKdetails where Filenumber=" & FileNumber & "And Current=True", dbOpenDynaset, dbSeeChanges)

On Error Resume Next

With rstBKdetails
If .RecordCount > 0 Then
AttyID = !AttorneyID
TrusteeID = !Trustee
End If
End With
rstBKdetails.Close
On Error GoTo Err_cmdChangeToFC_Click


If TrusteeID <> "" Then
Set rstTrustees = CurrentDb.OpenRecordset("select * from BKTrustees where ID=" & TrusteeID, dbOpenDynaset, dbSeeChanges)
On Error Resume Next
With rstTrustees
TrusteeFirst = !First
TrusteeLast = !Last
TrusteeAddress = !Address
If Not IsNull(!TrusteeAddress2) Then
TrusteeAddress2 = !Address2
End If
TrusteeCity = !City
TrusteeState = !State
TrusteeZip = !Zip
End With
rstTrustees.Close
End If
On Error GoTo Err_cmdChangeToFC_Click

If AttyID <> "" Then
Set rstNames = CurrentDb.OpenRecordset("select * from Names", dbOpenDynaset, dbSeeChanges)
Set rstBKAtty = CurrentDb.OpenRecordset("select * from BKAttorneys where attyid=" & AttyID, dbOpenDynaset, dbSeeChanges)
On Error Resume Next
With rstNames
.AddNew
!FileNumber = FileNumber
!First = rstBKAtty!FirstName
!Last = rstBKAtty!LastName
!Company = rstBKAtty!AttorneyFirm
!Address = rstBKAtty!Address
!City = rstBKAtty!City
!State = rstBKAtty!State
!Zip = rstBKAtty!Zip
!NoticeType = 10
.Update
End With
rstNames.Close
End If

If TrusteeID <> "" Then
Set rstNames = CurrentDb.OpenRecordset("select * from Names", dbOpenDynaset, dbSeeChanges)
On Error Resume Next
With rstNames
.AddNew
!FileNumber = FileNumber
!Company = "BK Trustee"
!First = TrusteeFirst
!Last = TrusteeLast
!Address = TrusteeAddress
!Address2 = TrusteeAddress2
!City = TrusteeCity
!State = TrusteeState
!Zip = TrusteeZip
!NoticeType = 14
.Update
End With
rstNames.Close
End If

CaseTypeID = 1
Call AddStatus(FileNumber, Now(), "File type changed to FC")
ReferralDate = Null


On Error GoTo Err_cmdChangeToFC_Click
End If


cmdDetails.SetFocus
Call Form_Current

Exit_cmdChangeToFC_Click:
    Exit Sub

Err_cmdChangeToFC_Click:
    MsgBox Err.Description
    Resume Exit_cmdChangeToFC_Click
    
End Sub

Private Sub cmdChangeToBK_Click()

On Error GoTo Err_cmdChangeToBK_Click
CaseTypeID = 2
Call AddStatus(FileNumber, Now(), "File type changed to BK")
cmdDetails.SetFocus
Call Form_Current

Exit_cmdChangeToBK_Click:
    Exit Sub

Err_cmdChangeToBK_Click:
    MsgBox Err.Description
    Resume Exit_cmdChangeToBK_Click
    
End Sub

Private Sub cmdChangeToEV_Click()

On Error GoTo Err_cmdChangeToEV_Click

If CaseTypeID = 2 Then  ' BK
    If DCount("*", "rqryNeedToInvoiceBankruptcy", "File=" & FileNumber) > 0 Then
        MsgBox "The case type cannot be changed because an invoice needs to be done.  See the Accounting department for assistance.", vbCritical
        Exit Sub
    End If
End If

CaseTypeID = 7
Call AddStatus(FileNumber, Now(), "File type changed to EV")
cmdDetails.SetFocus
Call Form_Current

Exit_cmdChangeToEV_Click:
    Exit Sub

Err_cmdChangeToEV_Click:
    MsgBox Err.Description
    Resume Exit_cmdChangeToEV_Click
    
End Sub

Private Sub cmdViewScan_Click()
Dim Filespec As String

On Error GoTo Err_cmdViewScan_Click

Filespec = Dir$(ClosedScanLocation & FileNumber & "*.pdf")
Filespec = Dir$()       ' see if there's another file

If Filespec = "" Then
    Call StartDoc(ClosedScanLocation & FileNumber & ".pdf")    ' only 1 file
Else
    DoCmd.OpenForm "frmScans", , , , , , FileNumber
End If

Exit_cmdViewScan_Click:
    Exit Sub

Err_cmdViewScan_Click:
    MsgBox Err.Description
    Resume Exit_cmdViewScan_Click
    
End Sub

Private Sub cmdChangeToREO_Click()

On Error GoTo Err_cmdChangeToREO_Click

If CaseTypeID = 2 Then  ' BK
    If DCount("*", "rqryNeedToInvoiceBankruptcy", "File=" & FileNumber) > 0 Then
        MsgBox "The case type cannot be changed because an invoice needs to be done.  See the Accounting department for assistance.", vbCritical
        Exit Sub
    End If
End If

CaseTypeID = 9
Call AddStatus(FileNumber, Now(), "File type changed to REO")
cmdDetails.SetFocus
Call Form_Current

Exit_cmdChangeToREO_Click:
    Exit Sub

Err_cmdChangeToREO_Click:
    MsgBox Err.Description
    Resume Exit_cmdChangeToREO_Click
    
End Sub

Private Sub cmdChangeToPND_Click()
On Error GoTo Err_cmdChangeToREO_Click

Dim rsMyRS As Recordset
Dim sqlString As String

sqlString = "SELECT Names.*"
sqlString = sqlString + " FROM [Names]"
sqlString = sqlString + " WHERE Names.FileNumber =" & Forms![Case List]!FileNumber & ";"

Set rsMyRS = CurrentDb.OpenRecordset(sqlString, dbOpenDynaset, dbSeeChanges)


If CaseTypeID <> 2 Then  ' BK
    
        MsgBox "This can only be done for cases in active bankruptcy.", vbCritical
        Exit Sub
    End If

MsgBox "Case has been changed to Pending"
CaseTypeID = 11
Call AddStatus(FileNumber, Now(), "File type changed to Pending")

If Not (rsMyRS.BOF And rsMyRS.EOF) Then rsMyRS.MoveFirst
Do While Not rsMyRS.EOF
    If rsMyRS!NoticeType = 10 Then
        rsMyRS.Edit
        rsMyRS!NoticeType = 16
        rsMyRS.Update
        rsMyRS.MoveNext
    Else
        rsMyRS.MoveNext
    End If
Loop

cmdDetails.SetFocus
Call Form_Current
rsMyRS.Close
Set rsMyRS = Nothing

Exit_cmdChangeToREO_Click:
    Exit Sub

Err_cmdChangeToREO_Click:
    MsgBox Err.Description
    Resume Exit_cmdChangeToREO_Click
    
End Sub


Private Sub cmdInvestorMERS_Click()

On Error GoTo Err_cmdInvestorMERS_Click

Investor = "Mortgage Electronic Registration Systems, Inc. As Nominee for the Beneficiary"
If ClientID = 97 Then    ' EMC
    InvestorAddr = "2780 Lake Vista Drive" & vbNewLine & "Lewisville, TX 75067-3884"
Else
    InvestorAddr = "P.O. Box 2026" & vbNewLine & "Flint, MI 48501-2026"
End If

Exit_cmdInvestorMERS_Click:
    Exit Sub

Err_cmdInvestorMERS_Click:
    MsgBox Err.Description
    Resume Exit_cmdInvestorMERS_Click
    
End Sub

Private Sub cmdAddDoc_Click()
DoCmd.SetWarnings False

Dim rstdocs As Recordset
Dim strSQL As String
'creat folder and subfolder for ssn

Dim fso
Dim ss As String
ss = "SSN"
'8/27/14 SA
FileNO = FileNumber

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
Dim GroupCode As String, DocType As String, rstDoc As Recordset, DocDateInput As String, DocDate As Date, UpdateFlag As Boolean, UpdateCase As String
Dim strSCRAQueueFiles As Recordset
Dim rstqueueCount As Integer

On Error GoTo Err_cmdAddDoc_Click

Me.Refresh

    If CurrentProject.AllForms("scra search info").IsLoaded = True Then
        MsgBox "Please close the SCRA Borrower info window first", vbCritical
        Exit Sub
    End If

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
''''''''''''''''''

If IsNumeric(SCRAID) Then

Select Case SCRAID
    Case 10, 11, 12
    If Forms!quescra9!cmdUploadSCRA.Enabled = True Then 'sarabs
        If Forms![Case List]!Indicatorbox = 1 Then
        DoCmd.OpenForm "Select Document Type", , , , , acDialog
        End If
    'Military Affidavit
   ' SelectedDocType = 1398
    Else
    'PACER search
        selecteddoctype = 1010
    End If
    Case 1, 2, 3, 31, 32, 33, 4, 5, 6, 7, 8, 9, 14, 15, 16, 111, 61, 65, 35, 36, 37, 38, 125, 126, 127, 91, 34, 128, 93, 94, 95, 96, 97, 98, 99, 34, 45, 39, 38
    If Forms![Case List]!Indicatorbox = 1 Then
    DoCmd.OpenForm "Select Document Type", , , , , acDialog
    End If
    
    
    'Military Affidavit
    'SelectedDocType = 1398
    Case 13 'VA Appraisal
    selecteddoctype = 1449
End Select

UpdateCase = MsgBox("Will this document be uploaded to the Client?", vbYesNoCancel)
        If UpdateCase = vbYes Then
        UpdateFlag = True
        ElseIf UpdateCase = vbCancel Then
        Exit Sub
        End If

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

If selecteddoctype = 1516 Then
FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & ss & "\" & newfilename
Else
FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & newfilename
End If

DoCmd.SetWarnings False
Dim strSQLValues As String: strSQLValues = ""
strSQL = ""
strSQLValues = FileNumber & "," & selecteddoctype & ",'" & GroupCode & "'," & GetStaffID() & ",'" & DocDate & "','" & Replace(newfilename, "'", "''") & "','" & Replace(newfilename, "'", "''") & "'"
'Debug.Print strSQLValues
strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
'Debug.Print strSQL
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings True

If MsgBox("New document " & newfilename & " accepted.  OK to delete " & Filespec & "?", vbQuestion + vbYesNo) = vbYes Then Kill Filespec

Dim rstqueue As Recordset, rstJnl As Recordset, Status As String, jnltxt As String, rstEV As Recordset
Dim rstLocalQueue As Recordset
Dim rstLocalQueue2 As Recordset


Select Case SCRAID

Case 1
Status = "SCRA Check- First Legal, completed"
Case 111
Status = "SCRA Check- First Legal JPM/Wells, completed"
Case 2
Status = "SCRA Check- Docketing, completed"


Case 3
If Forms![Case List]!ClientID = 543 Then
Status = "SCRA Check- Sale Date ResCap 7 day, completed"
Else
Status = "SCRA Check- Sale Date , completed"
End If



Case 31, 45
Status = "SCRA Check- Sale Date 7 day, completed"
Case 32
Status = "SCRA Check- Sale Date JPM 3 day, completed"
Case 33
Status = "SCRA Check- 1 Day prior to Sale, completed"
Case 4
Status = "SCRA Check- Post Sale, Completed"
Case 42
Status = "SCRA Check- Post Sale, Completed"
Case 43, 44
Status = "SCRA Check- Post Sale, Completed"
Case 5
Status = "SCRA Check- Post Sale, Completed"
Case 6
Status = "SCRA Check- Ratification, Completed"
Case 7
Status = "SCRA Check- Deeds Sent, Completed"
Case 8
Status = "SCRA Check- Sale , Completed"
Case 9
Status = "SCRA Check- DIL Disposition, Completed"
Case 61
Status = "SCRA Check- Post Sale, Completed"
Case 34
Status = "SCRA Check- 2 day Sale, Completed"
Case 65
Status = "SCRA Check- Ratification, Completed"

Case 35
Status = "SCRA Check- Sale Date BOA 7 day, completed"

Case 36
Status = "SCRA Check- Sale Date PHH 1 day, completed"

Case 37
Status = "SCRA Check- Sale Date , completed"

Case 38
Status = "SCRA Check- Day of Sale, completed"

Case 39
Status = "SCRA Check- 1 Day Before Sale, completed"

Case 125
Status = "SCRA Check- New Referral, Completed"
Case 126
Status = "SCRA Check- New Referral, Completed"


Case 127
Status = "SCRA Check- Borrower Served, Completed"

Case 128
Status = "SCRA Check- Title Received, Completed"

Case 91
Status = "SCRA Check - Sale 40 days, completed"

Case 93
Status = "SCRA Check - Sale 14 Days, completed"

Case 94
Status = "SCRA Check - Sale 10 Days, completed"

Case 95
Status = "SCRA Check - Sale 22 Days, completed"


Case 96
Status = "SCRA Check - Sent Complaint To Court, completed"

Case 97
Status = "SCRA Check - Motion for Judgment of Foreclosure, completed"

Case 98
Status = "SCRA Check - Title Received, Completed"
Case 99
Status = "SCRA Check - 1 day prior to sale, Completed"


Case 13
Status = "VA Appraisal Ordered"
Case 14
Status = "SCRA Check- BK POC"
Case 15
Status = "SCRA Check- BK 362"
Case 16
Status = "SCRA Check- BK NOD"
Case 10
Set rstEV = CurrentDb.OpenRecordset("select * from evdetails where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
rstEV.Edit
If selecteddoctype = 1010 Then
Status = "PACER Check- Eviction Referral, Completed"
If MsgBox("Did RA perform the PACER check?", vbYesNo) = vbYes Then
Dim rstNames As Recordset, ctr As Integer
Set rstNames = CurrentDb.OpenRecordset("select * from Names where filenumber=" & FileNumber & " and noteholder=yes", dbOpenDynaset, dbSeeChanges)
With rstNames
.MoveLast
ctr = .RecordCount
.MoveFirst
.Close
End With
AddInvoiceItem FileNumber, "EV-Search", "PACER search", (25 * ctr), 0, True, True, True, False
rstEV!ReferredPACER = "R"
Else
rstEV!ReferredPACER = "C"
End If
Else
Status = "SCRA Check- Eviction Referral, Completed"
If MsgBox("Did RA perform the SCRA check?", vbYesNo) = vbYes Then

Set rstNames = CurrentDb.OpenRecordset("select * from Names where filenumber=" & FileNumber & " and noteholder=yes", dbOpenDynaset, dbSeeChanges)
With rstNames
.MoveLast
ctr = .RecordCount
.MoveFirst
.Close
End With

AddInvoiceItem FileNumber, "EV-Search", "SCRA search", (50 * ctr), 0, True, True, True, False
rstEV!ReferredSCRA = "R"
Else
rstEV!ReferredSCRA = "C"
End If
rstEV.Update
End If
Case 11
Set rstEV = CurrentDb.OpenRecordset("select * from evdetails where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
rstEV.Edit
If selecteddoctype = 1010 Then
Status = "PACER Check- Eviction Hearing, Completed"
If MsgBox("Did RA perform the PACER check?", vbYesNo) = vbYes Then

Set rstNames = CurrentDb.OpenRecordset("select * from Names where filenumber=" & FileNumber & " and noteholder=yes", dbOpenDynaset, dbSeeChanges)
With rstNames
.MoveLast
ctr = .RecordCount
.MoveFirst
.Close
End With
AddInvoiceItem FileNumber, "EV-Search", "PACER search", (25 * ctr), 0, True, True, True, False
rstEV!ReferredPACER = "R"
Else
rstEV!ReferredPACER = "C"
End If
Else
Status = "SCRA Check- Eviction Hearing, Completed"
If MsgBox("Did RA perform the SCRA check?", vbYesNo) = vbYes Then

Set rstNames = CurrentDb.OpenRecordset("select * from Names where filenumber=" & FileNumber & " and noteholder=yes", dbOpenDynaset, dbSeeChanges)
With rstNames
.MoveLast
ctr = .RecordCount
.MoveFirst
.Close
End With
AddInvoiceItem FileNumber, "EV-Search", "SCRA search", (50 * ctr), 0, True, True, True, False
rstEV!ReferredSCRA = "R"
Else
rstEV!ReferredSCRA = "C"
End If
End If
rstEV.Update
Case 12
Set rstEV = CurrentDb.OpenRecordset("select * from evdetails where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
rstEV.Edit
If selecteddoctype = 1010 Then
Status = "PACER Check- Eviction Lockout, Completed"
If MsgBox("Did RA perform the PACER check?", vbYesNo) = vbYes Then

Set rstNames = CurrentDb.OpenRecordset("select * from Names where filenumber=" & FileNumber & " and noteholder=yes", dbOpenDynaset, dbSeeChanges)
With rstNames
.MoveLast
ctr = .RecordCount
.MoveFirst
.Close
End With
AddInvoiceItem FileNumber, "EV-Search", "PACER search", (25 * ctr), 0, True, True, True, False
rstEV!HearingPACER = "R"
Else
rstEV!HearingPACER = "C"
End If
Else
Status = "SCRA Check- Eviction Lockout, Completed"
If MsgBox("Did RA perform the SCRA check?", vbYesNo) = vbYes Then

Set rstNames = CurrentDb.OpenRecordset("select * from Names where filenumber=" & FileNumber & " and noteholder=yes", dbOpenDynaset, dbSeeChanges)
With rstNames
.MoveLast
ctr = .RecordCount
.MoveFirst
.Close
End With
AddInvoiceItem FileNumber, "EV-Search", "SCRA search", (50 * ctr), 0, True, True, True, False
rstEV!LockoutSCRA = "R"
Else
rstEV!LockoutSCRA = "C"
End If
End If
rstEV.Update
End Select

If UpdateFlag = True Then
    If DLookup("emailsearches", "clientlist", "clientid=" & ClientID) = False Then
    Status = Status & " and uploaded to client"
    Else
    Status = Status & " and emailed to client"
    End If
End If

AddStatus FileNumber, Date, Status

jnltxt = Status


DoCmd.SetWarnings False
strinfo = jnltxt
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True



If UpdateFlag = True Then
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)

If CaseTypeID = 1 Then
'    If IsLoadedF("queSCRAFC") = True Then  -- stop on 09/08

Set rstLocalQueue = CurrentDb.OpenRecordset("Select Completed FROM SCRAqueuefiles where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
rstLocalQueue.Edit
 '   End If
    
  
    
    If IsLoadedF("queSCRAFCNew") = True Then
    'Dim rstLocalQueue2 As Recordset
    Set rstLocalQueue2 = CurrentDb.OpenRecordset("Select * FROM SCRA_ALL_Q where File=" & FileNumber, dbOpenDynaset)
    Dim rstSCRAUpdate As Recordset
    Set rstSCRAUpdate = CurrentDb.OpenRecordset("SCRA_All_update", dbOpenDynaset, dbSeeChanges)
    With rstSCRAUpdate
    .AddNew
    !File = FileNumber
    !Client = rstLocalQueue2!Client
    !Stage = rstLocalQueue2!Stage
    !State = rstLocalQueue2!State
    !RefDate = rstLocalQueue2!RefDate
    !DueDate = rstLocalQueue2!DueDate
    !StageID = rstLocalQueue2!StageID
    !Who = GetFullName
    !DateCompleted = Now
    .Update
     End With
    'Set rstSCRAUpdate = Nothing

        
    'rstLocalQueue2.Edit
    
    
    End If
    
End If







With rstqueue
.Edit
Select Case SCRAID
Case 1
!SCRAComplete1 = Now
!SCRAUser1 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update

' If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
    If IsLoadedF("queSCRAFCNew") = True Then
    rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 
Case 111
!SCRAComplete1a = Now
!SCRAUser1a = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If

.Update

'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
    rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
    
Case 2
!SCRAComplete2 = Now
!SCRAUser2 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update

'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 3
!SCRAComplete3 = Now
!SCRAUser3 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update

'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 31
!SCRAComplete3a = Now
!SCRAUser3a = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
    rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 32
!SCRAComplete3b = Now
!SCRAUser3b = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 33, 45
!SCRAComplete3c = Now
!SCRAUser3c = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 34
!SCRAComplete3d = Now
!SCRAuser3d = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 
Case 39
!SCRAComplete3e = Now
!SCRAUser3e = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 
Case 4
If rstLocalQueue2!State = "VA" Then
!SCRAComplete4a = Now
!SCRAUser4a = GetStaffID
Else
!SCRAComplete4b = Now
!SCRAUser4b = GetStaffID
End If

'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
    rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 42
!SCRAComplete4b = Now
!SCRAUser4b = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If

.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 43
!SCRAComplete4c = Now
!SCRAUser4c = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 
 
Case 44
!SCRAComplete3f = Now
!SCRAUser3f = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 
 
 
Case 5
!SCRAComplete4b = Now
!SCRAUser4b = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 61
!SCRAComplete5a = Now
!SCRAUser5a = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 65
!SCRAComplete5_5 = Now
!SCRAUser5_5 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 Case 125
!SCRAComplete125 = Now
!SCRAUser125 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 
 Case 35, 36, 37, 38
!SCRAComplete3a = Now
!SCRAUser3a = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 
  Case 126
!SCRAComplete126 = Now
!SCRAUser126 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 Case 127
!SCRAComplete127Borroer = Now
!SCRACompelte127 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 128
!SCRAComplete3c = Now
!SCRAUser3c = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
    

Case 93
!SCRAComplete3b = Now
!SCRAUser3b = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 95
!SCRAComplete3b = Now
!SCRAUser3b = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 96
!SCRASentComplaintToCortCompleted = Now
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If

Case 97
!SCRAJudgmentEnteredCompleted = Now
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If

 
Case 98
DoCmd.SetWarnings False

strSQL = "UPDATE TitleReceivedArchive SET " & " SCRAsearch = #" & Now() & "# , SCRASearchBy = '" & GetFullName() & _
    "' WHERE FileNumber = " & Forms!queSCRAFCNew!lstFiles & " AND TitleRecieved = (#" & Forms!queSCRAFCNew!lstFiles.Column(5) & "#)"
    strSQL = ""
DoCmd.SetWarnings True
    rstLocalQueue!Completed = True
    rstLocalQueue.Update
    If IsLoadedF("queSCRAFCNew") = True Then
     rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
     Forms!queSCRAFCNew!QueueCount = rstqueueCount
     Forms!queSCRAFCNew.Refresh
    End If



Case 99
!SCRAComplete3c = Now
!SCRAUser3c = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
    

 
Case 94
!SCRAComplete3c = Now
!SCRAUser3c = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 
 
Case 6
!SCRAComplete5 = Now
!SCRAUser5 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 7
!SCRAComplete6 = Now
!SCRAUser6 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 8
!SCRAComplete7 = Now
!SCRAUser7 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If

.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
 Case 9
!SCRAComplete8 = Now
!SCRAUser8 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
 
Case 90
!SCRAComplete9 = Now
!SCRAUser9 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If
    


Case 91
!SCRAComplete3 = Now
!SCRAUser3 = GetStaffID
'If IsLoadedF("queSCRAFC") = True Then
rstLocalQueue!Completed = True
rstLocalQueue.Update
'End If
.Update
'If IsLoadedF("queSCRAFC") = True Then Forms!queSCRAFC.Refresh
 If IsLoadedF("queSCRAFCNew") = True Then
 rstqueueCount = DCount("file", "SCRA_ALL_Q", "completed= 0")
    Forms!queSCRAFCNew!QueueCount = rstqueueCount
    Forms!queSCRAFCNew.Refresh
    End If


Set rstSCRAUpdate = Nothing
 
Case 14
!SCRABKComplete1 = Now
!SCRABKUser1 = GetStaffID
.Update
Forms!quescraBK.Refresh
Case 15
!SCRABKComplete2 = Now
!SCRABKUser2 = GetStaffID
.Update
Forms!quescraBK.Refresh
Case 16
!SCRABKComplete3 = Now
!SCRABKUser3 = GetStaffID
.Update
Forms!quescraBK.Refresh
Case 10
    If selecteddoctype = 1010 Then
    !SCRAPACERreferralComplete = Now
    !SCRAPACERreferralUser = GetStaffID
    Forms!quescra9!cmdUploadSCRA.Enabled = True
    Forms!quescra9!cmdUploadSCRA.SetFocus
    Forms!quescra9!cmdUploadPACER.Enabled = False
    Else
    Forms!quescra9!cmdUploadPACER.Enabled = True
    Forms!quescra9!cmdUploadPACER.SetFocus
    Forms!quescra9!cmdUploadSCRA.Enabled = False
    End If
    .Update
    Forms!quescra9.Refresh
Case 11
    If selecteddoctype = 1010 Then
    !SCRAPACERHearingcomplete = Now
    !SCRAPACERHearinguser = GetStaffID
    Forms!quescra9!cmdUploadSCRA.Enabled = True
    Forms!quescra9!cmdUploadSCRA.SetFocus
    Forms!quescra9!cmdUploadPACER.Enabled = False
    Else
    Forms!quescra9!cmdUploadPACER.Enabled = True
    Forms!quescra9!cmdUploadPACER.SetFocus
    Forms!quescra9!cmdUploadSCRA.Enabled = False
    End If
    .Update
    Forms!quescra9.Refresh
Case 12
    If selecteddoctype = 1010 Then
    !SCRAPACERLockoutcomplete = Now
    !SCRAPACERLockoutuser = GetStaffID
    Forms!quescra9!cmdUploadSCRA.Enabled = True
    Forms!quescra9!cmdUploadSCRA.SetFocus
    Forms!quescra9!cmdUploadPACER.Enabled = False
    Else
    Forms!quescra9!cmdUploadPACER.Enabled = True
    Forms!quescra9!cmdUploadPACER.SetFocus
    Forms!quescra9!cmdUploadSCRA.Enabled = False
    End If
    .Update
    Forms!quescra9.Refresh
End Select
End With
Set rstqueue = Nothing
End If

If SCRAID = 13 Then
SCRAID = ""
Forms!foreclosuredetails!VAAppraisal = Date
AddStatus FileNumber, Date, "Ordered VA Appraisal"
End If
DoCmd.Close
Exit Sub
'''''''''''''''''
Else

If SCRAID = "FLMA" Then
'Final LMA
selecteddoctype = 1345
'DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage for mailing the FLMA|FC-FLMA|FLMA Postage"
DoCmd.OpenForm "GetPostageFLMASEntToCourt", , , , , acDialog, "Enter total postage for mailing the FLMA|FC-FLMA|FLMA Postage|FLMA|FC-FLMA|FLMA Overnight Postage"

Forms!foreclosuredetails!cmdWizComplete.Enabled = True
ElseIf SCRAID = "BorrowerServed" Then
selecteddoctype = 77 'Bill
Else
DoCmd.OpenForm "Select Document Type", , , , , acDialog
End If

If selecteddoctype = 0 Then Exit Sub

GroupCode = Nz(DLookup("GroupCode", "DocumentTitles", "ID=" & selecteddoctype))
'If GroupCode = "" Then
    newfilename = DLookup("Title", "DocumentTitles", "ID=" & selecteddoctype) & " " & Format$(Now(), "yyyymmdd hhnnss") & fileextension
'Else
'    NewFilename = GroupDelimiter & GroupCode & GroupDelimiter & DLookup("Title", "DocumentTitles", "ID=" & SelectedDocType) & " " & Format$(Now(), "yyyymmdd hhnn") & FileExtension
'End If



Select Case selecteddoctype


Case 119
Dim rstFCDIL As Recordset
Set rstFCDIL = CurrentDb.OpenRecordset("Select * from FCDIL where Filenumber =" & FileNumber)
If Not rstFCDIL.EOF Then
With rstFCDIL
.Edit
!CertOfPubField = Date
.Update
End With
Else
With rstFCDIL
.AddNew
!FileNumber = FileNumber
!CertOfPubField = Date
.Update
End With
End If
AddStatus FileNumber, Date, "Cert Of Pub uploaded "

Set rstFCDIL = Nothing



Case 1, 591   'If title or title update, update fc record
Call GeneralMissingDoc(FileNumber, 591, False, False, False, False, False, , True)
Call GeneralMissingDoc(FileNumber, 1, False, False, False, False, False, , True)
Call GeneralMissingDoc(FileNumber, 1, False, False, True, False, False)
Call GeneralMissingDoc(FileNumber, 591, False, False, True, False, False)
    Dim rstFCdetails As Recordset
    AddStatus FileNumber, Date, "Received title"
    If CaseTypeID = 1 Then
    If ClientID <> 328 Then 'SLS does own title
    
    'change made on 2_27_15
    DoCmd.OpenForm "GetTitleSearchFee", , , , , acDialog, "Enter Title Search costs, zero if none|FC-TC1|Title Search|Abstractor"

    
    'DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter Title Search costs, zero if none|FC-TC1|Title Search|Abstractor"
        'If JurisdictionID = 4 Or JurisdictionID = 18 Then 'PG and Balt City judgment search
        'DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "PG County or Balt City File:  Enter Judgment Search costs, zero if none|FC-TC1|Judgment Search|Title"
        'End If
    End If
    If CurrentProject.AllForms("foreclosuredetails").IsLoaded = True Then
    DoCmd.Close acForm, "foreclosuredetails"
    Set rstFCdetails = CurrentDb.OpenRecordset("select * from fcdetails where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    With rstFCdetails
    .Edit
    !TitleBack = Date
    !TitleReviewToClient = Null
    .Update
    .Close
    End With
    
    DoCmd.SetWarnings False
    Dim rstsql As String
    rstsql = "Insert InTo TitleReceivedArchive (FileNumber, TitleRecieved, DateEntered) Values ( " & FileNumber & ", '" & Date & "' , '" & Now() & "')"
    DoCmd.RunSQL rstsql
    DoCmd.SetWarnings True


    
    
    End If
    Call Details(1)

     
    End If

'Postage for BK docs
Case 1250, 620, 1375, 1376, 1027, 1251, 500, 1218, 69, 515, 70, 1259, 387, 1372
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage for mailing the BK document|BK-Misc|Bankruptcy Document Postage"
Case 134 'IRS notice
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage for mailing the IRS notices|FC-IRS|IRS Postage- Regular and Cert"
Case 1269 'Mediation packet
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage and/or overnight costs|FC-Med|Mediation Package Postage/Overnight Costs"
Case 162, 1498 'Payoff/Reinstatement Quote
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage and/or overnight costs|FC-LM|Reinstatement Quote Postage/Overnight Costs"
Case 1506 'Debt dispute response
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage and/or overnight costs|FC-Debt|Debt Dispute Response Postage/Overnight Costs"
Case 351 'Lien cert
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage for mailing the Lien Cert|FC-Lien|Lien Certificate Postage"
Case 631, 771 'NiSi order
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter overnight costs for sending NiSi docs to court|FC-NiSi|NiSi to Court Overnight Costs"
Case 238 'NiSi
DoCmd.OpenForm "Getfeenew", , , , , acDialog, "Enter NiSi costs|FC-NiSi|NiSi advertising costs|Advertising"
'HUD Occ Letter
Case 1440
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter postage for mailing the HUD Occ Package|FC-HUD|HUD OCC Letter Postage"
Case 1444 'Process Server costs for Provest
DoCmd.OpenForm "Getfeenew", , , , , acDialog, "Enter Service costs for Provest|FC-Misc|Process Server costs- Provest|Process Server"
'DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter actual postage and/or overnight costs for sending the Affidavit of Service|FC-SVC|Affidavit of Service postage/overnight costs"
'Removed Per John Aks 9/27 MC

Case 196 'Notice of Foreclosure sale
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage for Notice of Foreclosure sale |FC-NOS|Notice of Sale Postage"

Case 9 'Lost Note Letter
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter postage for mailing the Lost Note Letter|FC-LNA|Lost Note Letter Postage"
Case 742 'Notice to Borrower
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total regular & cert postage|FC-NOT|Notice to Borrower Postage"

'Case 362 'Assignment
'DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage and/or overnight costs|FC-Assng|Assignment Postage/Overnight Costs"

Case 105 'Report of Sale packet
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage and/or overnight costs|FC-ROS|Report of Sale Postage/Overnight Costs"

Case 645 'SOT
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter overnight costs for the SOT|FC-SOT|Substitution of Trustee Overnight Costs"
Case 283 'Atty correspondence
If MsgBox("Do you need to enter postage or overnight costs?", vbYesNo) = vbYes Then
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter overnight or postage costs for the Attorney Correspondence|FC-Misc|Attorney Correspondence postage or overnight costs"
End If
Case 1194 'NOI
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage for mailing the NOI|FC-NOI|45 Day NOI Postage"
Case 1307 'BK funds to client
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter overnight costs for sending funds to client|BK-Misc|BK Funds to Client Overnight Costs"
Case 535, 145 'Affidavit of Service, Notice to Occupant
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage and/or overnight costs for sending the Affidavit of Service|FC-SVC|Affidavit of Service postage/overnight costs"
Case 461, 72 ' Deed in Lieu
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter postage for mailing the DIL|FC-DIL|DIL Postage"
Case 1499 'Deed in Lieu to client
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter overnight costs for sending the DIL to client|FC-DIL|DIL overnight costs to client"
Case 1500 'Deed in Lieu invoices
DoCmd.OpenForm "Getfeenew", , , , , acDialog, "Enter total DIL costs from this bill|FC-DIL|DIL invoices|"
Case 97 'Docket Package overnight
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter overnight costs for the Docket package|FC-DKT|Docket Package Overnight Cost"
If Forms![Case List]!JurisdictionID = 18 Then DoCmd.OpenForm "GetPostageJuris", , , , , acDialog, "Enter postage for Prince George's registration mailing |FC-DKT|Prince George's registration Mailing"
If Forms![Case List]!State = "MD" Then
Select Case DLookup("City", "FCDetails", "FileNumber = " & FileNumber & " And Current = True")
    Case "Annapolis"
    DoCmd.OpenForm "GetPostageCity", , , , , acDialog, "Enter postage for Annapolis registration mailing |FC-DKT|Annapolis registration Mailing"
    Case "Poolesville"
    DoCmd.OpenForm "GetPostageCity", , , , , acDialog, "Enter postage for Poolesville registration mailing |FC-DKT|Poolesville registration Mailing"
    Case "College Park"
    DoCmd.OpenForm "GetPostageCity", , , , , acDialog, "Enter postage for College Park registration mailing |FC-DKT|College Park registration Mailing"
    Case "Salisbury"
    DoCmd.OpenForm "GetPostageCity", , , , , acDialog, "Enter postage for Salisbury registration mailing |FC-DKT|Salisbury registration Mailing"
    Case "Laurel"
    DoCmd.OpenForm "GetPostageCity", , , , , acDialog, "Enter postage for Laurel registration mailing |FC-DKT|Laurel registration Mailing"
    End Select
End If





Case 1310 'FLMA overnight
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter OVERNIGHT costs excluding postage for the FLMA|FC-FLMA|FLMA Overnight Cost"
'DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter only the total regular and cert postage for mailing the FLMA|FC-FLMA|FLMA Postage"
Case 134 'IRS Notice
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage for mailing the IRS Notices|FC-IRS|IRS Notices"
Case 24 'Title Claim
If MsgBox("Do you need to enter postage or overnight costs?", vbYesNo) = vbYes Then
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage and/or overnight costs|FC-Title|Title Claim Postage/Overnight Costs"
End If
Case 1269 'Mediation packet
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total overnight costs for the Mediation packet|FC-MED|Mediation Overnight Cost"

Case 860  'Loan Modification Agreement
Call GeneralMissingDoc(FileNumber, 860, False, False, True, False, True)
'DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage or overnight costs for the Loan Modification Agreement|FC-OTH|Loan Mod Agreement Postage/Overnight Cost"

Case 1105  'MOD agreement
Call GeneralMissingDoc(FileNumber, 1105, False, False, True, False, True)

Case 1329 'Loss Mit App
If MsgBox("Do you need to enter postage or overnight costs?", vbYesNo) = vbYes Then
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage or overnight costs for the Loss Mitigation Application|FC-Loss|Loss Mitigation postage/overnight costs"
End If
Case 93  'Debt Dispute
If MsgBox("Do you need to enter postage or overnight costs?", vbYesNo) = vbYes Then
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage or overnight costs for the Debt Dispute|FC-Oth|Debt Dispute Postage/Overnight Cost"
End If
Case 1400  'Forbearance Agreement
If MsgBox("Do you need to enter postage or overnight costs?", vbYesNo) = vbYes Then
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage or overnight costs for the Forbearance Agreement|FC-Oth|Forbearance Agreement Postage/Overnight Cost"
End If
Case 5 'HUD title Policy
If CaseTypeID = 2 Or CaseTypeID = 3 Then
DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter the cost of the HUD title policy|FC-Title|HUD Title Policy|"
End If
Case 1477 'Property Reg
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage or overnight costs for the Property Registration|FC-Oth|Property Registration Postage/Overnight Cost"

Case 1495 'Deed calculation to back up line item bill for all transfer taxes
If DCount("", "rqryDeedDocCount", "Filenumber" & FileNumber) = 0 Then
DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter Total Amount of Deed Calculation|FC-TAX|Total Deed Recordation Costs|court"
ElseIf MsgBox("A deed calculation has already been sent to invoicing.  Are you sure you want this sent to invoicing?") = vbYes Then
DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter Total Amount of Deed Calculation|FC-TAX|Total Deed Recordation Costs|court"
End If
Case 75 ' Real property taxes or Stormwater tax
DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter the amount of Real Property Tax or Stormwater due|FC-TAX|Real Property/Stormwater Tax|court"
Case 738 'Water bill lien
DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter the amount of Water Bill or Lien Due|FC-TAX|Water Bill/Lien|"
Case 1496 'Liens
DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter the amount of Lien Due|FC-TAX|City Services/Liens|"
Case 1497 'EV Co-Counsel Bill
DoCmd.OpenForm "GetHours", , , , , acDialog, "EV-OTH"
Case 37, 844, 303 'Motions to Vacate or Reconsider
If MsgBox("Do you need to enter fees/costs associated with this Motion?") = vbYes Then
DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter Total Amount of Fees/Costs|FC-Motion|Attorney Fee- Motion to Vacate or Reconsider|"
End If
Case 96 'Trustees Deed
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage and/or overnight costs|FC-Misc|Trustees' Deed Postage/Overnight Costs"
Case 205




Case 124
       Call GeneralMissingDoc(FileNumber, 124, True, False, False, False, False)




'DoCmd.SetWarnings False
'
'    strSQl = "UPDATE DemandDocsNeeded SET " & " DocReceived = #" & Now() & "# , docreceivedby = " & GetStaffID & _
'    " WHERE FileNumber = " & FileNumber & " AND DocName = ('" & "Waiting for client demand" & "')" & " And IsNull(DocReceived)"
'    DoCmd.RunSQL strSQl
'    strSQl = ""
'    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "Client Demand Letter upload to system" & "',1 )"
'    DoCmd.RunSQL strSQLJournal
'    strSQLJournal = ""
'
'    Forms!Journal.Requery
'
'    Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DemandDocsNeeded where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
'    If rstdocs.EOF Then
'
'    strSQl = "UPDATE wizardqueuestats SET DemandDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
'            DoCmd.RunSQL strSQl
'            strSQl = ""
'    Set rstQueue = Nothing
'    End If
'
'    DoCmd.SetWarnings True
'
    
    
    
    
Case 1493
    Call GeneralMissingDoc(FileNumber, 1493, False, False, False, False, False, , True)
   Call GeneralMissingDoc(FileNumber, 1493, False, False, True, False, False, 1)
     If ClientID <> 446 And ClientID <> 385 Then      ' Not  BOA nither Nation
        Call GeneralMissingDoc(FileNumber, 1493, False, True, False, False, False, 1)
     End If
    
     
Case 1523
    Call GeneralMissingDoc(FileNumber, 1523, False, False, False, False, False, , True)
    Call GeneralMissingDoc(FileNumber, 1523, False, False, True, False, False, 1)
      If ClientID <> 446 And ClientID <> 385 Then      ' Not  BOA nither Nation
        Call GeneralMissingDoc(FileNumber, 1523, False, True, False, False, False, 1)
     End If


  
    
Case 1554
     Call GeneralMissingDoc(FileNumber, 1554, False, True, False, False, False)


Case 1549
    Call GeneralMissingDoc(FileNumber, 1549, True, False, False, True, False)


    
Case 988 ' These for NOI missing docs SA 08/24/14
    Call GeneralMissingDoc(FileNumber, 988, True, False, False, False, False)



Case 1553 ' These for NOI missing docs SA 08/24/14
    Call GeneralMissingDoc(FileNumber, 1553, False, False, False, True, False)


Case 1550
    Call GeneralMissingDoc(FileNumber, 1550, True, False, False, False, False)


Case 1371
        Call GeneralMissingDoc(FileNumber, 1371, False, True, True, False, True)
    
    Case 4
        Call GeneralMissingDoc(FileNumber, 4, False, True, True, False, True)
        Call GeneralMissingDoc(FileNumber, 4, False, False, False, False, False, , True)
    Case 1517
        Call GeneralMissingDoc(FileNumber, 1517, False, True, True, False, True)
        Call GeneralMissingDoc(FileNumber, 1517, False, False, False, False, False, , True)
    Case 1450
        Call GeneralMissingDoc(FileNumber, 1450, False, True, True, False, True)
    
    Case 1522
        Call GeneralMissingDoc(FileNumber, 1522, False, True, True, False, True)
    
    Case 962
        Call GeneralMissingDoc(FileNumber, 962, False, False, True, False, True)
    
    Case 1526
        Call GeneralMissingDoc(FileNumber, 1526, False, False, True, False, True)
    
    Case 1511
        Call GeneralMissingDoc(FileNumber, 1511, False, False, True, False, False)
        Call GeneralMissingDoc(FileNumber, 1511, False, False, False, False, False, , True)
    
    Case 1361
        Call GeneralMissingDoc(FileNumber, 1361, False, False, True, False, False)
        Call GeneralMissingDoc(FileNumber, 1361, False, False, False, False, False, , True)
    
    Case 1525
        Call GeneralMissingDoc(FileNumber, 1525, False, False, True, False, False)
        Call GeneralMissingDoc(FileNumber, 1525, False, False, False, False, False, , True)
    
    Case 1
        Call GeneralMissingDoc(FileNumber, 1, False, False, True, False, False)
        Call GeneralMissingDoc(FileNumber, 1, False, False, False, False, False, , True)
    Case 591
        Call GeneralMissingDoc(FileNumber, 591, False, False, True, False, False)
        Call GeneralMissingDoc(FileNumber, 591, False, False, False, False, False, , True)
    
    Case 288
        Call GeneralMissingDoc(FileNumber, 288, False, False, True, False, True)
        Call GeneralMissingDoc(FileNumber, 288, False, False, False, False, False, , True)
    
    Case 362
        Call GeneralMissingDoc(FileNumber, 362, False, False, True, False, False)
        Call GeneralMissingDoc(FileNumber, 362, False, False, False, False, False, , True)
    
        
    Case 299
        Call GeneralMissingDoc(FileNumber, 299, False, False, True, False, False)
        
    Case 223
        Call GeneralMissingDoc(FileNumber, 223, False, False, False, False, True)
    
    Case 464
        Call GeneralMissingDoc(FileNumber, 464, False, False, False, False, True)
        
    Case 592
        Call GeneralMissingDoc(FileNumber, 592, False, False, False, False, True)
        
    Case 1345
        Call GeneralMissingDoc(FileNumber, 1345, False, False, False, False, True)

    Case 1484
        Call GeneralMissingDoc(FileNumber, 1484, False, False, False, False, True)


End Select
End If

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

DoCmd.SetWarnings False
strinfo = "Added SSN Document "
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True



DoCmd.SetWarnings False
strSQLValues = ""
strSQL = ""
strSQLValues = FileNumber & "," & selecteddoctype & ",'" & GroupCode & "'," & GetStaffID() & ",'" & DocDate & "','" & Replace(newfilename, "'", "''") & "','" & Replace(newfilename, "'", "''") & "'"
'Debug.Print strSQLValues
strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
'Debug.Print strSQL
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings True


Call UpdateDocumentList

Case Else

'If SelectedDocType <> 1398 And SelectedDocType <> 1010 And SelectedDocType <> 1449 Then
FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & newfilename



DoCmd.SetWarnings False
strSQLValues = ""
strSQL = ""
strSQLValues = FileNumber & "," & selecteddoctype & ",'" & GroupCode & "'," & GetStaffID() & ",'" & DocDate & "','" & Replace(newfilename, "'", "''") & "','" & Replace(newfilename, "'", "''") & "'"
'Debug.Print strSQLValues
strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
'Debug.Print strSQL
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings True


Call UpdateDocumentList
'End If




End Select




If selecteddoctype = 1105 Or selecteddoctype = 860 Or selecteddoctype = 1493 Or selecteddoctype = 288 Or selecteddoctype = 1450 Or selecteddoctype = 4 Or _
selecteddoctype = 1371 Or selecteddoctype = 362 Or selecteddoctype = 299 Or selecteddoctype = 591 Or selecteddoctype = 1 Or selecteddoctype = 1371 Or selecteddoctype = 962 Or selecteddoctype = 1523 Then


If MsgBox("New document " & newfilename & " accepted.  OK to delete " & Filespec & "?", vbQuestion + vbYesNo) = vbYes Then Kill Filespec
'prompt for NOI document missing queue update
  
 
    'prompt for Restart document missing queue update
'        If Not IsNull(DLookup("FileNumber", "RestartDocumentMissing", "FileNumber=" & filenumber & " and IsNull(DocReceived)")) Then
'        DoCmd.OpenForm "MissingDocsListRestart"
'        Forms!MissingDocsListRestart!FileNbr = filenumber
'        End If

   

Select Case selecteddoctype

        
Dim stDocName As String
    
 
    Case 1513
    stDocName = "Referral-SSN"
    Case 1514
    stDocName = "Pacer-SSN"
    Case 1515
    stDocName = "Screen Print-SSN"
    Case 1516
    stDocName = "SCRA-SSN"
    Case 1517
    stDocName = "Note-SSN"
    Case 1518
    stDocName = "Deed of Trust-SSN"
    Case 1519
    stDocName = "Loan Application-SSN"
    Case 1520
    stDocName = "Orgination Package-SSN"
    Case 1521
    stDocName = "LOA-SSN"
    Case 1522
    stDocName = "Collateral Docs-SSN"
  '  Case 1523
'    stDocName = "Client Figures - SSN"
    Case 1524
    stDocName = "Death Search -SSN"
    Case 1525
    stDocName = "Skip Trace - SSN"
    
    Case 1528
    stDocName = "CFK agreement - SSN"
    Case 1557
    stDocName = "Judgment-SSN"
    Case 1558
    stDocName = "Death Cert-SSN"
    
End Select




Else
If (selecteddoctype <> 1345 And selecteddoctype <> 1549 And selecteddoctype <> 1550 And selecteddoctype <> 124 And selecteddoctype <> 1556 And selecteddoctype <> 1555 And selecteddoctype <> 1546 And selecteddoctype <> 1562 And selecteddoctype <> 1569 And selecteddoctype <> 1570 And selecteddoctype <> 1571 And selecteddoctype <> 1572) Then

If MsgBox("New document " & newfilename & " accepted.  OK to delete " & Filespec & "?", vbQuestion + vbYesNo) = vbYes Then Kill Filespec

  
  'added on 2/5/15 for not show missing file forms
   If selecteddoctype <> 1525 Then
        If Not IsNull(DLookup("FileNumber", "RestartDocumentMissing", "FileNumber=" & FileNumber & " and IsNull(DocReceived)")) Then
        DoCmd.OpenForm "MissingDocsListRestart"
        Forms!MissingDocsListRestart!FileNbr = FileNumber
        End If

               
     End If                 '  DoCmd.OpenForm "MissingDocsListDemand"
                      '  End If
End If


End If

Select Case selecteddoctype

Case 1555
    DoCmd.SetWarnings False
    strSQL = ""
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "EV-Service Affidavit document uploaded " & "',1 )"
    DoCmd.RunSQL strSQLJournal
    strSQLJournal = ""
    DoCmd.SetWarnings True
    Forms!Journal.Requery
    

Case 1556
    DoCmd.OpenForm "GetFeeNew", , , , , acDialog, "Enter the cost of the Service Affidavit|EV|EV Service Affidavit|EV-Service-Affidavit"

    DoCmd.SetWarnings False
    strSQL = ""
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "EV-Invoiced Service Affidavit document uploaded " & "',2 )"
    DoCmd.RunSQL strSQLJournal
    strSQLJournal = ""
    DoCmd.SetWarnings True
    Forms!Journal.Requery



Case 167

DoCmd.OpenForm "GetFeeNew1", , , , , acDialog, "Enter Advertising(Newspapers) cost|FC-ADV|Advertising|Advertising"
AddInvoiceItem Me.FileNumber, "FC-ADV", "Advertising", Format$(Me.txtAmt, "Currency"), Me.txt_Vendor, False, False, False, True



Case 1546, 1562  'Litigation Billing Docs project 1/10/15 SA

    DoCmd.SetWarnings False
        Dim DocIDNo As Long
        Dim clientShor As String
        clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)
        DocIDNo = GetLastDocIDNo(GetStaffID(), selecteddoctype, FileNumber)
        strSQL = "Insert into Accou_LitigationBillingQueue (CaseFile, ProjectName, ClientShortName, Client, DIQ, Hold, MangNotic, DocIndexID, DocumentId, StaffID, StaffName) Values (" & FileNumber & ", ' " & Forms![Case List]!PrimaryDefName & "' , '" & _
                clientShor & " '," & Forms![Case List]!ClientID & ", Now(), '','', " & DocIDNo & ", " & selecteddoctype & ", " & GetStaffID() & ", '" & GetFullName() & "'" & ")"
        
        DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True



Case 1569  'Client Figures  Nationstar ticket no. 4/15/2015 SA

If Forms![Case List]!ClientID = 385 Then
 Call GeneralMissingDoc(FileNumber, 1569, False, True, False, False, False, 385)
End If

Case 1571   'Client Figures  Nationstar ticket no. 4/15/2015 SA

If Forms![Case List]!ClientID = 385 Then
 Call GeneralMissingDoc(FileNumber, 1571, False, True, False, False, False, 385)
End If



Case 1570, 1572 'Client Figures  BOA  ticket no. 4/15/2015 SA
    
    If Forms![Case List]!ClientID = 446 Then
         Call GeneralMissingDoc(FileNumber, 1570, False, True, False, False, False, 446)
    End If

Case 1572 'Client Figures  BOA  ticket no. 4/15/2015 SA
    
    If Forms![Case List]!ClientID = 446 Then
         Call GeneralMissingDoc(FileNumber, 1572, False, True, False, False, False, 446)
    End If
    


End Select




Exit_cmdAddDoc_Click:
    Exit Sub

Err_cmdAddDoc_Click:


DoCmd.SetWarnings True

        MsgBox Err.Description
        Resume Exit_cmdAddDoc_Click
    'End If
End Sub

Private Sub cmdViewFolder_Click()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else

    On Error GoTo Err_cmdViewFolder_Click
    
    Shell "Explorer """ & DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\""", vbNormalFocus
    
Exit_cmdViewFolder_Click:
        Exit Sub
    
Err_cmdViewFolder_Click:
        MsgBox Err.Description
        Resume Exit_cmdViewFolder_Click
End If

    
End Sub

Private Sub cmdFileLabel_Click()
Dim R As Recordset

On Error GoTo Err_cmdFileLabel_Click

Set R = CurrentDb.OpenRecordset("SELECT * FROM qryFileLabel WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If R.EOF Then
    MsgBox "Cannot print file label because required information is missing.  Check property address, client, and jurisdiction.", vbCritical
Else
    If StartLabel() Then
        Print #6, "|FONTNAME Arial"
        Print #6, "|FONTSIZE 10"
        Print #6, "|FONTBOLD 1"
        Print #6, "|TEXT " & R!ShortClientName & " v. " & R!PrimaryDefName
        Print #6, "|RIGHT #" & FileNumber
        Print #6, "|NEWLINE"
        Print #6, "#" & R!LoanNumber & "   " & R!PropertyAddress & IIf(Len(R![Fair Debt] & "") = 0, "", ", " & R![Fair Debt])
        Print #6, "|RIGHT " & R!Jurisdiction & ", " & R!State
        Call FinishLabel
        MsgBox "File label has been printed", vbInformation
    End If
End If
R.Close

Exit_cmdFileLabel_Click:
    Exit Sub

Err_cmdFileLabel_Click:
    MsgBox Err.Description
    Resume Exit_cmdFileLabel_Click
    
End Sub


Public Sub UpdateAbstractor()

Dim li_AbstractorID As Integer

li_AbstractorID = Nz(DLookup("[ClientAbstrator]", "[ClientList]", "[ClientID] = " & [ClientID]), 0)
If (li_AbstractorID = 0) Then
  li_AbstractorID = Nz(DLookup("[Abstractor]", "[JurisdictionList]", "[JurisdictionID] = " & [JurisdictionID]), 0)
End If

CaseAbsractor = li_AbstractorID



End Sub

Public Sub cmdConflicts_Click()

Call CheckConflicts(Me.FileNumber)
Me.sfrmConflicts.Requery

End Sub

Private Sub Command205_Click()
On Error GoTo Err_Command205_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command205_Click:
    Exit Sub

Err_Command205_Click:
    MsgBox Err.Description
    Resume Exit_Command205_Click
    
End Sub

