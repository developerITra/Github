VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_wizRestartCaseList1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const GroupDelimiter = ";"

Private Sub Active_AfterUpdate()

If Active Then
    ' OnStatusReport.Enabled = True
Else
    If DCount("*", "qryNeedToInvoiceBK", "FileNumber=" & FileNumber) > 0 Then
        MsgBox "This file cannot be closed because an invoice needs to be done.  See the Accounting department for assistance.", vbCritical
        Exit Sub
    End If
    If IsNull(CloseDate) Then
        Select Case MsgBox("Do you really want to close this file?", vbQuestion + vbYesNoCancel)
            Case vbYes
                OnStatusReport = False
                ' OnStatusReport.Enabled = False
                CloseDate = Now
                ClosedBy = GetLoginName()
                ClosedNumber = ReserveNextClosedNumber()
            Case vbNo, vbCancel
                Active = True
        End Select
    End If
    Call AddStatus(FileNumber, Now(), "File closed")
End If
End Sub

Private Sub CaseTypeID_AfterUpdate()
Call UpdateCaption
End Sub


Private Sub cboSortby_AfterUpdate()
  UpdateDocumentList
  
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

'Private Sub cmdAddtoQueue_Click()
'Dim rstFCdetails As Recordset, rstCase As Recordset, Reason As Long
'Set rstCase = CurrentDb.OpenRecordset("select * from caselist where filenumber = " & FileNumber, dbOpenDynaset)
''Set rstqueue = CurrentDb.OpenRecordset("select RestartReasonRestartQueue from wizardqueuestats where filenumber = " & FileNumber, dbOpenDynaset)
'Set rstFCdetails = CurrentDb.OpenRecordset("SELECT DispositionRescinded FROM FCDetails WHERE FileNumber = " & FileNumber & " AND Current = True", dbOpenDynaset, dbSeeChanges)
'
'If CaseTypeID = 2 Then
'Reason = 1
'    Call RestartRSICompletionUpdate(FileNumber, Reason)
'    MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it is an active Bankruptcy."
'Else
'Select Case txtDisposition & ""
'Case 1
'    If IsNull(rstFCdetails!DispositionRescinded) Then
'    Reason = 2
'    Call RestartRSICompletionUpdate(FileNumber, Reason)
'    MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it is a Buy In with no rescinded date."
'    End If
'Case 2
'    If IsNull(rstFCdetails!DispositionRescinded) Then
'    Reason = 2
'    Call RestartRSICompletionUpdate(FileNumber, Reason)
'    MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it is a 3rd Party with no rescinded date."
'    End If
'
'Case ""
'    Reason = 3
'    Call RestartRSICompletionUpdate(FileNumber, Reason)
'    MsgBox "File " & FileNumber & " has been added to the RSI Review Queue because it has no disposition."
'
'Case Else
'
'    With rstCase
'    .Edit
'    !Active = True
'    !OnStatusReport = True
'    !RestartReceived = Now
'    .Update
'    End With
'    'AddStatus FileNumber, Date, "Restart Pending"
'
'
'''''''''''''''
'If MsgBox("Are you sure you want to add another Foreclosure?  If so, you MUST complete the wizard by pressing the 'Complete Wizard' button that will appear on the Details tab", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
'
''Emulate Add FC here
'If (Nz(Disposition) = 2) Or (Nz(Disposition) = 1) Then
'    If PrivAdmin Then
'        If MsgBox("The property has already been sold! Are you sure you want to add another Foreclosure?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
'    Else
'        MsgBox "You cannot add another Foreclosure because the property has already been sold.  (Management can override this for you.)", vbCritical
'        Exit Sub
'    End If
'End If
'
''Convert these to rst
'Dim rstCaseList As Recordset, FileNum As Long, fc As Recordset
'Set rstCaseList = CurrentDb.OpenRecordset("SELECT * FROM CaseList WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
'With rstCaseList
'.Edit
'!ReferralDate = Date
'!ReferralDocsReceived = Null
'!RestartReceived = Null
'.Update
'End With
'rstCaseList.Close
''Forms![Case List]!Active = True
''Forms![Case List]!OnStatusReport = True
'
'Call AddStatus(FileNumber, Now(), "Referral Date")
'FileNum = FileNumber
'Set fc = CurrentDb.OpenRecordset("SELECT * FROM FCDetails WHERE FileNumber = " & FileNum & " AND Current = True", dbOpenDynaset, dbSeeChanges)
'Me.AllowAdditions = True
'DoCmd.GoToRecord , , acNewRec
'
'Me.AllowEdits = True
'
'
'    If Not fc.EOF Then
'
'        FileNumber = FileNum
'        NewFC = Date
'        Referral = Date
'        PrimaryFirstName = fc("PrimaryFirstName")
'        PrimaryLastName = fc("PrimaryLastName")
'        SecondaryFirstName = fc("SecondaryFirstName")
'        SecondaryLastName = fc("SecondaryLastName")
'        PropertyAddress = fc("PropertyAddress")
'        City = fc("City")
'        State = fc("State")
'        ZipCode = fc("ZipCode")
'        TaxID = fc("TaxID")
'        optLeasehold = fc("Leasehold")
'        GroundRentAmount = fc("GroundRentAmount")
'        GroundRentPayable = fc("GroundRentPayable")
'        LegalDescription = fc("LegalDescription")
'        ShortLegal = fc("ShortLegal")
'        Comment = fc("Comment")
'        DOT = fc("DOT")
'        DOTdate = fc("DOTdate")
'        OriginalTrustee = fc("OriginalTrustee")
'        OriginalBeneficiary = fc("OriginalBeneficiary")
'        Liber = fc("Liber")
'        Folio = fc("Folio")
'        OriginalMortgagors = fc("OriginalMortgagors")
'        OriginalPBal = fc("OriginalPBal")
'        RemainingPBal = fc("RemainingPBal")
'        LoanNumber = fc("LoanNumber")
'        LoanType = fc("LoanType")
'        LienPosition = fc("LienPosition")
'        FHALoanNumber = fc("FHALoanNumber")
'        FNMALoanNumber = fc("FNMALoanNumber")
'        AbstractorCaseNumber = fc("AbstractorCaseNumber")
'        CourtCaseNumber = fc("CourtCaseNumber")
'        TitleThru = fc("TitleThru")
'        If State <> "DC" Then Docket = fc("Docket")
'        TitleReviewTo = fc("TitleReviewTo")
'        TitleReviewOf = fc("TitleReviewOf")
'        TitleReviewFax = fc("TitleReviewFax")
'        TitleReviewNameOf = fc("TitleReviewNameOf")
'        TitleReviewLiens = fc("TitleReviewLiens")
'        TitleReviewJudgments = fc("TitleReviewJudgments")
'        TitleReviewTaxes = fc("TitleReviewTaxes")
'        TitleReviewStatus = fc("TitleReviewStatus")
'        TitleClaim = fc("TitleClaim")
'        TitleClaimSent = fc("TitleClaimSent")
'        TitleClaimResolved = fc("TitleClaimResolved")
'        DeedAppReceived = fc!DeedAppReceived
'        DeedAppDate = fc!DeedAppDate
'        DeedAppSentToRecord = fc!DeedAppSentToRecord
'        DeedAppRecorded = fc!DeedAppRecorded
'        DeedAppLiber = fc!DeedAppLiber
'        DeedAppFolio = fc!DeedAppFolio
'        SentToDocket = fc!SentToDocket
'        Docket = fc!Docket
'        ServiceSent = fc!ServiceSent
'        BorrowerServed = fc!BorrowerServed
'        IRSLiens = fc!IRSLiens
'        NOI = fc!NOI
'        DOTrecorded = fc!DOTrecorded
'        LastPaymentDated = fc!LastPaymentDated
'        AmountOwedNOI = fc!AmountOwedNOI
'        DateOfDefault = fc!DateOfDefault
'        SecuredParty = fc!SecuredParty
'        SecuredPartyPhone = fc!SecuredPartyPhone
'        TypeOfDefault = fc!TypeOfDefault
'        OtherDefault = fc!OtherDefault
'        MortgageLender = fc!MortgageLender
'        MortgageLenderLicense = fc!MortgageLenderLicense
'        MortgageOriginator = fc!MortgageOriginator
'        MortgageOriginatorLicense = fc!MortgageOriginatorLicense
'        AccelerationLetter = fc!AccelerationLetter
'        Current = True
'        FairDebt = fc!FairDebt
'        NewFC = Now()
'        If StaffID = 0 Then Call GetLoginName
'        NewFCBy = StaffID
'
'
'        If (State = "VA") Then
'          AddInvoiceItem FileNum, "FC-REF", "Attorney fee", 600, True, False, False, False
'        End If
'        DoCmd.RunCommand acCmdSaveRecord
'        Do While Not fc.EOF     ' make all previously current records not current
'            fc.Edit
'            fc("Current") = False
'            fc.Update
'            fc.MoveNext
'        Loop
'    End If
'    fc.Close
'   ' Me!Current = True
'DoCmd.Close
'        If MsgBox("Will the Acceleration Letter, NOI, or Fair Debt Letter need to be re-sent?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
'        DoCmd.OpenForm "OptionalValues", , , "filenumber=" & FileNum
'        Else: GoTo ExitProc
'        End If
'
'
'Exit Sub
'
'
'''''''''''''
'
'
'
'
'''''''''''
'DoCmd.OpenForm "wizIntakeRestart"
'Forms!wizIntakeRestart!txtFileNumber = FileNumber
'
'Dim rstNames As Recordset, ctr As Integer
'
'Set rstCase = CurrentDb.OpenRecordset("SELECT * FROM CaseList WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
'Set rstFCdetails = CurrentDb.OpenRecordset("SELECT * FROM FCDetails WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
'Set rstNames = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE (FileNumber=" & FileNumber & " AND Mortgagor = True) OR (FileNumber=" & FileNumber & " AND Owner = True)OR (FileNumber=" & FileNumber & " AND Noteholder = True)", dbOpenDynaset, dbSeeChanges)
'
'With rstCase
'    Forms!wizIntakeRestart!txtReferralDate = !ReferralDate
'    Forms!wizIntakeRestart!txtProjectName = !PrimaryDefName
'    Forms!wizIntakeRestart!cbxClient = !ClientID
'    Forms!wizIntakeRestart!cbxJurisdictionID = !JurisdictionID
'End With
'
'With rstFCdetails
'    Forms!wizIntakeRestart!txtLoanNumber = !LoanNumber
'    Forms!wizIntakeRestart!txtPropertyAddress = !PropertyAddress
'    Forms!wizIntakeRestart!txtCity = !City
'    Forms!wizIntakeRestart!txtState = !State
'    Forms!wizIntakeRestart!txtZipCode = !ZipCode
'End With
'
'
'With rstNames
'If rstNames.RecordCount > 0 Then
'.MoveLast
'ctr = .RecordCount
'.MoveFirst
'Forms!wizIntakeRestart!txtFirstName1 = !First
'Forms!wizIntakeRestart!txtLastName1 = !Last
'If Not IsNull(!SSN) Then Forms!wizIntakeRestart!txtSSN1 = !SSN
'If ctr > 1 Then
'.MoveNext
'Forms!wizIntakeRestart!txtFirstName2 = !First
'Forms!wizIntakeRestart!txtLastName2 = !Last
'If Not IsNull(!SSN) Then Forms!wizIntakeRestart!txtSSN2 = !SSN
'End If
'If ctr > 2 Then
'.MoveNext
'Forms!wizIntakeRestart!txtFirstName3 = !First
'Forms!wizIntakeRestart!txtLastName3 = !Last
'If Not IsNull(!SSN) Then Forms!wizIntakeRestart!txtSSN3 = !SSN
'End If
'If ctr = 4 Then
'.MoveNext
'Forms!wizIntakeRestart!txtFirstName4 = !First
'Forms!wizIntakeRestart!txtLastName4 = !Last
'If Not IsNull(!SSN) Then Forms!wizIntakeRestart!txtSSN4 = !SSN
'End If
'End If
'End With
'rstNames.Close
''Recode closing argument
'DoCmd.Close acForm, Me.Name
'DoCmd.Close acForm, "DocsWindow"
'DoCmd.Close acForm, "Journal"

'End Sub
Private Sub cmdNewBillSheet_Click()
On Error GoTo Err_cmdNewBillSheet_Click

  DoCmd.OpenReport "Bill Sheet New", acViewPreview

Exit_cmdNewBillSheet_Click:
  Exit Sub
  
Err_cmdNewBillSheet_Click:
  MsgBox Err.Description
  Resume Exit_cmdNewBillSheet_Click
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

Private Sub cmdCancel_Click()
If MsgBox("Are you sure you want to cancel without a journal note entered?", vbYesNo) = vbNo Then
Exit Sub
End If
DoCmd.Close
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
    If lstDocs.Selected(i) Then StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)
Next i

Exit_cmdView_Click:
    Exit Sub

Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click
    
End Sub

 Sub cmdDetails_Click()
'Removed Private clause to call from FC print close event
If IsLoadedF("wizRestartFCdetails1") = False Then
Call Details(CaseTypeID)
Else

    If EditFormRSI = True Then
    EditFormRSI = False
    Call Details(CaseTypeID)
    End If

End If
End Sub

Private Sub Details(CaseType As Long)
Dim stDocName As String
Dim stLinkCriteria As String
Dim Details As Recordset

On Error GoTo Err_cmdDetails_Click



If CaseType > 1 And CaseType < 11 Then
MsgBox "You can only use this wizard for foreclosures, please verify the case type", vbCritical
Exit Sub
End If

        stDocName = "wizRestartFCdetails1"
        stLinkCriteria = "[FileNumber]=" & Me![FileNumber] & " AND Current = True"
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM FCDetails WHERE FileNumber = " & Me!FileNumber, dbOpenSnapshot)

If Details.EOF Then
    If IsNull(ReferralDate) Then
        MsgBox "Referral Date is required", vbCritical
        Exit Sub
    End If
    Call AddDetailRecord(CaseType, Me!FileNumber, ReferralDate)
    
        Details.Close
        Set Details = CurrentDb.OpenRecordset("SELECT FileNumber FROM FCDetails WHERE FileNumber = " & Me!FileNumber, dbOpenSnapshot)
        If Details.EOF Then Call AddDetailRecord(1, Me!FileNumber, ReferralDate)
    End If

Details.Close

DoCmd.OpenForm stDocName, , , stLinkCriteria


Exit_cmdDetails_Click:
    Exit Sub

Err_cmdDetails_Click:
    MsgBox Err.Description
    Resume Exit_cmdDetails_Click
    
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

Private Sub Form_Close()
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "DocsWindow"
Call ReleaseFile(FileNumber)
If EMailStatus = 1 Then MsgBox "Reminder: EMail is still active", vbInformation
End Sub

Private Sub cmdClose_Click()
Dim JnlNote As String
On Error GoTo Err_cmdClose_Click
JnlNote = "This file has been checked for conflicts via Restart wizard"
'Dim lrs As Recordset
'Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'    lrs.AddNew
'    lrs![FileNumber] = FileNumber
'    lrs![JournalDate] = Now
'    lrs![Who] = GetFullName()
'    lrs![Info] = JnlNote & vbCrLf
'    lrs![Color] = 1
'    lrs.Update
'    lrs.Close
    
DoCmd.SetWarnings False
strinfo = JnlNote & vbCrLf
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

DoCmd.Close



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
    Case 2
        GroupName = "B"
End Select

lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name],DocIndex.doctitleid AS DocType, DocIndex.Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='" & GroupName & "' AND Filespec IS NOT NULL and DeleteDate is null ORDER BY " & Me.cboSortby
lstDocs.Requery



Exit Sub

UpdateDocumentListErr:
    MsgBox Err.Description, vbCritical
    Exit Sub
    
End Sub
Private Sub UpdateDocumentListOrderDate()
Dim GroupName As String

On Error GoTo UpdateDocumentListErr

Select Case optDocType
    Case 1
        GroupName = ""
    Case 2
        GroupName = "B"
End Select

lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name] FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='" & GroupName & "' AND Filespec IS NOT NULL and DeleteDate is null" ' ORDER BY Datestamp Desc"
lstDocs.Requery

Exit Sub

UpdateDocumentListErr:
    MsgBox Err.Description, vbCritical
    Exit Sub
    
End Sub
Private Sub Form_Open(Cancel As Integer)

Active.Locked = Not PrivCloseFiles
OnStatusReport.Locked = Not PrivCloseFiles

cmdDeleteDoc.Enabled = PrivDeleteDocs



    Me.AllowEdits = False
    
    Detail.BackColor = 8421631
If Not IsNull(Forms!wizRestartCaseList1.txtDisposition) Then
Forms!wizRestartCaseList1.lblDisposition.Visible = True
Forms!wizRestartCaseList1.lblDisposition.Caption = "The file has a disposition of:  " & DLookup("Disposition", "FCDisposition", "ID=" & Forms!wizRestartCaseList1.txtDisposition)
Else
Forms!wizRestartCaseList1.lblDisposition.Caption = "The file has NO disposition and CANNOT be restarted at this time"
End If
Call UpdateCaption
cmdDetails.SetFocus


cmdViewScan.Enabled = (Dir$(ClosedScanLocation & FileNumber & "*.pdf") <> "")
Call UpdateDocumentList
If IsNull(JurisdictionID) Then MsgBox "Jurisdiction is missing!", vbExclamation

'added on 9/11/2014

If DCount("*", "CIVDetails", "FCFileNumber= " & FileNumber) > 0 Then
    MsgBox "CAUTION! Litigation in progress, see an attorney!", vbExclamation
    Me.Detail.BackColor = vbYellow
End If


End Sub

Private Sub JurisdictionID_AfterUpdate()
  
  
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

Private Sub optDocType_AfterUpdate()
Call UpdateDocumentList
End Sub



Private Sub PrimaryDefName_AfterUpdate()
Call UpdateCaption
End Sub



Private Sub ReferralDate_AfterUpdate()
Call AddStatus(FileNumber, ReferralDate, "Referral Date")
End Sub

Private Sub ReferralDocsReceived_DblClick(Cancel As Integer)
ReferralDocsReceived = Date
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
RestartReceived = Date
Call RestartReceived_AfterUpdate
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

Private Sub cmdAddDoc_Click()



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



Call UpdateDocumentList
If MsgBox("New document " & newfilename & " accepted.  OK to delete " & Filespec & "?", vbQuestion + vbYesNo) = vbYes Then Kill Filespec


If selecteddoctype = 205 Then
If Forms!foreclosuredetails!WizardSource = "Title" Then
Forms!foreclosuredetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!cmdcloserestart.Visible = False
End If
End If




Select Case selecteddoctype
Dim rstFCdetails As Recordset
Case 1   'If title or title update, update fc records

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


End Select






Exit_cmdAddDoc_Click:
    Exit Sub

Err_cmdAddDoc_Click:

        MsgBox Err.Description
        Resume Exit_cmdAddDoc_Click


' old coding
'Dim fso
'Dim ss As String
'ss = "SSN"
'Set fso = CreateObject("Scripting.FileSystemObject")
'
'    If Not (fso.FolderExists(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\")) Then
'        fso.CreateFolder (DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\")
'    End If
'
'    If Not (fso.FolderExists(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & ss & "\")) Then
'        fso.CreateFolder (DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & ss & "\")
'        Dim objFSO
'        Dim objFolder
'        Set objFSO = CreateObject("Scripting.FileSystemObject")
'        Set objFolder = objFSO.GetFolder(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\")
'
'            If objFolder.Attributes = objFolder.Attributes And 2 Then
'               objFolder.Attributes = objFolder.Attributes Xor 2
'            End If
'    End If
'
'
'
'Dim Filespec As String, FileExtension As String, Path As String, FileName As String, NewFilename As String, i As Integer, Prompt As String
'Dim GroupCode As String, DocType As String, rstDoc As Recordset, DocDateInput As String, DocDate As Date
'
'On Error GoTo Err_cmdAddDoc_Click
'
'Me.Refresh
'
'
'
'Filespec = OpenFile(Me)
'If Filespec = "" Then Exit Sub
'
'For i = Len(Filespec) To 0 Step -1
'    If Asc(Mid$(Filespec, i, 1)) <> 0 Then Exit For
'Next i
'If i = 0 Then
'    MsgBox "Invalid file specification: " & Filespec, vbCritical
'    Exit Sub
'End If
'Filespec = Left$(Filespec, i)
'
'For i = Len(Filespec) To 0 Step -1
'    If Mid$(Filespec, i, 1) = "." Then Exit For
'Next i
'If i = 0 Then
'    MsgBox "Invalid file specification: " & Filespec, vbCritical
'    Exit Sub
'End If
'FileExtension = Mid$(Filespec, i)
'
'For i = Len(Filespec) To 0 Step -1
'    If Mid$(Filespec, i, 1) = "\" Then Exit For
'Next i
'If i = 0 Then
'    MsgBox "Invalid file specification: " & Filespec, vbCritical
'    Exit Sub
'End If
'
'Path = Left$(Filespec, i)
'FileName = Mid$(Filespec, i + 1)
'DoCmd.OpenForm "Select Document Type", , , , , acDialog
'If SelectedDocType = 0 Then Exit Sub
'
''MsgBox (SelectedDocType)
'GroupCode = Nz(DLookup("GroupCode", "DocumentTitles", "ID=" & SelectedDocType))
''If GroupCode = "" Then
'    NewFilename = DLookup("Title", "DocumentTitles", "ID=" & SelectedDocType) & " " & Format$(Now(), "yyyymmdd hhnn") & FileExtension
''Else
''    NewFilename = GroupDelimiter & GroupCode & GroupDelimiter & DLookup("Title", "DocumentTitles", "ID=" & SelectedDocType) & " " & Format$(Now(), "yyyymmdd hhnn") & FileExtension
''End If
'
'If dir$(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & NewFilename) <> "" Then
'    MsgBox NewFilename & " already exists.", vbCritical
'    Exit Sub
'End If
'
'If dir$(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & ss & "\") <> "" Then
'    MsgBox NewFilename & " already exists.", vbCritical
'    Exit Sub
'End If
'
'
'If PrivDocDate Then
'    DocDateInput = InputBox$("Enter scan date:", , Format$(Date, "m/d/yyyy"))
'    If DocDateInput = "" Then Exit Sub
'    If Not IsDate(DocDateInput) Then
'        MsgBox ("Invalid or unrecognized date"), vbCritical
'        Exit Sub
'    End If
'    DocDate = CVDate(DocDateInput)
'    If DocDate = Date Then DocDate = Now()  ' if user took default (today) then also store the time
'Else
'    DocDate = Now()
'End If
'
'
'MsgBox (SelectedDocType)
'
'
'Select Case SelectedDocType
'
'Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528
'FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & ss & "\" & NewFilename
'
'DoCmd.SetWarnings False
'strInfo = "Added SSN Document "
'strInfo = Replace(strInfo, "'", "''")
'strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strInfo & "',1 )"
'DoCmd.RunSQL strSQLJournal
'DoCmd.SetWarnings True
'
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
'
'
'Call UpdateDocumentList
'
' MsgBox (SelectedDocType)
'
'
'Case Else
'
' MsgBox (SelectedDocType)
'
'If SelectedDocType = 1398 Then  'And SelectedDocType <> 1010 And SelectedDocType <> 1449 Then
'FileCopy Filespec, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & NewFilename
'
'
'
'
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
'
'
'Call UpdateDocumentList
'End If
'End Select
'
'If MsgBox("New document " & NewFilename & " accepted.  OK to delete " & Filespec & "?", vbQuestion + vbYesNo) = vbYes Then Kill Filespec
''prompt for NOI document missing queue update
'If Not IsNull(DLookup("FileNbr", "DocumentMissing", "FileNbr=" & FileNumber)) Then
'DoCmd.OpenForm "MissingDocsList"
'End If
'
'
'Exit_cmdAddDoc_Click:
'    Exit Sub
'
'Err_cmdAddDoc_Click:
'    If Err.Number = 76 Then     ' path not found
'        MkDir DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\"
'        Resume
'    Else
'        MsgBox Err.Description
'        Resume Exit_cmdAddDoc_Click
'    End If
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



