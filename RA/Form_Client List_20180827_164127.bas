VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Client List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboSortby_AfterUpdate()
  UpdateDocumentList
End Sub

Sub cbxSelect_AfterUpdate()
    ' Find the record that matches the control.
    Me.RecordsetClone.FindFirst "[ClientID] = " & Me![cbxSelect]
    Me.Bookmark = Me.RecordsetClone.Bookmark
    
    
 
End Sub

Private Sub ClientAbstractor_AfterUpdate()
   ' update case abstractor for all cases where client abstractor is updated
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("update CASELIST set CaseAbstractor = " & [ClientAbstractor] & " where ClientID = " & Me.ClientID)
    DoCmd.SetWarnings True
    
End Sub



Private Sub cmdAttachedDoc_Click()
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

AttachmentPath = """" & ClientDocLocation & Format$(ClientID, "0000") & "\" & lstDocs.Column(3) & """"

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

Private Sub cmdClose_Click()
On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdAddNew_Click()

On Error GoTo Err_cmdAddNew_Click
DoCmd.GoToRecord , , acNewRec

Exit_cmdAddNew_Click:
    Exit Sub

Err_cmdAddNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdAddNew_Click
    
End Sub

Private Sub cmdDelete_Click()

On Error GoTo Err_cmdDelete_Click
If MsgBox("Really delete this client?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
DoCmd.RunCommand acCmdDeleteRecord

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
    
End Sub

Private Sub Command425_Click()


End Sub

Private Sub DCPostSale_AfterUpdate()
If Not IsNull(txt_DCReferralPct) And Not IsNull(txt_DCComplaintFiledPct) And Not IsNull(txt_DCServiceCompPct) And Not IsNull(txt_DCJudgementEnteredPct) And Not IsNull(txt_DCSalePct) And Not IsNull(DCPostSale) And (txt_DCReferralPct + txt_DCComplaintFiledPct + txt_DCServiceCompPct + txt_DCJudgementEnteredPct + txt_DCSalePct + DCPostSale) <> 1 Then
    MsgBox ("Check DC column under Milestone Billing page, All should add up to 100%")
End If
End Sub

'Private Sub cmdScraR_Click()
'DoCmd.OpenForm "ClientRules", , , Forms![Client List]!ClientID = ruleid, , , "SCRA Rulls"
'Forms!ClietRules!DetailsR.ControlSource = SCRAR
'
'End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'If Not PrivAdmin Then
'    Cancel = 1
'    Me.Undo
'    Call cbxSelect_AfterUpdate
'    MsgBox "You are not authorized to make changes", vbCritical
'End If
'Removed on 11/28, contradicts form_open privileges
End Sub

Private Sub Form_Current()
Call UpdateDocumentList

If PrivClientFeeCost Then
Dim ctrl As Control
'For Each ctrl In Me.sfrmClientContacts.Form.Controls
'If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then
'ctrl.Locked = True
'End If
'If TypeOf ctrl Is CommandButton Then ctrl.Enabled = False
'Next
'
''
        For Each ctrl In Me.sfrmBKClientDeadlines.Form.Controls
        If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then
        ctrl.Locked = False
        End If
        Next
        '
        '
        For Each ctrl In Me.sfrmFCClientDeadlines.Form.Controls
        If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then
        ctrl.Locked = False
        End If
        Next
        '
        For Each ctrl In Me.sfrmEVClientDeadlines.Form.Controls
        If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then
        ctrl.Locked = False
        End If
        Next
'fc

Me.Text116.Locked = False
Me.Text118.Locked = False
Me.Text120.Locked = False
Me.Text364.Locked = False
Me.Text366.Locked = False
Me.Text538.Locked = False
Me.ForbearanceFee.Locked = False
Me.ForbearanceFee.Locked = False
Me.Text419.Locked = False
Me.Text534.Locked = False
Me.FeeVAReferral.Locked = False
Me.FeeDCReferral.Locked = False
Me.FeeMDReferral.Locked = False
Me.TitleClaim.Locked = False
Me.Text275.Locked = False
Me.Text277.Locked = False
Me.Text113.Locked = False
Me.Text103.Locked = False
Me.Text176.Locked = False
Me.Text135.Locked = False
Me.cmdLostNoteCost.Locked = False
Me.Text185.Locked = False
Me.SkipTraceCost.Locked = False
Me.Text495.Locked = False
Me.Combo497.Locked = False

'bk
Me.Text288.Locked = False
Me.Text290.Locked = False
Me.Text299.Locked = False
Me.Text316.Locked = False
Me.Text318.Locked = False
Me.Text334.Locked = False
Me.Text342.Locked = False
Me.Text320.Locked = False
Me.Text332.Locked = False
Me.VAOrder.Locked = False
Me.Text532.Locked = False
Me.Text294.Locked = False
Me.Text309.Locked = False
Me.NonStandardevFee.Locked = False
Me.Text322.Locked = False
Me.Text324.Locked = False
Me.Text344.Locked = False
Me.Text326.Locked = False
Me.Text328.Locked = False
Me.Text338.Locked = False
Me.Text340.Locked = False
Me.Text336.Locked = False
Me.Text330.Locked = False
Me.NODObj.Locked = False
Me.NOFCFee.Locked = False
Me.NOFObjCFee.Locked = False
Me.Text311.Locked = False
Me.Text313.Locked = False
Me.Text528.Locked = False
Me.Text530.Locked = False
Me.Text509.Locked = False
Me.Combo511.Locked = False
Me.Text517.Locked = False
Me.Combo521.Locked = False

'millston
Me.VAReferralPct.Locked = False
Me.VA1stActionPct.Locked = False
Me.VASalePct.Locked = False
Me.VASalePct.Locked = False
Me.Check219.Locked = False
Me.MDReferralPct.Locked = False
Me.MDComplaintFiledPct.Locked = False
Me.MDServiceCompPct.Locked = False
Me.MDSalePct.Locked = False
Me.MDPostSale.Locked = False
Me.MDJudgmentEnteredPct.Locked = False
Me.txt_DCReferralPct.Locked = False
Me.txt_DCComplaintFiledPct.Locked = False
Me.txt_DCServiceCompPct.Locked = False
Me.txt_DCJudgementEnteredPct.Locked = False
Me.txt_DCSalePct.Locked = False
Me.DCPostSale.Locked = False


End If


Me.cbxSelect.Locked = False


End Sub

Private Sub Form_Open(Cancel As Integer)

Me.AllowAdditions = PrivClients 'PrivAdmin
Me.AllowDeletions = PrivClients 'PrivAdmin
cmdAddNew.Enabled = PrivClients 'PrivAdmin Or PrivClients
cmdDelete.Enabled = PrivClients 'PrivAdmin
Me.Active.Enabled = PrivClients 'PrivAdmin Or PrivClients
Me.RoseEquelClient.Enabled = PrivClients
Me.cmdAddDoc.Enabled = PrivClients
'Me.cmdViewDocFolder.Enabled = PrivClients
Me.cmdDeleteDoc.Enabled = PrivClients
Me.cmdAttachedDoc.Enabled = PrivClients
'cmdDeleteDoc.Enabled = PrivDeleteDocs

If Not PrivClients Then
Dim ctrl As Control
For Each ctrl In Me.Form.Controls
If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Then
ctrl.Locked = True
End If
Next

For Each ctrl In Me.sfrmClientContacts.Form.Controls
If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then
ctrl.Locked = True
End If
If TypeOf ctrl Is CommandButton Then ctrl.Enabled = False
Next

'
For Each ctrl In Me.sfrmBKClientDeadlines.Form.Controls
If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then
ctrl.Locked = True
End If
Next


For Each ctrl In Me.sfrmFCClientDeadlines.Form.Controls
If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then
ctrl.Locked = True
End If
Next

For Each ctrl In Me.sfrmEVClientDeadlines.Form.Controls
If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then
ctrl.Locked = True
End If
Next


End If



End Sub

Private Sub cmdAddDoc_Click()
Dim Filespec As String, fileextension As String, Path As String, FileName As String, newfilename As String, i As Integer, Prompt As String
Dim rstClientDoc As Recordset, DocDateInput As String, DocDate As Date

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

newfilename = FileName

If Dir$(ClientDocLocation & "\" & Format$(ClientID, "0000") & "\" & newfilename) <> "" Then
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

FileCopy Filespec, ClientDocLocation & Format$(ClientID, "0000") & "\" & newfilename

Set rstClientDoc = CurrentDb.OpenRecordset("DocIndexClients", dbOpenDynaset, dbSeeChanges)
With rstClientDoc
    .AddNew
    !ClientID = ClientID
    !StaffID = GetStaffID()
    !DateStamp = DocDate
    !Filespec = newfilename
    !Notes = newfilename
    .Update
    .Close
End With

Call UpdateDocumentList
If MsgBox("New document " & newfilename & " accepted.  OK to delete " & Filespec & "?", vbQuestion + vbYesNo) = vbYes Then Kill Filespec

Exit_cmdAddDoc_Click:
    Exit Sub

Err_cmdAddDoc_Click:
    If Err.Number = 76 Then     ' path not found
        MkDir ClientDocLocation & "\" & Format$(ClientID, "0000") & "\"
        Resume
    Else
        MsgBox Err.Description
        Resume Exit_cmdAddDoc_Click
    End If
End Sub

Private Sub UpdateDocumentList()

On Error GoTo UpdateDocumentListErr

lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name] FROM DocIndexClients LEFT JOIN Staff ON DocIndexClients.StaffID=Staff.ID WHERE ClientID=" & ClientID & " AND Filespec IS NOT NULL AND DeleteDate IS NULL ORDER BY " & Me.cboSortby
lstDocs.Requery

Exit Sub

UpdateDocumentListErr:
    MsgBox Err.Description, vbCritical
    Exit Sub
    
End Sub

Private Sub cmdViewDocFolder_Click()

On Error GoTo Err_cmdViewFolder_Click

Shell "Explorer """ & ClientDocLocation & Format$(ClientID, "0000") & "\""", vbNormalFocus

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
DoCmd.RunSQL ("UPDATE DocIndexClients set DeleteDate = Now(), DeleteStaff = '" & ls_LoginName & "' WHERE DocID = " & lstDocs.Column(0))

DoCmd.SetWarnings True

Call UpdateDocumentList

Exit_cmdDeleteDoc_Click:
  Exit Sub
  
Err_cmdDeleteDoc_Click:
  MsgBox Err.Description
  Resume Exit_cmdDeleteDoc_Click
  
End Sub

Private Sub cmdSelAllDoc_Click()
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

Private Sub cmdInvertDocSel_Click()
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

Private Sub cmdViewDoc_Click()
Dim i As Long

On Error GoTo Err_cmdView_Click

For i = 0 To lstDocs.ListCount - 1
    If lstDocs.Selected(i) Then StartDoc ClientDocLocation & Format$(ClientID, "0000") & "\" & lstDocs.Column(3, i)
Next i

Exit_cmdView_Click:
    Exit Sub

Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click
    
End Sub

Private Sub lstDocs_DblClick(Cancel As Integer)
Call cmdViewDoc_Click
End Sub


Private Sub MDComplaintFiledPct_AfterUpdate()
If Not IsNull(MDReferralPct) And Not IsNull(MDComplaintFiledPct) And Not IsNull(MDServiceCompPct) And Not IsNull(MDSalePct) And Not IsNull(MDPostSale) And (MDReferralPct + MDComplaintFiledPct + MDServiceCompPct + MDSalePct + MDPostSale) <> 1 Then

    MsgBox ("Check MD column under Milestone Billing page, All should add up to 100%")
End If

End Sub

Private Sub MDPostSale_AfterUpdate()
If Not IsNull(MDReferralPct) And Not IsNull(MDComplaintFiledPct) And Not IsNull(MDServiceCompPct) And Not IsNull(MDSalePct) And Not IsNull(MDPostSale) And (MDReferralPct + MDComplaintFiledPct + MDServiceCompPct + MDSalePct + MDPostSale) <> 1 Then

    MsgBox ("Check MD column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub MDReferralPct_AfterUpdate()
If Not IsNull(MDReferralPct) And Not IsNull(MDComplaintFiledPct) And Not IsNull(MDServiceCompPct) And Not IsNull(MDSalePct) And Not IsNull(MDPostSale) And (MDReferralPct + MDComplaintFiledPct + MDServiceCompPct + MDSalePct + MDPostSale) <> 1 Then

    MsgBox ("Check MD column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub MDSalePct_AfterUpdate()
If Not IsNull(MDReferralPct) And Not IsNull(MDComplaintFiledPct) And Not IsNull(MDServiceCompPct) And Not IsNull(MDSalePct) And Not IsNull(MDPostSale) And (MDReferralPct + MDComplaintFiledPct + MDServiceCompPct + MDSalePct + MDPostSale) <> 1 Then

    MsgBox ("Check MD column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub MDServiceCompPct_AfterUpdate()
If Not IsNull(MDReferralPct) And Not IsNull(MDComplaintFiledPct) And Not IsNull(MDServiceCompPct) And Not IsNull(MDSalePct) And Not IsNull(MDPostSale) And (MDReferralPct + MDComplaintFiledPct + MDServiceCompPct + MDSalePct + MDPostSale) <> 1 Then

    MsgBox ("Check MD column under Milestone Billing page, All should add up to 100%")
End If
End Sub

'Private Sub SelectFile1_Click()
'SelectFile1 = Application.FileDialog(msoFileDialogFilePicker)
'End Sub
'
'Private Sub SelectFile2_Click()
'SelectFile2 = Application.GetOpenFileName
'End Sub
'
'Private Sub SelectFile3_Click()
'SelectFile3 = Application.GetOpenFileName
'End Sub
Private Sub RoseEquelClient_Click()
Dim C As Recordset


Set C = CurrentDb.OpenRecordset("SELECT * FROM ClientList WHERE ClientID = " & ClientID, dbOpenSnapshot)
If Not C.EOF Then
    
    
      ClientNameAsInvestor = C("LongClientName")
    
End If
C.Close


End Sub
Private Sub Command443_Click()
On Error GoTo Err_Command443_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ClientRules"
    
    stLinkCriteria = "[ruleid]=" & Me![ClientID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command443_Click:
    Exit Sub

Err_Command443_Click:
    MsgBox Err.Description
    Resume Exit_Command443_Click
    
End Sub
Private Sub Command445_Click()
On Error GoTo Err_Command445_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim rstClientrules As Recordset
    'If DLookup("ClientNo", "Rules", "ClientID = " & ClientNo) = False Then
    Set rstClientrules = CurrentDb.OpenRecordset("Select * From Rules where ClientNo = " & ClientID, dbOpenDynaset, dbSeeChanges)
    If rstClientrules.EOF Then
    With rstClientrules
    .AddNew
    !ClientNo = ClientID
    .Update
    End With
    rstClientrules.Close
    
    End If
    

    stDocName = "ClientRules"
    
    stLinkCriteria = "[ClientNo]=" & Me![ClientID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command445_Click:
    Exit Sub

Err_Command445_Click:
    MsgBox Err.Description
    Resume Exit_Command445_Click
    
End Sub

Private Sub txt_DCComplaintFiledPct_AfterUpdate()
If Not IsNull(txt_DCReferralPct) And Not IsNull(txt_DCComplaintFiledPct) And Not IsNull(txt_DCServiceCompPct) And Not IsNull(txt_DCJudgementEnteredPct) And Not IsNull(txt_DCSalePct) And Not IsNull(DCPostSale) And (txt_DCReferralPct + txt_DCComplaintFiledPct + txt_DCServiceCompPct + txt_DCJudgementEnteredPct + txt_DCSalePct + DCPostSale) <> 1 Then
    MsgBox ("Check DC column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub txt_DCJudgementEnteredPct_AfterUpdate()
If Not IsNull(txt_DCReferralPct) And Not IsNull(txt_DCComplaintFiledPct) And Not IsNull(txt_DCServiceCompPct) And Not IsNull(txt_DCJudgementEnteredPct) And Not IsNull(txt_DCSalePct) And Not IsNull(DCPostSale) And (txt_DCReferralPct + txt_DCComplaintFiledPct + txt_DCServiceCompPct + txt_DCJudgementEnteredPct + txt_DCSalePct + DCPostSale) <> 1 Then
    MsgBox ("Check DC column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub txt_DCReferralPct_AfterUpdate()
If Not IsNull(txt_DCReferralPct) And Not IsNull(txt_DCComplaintFiledPct) And Not IsNull(txt_DCServiceCompPct) And Not IsNull(txt_DCJudgementEnteredPct) And Not IsNull(txt_DCSalePct) And Not IsNull(DCPostSale) And (txt_DCReferralPct + txt_DCComplaintFiledPct + txt_DCServiceCompPct + txt_DCJudgementEnteredPct + txt_DCSalePct + DCPostSale) <> 1 Then
    MsgBox ("Check DC column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub txt_DCSalePct_AfterUpdate()
If Not IsNull(txt_DCReferralPct) And Not IsNull(txt_DCComplaintFiledPct) And Not IsNull(txt_DCServiceCompPct) And Not IsNull(txt_DCJudgementEnteredPct) And Not IsNull(txt_DCSalePct) And Not IsNull(DCPostSale) And (txt_DCReferralPct + txt_DCComplaintFiledPct + txt_DCServiceCompPct + txt_DCJudgementEnteredPct + txt_DCSalePct + DCPostSale) <> 1 Then
    MsgBox ("Check DC column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub txt_DCServiceCompPct_AfterUpdate()
If Not IsNull(txt_DCReferralPct) And Not IsNull(txt_DCComplaintFiledPct) And Not IsNull(txt_DCServiceCompPct) And Not IsNull(txt_DCJudgementEnteredPct) And Not IsNull(txt_DCSalePct) And Not IsNull(DCPostSale) And (txt_DCReferralPct + txt_DCComplaintFiledPct + txt_DCServiceCompPct + txt_DCJudgementEnteredPct + txt_DCSalePct + DCPostSale) <> 1 Then
    MsgBox ("Check DC column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub VA1stActionPct_AfterUpdate()
If Not IsNull(VAReferralPct) And Not IsNull(VA1stActionPct) And Not IsNull(VASalePct) And Not IsNull(VAPostsale) And (VAReferralPct + VA1stActionPct + VASalePct + VAPostsale <> 1) Then
    MsgBox ("Check VA column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub VAPostsale_AfterUpdate()
If Not IsNull(VAReferralPct) And Not IsNull(VA1stActionPct) And Not IsNull(VASalePct) And Not IsNull(VAPostsale) And (VAReferralPct + VA1stActionPct + VASalePct + VAPostsale <> 1) Then
    MsgBox ("Check VA column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub VAReferralPct_AfterUpdate()
If Not IsNull(VAReferralPct) And Not IsNull(VA1stActionPct) And Not IsNull(VASalePct) And Not IsNull(VAPostsale) And (VAReferralPct + VA1stActionPct + VASalePct + VAPostsale <> 1) Then
    MsgBox ("Check VA column under Milestone Billing page, All should add up to 100%")
End If
End Sub

Private Sub VASalePct_AfterUpdate()
If Not IsNull(VAReferralPct) And Not IsNull(VA1stActionPct) And Not IsNull(VASalePct) And Not IsNull(VAPostsale) And (VAReferralPct + VA1stActionPct + VASalePct + VAPostsale <> 1) Then
    MsgBox ("Check VA column under Milestone Billing page, All should add up to 100%")
End If
End Sub
