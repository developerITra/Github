VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_BankruptcyDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim cntInv As Integer
Dim rstbk As Recordset
Dim AttorneyDisconnectedRS As ADODB.Recordset
Option Explicit

Private Sub AffDefaultInvoice()
'
DefObjectionFiled = Null
DefObjectionAnswerFiled = Null


If (Not IsNull(DefObjectionHearing)) Then

  If (DefObjectionHearing > Date) Then
    DefObjectionHearing = Null
  End If
  
End If

DefObjectionStatus = Null
DefObjectionStatusDate = Null
FiledInError = Null

Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
    FeeAmount = Nz(DLookup("AffidavitDefaultFee", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-AOD", "Filed Affidivit of Default Fee", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-AOD", "Filed Affidavit of Default Fee", 1, 0, True, True, False, False
    End If

End Sub

Private Sub AsOfDate_AfterUpdate()
AddStatus FileNumber, AsOfDate, "As Of Date"
End Sub

Private Sub AsOfDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
AsOfDate = Date
Call AsOfDate_AfterUpdate
End If

End Sub

Private Sub cbxNOFCObjectionStatus_AfterUpdate()
If Not IsNull(cbxNOFCObjectionStatus) Then
    Me.txtNOFCObjectionStatusDate.Value = Date
Else
    Me.txtNOFCObjectionStatusDate.Value = Null
End If

If Len(Me.cbxNOFCObjectionStatus) = 0 Then Me.txtNOFCObjectionStatusDate.Value = Null

If Me.cbxNOFCObjectionStatus = 3 Then Me.txtNOFCHearingDate.Enabled = True

AddStatus FileNumber, Now(), "NOFC Ojbection Status Entered (" & Me.cbxNOFCObjectionStatus.Column(1) & ")"

End Sub

Private Sub cbxNOFCStatus_AfterUpdate()
If Not IsNull(cbxNOFCStatus) Then
    Me.txtNOFCStatusDate.Value = Date
Else
    Me.txtNOFCStatusDate.Value = Null
End If

If Len(Me.cbxNOFCStatus) = 0 Then Me.txtNOFCStatusDate.Value = Null

AddStatus FileNumber, Now(), "NOFC Status Entered (" & Me.cbxNOFCStatus.Column(1) & ")"

End Sub

Private Sub cbxNoticeDisposition_AfterUpdate()

If Not IsNull(cbxNoticeDisposition) Then
txtNoticeDispositionDate.Value = Date

Else
txtNoticeDispositionDate.Value = Null
End If
If Len(cbxNoticeDisposition) = 0 Then txtNoticeDispositionDate.Value = Null

AddStatus FileNumber, Now(), "Notice Status Entered (" & cbxNoticeDisposition.Column(1) & ")"


End Sub

Private Sub cbxPaymentDisposition_Click()

If Not IsNull(cbxPaymentDisposition) Then
TtlPayChgLtrSent.Value = Date

Else
TtlPayChgLtrSent.Value.Value = Null
End If
If Len(cbxPaymentDisposition) = 0 Then TtlPayChgLtrSent.Value = Null

AddStatus FileNumber, Now(), "Payment Status Entered (" & cbxPaymentDisposition.Column(1) & ")"

End Sub

Private Sub cbxPlanDisposition_AfterUpdate()

If Not IsNull(cbxPlanDisposition) Then
PlanReviewed.Enabled = True
PlanReviewed.Value = Date

Else
PlanReviewed.Value = Null
End If

If Len(cbxPlanDisposition) = 0 Then
    PlanReviewed.Value = Null
Else
    PlanReviewed.Enabled = True
End If
AddStatus FileNumber, Now(), "Plan Status Entered (" & cbxPlanDisposition.Column(1) & ")"

End Sub

Private Sub cbxPOCDisposition_AfterUpdate()

If Not IsNull(cbxPOCDisposition) Then
txtPOCDispositionDate.Value = Date

Else
txtPOCDispositionDate.Value = Null
End If
If Len(cbxPOCDisposition) = 0 Then txtPOCDispositionDate.Value = Null

AddStatus FileNumber, Now(), "POC Status Entered (" & cbxPOCDisposition.Column(1) & ")"

End Sub

Private Sub CDHearing_BeforeUpdate(Cancel As Integer)
If Not IsNull(CDHearing) Then

    If HearingCheking(CDHearing, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(CDHearing, 2) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(CDHearing, 3) = 1 Then
    Cancel = 1
    End If
End If

End Sub

Private Sub CmdDeleRef_Click()
TtlPayChgLtrReferral = Null
AddStatus FileNumber, Date, "Removed Payment Change Disposition"
TtlPayChgLtrSent = Null
Me.txtEffectiveDate = Null
'Me.txtFiledByDate
Me.cbxPaymentDisposition = Null


End Sub

Private Sub cmdNewObjection_Click()

Me.txtNOFCHearingDate = Null
Me.cbxNOFCObjectionStatus = Null
Me.txtNOFCObjectionStatusDate = Null
AddStatus FileNumber, txtNOFCHearingDate, "Removed Notice of Final Cure Hearing Date"
End Sub

Private Sub cmdTOC_Click()
On Error GoTo Err_cmdTOC_Click
'If MsgBox("Really do a new Fee and Cost?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub

txtTOCReferral = Date
txtTOCFiled = Null
Call txtTOCReferral_Change

Exit_cmdTOC_Click:
    Exit Sub

Err_cmdTOC_Click:
    MsgBox Err.Description
    Resume Exit_cmdTOC_Click
    
End Sub

Private Sub ComAddName_Click()
DoCmd.OpenForm "sfrmNamesUpdate", , , , acFormAdd

End Sub

Private Sub Combo315_AfterUpdate()

'If Not IsNull(Combo315) Then
    'txt362StatusDate.Value = Date
    'If Len(HearingCalendarEntryID) <> 0 Then
        'Call DeleteCalendarEvent(HearingCalendarEntryID)
        'HearingCalendarEntryID = Null
    'End If


'Else
   ' txt362StatusDate.Value = Null
'End If

'If Len(Combo315) = 0 Then txt362StatusDate.Value = Null

'AddStatus FileNumber, Now(), "Objection Status Entered (" & Combo315.Column(1) & ")"


'---------------------------------------------
'Dim cntInv As Integer
'Dim rstbk As Recordset

Set rstbk = CurrentDb.OpenRecordset("Select * FROM rqryNeedToInoiceBanruptcyChecking where Filenumber=" & FileNumber & " And ( Disposition ='362' or Disposition ='362/SR' ) ", dbOpenDynaset, dbSeeChanges)
                        
If Not rstbk.EOF Then
    rstbk.MoveLast
    cntInv = rstbk.RecordCount
                    
    If cntInv = 2 Then
        MsgBox ("Please notice that " & cntInv & " of 362 Invoice still did not invoiced yet. File can not be closed")
        cntInv = 0
        'Combo315 = ""
        
    Exit Sub
    End If
End If


If Not IsNull(Combo315) Then
    If txt362StatusDate.Visible = True Then
        txt362StatusDate.Value = Date
    Else
        [362StatusDate1].Value = Date
    End If
'End If

    If Len(HearingCalendarEntryID) <> 0 Then
        Call DeleteCalendarEvent(HearingCalendarEntryID)
        HearingCalendarEntryID = Null
    End If


Else
'txt362StatusDate.Value = Null

    If txt362StatusDate.Visible = True Then
        txt362StatusDate.Value = Date
    Else
        [362StatusDate1].Value = Date
    End If
End If

If Len(Combo315) = 0 Then

    If txt362StatusDate.Visible = True Then
        txt362StatusDate.Value = Date
    Else
        [362StatusDate1].Value = Date
    End If
End If
'txt362StatusDate.Value = Null

AddStatus FileNumber, Now(), "Objection Status Entered (" & Combo315.Column(1) & ")"
'added 2/12/15
cmdNew362.Enabled = True
Me.Requery
'End If
'362StatusDateDate

MsgBox ("Disposition Status has beed set")
cntInv = 0

End Sub

Private Sub Command341_Click()
Box340.Visible = True
sfrmPostPetition.Visible = False
sfrmFailedPayment.Visible = True

End Sub

Private Sub Command344_Click()
Box340.Visible = False
sfrmFailedPayment.Visible = False
sfrmPostPetition.Visible = True

End Sub

Private Sub Command357_Click()

If Not IsNull([NODReferral6]) Then
MsgBox (" You already have six referal, you are not able to Add more")
Else
    If Not IsNull([NODReferral5]) Then
    Label352.Visible = True
    NODReferral6.Visible = True
    [6rdAff].Visible = True
    Default6Cured.Visible = True
    Else
        If Not IsNull([NODReferral4]) Then
        Label349.Visible = True
        NODReferral5.Visible = True
        [5rdAff].Visible = True
        Default5Cured.Visible = True
        Else
            If Not IsNull([NODReferral3]) Then
            Label346.Visible = True
            NODReferral4.Visible = True
            [4rdAff].Visible = True
            Default4Cured.Visible = True
            Else
                If Not IsNull([NODReferral2]) Then
                Label54.Visible = True
                NODReferral3.Visible = True
                [3rdAff].Visible = True
                Default3Cured.Visible = True
                Else
                    If Not IsNull([NODReferral1]) Then
                    Label53.Visible = True
                    NODReferral2.Visible = True
                    [2ndAff].Visible = True
                    Default2Cured.Visible = True
                    Else
                    Label52.Visible = True
                    NODReferral1.Visible = True
                    Affidavit.Visible = True
                    Default1Cured.Visible = True
End If
End If
End If
End If
End If
End If

End Sub

Private Sub CommdEdit_Click()

DoCmd.OpenForm "sfrmNamesUpdate", , , WhereCondition:="ID= " & Forms!BankruptcyDetails!sfrmNames!ID


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
'.cmdDelete.Enabled = True
'.cmdNoNotice.Enabled = True
'.cbxNotice.Enabled = True
'.cbxNotice.Locked = False
'
'End With
''Exit Sub
''Else
'


End Sub

Private Sub ConvDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ConvDate = Date
    Call ConvDate_AfterUpdate
End If
End Sub

Private Sub Ctl362Referral1_AfterUpdate()

Dim cntInv As Integer

Set rstbk = CurrentDb.OpenRecordset("Select * FROM rqryNeedToInoiceBanruptcyChecking where Filenumber=" & FileNumber & " And ( Disposition ='362' or Disposition ='362/SR' ) ", dbOpenDynaset, dbSeeChanges)
 If Not rstbk.EOF Then
 
    rstbk.MoveLast
    cntInv = rstbk.RecordCount
                    
    If cntInv = 2 Then
        MsgBox ("Please notice that " & cntInv & " of 362 Invoice still did not invoiced yet")
    cntInv = 0
    Exit Sub
    End If
                        
End If

AddStatus FileNumber, Ctl362Referral1, "Motion for Relief Referral Received"
'AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Filing Fee", Nz(DLookup("IValue", "DB", "Name='MFRfiling'")), 0, False, True, False, True
'Select Case Nz(Chapter)
'    Case 7
'        FeeAmount = Nz(DLookup("Fee362Ch7", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
'    Case 13
'        FeeAmount = Nz(DLookup("Fee362Ch13", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
'End Select
'If FeeAmount > 0 Then
'    AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Attorney Fee", FeeAmount, 0, True, True, False, False
'Else
'    AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Attorney Fee", GetFeeAmount("Motion for Relief Attorney Fee"), 0, True, True, False, False
'End If
 AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Attorney Fee", 400, 0, True, True, False, False

End Sub

Private Sub Ctl362Referral1_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Ctl362Referral)
End Sub

Private Sub Ctl362Referral1_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Ctl362Referral = Date
    Call Ctl362Referral1_AfterUpdate
End If
End Sub

Private Sub Ctl4rdAff_AfterUpdate()
AddStatus FileNumber, Ctl4rdAff, "NOD4 filed"
End Sub

Private Sub Ctl4rdAff_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Ctl4rdAff)

End Sub

Private Sub Ctl4rdAff_DblClick(Cancel As Integer)
Ctl4rdAff = Date
Call Ctl4rdAff_AfterUpdate

End Sub

Private Sub Ctl5rdAff_AfterUpdate()
AddStatus FileNumber, Ctl5rdAff, "NOD5 filed"
End Sub

Private Sub Ctl5rdAff_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Ctl5rdAff)

End Sub

Private Sub Ctl5rdAff_DblClick(Cancel As Integer)
Ctl5rdAff = Now()
Call Ctl5rdAff_AfterUpdate
End Sub

Private Sub Ctl6rdAff_AfterUpdate()
AddStatus FileNumber, Ctl6rdAff, "NOD6 filed"
End Sub

Private Sub Ctl6rdAff_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Ctl6rdAff)
End Sub

Private Sub Ctl6rdAff_DblClick(Cancel As Integer)
Ctl6rdAff = Now()
Call Ctl6rdAff_AfterUpdate
End Sub

Private Sub DateofFiling_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DateofFiling = Date
    Call DateofFiling_AfterUpdate
End If

End Sub

Private Sub Default4Cured_AfterUpdate()
AddStatus FileNumber, Date, "Debtor cured 4rd default"
End Sub

Private Sub Default4Cured_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Default4Cured)
End Sub

Private Sub Default4Cured_DblClick(Cancel As Integer)
Default4Cured = Date
Call Default4Cured_AfterUpdate
End Sub

Private Sub Default5Cured_AfterUpdate()
AddStatus FileNumber, Date, "Debtor cured 5rd default"
End Sub

Private Sub Default5Cured_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Default5Cured)
End Sub

Private Sub Default5Cured_DblClick(Cancel As Integer)
Default5Cured = Now()
Call Default5Cured_AfterUpdate
End Sub

Private Sub Default6Cured_AfterUpdate()
AddStatus FileNumber, Date, "Debtor cured 6rd default"
End Sub

Private Sub Default6Cured_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Default6Cured)

End Sub

Private Sub Default6Cured_DblClick(Cancel As Integer)
Default6Cured = Now()
Call Default6Cured_AfterUpdate
End Sub

Private Sub DefObjectionHearing_BeforeUpdate(Cancel As Integer)
If Not IsNull(DefObjectionHearing) Then

    If HearingCheking(DefObjectionHearing, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(DefObjectionHearing, 2) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(DefObjectionHearing, 3) = 1 Then
    Cancel = 1
    End If

End If


End Sub

Private Sub Form_Activate()
Debug.Print "Activate"


End Sub

Private Sub Form_Load()
Debug.Print "Load"
End Sub

Private Sub Form_Open(Cancel As Integer)
Debug.Print "Open"
    If FileReadOnly Or EditDispute Then
    
        Dim ctl As Control
        Dim lngI As Long
        Dim bSkip As Boolean
    
        For Each ctl In Form.Controls
        Select Case ctl.ControlType
        Case acTextBox, acListBox, acOptionGroup, acCheckBox, acSubform, acOptionButton, acComboBox
             bSkip = False
'                If ctl.Name = "lstDocs" Then bSkip = True
                If Not bSkip Then ctl.Locked = True
'
                
        Case acCommandButton
            bSkip = False
                If ctl.Name = "cmdclose" Then bSkip = True
                If ctl.Name = "cmdselectfile" Then bSkip = True
'                If ctl.Name = "cmdGoToFile" Then bSkip = True
'                If ctl.Name = "cmdClose" Then bSkip = True
'                If ctl.Name = "cmdSelectFile" Then bSkip = True
                If Not bSkip Then ctl.Enabled = False
             
'        Case acComboBox
'            bSkip = False
'            If ctl.Name = "cbxDetails" Then bSkip = True
'            If Not bSkip Then ctl.Locked = True
        
        
        End Select
        Next
    End If

Dim cnn As ADODB.Connection
Dim strSQL As String

On Error GoTo ErrHandler
Set AttorneyDisconnectedRS = New ADODB.Recordset
Set cnn = Application.CurrentProject.Connection

strSQL = "Select * From BKAttorneys order BY LastName"

With AttorneyDisconnectedRS

.CursorLocation = adUseClient
.Open strSQL, cnn, adOpenStatic, adLockBatchOptimistic

Set cnn = Nothing
End With

Set Me.cbxBKAtty.Recordset = AttorneyDisconnectedRS
Exit Sub

ErrHandler:
MsgBox Err.Number & ": " & Err.Description, vbOKOnly, "Error"

End Sub



Private Sub Hearing_BeforeUpdate(Cancel As Integer)
If Not IsNull(Hearing) Then

    If HearingCheking(Hearing, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(Hearing, 2) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(Hearing, 3) = 1 Then
    Cancel = 1
    End If

End If


'Cancel = HearingCheking(Hearing, 2)
'Cancel = HearingCheking(Hearing, 3)

'If Weekday(Hearing) = vbSunday Or Weekday(Hearing) = vbSaturday Then
'    MsgBox "Exceptions Hearing date cannot be Saturday or Sunday", vbCritical
'   Cancel = 1
'End If
'
'    If Hour(Hearing) < 8 Or Hour(Hearing) > 18 Then
'        MsgBox "Invalid Exceptions Hearing time: " & Format$(Hearing, "h:nn am/pm")
'        Cancel = 1
'    End If


    


End Sub

Private Sub LoanModHearingDate_BeforeUpdate(Cancel As Integer)
If Not IsNull(LoanModHearingDate) Then

    If HearingCheking(LoanModHearingDate, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(LoanModHearingDate, 2) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(LoanModHearingDate, 3) = 1 Then
    Cancel = 1
    End If
    
End If

End Sub

Private Sub MLEffectiveDate_AfterUpdate()
If Not IsNull(MLEffectiveDate) Then
  AddStatus FileNumber, MLEffectiveDate, "ML Effective Date entered"
End If
End Sub

Private Sub MLEffectiveDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
MLEffectiveDate = Now()
Call MLEffectiveDate_AfterUpdate
End If

End Sub

Private Sub NODReferral1_AfterUpdate()
AddStatus FileNumber, NODReferral1, "NOD Referral Received"
DefaultCured = 0
Call AffDefaultInvoice
End Sub

Private Sub NODReferral1_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    NODReferral1 = Date
    Call NODReferral1_AfterUpdate
End If
End Sub

Private Sub NODReferral2_AfterUpdate()
AddStatus FileNumber, NODReferral2, "NOD2 Referral Received"
Call AffDefaultInvoice
End Sub

Private Sub NODReferral2_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    NODReferral2 = Date
    Call NODReferral2_AfterUpdate
End If
End Sub

Private Sub NODReferral3_AfterUpdate()
AddStatus FileNumber, NODReferral3, "NOD3 Referral Received"
Call AffDefaultInvoice
End Sub

Private Sub NODReferral3_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    NODReferral3 = Date
    Call NODReferral3_AfterUpdate
End If
End Sub

Private Sub ServicerRelease_AfterUpdate()
Dim Status As String, rstJnl As Recordset
If Not IsNull(ServicerRelease) Then servicereffective = InputBox("Please enter the effective date")
Status = "Servicer Release notified on " & ServicerRelease & "; effective " & servicereffective
AddStatus FileNumber, Now(), Status
'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
With rstJnl
.AddNew
!FileNumber = FileNumber
!JournalDate = Now
!Who = GetFullName
!Info = Status
!Color = 2
.Update
End With
Set rstJnl = Nothing
End Sub
Private Sub Affidavit_AfterUpdate()
AddStatus FileNumber, Affidavit, "NOD filed"
'DefaultCured = 0
'Call AffDefaultInvoice
End Sub

Private Sub Affidavit_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Affidavit)
End Sub

Private Sub Affidavit_DblClick(Cancel As Integer)
Affidavit = Now()
Call Affidavit_AfterUpdate
End Sub

Private Sub AssignBy_AfterUpdate()
Call AssignmentVisuals
End Sub

Private Sub AssignByDOT_AfterUpdate()
Call AssignmentVisuals
End Sub

'Private Sub TitleAssignNeededdate_AfterUpdate()
'  AddStatus FileNumber, Me.TitleAssignNeededDate, "Assignment Drafted"
'
'  AddInvoiceItem FileNumber, "BK-POC", "Assignment Drafted", GetFeeAmount("Assignment Drafted"), False, True, False, True
'  Dim ClientID As Integer
'If Not IsNull(CDDefReferralRecd) = True Then
'            FeeAmount = Nz(DLookup("CramdownOrderFee", "ClientList", "ClientID=" & ClientID))
'            If FeeAmount > 0 Then
'                AddInvoiceItem FileNumber, "BK-CD", "Cramdown Order Fee", FeeAmount, True, True, False, False
'            Else
'                AddInvoiceItem FileNumber, "BK-CD", "Cramdown Order Fee", 1, True, True, False, False
'            End If
'End If
'End Sub
'
'
'Private Sub TitleAssignNeededdate_DblClick(Cancel As Integer)
'  Me.TitleAssignNeededDate = Date
'  Call TitleAssignNeededdate_AfterUpdate
'End Sub
'
'
'Private Sub AssignmentReceived_AfterUpdate()
'  AddStatus FileNumber, Me.AssignmentReceived, "Assignment Received"
'
'End Sub
'
'Private Sub AssignmentReceived_BeforeUpdate(Cancel As Integer)
'Cancel = CheckFutureDate(AssignmentReceived)
'End Sub
'
'Private Sub AssignmentReceived_DblClick(Cancel As Integer)
'  Me.AssignmentReceived = Date
'  Call AssignmentReceived_AfterUpdate
'End Sub
'
'Private Sub AssignmentSentToCourt_AfterUpdate()
'  AddStatus FileNumber, Me.AssignmentSentToCourt, "Assignment Sent To Court"
'
'End Sub
'
'Private Sub AssignmentSentToCourt_BeforeUpdate(Cancel As Integer)
'Cancel = CheckFutureDate(AssignmentSentToCourt)
'End Sub
'
'Private Sub AssignmentSentToCourt_DblClick(Cancel As Integer)
'  Me.AssignmentSentToCourt = Date
'  Call AssignmentSentToCourt_AfterUpdate
'End Sub

Private Sub BarDateDeadline_AfterUpdate()
AddStatus FileNumber, BarDateDeadline, "Bar Date Deadline"
End Sub

Private Sub BKTab_Change()
If BKTab.Value = 10 Then sfrmStatus.Requery
End Sub

Private Sub CaseNo_AfterUpdate()
AddStatus FileNumber, Now(), "Debtor filed Chapter " & Chapter & " Bankruptcy on " & _
    Format$(DateofFiling, "m/d/yyyy") & " in the USBC for " & _
    District.Column(2) & ", " & District.Column(3) & ", Case Number " & CaseNo
Select Case Right(CaseNo, 2)
    Case "JS"
        Courtroom = "Courtroom 9D"
    Case "SD"
        Courtroom = "Courtroom 9C"
    Case "DK"
        Courtroom = "Courtroom 3C"
    Case "PM"
        Courtroom = "Courtroom 3D"
End Select
End Sub

Private Function AttyInfoToggle(iCase As Integer) As Boolean
Dim bOnOff As Boolean
' 2012.02.23
Select Case iCase
Case 0 ' Query:  True if any info
    '2012.03.12 Added test for null DaveW
    If Not IsNull(AttorneyID) Then AttyInfoToggle = True
    'AttyInfoToggle = "" <> Trim(AttorneyLastName.Value)
    'Trim (AttorneyFirstName.Value & AttorneyLastName.Value _
    '        & AttorneyFirm.Text & AttorneyAddress.Value & AttorneyCity.Value _
    '        & AttorneyState.Text & AttorneyZip.Value _
    '        & AttorneyPhone.Text & AttorneyFax.Value)
    Exit Function
Case -1 '  Clear
    AttorneyFirstName.Value = ""
    AttorneyLastName.Value = ""
    AttorneyFirm.Value = ""
    AttorneyAddress.Value = ""
    AttorneyCity.Value = ""
    AttorneyState.Value = ""
    AttorneyZip.Value = ""
    AttorneyPhone.Value = ""
    AttorneyFax.Value = ""
Case Else
    bOnOff = (2 = iCase)
    AttorneyFirstName.Enabled = bOnOff
    AttorneyLastName.Enabled = bOnOff
    AttorneyFirm.Enabled = bOnOff
    AttorneyAddress.Enabled = bOnOff
    AttorneyCity.Enabled = bOnOff
    AttorneyState.Enabled = bOnOff
    AttorneyZip.Enabled = bOnOff
    AttorneyPhone.Enabled = bOnOff
    AttorneyFax.Enabled = bOnOff
    cbxBKAtty.Enabled = bOnOff
    'If bOnOff Then cbxBKAtty.Enabled = True
End Select
    
End Function

Private Sub ckProSe_Click()

' 2012.02.23 DaveW
' Per Tina G:  Add checkbox pro se which greys-out rest of atty choices
If ckProSe Then
    ' Turnin ProSe on:
    If AttyInfoToggle(0) _
    Then
        If vbNo = MsgBox("Are you sure you want to clear the existing Attorney details?", vbYesNo) _
        Then
            ckProSe = False
            Exit Sub
        End If
        
    End If
    AttyInfoToggle (-1) ' clear
    AttyInfoToggle (1) ' disable
Else
    AttyInfoToggle (2) ' enable
End If
End Sub

Private Sub cbxBKAtty_AfterUpdate()
Dim rstBKAtty As Recordset
On Error GoTo ErrHandler

Dim strCriteria As String

strCriteria = "AttyID = '" & Me!cbxBKAtty & "'"

Debug.Print strCriteria

With AttorneyDisconnectedRS

If Not (.EOF) Then

.Find strCriteria
AttorneyFirstName = !FirstName
AttorneyLastName = !LastName
AttorneyFirm = !AttorneyFirm
AttorneyAddress = !Address
AttorneyCity = !City
AttorneyState = !State
AttorneyZip = !Zip
AttorneyPhone = !Phone
AttorneyFax = !FAX
End If
End With
Exit Sub

ErrHandler:
MsgBox Err.Number & ": " & Err.Description, vbOKOnly, "Error"
'If IsNull(cbxBKAtty) Then Exit Sub
'Set rstBKAtty = CurrentDb.OpenRecordset("SELECT * FROM BKAttorneys WHERE AttyID=" & cbxBKAtty, dbOpenDynaset, dbSeeChanges)
'With rstBKAtty
'.Edit
'AttorneyFirstName = !FirstName
'AttorneyLastName = !LastName
'AttorneyFirm = !AttorneyFirm
'AttorneyAddress = !Address
'AttorneyCity = !City
'AttorneyState = !State
'AttorneyZip = !Zip
'AttorneyPhone = !Phone
'AttorneyFax = !FAX
'.Update
'.Close
'End With
'Call AttyInfoToggle(2)

ckProSe = False
End Sub

Private Sub CDAnswerFiled_AfterUpdate()
AddStatus FileNumber, CDAnswerFiled, "Cramdown Answer Filed"

End Sub

Private Sub CDAnswerFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(CDAnswerFiled)
End Sub

Private Sub CDAnswerFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
CDAnswerFiled = Date
AddStatus FileNumber, CDAnswerFiled, "Cramdown Answer Filed"
End If

End Sub

Private Sub CDDefReferralRecd_AfterUpdate()
AddStatus FileNumber, CDDefReferralRecd, "Defense of Cramdown Referral Received"
'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)


If Not IsNull(CDDefReferralRecd) = True Then
            FeeAmount = Nz(DLookup("CramdownOrderFee", "ClientList", "ClientID=" & ClientID))
            If FeeAmount > 0 Then
                AddInvoiceItem FileNumber, "BK-CD", "Cramdown Order Fee", FeeAmount, 0, True, True, False, False
            Else
                AddInvoiceItem FileNumber, "BK-CD", "Cramdown Order Fee", 1, 0, True, True, False, False
            End If
End If

End Sub

Private Sub CDDefReferralRecd_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(CDDefReferralRecd)
End Sub

Private Sub CDDefReferralRecd_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
CDDefReferralRecd = Date
Call CDDefReferralRecd_AfterUpdate
End If

End Sub


Private Sub CDOrderEntered_AfterUpdate()

If Not IsNull(CDOrderEntered) Then
  AddStatus FileNumber, CDOrderEntered, "Cramdown Order Entered"

  

End If
End Sub

Private Sub CDOrderEntered_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
CDOrderEntered = Date
Call CDOrderEntered_AfterUpdate
End If

End Sub

Private Sub CDRespDeadline_AfterUpdate()
AddStatus FileNumber, CDRespDeadline, "Cramdown Response Deadline"
End Sub

Private Sub CDRespDeadline_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(CDRespDeadline)
End Sub

Private Sub Chapter_AfterUpdate()
Trustee.Requery
Call PlanEnabled
End Sub

Private Sub Closed_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Closed)
End Sub

Private Sub Closed_DblClick(Cancel As Integer)


Set rstbk = CurrentDb.OpenRecordset("Select * FROM rqryNeedToInoiceBanruptcyChecking where Filenumber=" & FileNumber & " And ( Disposition ='362' or Disposition ='362/SR' ) ", dbOpenDynaset, dbSeeChanges)
                        
If Not rstbk.EOF Then
    rstbk.MoveLast
    cntInv = rstbk.RecordCount
                    
    If cntInv = 2 Then
        MsgBox ("Please notice that " & cntInv & " of 362 Invoice still did not invoiced yet. File can not be closed")
        cntInv = 0
        'Closed = ""
    Exit Sub
    End If
End If


If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
Closed = Now()
AddStatus FileNumber, Closed, "Case Closed"
End If
End Sub

Private Sub cmdSetDisposition_Click()
On Error GoTo Err_cmdSetDisposition_Click

If IsNull(CDDisposition) And PrivSetDisposition Then
    Call SetDisposition(0)
End If

Exit_cmdSetDisposition_Click:
    Exit Sub

Err_cmdSetDisposition_Click:
    MsgBox Err.Description
    Resume Exit_cmdSetDisposition_Click
End Sub

Private Sub cmdSetFinalDisposition_Click()
If FileReadOnly Or EditDispute Then
   Exit Sub
End If


If Not PrivSetDisposition Then
MsgBox ("You do not have permission to enter a disposition, see your Manager")
Exit Sub
End If

On Error GoTo Err_cmdSetFinalDisposition_Click

If IsNull(BKDisposition) And PrivSetDisposition Then

    
    Call SetFinalDisposition(0)
'    If Sale > Date Then     ' if the sale is in the future then try to remove it from the shared calendar
'        If Not IsNull(Disposition) And Not IsNull(SaleCalendarEntryID) Then
'            Call DeleteCalendarEvent(SaleCalendarEntryID)
'            SaleCalendarEntryID = Null
'        End If
'    End If
End If

Exit_cmdSetFinalDisposition_Click:
    Exit Sub

Err_cmdSetFinalDisposition_Click:
    MsgBox Err.Description
    Resume Exit_cmdSetFinalDisposition_Click



End Sub

Private Sub SetFinalDisposition(DispositionID As Long)
Dim StatusText As String, FeeAmount As Currency


If DispositionID = 0 Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    SelectedDispositionID = 0
    DoCmd.OpenForm "SetDisposition", , , , , acDialog, "BK-FINAL"
Else
    Disposition = DispositionID
End If

If SelectedDispositionID > 0 Then   ' if it was actually set


    cmdClose.SetFocus
    cmdSetFinalDisposition.Enabled = False ' don't allow any changes
    
    BKDisposition = SelectedDispositionID
    BKDisposition.Requery
    BKDispositionDate = Date
    
    If StaffID = 0 Then Call GetLoginName
    BKDispositionStaffID = StaffID
    DoCmd.RunCommand acCmdSaveRecord
    
    'DispositionDesc.Requery
    BKDispositionInitials.Requery
    
    StatusText = DispositionDesc
    If StatusText <> "" Then AddStatus FileNumber, Now(), StatusText
    
    End If

End Sub

Private Sub ConvDate_AfterUpdate()
AddStatus FileNumber, ConvDate, "Converted from Chapter " & ConvChapter & " to Chapter " & Chapter

If Not IsNull(ConvDate) Or ConvDate <> "" Then
    Me.chconvert = True
Else
    Me.chconvert = False
End If

End Sub

Private Sub ConvDate_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ConvDate)
End Sub

Private Sub Ctl2ndAff_AfterUpdate()
AddStatus FileNumber, Ctl2ndAff, "NOD2 filed"
'Call AffDefaultInvoice
End Sub

Private Sub Ctl2ndAff_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Ctl2ndAff)
End Sub

Private Sub Ctl2ndAff_DblClick(Cancel As Integer)
Ctl2ndAff = Now()
Call Ctl2ndAff_AfterUpdate
End Sub

Private Sub Ctl341Date_AfterUpdate()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
AddStatus FileNumber, [341Date], "Meeting of Creditors"
End If
End Sub



Private Sub Ctl362_AfterUpdate()
Dim FeeAmount As Currency

AddStatus FileNumber, [362], "Filed motion for relief from automatic stay"

Select Case Nz(Chapter)
    Case 7
        FeeAmount = Nz(DLookup("Fee362Ch7", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
    Case 13
        FeeAmount = Nz(DLookup("Fee362Ch13", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
End Select
'added 8_20_15
If [362StatusDate1].Visible = True Then
   
   If (Combo315.Column(1) = "Cancelled" And [362StatusDate1] < [362]) Or Dismissed < [362] Or Closed < [362] Or Discharged < [362] Then
   Exit Sub
   End If

ElseIf txt362StatusDate.Visible = True Then
   
    If (Combo315.Column(1) = "Cancelled" And [txt362StatusDate] < [362]) Or Dismissed < [362] Or Closed < [362] Or Discharged < [362] Then
    Exit Sub
   End If
End If


AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Filing Fee", Nz(DLookup("IValue", "DB", "Name='MFRfiling'")), 0, False, True, False, True

'If FeeAmount > 0 Then
    AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Attorney Fee Balance", FeeAmount - 400, 0, True, True, False, False
    'AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Attorney Fee", 400, 0, True, True, False, False
'Else
    'AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Attorney Fee", GetFeeAmount("Motion for Relief Attorney Fee"), 0, True, True, False, False
'End If


End Sub

Private Sub Ctl362_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Ctl362)
End Sub

Private Sub Ctl362_CoDebtor_AfterUpdate()
AddStatus FileNumber, Ctl362_CoDebtor, "Motion from CoDebtor for Relief Referral Received"
End Sub

Private Sub Ctl362_CoDebtor_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Ctl362_CoDebtor)
End Sub

Private Sub Ctl362_CoDebtor_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Ctl362_CoDebtor = Now()
    Call Ctl362_CoDebtor_AfterUpdate
End If

End Sub

Private Sub Ctl362_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Ctl362 = Now()
    Call Ctl362_AfterUpdate
End If
End Sub

Private Sub Ctl362Referral_AfterUpdate()

AddStatus FileNumber, Ctl362Referral, "Motion for Relief Referral Received"
'AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Filing Fee", Nz(DLookup("IValue", "DB", "Name='MFRfiling'")), 0, False, True, False, True
'Select Case Nz(Chapter)
'    Case 7
'        FeeAmount = Nz(DLookup("Fee362Ch7", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
'    Case 13
'        FeeAmount = Nz(DLookup("Fee362Ch13", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
'End Select
'If FeeAmount > 0 Then
'    AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Attorney Fee", FeeAmount, 0, True, True, False, False
'Else
'    AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Attorney Fee", GetFeeAmount("Motion for Relief Attorney Fee"), 0, True, True, False, False
'End If

 AddInvoiceItem FileNumber, "BK-MFR", "Motion for Relief Attorney Fee", 400, 0, True, True, False, False
End Sub

Private Sub Ctl362Referral_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Ctl362Referral)
End Sub

Private Sub Ctl362Referral_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Ctl362Referral = Date
    Call Ctl362Referral_AfterUpdate
End If
End Sub

Private Sub Ctl3rdAff_AfterUpdate()
'Call AffDefaultInvoice
AddStatus FileNumber, Ctl3rdAff, "NOD3 filed"
End Sub

Private Sub Ctl3rdAff_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Ctl3rdAff)
End Sub

Private Sub Ctl3rdAff_DblClick(Cancel As Integer)
Ctl3rdAff = Now()
Call Ctl3rdAff_AfterUpdate
End Sub

Private Sub Current_AfterUpdate()
Current.Locked = Current.Value
End Sub

Private Sub DateofFiling_AfterUpdate()

' if Bk filed, if Eviction Hearing exists and after today, delete from calendar and clear out Eviction hearing date

Dim EVHearingDate As Variant

EVHearingDate = DLookup("HearingDate", "[EVDetails]", "FileNumber = " & Me.FileNumber)
If (EVHearingDate > Date) Then

  Dim EVHearingCalendarEntryID As Variant
  EVHearingCalendarEntryID = DLookup("HearingCalendarEntryID", "[EVDetails]", "FileNumber = " & FileNumber)
  
  If Not IsNull(EVHearingCalendarEntryID) Then
    DeleteCalendarEvent (Nz(EVHearingCalendarEntryID))
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("update EVDetails set HearingDate = null, HearingCalendarEntryID = null where FileNumber = " & FileNumber)
    DoCmd.SetWarnings True
    
  End If

End If

' if Bk filed, if Lockout Date exists and after today, delete from calendar and clear out Lockout date
Dim EVLockoutDate As Variant

EVLockoutDate = DLookup("LockoutDate", "[EVDetails]", "FileNumber = " & Me.FileNumber)
If (EVLockoutDate > Date) Then

  Dim EVLockoutDateCalendarEntryID As Variant
  EVLockoutDateCalendarEntryID = DLookup("LockoutDateCalendarEntryID", "[EVDetails]", "FileNumber = " & FileNumber)
  
  If Not IsNull(EVLockoutDateCalendarEntryID) Then
    DeleteCalendarEvent (Nz(EVLockoutDateCalendarEntryID))
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("update EVDetails set LockoutDate = null, LockoutDateCalendarEntryID = null where FileNumber = " & FileNumber)
    DoCmd.SetWarnings True
    
  End If

End If


End Sub

Private Sub DateofFiling_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DateofFiling)
End Sub

Private Sub DebtLastContact_AfterUpdate()
AddStatus FileNumber, DebtLastContact, "Debt Last Contact"
End Sub

Private Sub DebtLastContact_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DebtLastContact)
End Sub

Private Sub DebtLastContact_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
DebtLastContact = Date
Call DebtLastContact_AfterUpdate
End If

End Sub

Private Sub DebtPaidThru_AfterUpdate()
AddStatus FileNumber, DebtPaidThru, "Debt Paid Thru"
End Sub

Private Sub DebtPaidThru_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DebtPaidThru)
End Sub

Private Sub DebtPaidThru_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
DebtPaidThru = Date
Call DebtPaidThru_AfterUpdate
End If

End Sub

Private Sub Default1Cured_AfterUpdate()
AddStatus FileNumber, Date, "Debtor cured 1st default"
End Sub

Private Sub Default1Cured_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Default1Cured)
End Sub

Private Sub Default1Cured_DblClick(Cancel As Integer)
Default1Cured = Date
Call Default1Cured_AfterUpdate
End Sub

Private Sub Default2Cured_AfterUpdate()
AddStatus FileNumber, Date, "Debtor cured 2nd default"
End Sub

Private Sub Default2Cured_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Default2Cured)
End Sub

Private Sub Default2Cured_DblClick(Cancel As Integer)
Default2Cured = Date
Call Default2Cured_AfterUpdate
End Sub

Private Sub Default3Cured_AfterUpdate()
AddStatus FileNumber, Date, "Debtor cured 3rd default"
End Sub

Private Sub Default3Cured_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Default3Cured)
End Sub

Private Sub Default3Cured_DblClick(Cancel As Integer)
Default3Cured = Date
Call Default3Cured_AfterUpdate
End Sub

Private Sub DefObjectionAnswerFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DefObjectionAnswerFiled)
End Sub

Private Sub DefObjectionFiled_AfterUpdate()
If Not IsNull(DefObjectionFiled) Then AddInvoiceItem FileNumber, "BK/NOD-Objection", "Filed Objection to Default", Nz(DLookup("NODObj", "ClientList", "ClientID=" & Forms![Case List]!ClientID)), 0, True, True, False, False
End Sub

Private Sub DefObjectionFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DefObjectionFiled)
End Sub

Private Sub DefObjectionStatus_AfterUpdate()
  If (Not IsNull(DefObjectionStatus)) Then
    DefObjectionStatusDate = Date
    If Len(DefObjectionHearingCalendarEntryID) <> 0 Then
    Call DeleteCalendarEvent(DefObjectionHearingCalendarEntryID)
    DefObjectionHearingCalendarEntryID = Null
    End If
    
    Else
    DefObjectionStatusDate = Null
    End If
  If Len(DefObjectionStatus) = 0 Then DefObjectionStatusDate = Null
  AddStatus FileNumber, Date, "Default objection status: " & DefObjectionStatus.Column(1)
End Sub

Private Sub Discharged_AfterUpdate()

'added 2/18/15
Set rstbk = CurrentDb.OpenRecordset("Select * FROM rqryNeedToInoiceBanruptcyChecking where Filenumber=" & FileNumber & " And ( Disposition ='362' or Disposition ='362/SR' ) ", dbOpenDynaset, dbSeeChanges)
                        
If Not rstbk.EOF Then
    rstbk.MoveLast
    cntInv = rstbk.RecordCount
                    
    If cntInv = 2 Then
        MsgBox ("Please notice that " & cntInv & " of 362 Invoice still did not invoiced yet. File can not be closed")
        cntInv = 0
        'Discharged = ""
    Exit Sub
    End If
End If


AddStatus FileNumber, Discharged, "Discharged"
End Sub

Private Sub Discharged_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Discharged)
End Sub

Private Sub Discharged_DblClick(Cancel As Integer)

'added 2/18/15
Set rstbk = CurrentDb.OpenRecordset("Select * FROM rqryNeedToInoiceBanruptcyChecking where Filenumber=" & FileNumber & " And ( Disposition ='362' or Disposition ='362/SR' ) ", dbOpenDynaset, dbSeeChanges)
                        
If Not rstbk.EOF Then
    rstbk.MoveLast
    cntInv = rstbk.RecordCount
                    
    If cntInv = 2 Then
        MsgBox ("Please notice that " & cntInv & " of 362 Invoice still did not invoiced yet. File can not be closed")
        cntInv = 0
        'Discharged = ""
    Exit Sub
    End If
End If


If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
Discharged = Date
AddStatus FileNumber, Discharged, "Discharged"
End If
End Sub

Private Sub Dismissed_AfterUpdate()

'added 2/18/15
Set rstbk = CurrentDb.OpenRecordset("Select * FROM rqryNeedToInoiceBanruptcyChecking where Filenumber=" & FileNumber & " And ( Disposition ='362' or Disposition ='362/SR' ) ", dbOpenDynaset, dbSeeChanges)
                        
If Not rstbk.EOF Then
    rstbk.MoveLast
    cntInv = rstbk.RecordCount
                    
    If cntInv = 2 Then
        MsgBox ("Please notice that " & cntInv & " of 362 Invoice still did not invoiced yet. File can not be closed")
        'Dismissed = ""
        cntInv = 0
    Exit Sub
    End If
End If


AddStatus FileNumber, Dismissed, "Case dismissed"
End Sub

Private Sub Dismissed_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Dismissed)
End Sub

Private Sub Dismissed_DblClick(Cancel As Integer)
'added 2/18/15
Set rstbk = CurrentDb.OpenRecordset("Select * FROM rqryNeedToInoiceBanruptcyChecking where Filenumber=" & FileNumber & " And ( Disposition ='362' or Disposition ='362/SR' ) ", dbOpenDynaset, dbSeeChanges)
                        
If Not rstbk.EOF Then
    rstbk.MoveLast
    cntInv = rstbk.RecordCount
                    
    If cntInv = 2 Then
        MsgBox ("Please notice that " & cntInv & " of 362 Invoice still did not invoiced yet. File can not be closed")
        'Dismissed = ""
        cntInv = 0
    Exit Sub
    End If
    
    
End If



If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
Dismissed = Now()
AddStatus FileNumber, Dismissed, "Case dismissed"
End If
End Sub

Private Sub District_AfterUpdate()
Call UpdateHearingLocations
End Sub

Private Sub EnteredAppearance_AfterUpdate()
AddStatus FileNumber, EnteredAppearance, "Line filed entering appearance"
End Sub

Private Sub EnteredAppearance_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(EnteredAppearance)
End Sub

Private Sub EnteredAppearance_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    EnteredAppearance = Now()
    AddStatus FileNumber, EnteredAppearance, "Line filed entering appearance"
End If
End Sub

Private Sub FiledInError_AfterUpdate()
AddStatus FileNumber, FiledInError, "Filed In Error"
End Sub

Private Sub FiledInError_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(FiledInError)
End Sub

Private Sub FiledInError_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    FiledInError = Now()
    AddStatus FileNumber, FiledInError, "Filed In Error"
End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
If (Nz(Hearing) <> Nz(Hearing.OldValue)) Then
  HearingCalendarEntryID = UpdateCalendar(Hearing.OldValue, Hearing, Nz(HearingCalendarEntryID), "362")
End If

If (Nz(txtNOFCHearingDate) <> Nz(txtNOFCHearingDate.OldValue)) Then
     NOFCObjectionHearingDateCalendarEntryID = UpdateCalendar(txtNOFCHearingDate.OldValue, txtNOFCHearingDate, Nz(NOFCObjectionHearingDateCalendarEntryID), "NOFC Hearing")
End If

If (Nz(DefObjectionHearing) <> Nz(DefObjectionHearing.OldValue)) Then
  DefObjectionHearingCalendarEntryID = UpdateCalendar(DefObjectionHearing.OldValue, DefObjectionHearing, Nz(DefObjectionHearingCalendarEntryID), "AFF")
End If

If (Nz(POCObjHearing) <> Nz(POCObjHearing.OldValue)) Then
  POSObjHearingCalendarEntryID = UpdateCalendar(POCObjHearing.OldValue, POCObjHearing, Nz(POSObjHearingCalendarEntryID.OldValue), "POC")
End If

If (Nz(PlanConfHearing) <> Nz(PlanConfHearing.OldValue)) Then
  PlanConfHearingCalendarEntryID = UpdateCalendar(PlanConfHearing.OldValue, PlanConfHearing, Nz(PlanConfHearingCalendarEntryID), "PLAN")
End If

If (Nz(CDHearing) <> Nz(CDHearing.OldValue)) Then
  CDHearingCalendarEntryID = UpdateCalendar(CDHearing.OldValue, CDHearing, Nz(CDHearingCalendarEntryID), "CRAMDOWN Hearing")
End If

If (Nz(CDSchedulingConf) <> Nz(CDSchedulingConf.OldValue)) Then
  CDSchedulingConfCalendarEntryID = UpdateCalendar(CDSchedulingConf.OldValue, CDSchedulingConf, Nz(CDSchedulingConfCalendarEntryID), "CRAMDOWN Scheduling Conference")
End If

If (Nz(LoanModHearingDate) <> Nz(LoanModHearingDate.OldValue)) Then
  Me.LoanModHearingCalendarEntryID = UpdateCalendar(LoanModHearingDate.OldValue, LoanModHearingDate, Nz(LoanModHearingCalendarEntryID), "Loan Modification Hearing")
End If


End Sub

Private Sub Form_Close()

Set AttorneyDisconnectedRS = Nothing
DoCmd.Restore
End Sub

Private Sub Form_Current()
If Not IsNull([NODReferral1]) Then
Label52.Visible = True
NODReferral1.Visible = True
Affidavit.Visible = True
Default1Cured.Visible = True
End If

If Not IsNull([NODReferral2]) Then
Label53.Visible = True
NODReferral2.Visible = True
[2ndAff].Visible = True
Default2Cured.Visible = True
End If

If Not IsNull([NODReferral3]) Then
Label54.Visible = True
NODReferral3.Visible = True
[3rdAff].Visible = True
Default3Cured.Visible = True
End If

If Not IsNull([NODReferral4]) Then
Label346.Visible = True
NODReferral4.Visible = True
[4rdAff].Visible = True
Default4Cured.Visible = True
End If

If Not IsNull([NODReferral5]) Then
Label349.Visible = True
NODReferral5.Visible = True
[5rdAff].Visible = True
Default5Cured.Visible = True
End If

If Not IsNull([NODReferral6]) Then
Label352.Visible = True
NODReferral6.Visible = True
[6rdAff].Visible = True
Default6Cured.Visible = True
End If

If Not IsNull(txtNOFCHearingDate) Then Me.txtNOFCHearingDate.Enabled = False
If Me.cbxNOFCObjectionStatus = 3 Then Me.txtNOFCHearingDate.Enabled = True
If Not IsNull(cbxPlanDisposition) Then PlanReviewed.Enabled = True

Dim bk As Recordset, rstBKAtty As Recordset
Dim bIsNew As Boolean

Me.Caption = "Bankruptcy File " & Me![FileNumber] & " " & Forms![Case List]![PrimaryDefName]


If FileReadOnly Or EditDispute Then
    Me.AllowEdits = False
    cmdNewBankruptcy.Enabled = False
    cmdPrint.Enabled = False
    sfrmPropAddr.Form.AllowEdits = False
    sfrmComments.Form.AllowEdits = False
    sfrmNames.Form.AllowEdits = False
    sfrmNames.Form.AllowAdditions = False
    sfrmNames.Form.AllowDeletions = False
    sfrmNames!cmdCopy.Enabled = False
    sfrmNames!cmdTenant.Enabled = False
    sfrmNames!cmdDelete.Enabled = False
    CommdEdit.Enabled = False
    sfrmNames!cmdNoNotice.Enabled = False
    sfrmBKDebtPre.Form.AllowEdits = False
    sfrmBKDebtPre.Form.AllowAdditions = False
    sfrmBKDebtPre.Form.AllowDeletions = False
    sfrmBKDebtPost.Form.AllowEdits = False
    sfrmBKDebtPost.Form.AllowAdditions = False
    sfrmBKDebtPost.Form.AllowDeletions = False
    sfrmStatus.Form.AllowEdits = False
    sfrmStatus.Form.AllowAdditions = False
    sfrmStatus.Form.AllowDeletions = False
    Detail.BackColor = ReadOnlyColor
    lblShowCurrent.BackColor = ReadOnlyColor
    ckProSe.Enabled = False
    CommdEdit.Enabled = False
    ComAddName.Enabled = False
    
Else
    Me.AllowEdits = True
    'cmdNewBankruptcy.Enabled = True
    cmdPrint.Enabled = True
'    sfrmPropAddr.Form.AllowEdits = True SA 10/05
'    sfrmComments.Form.AllowEdits = True
'    sfrmNames.Form.AllowEdits = True
'    sfrmNames.Form.AllowAdditions = True
'    sfrmNames.Form.AllowDeletions = True
'    sfrmNames!cmdCopy.Enabled = True
'    sfrmNames!cmdTenant.Enabled = True
'    sfrmNames!cmdDelete.Enabled = True
'    sfrmNames!cmdNoNotice.Enabled = True
    sfrmBKDebtPre.Form.AllowEdits = True
    sfrmBKDebtPre.Form.AllowAdditions = True
    sfrmBKDebtPre.Form.AllowDeletions = True
    sfrmBKDebtPost.Form.AllowEdits = True
    sfrmBKDebtPost.Form.AllowAdditions = True
    sfrmBKDebtPost.Form.AllowDeletions = True
    sfrmStatus.Form.AllowEdits = True
    sfrmStatus.Form.AllowAdditions = True
    sfrmStatus.Form.AllowDeletions = True
    Detail.BackColor = -2147483633
    lblShowCurrent.BackColor = -2147483633
    ckProSe.Enabled = True
    
    'If IsNull(AttorneyLastName) Then cbxBKAtty.Enabled = True
    
    If (IsNull(CDDisposition) And PrivSetDisposition) Then
      cmdSetDisposition.Enabled = True
            
    End If
    
    If (IsNull(BKDisposition) And PrivSetDisposition) Then
      cmdSetFinalDisposition.Enabled = True
      
    End If
    
    
End If

'2012.02.27

If Not IsNull(cbxBKAtty) Then
Set rstBKAtty = CurrentDb.OpenRecordset("SELECT * FROM BKAttorneys WHERE AttyID=" & cbxBKAtty, dbOpenDynaset, dbSeeChanges)
Call AttyInfoToggle(2)
ckProSe = False
With rstBKAtty
.Edit
AttorneyFirstName = !FirstName
AttorneyLastName = !LastName
AttorneyFirm = !AttorneyFirm
AttorneyAddress = !Address
AttorneyCity = !City
AttorneyState = !State
AttorneyZip = !Zip
AttorneyPhone = !Phone
AttorneyFax = !FAX
.Update
.Close
End With
Else
Call AttyInfoToggle(1)
ckProSe = True
End If



bIsNew = Me.NewRecord 'Do not show Missing District message on New (2012.03.12)
If bIsNew Then    ' fill in info from previous FC, if any
    Set bk = CurrentDb.OpenRecordset("SELECT * FROM BKDetails WHERE FileNumber = " & FileNumber & " AND Current = True", dbOpenDynaset, dbSeeChanges)
    If Not bk.EOF Then
        'PrimaryFirstName = fc("PrimaryFirstName")
        'PrimaryLastName = fc("PrimaryLastName")
        
        Do While Not bk.EOF     ' make all previously current records not current
            bk.Edit
            bk("Current") = False
            bk.Update
            bk.MoveNext
        Loop
    End If
    bk.Close
    Me!Current = True           ' and make this record current
End If
Call SetPropertyType
Call PlanEnabled
If Not bIsNew Then
    If IsNull(District) Then MsgBox "District is missing!", vbCritical
End If
Call UpdateHearingLocations
Current.Locked = Current.Value

If IsNull([362]) Then
    [362].Locked = False
    [362].BackStyle = 1
    cmdNew362.Enabled = False
Else
    [362].Locked = True
    [362].BackStyle = 0
    cmdNew362.Enabled = True
End If

If Not IsNull([Modify]) Then cmdNew362.Enabled = True
If Not IsNull([Terminating]) Then cmdNew362.Enabled = True
If Not IsNull([362Disposition]) Then cmdNew362.Enabled = True

Call AssignmentVisuals

End Sub

Private Sub UpdateHearingLocations()
If IsNull(District) Then
    HearingLocation.RowSource = ""
Else
    HearingLocation.RowSource = "SELECT ID, HearingAddress FROM HearingLocations WHERE DistrictID=" & District
End If
End Sub
Private Sub Hearing_AfterUpdate()

AddStatus FileNumber, Now(), "Hearing scheduled for " & Format$(Hearing, "m/d/yyyy h:nn am/pm")
End Sub


Private Sub LoanModFiled_AfterUpdate()
If Not IsNull(LoanModFiled) Then
  AddStatus FileNumber, LoanModFiled, "Loan Modification Order Filed"
End If
End Sub

Private Sub LoanModFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
LoanModFiled = Now()
Call LoanModFiled_AfterUpdate
End If

End Sub

Private Sub LoanModOrderEntered_AfterUpdate()
If Not IsNull(LoanModOrderEntered) Then
  AddStatus FileNumber, LoanModOrderEntered, "Loan Modification Order Entered"

End If
End Sub

Private Sub LoanModOrderEntered_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
LoanModOrderEntered = Now()
Call LoanModOrderEntered_AfterUpdate
End If
End Sub

Private Sub LoanModReferralRecd_AfterUpdate()
If Not IsNull(LoanModReferralRecd) Then
  AddStatus FileNumber, LoanModReferralRecd, "Loan Modification Order Referral Received"

'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)

            FeeAmount = Nz(DLookup("LoanModRecdFee", "ClientList", "ClientID=" & ClientID))
            If FeeAmount > 0 Then
                AddInvoiceItem FileNumber, "BK-MISC", "Loan Modification Referral Fee", FeeAmount, 0, True, True, False, False
            Else
                AddInvoiceItem FileNumber, "BK-MISC", "Loan Modification Referral Fee", 1, 0, True, True, False, False
            End If
End If
End Sub

Private Sub LoanModReferralRecd_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
LoanModReferralRecd = Now()
Call LoanModReferralRecd_AfterUpdate
End If


End Sub

Private Sub LoanModwithdrawn_AfterUpdate()
If Not IsNull(LoanModWithdrawn) Then
  AddStatus FileNumber, LoanModWithdrawn, "Loan Modification Order Withdrawn"
End If
End Sub

Private Sub LoanModwithdrawn_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
LoanModWithdrawn = Now()
Call LoanModwithdrawn_AfterUpdate
End If

End Sub

Private Sub Modify_AfterUpdate()
AddStatus FileNumber, Modify, "Order modifying automatic stay"
End Sub

Private Sub Modify_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Modify)
End Sub

Private Sub Modify_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Modify = Now()
    AddStatus FileNumber, Modify, "Order modifying automatic stay"
End If
End Sub

Private Sub MotionDenied_AfterUpdate()
If Not IsNull(MotionDenied) Then AddStatus FileNumber, MotionDenied, "Motion Denied"
End Sub

Private Sub MotionDenied_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(MotionDenied)
End Sub

Private Sub MotionDenied_DblClick(Cancel As Integer)
MotionDenied = Date
AddStatus FileNumber, MotionDenied, "Motion Denied"
End Sub

Private Sub MotionWithdrawn_AfterUpdate()
If Not IsNull(MotionWithdrawn) Then AddStatus FileNumber, MotionWithdrawn, "Motion Withdrawn"
End Sub

Private Sub MotionWithdrawn_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(MotionWithdrawn)
End Sub

Private Sub MotionWithdrawn_DblClick(Cancel As Integer)
MotionWithdrawn = Date
AddStatus FileNumber, MotionWithdrawn, "Motion Withdrawn"
End Sub

Private Sub NODReferral4_AfterUpdate()
AddStatus FileNumber, NODReferral4, "NOD4 Referral Received"
Call AffDefaultInvoice
End Sub

Private Sub NODReferral4_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    NODReferral4 = Date
    Call NODReferral4_AfterUpdate
End If
End Sub

Private Sub NODReferral5_AfterUpdate()
AddStatus FileNumber, NODReferral5, "NOD5 Referral Received"
Call AffDefaultInvoice
End Sub

Private Sub NODReferral5_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    NODReferral5 = Date
    Call NODReferral5_AfterUpdate
End If
End Sub

Private Sub NODReferral6_AfterUpdate()
AddStatus FileNumber, NODReferral6, "NOD6 Referral Received"
Call AffDefaultInvoice
End Sub

Private Sub NODReferral6_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    NODReferral6 = Date
    Call NODReferral6_AfterUpdate
End If
End Sub

Private Sub NoticeTerminating_AfterUpdate()
AddStatus FileNumber, NoticeTerminating, "Notice Terminating"
'  ***Move this one to the Fees table reference***
'AddInvoiceItem FileNumber, "BK-AOD", "Notice of Termination Filed", 50, 0, True, True, False, False
End Sub

Private Sub NoticeTerminating_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(NoticeTerminating)
End Sub

Private Sub NoticeTerminating_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
NoticeTerminating = Date
'Call NoticeTerminating_AfterUpdate
End If
End Sub

Private Sub ObjectionDeadline_AfterUpdate()
AddStatus FileNumber, Date, "Deadline for objections is " & ObjectionDeadline
End Sub

Private Sub ODC_AfterUpdate()
AddStatus FileNumber, ODC, "Sent notice of motion"
End Sub

Private Sub ODC_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ODC)
End Sub

Private Sub ODC_CoDebtor_AfterUpdate()
AddStatus FileNumber, ODC_CoDebtor, "Sent CoDebtor notice of motion"
End Sub

Private Sub ODC_CoDebtor_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ODC_CoDebtor)
End Sub

Private Sub ODC_CoDebtor_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ODC_CoDebtor = Date
    Call ODC_CoDebtor_AfterUpdate
End If
End Sub

Private Sub ODC_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ODC = Now()
    AddStatus FileNumber, ODC, "Sent notice of motion"
End If
End Sub

Private Sub optRealEstate_AfterUpdate()
Call SetPropertyType
End Sub

Private Sub PlanConfHearing_AfterUpdate()
AddStatus FileNumber, Date, "Confirmation hearing scheduled for " & Format$(PlanConfHearing, "m/d/yyyy")
End Sub

Private Sub PlanConfHearing_BeforeUpdate(Cancel As Integer)

If Not IsNull(PlanConfHearing) Then

    If HearingCheking(PlanConfHearing, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(PlanConfHearing, 2) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(PlanConfHearing, 3) = 1 Then
    Cancel = 1
    End If
    
End If

End Sub

Private Sub PlanConfirmDate_AfterUpdate()
AddStatus FileNumber, PlanConfirmDate, "Plan Confirmation Date set"
End Sub

Private Sub PlanConfirmDate_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(PlanConfirmDate)

End Sub

Private Sub PlanConfirmDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
PlanConfirmDate = Date
Call PlanConfirmDate_AfterUpdate
End If

End Sub

Private Sub PlanNotifyClient_AfterUpdate()
AddStatus FileNumber, PlanNotifyClient, "Plan Notify Client 6 months confirmation date set"
End Sub

Private Sub PlanNotifyClient_BeforeUpdate(Cancel As Integer)

If Not IsNull(PlanNotifyClient) Then
    If DateDiff("d", PlanConfirmDate, PlanNotifyClient) < 60 Then
        Cancel = 1
        MsgBox "Cannot enter date prior to 6 months of confirmation date.  See your manager.", vbCritical
    End If
End If
End Sub

Private Sub PlanNotifyClient_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
PlanNotifyClient = Date
Call PlanNotifyClient_AfterUpdate
End If


End Sub

Private Sub PlanObjDeadline_AfterUpdate()
AddStatus FileNumber, Date, "Plan Objection Deadline is " & Format$(PlanObjDeadline, "m/d/yyyy")
End Sub

Private Sub PlanObjFiled_AfterUpdate()
AddStatus FileNumber, PlanObjFiled, "Filed objection to plan"

End Sub

Private Sub PlanObjFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(PlanObjFiled)
End Sub

Private Sub PlanObjFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
PlanObjFiled = Date
Call PlanObjFiled_AfterUpdate
End If


End Sub

Private Sub PlanObjReceived_AfterUpdate()

If Not IsNull(PlanObjReceived) Then
AddStatus FileNumber, PlanObjReceived, "Plan Objection Received"
'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
    FeeAmount = Nz(DLookup("PlanObjFee", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-PR", "BK/Plan-ObjRec'd", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-PR", "BK/Plan-ObjRec'd", 1, 0, True, True, False, False
    End If
End If
End Sub

Private Sub PlanObjReceived_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(PlanObjReceived)
End Sub

Private Sub PlanObjReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
PlanObjReceived = Date
Call PlanObjReceived_AfterUpdate
End If

End Sub

Private Sub PlanObjStatus_AfterUpdate()
  If (Not IsNull(PlanObjStatus)) Then
   PlanObjStatusDate = Date
   Else
   PlanObjStatusDate = Null
   
  End If
  If Len(PlanObjStatus) = 0 Then PlanObjStatusDate = Null
  AddStatus FileNumber, Date, "Objection status: " & PlanObjStatus.Column(1)
End Sub

Private Sub PlanReferralRecd_AfterUpdate()
If Not IsNull(PlanReferralRecd) Then
AddStatus FileNumber, PlanReferralRecd, "Plan Referral Received"
'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
    FeeAmount = Nz(DLookup("PlanRecdFee", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-PR", "BK/Plan-Referral", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-PR", "BK/Plan-Referral", 1, 0, True, True, False, False
    End If
End If
End Sub

Private Sub PlanReferralRecd_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(PlanReferralRecd)
End Sub

Private Sub PlanReferralRecd_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
PlanReferralRecd = Date
Call PlanReferralRecd_AfterUpdate
End If


End Sub

Private Sub PlanReviewed_AfterUpdate()
AddStatus FileNumber, PlanReviewed, "Reviewed plan of reorganization"
End Sub


Private Sub PlanReviewed_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(PlanReviewed)
End Sub

Private Sub PlanReviewed_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
PlanReviewed = Date
Call PlanReviewed_AfterUpdate
End If

End Sub

Private Sub POC_AfterUpdate()
AddStatus FileNumber, POC, "Proof of claim filed"


'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
    FeeAmount = Nz(DLookup("POCfiledfee", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-POC", "Proof of Claim Filed", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-POC", "Proof of Claim Filed", 1, 0, True, True, False, False
    End If
    FeeAmount = Nz(DLookup("TtlUpdate", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-POC", "Title Update / Chain of title search", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-POC", "Title Update / Chain of title search", 1, 0, True, True, False, False
    End If
    FeeAmount = Nz(DLookup("Assignmentfee", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-POC", "Assignment Filing Fee", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-POC", "Assignment Filing Fee", 1, 0, True, True, False, False
    End If



End Sub

Private Sub POC_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(POC)
End Sub

Private Sub POC_DblClick(Cancel As Integer)
POC = Now()
Call POC_AfterUpdate
End Sub

Private Sub POCObj_AfterUpdate()
AddStatus FileNumber, POCObj, "Objection to Proof of Claim"
'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)

            FeeAmount = Nz(DLookup("POCobjFee", "ClientList", "ClientID=" & ClientID))
            If FeeAmount > 0 Then
                AddInvoiceItem FileNumber, "BK-POC", "BK/POC-Objection", FeeAmount, 0, True, True, False, False
            Else
                AddInvoiceItem FileNumber, "BK-POC", "BK/POC-Objection", 1, 0, True, True, False, False
            End If
               
End Sub

Private Sub POCObj_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(POCObj)
End Sub

Private Sub POCObj_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
POCObj = Date
Call POCObj_AfterUpdate
End If
End Sub

Private Sub POCObjHearing_AfterUpdate()
AddStatus FileNumber, Date, "POC Objection Hearing scheduled for " & POCObjHearing
End Sub

Private Sub POCObjHearing_BeforeUpdate(Cancel As Integer)

If Not IsNull(POCObjHearing) Then

    If HearingCheking(POCObjHearing, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(POCObjHearing, 2) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(POCObjHearing, 3) = 1 Then
    Cancel = 1
    End If
End If

End Sub

Private Sub POCReceived_AfterUpdate()
AddStatus FileNumber, POCReceived, "Proof of claim received"
'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
    FeeAmount = Nz(DLookup("POCrecdFee", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-POC", "BK/POC-Referral", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-POC", "BK/POC-Referral", 1, 0, True, True, False, False
    End If
End Sub

Private Sub POCReceived_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(POCReceived)
End Sub

Private Sub POCReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
POCReceived = Date
Call POCReceived_AfterUpdate
End If
End Sub

Private Sub POCRespFiled_AfterUpdate()
AddStatus FileNumber, POCRespFiled, "Response to Objection to Proof of Claim"
End Sub

Private Sub POCRespFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(POCRespFiled)
End Sub

Private Sub POCRespFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
POCRespFiled = Date
AddStatus FileNumber, POCRespFiled, "Response to Objection to Proof of Claim"
End If
End Sub

Private Sub POCStatus_AfterUpdate()
 If (Not IsNull(POCStatus)) Then
    POCStatusDate = Date
    Else
    POCStatusDate = Null
    End If
  If Len(POCStatus) = 0 Then POCStatusDate = Null
  AddStatus FileNumber, Date, "POC Status of Objection  " & POCStatus.Column(1)
End Sub

Private Sub POCStatus_Click()
'  If (Not IsNull(Me.POCStatus)) Then
'    Me.POCStatusDate = Date
'  End If
End Sub

Private Sub PosPetiFeesCost_Click()
[Post-Petition Pay History].Visible = False
[Pre-Petition Pay History].Visible = False
[Post-Petition Fee Breakdown Addendum].Visible = True
[Post Petition Taxes-Insurance Advances Addendum].Visible = False

End Sub

Private Sub PosPetTaxInsuAdvaAdde_Click()
[Post-Petition Pay History].Visible = False
[Pre-Petition Pay History].Visible = False
[Post-Petition Fee Breakdown Addendum].Visible = False
[Post Petition Taxes-Insurance Advances Addendum].Visible = True

End Sub

Private Sub PostPetiPayHistory_Click()
[Post-Petition Pay History].Visible = True
[Pre-Petition Pay History].Visible = False
[Post-Petition Fee Breakdown Addendum].Visible = False
[Post Petition Taxes-Insurance Advances Addendum].Visible = False

End Sub

Private Sub PrePetiPayHis_Click()
[Pre-Petition Pay History].Visible = True
[Post-Petition Pay History].Visible = False
[Post-Petition Fee Breakdown Addendum].Visible = False
[Post Petition Taxes-Insurance Advances Addendum].Visible = False


End Sub

Private Sub ReaffFiled_AfterUpdate()
AddStatus FileNumber, ReaffFiled, "Reaffirmation Filed"

If (Not IsNull(ReaffFiled)) Then
'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
    FeeAmount = Nz(DLookup("ReaffFiled", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-MISC", "Reaffirmation Filed", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-MISC", "Reaffirmation Filed", 1, 0, True, True, False, False
    End If

End If
End Sub

Private Sub ReaffFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ReaffFiled)
End Sub

Private Sub ReaffFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ReaffFiled = Date
Call ReaffFiled_AfterUpdate
End If

End Sub

Private Sub ReaffRecdExecuted_AfterUpdate()
AddStatus FileNumber, ReaffRecdExecuted, "Reaffirmation Received Executed"
'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
    FeeAmount = Nz(DLookup("ReaffExec", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-POC", "BK/Reaff-Executed", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-POC", "BK/Reaff-Executed", 1, 0, True, True, False, False
    End If

End Sub

Private Sub ReaffRecdExecuted_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ReaffRecdExecuted)
End Sub

Private Sub ReaffRecdExecuted_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ReaffRecdExecuted = Date
Call ReaffRecdExecuted_AfterUpdate
End If

End Sub

Private Sub ReaffReferralRecd_AfterUpdate()
AddStatus FileNumber, ReaffReferralRecd, "Reaffirmation Referral Received"
'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
    FeeAmount = Nz(DLookup("ReaffReferral", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-MISC", "BK/Reaff-Referral", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-MISC", "BK/Reaff-Referral", 1, 0, True, True, False, False
    End If

End Sub

Private Sub ReaffReferralRecd_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ReaffReferralRecd)
End Sub

Private Sub ReaffReferralRecd_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ReaffReferralRecd = Date
Call ReaffReferralRecd_AfterUpdate
End If


End Sub


Private Sub ReaffRetExecuted_AfterUpdate()
AddStatus FileNumber, ReaffRetExecuted, "Reaffirmation Return Executed"

End Sub

Private Sub ReaffRetExecuted_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ReaffRetExecuted)
End Sub

Private Sub ReaffRetExecuted_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ReaffRetExecuted = Date
AddStatus FileNumber, ReaffRetExecuted, "Reaffirmation Return Executed"
End If


End Sub

Private Sub ReaffSentToDD_AfterUpdate()
AddStatus FileNumber, ReaffSentToDD, "Reaffirmation Sent To D/D's Counsel"

If (Not IsNull([ReaffSentToDD])) Then
'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
    FeeAmount = Nz(DLookup("ReaffDDcounsel", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-MISC", "Reaffirmation Sent To Debtor/Debtor''s Counsel", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-MISC", "Reaffirmation Sent To Debtor/Debtor''s Counsel", 1, 0, True, True, False, False
    End If
End If

End Sub

Private Sub ReaffSentToDD_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ReaffSentToDD)
End Sub

Private Sub ReaffSentToDD_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ReaffSentToDD = Date
AddStatus FileNumber, ReaffSentToDD, "Reaffirmation Sent To D/D's Counsel"
Call ReaffSentToDD_AfterUpdate
End If

End Sub

Private Sub ReaffToClient_AfterUpdate()
AddStatus FileNumber, ReaffToClient, "Reaffirmation To Client For Review and Execute"
End Sub

Private Sub ReaffToClient_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ReaffToClient)
End Sub

Private Sub ReaffToClient_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ReaffToClient = Date
AddStatus FileNumber, ReaffToClient, "Reaffirmation To Client For Review and Execute"
End If


End Sub

Private Sub ShowCurrent_Click()
If ShowCurrent Then
    Me.Filter = "FileNumber = " & Me![FileNumber] & "AND Current = True"
Else
    Me.Filter = "FileNumber = " & Me![FileNumber]
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

Private Sub cmdPrint_Click()
On Error GoTo Err_cmdPrint_Click

If IsNull(District) Then
    MsgBox "You must select a district before you can print.", vbCritical
    Exit Sub
End If
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm "BankruptcyPrint", , , "BankruptcyID=" & Me!BankruptcyID

Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrint_Click
    
End Sub

Private Sub Terminating_AfterUpdate()
AddStatus FileNumber, Terminating, "Order terminating automatic stay"
End Sub

Private Sub Terminating_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Terminating)
End Sub

Private Sub Terminating_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Terminating = Now()
    AddStatus FileNumber, Terminating, "Order terminating automatic stay"
End If
End Sub

Private Sub cmdNewBankruptcy_Click()

On Error GoTo Err_cmdNewBankruptcy_Click
If MsgBox("Are you sure you want to add another Bankruptcy?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub

Forms![Case List]!ReferralDate = Date
Forms![Case List]!ReferralDocsReceived = Null

Me.AllowAdditions = True
DoCmd.GoToRecord , , acNewRec
Me.AllowAdditions = False

Exit_cmdNewBankruptcy_Click:
    Exit Sub

Err_cmdNewBankruptcy_Click:
    MsgBox Err.Description
    Resume Exit_cmdNewBankruptcy_Click
    
End Sub

Private Sub SetPropertyType()
If optRealEstate Then
    PropertyDesc.Enabled = False
    PropertyContract.Enabled = False
Else
    PropertyDesc.Enabled = True
    PropertyContract.Enabled = True
End If
End Sub

Private Sub cmdSelectFile_Click()

On Error GoTo Err_cmdSelectFile_Click

DoCmd.Close
DoCmd.OpenForm "Select File"

Exit_cmdSelectFile_Click:
    Exit Sub

Err_cmdSelectFile_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectFile_Click
    
End Sub

Private Sub PlanEnabled()
pgPlan.Enabled = Nz((Chapter = 11 Or Chapter = 13))
End Sub

Private Sub cmdNew362_Click()

On Error GoTo Err_cmdNew362_Click
Label190.Visible = True
'If MsgBox("Really do a new 362?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub

'Dim rstbk As Recordset
'Set rstbk = CurrentDb.OpenRecordset("Select * FROM rqryNeedToInoiceBanruptcyChecking where Filenumber=" & FileNumber & " And ( Disposition ='362' or Disposition ='362/SR' ) ", dbOpenDynaset, dbSeeChanges)
    'If Not rstbk.EOF Then
    'MsgBox (" Please notice that the 362 Invoice still did not invoiced yet")
    'Exit Sub
     'Else


            'AddStatus FileNumber, Now(), "New 362"
            '[362] = Null
            '[362].Locked = False
            '[362].BackStyle = 1
            '[362Referral] = Null
            '[ObjectionDeadline] = Null
            '[Hearing] = Null
            '[HearingCalendarEntryID] = Null
            'ODC = Null
            'Modify = Null
            'Terminating = Null
            'Affidavit = Null
            '[2ndAff] = Null
            '[3rdAff] = Null
            'Default1Cured = Null
            'Default2Cured = Null
            'Default3Cured = Null
            'NoticeTerminating = Null
            '[362Disposition] = Null
            '[362StatusDate] = Null

            '[362].SetFocus
            'cmdNew362.Enabled = False


    'End If
'Set rstbk = Nothing

'dim rstbk As Recordset
'Dim cntInv As Integer

Set rstbk = CurrentDb.OpenRecordset("Select * FROM rqryNeedToInoiceBanruptcyChecking where Filenumber=" & FileNumber & " And ( Disposition ='362' or Disposition ='362/SR' ) ", dbOpenDynaset, dbSeeChanges)
If Not rstbk.BOF Then
    rstbk.MoveLast
    cntInv = rstbk.RecordCount
Else
    cntInv = 0
End If

If cntInv = 2 Then
    MsgBox ("Please notice that " & cntInv & " of 362 Invoice still did not invoiced yet. Stop add new 362")
Exit Sub
End If
                        
                        
If cntInv <= 1 Then
    'If Not rstbk.EOF Then
                                    
    'If MsgBox("Please notice that there is " & cntInv & " of 362 Invoice still did not invoiced yet. Really do a new 362?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
    If MsgBox("Are you sure to add a new 362?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
                                  
                    
    If Not IsNull([362Referral]) And Not IsNull(txt362StatusDate) And IsNull([362StatusDate1]) And IsNull([362Referral1]) Then
        Me.[362Referral].Visible = False
        Me.txt362StatusDate.Visible = False
        Me.[362StatusDate1].Visible = True
        Me.[362Referral1].Visible = True
    End If
                                        
        AddStatus FileNumber, Now(), "New 362"
            [362] = Null
            [362].Locked = False
            [362].BackStyle = 1
            If Me.[362Referral].Visible = False Then
                [362Referral1] = Null
            Else
                [362Referral] = Null
            End If
            [ObjectionDeadline] = Null
            [Hearing] = Null
            [HearingCalendarEntryID] = Null
            ODC = Null
            Modify = Null
            Terminating = Null
            Affidavit = Null
            [2ndAff] = Null
            [3rdAff] = Null
            Default1Cured = Null
            Default2Cured = Null
            Default3Cured = Null
            NoticeTerminating = Null
            [362Disposition] = Null
                                                
        If Me.txt362StatusDate.Visible = False Then
           [362StatusDate1] = Null
        Else
           txt362StatusDate = Null
        End If
            [362].SetFocus
            cmdNew362.Enabled = False
    End If
Set rstbk = Nothing
                    
 cntInv = 0



Exit_cmdNew362_Click:
    Exit Sub

Err_cmdNew362_Click:
    MsgBox Err.Description
    Resume Exit_cmdNew362_Click
    
End Sub

Private Sub AssignmentVisuals()

AssignByDOT.Enabled = (Nz(AssignBy) = 1)
AssignByNote.Enabled = (Nz(AssignBy) = 2)
If (Nz(AssignBy) = 1 And Nz(AssignByDOT) = 2) Then
    MergerInfo.Enabled = True
    If IsNull(MergerInfo) Then MergerInfo = DLookup("MergerInfo", "ClientList", "ClientID=" & Forms![Case List]!ClientID)
Else
    MergerInfo.Enabled = False
End If
End Sub

Private Sub cmdAmmendedPOC_Click()

On Error GoTo Err_cmdAmmendedPOC_Click
If MsgBox("Really an ammended Proof of Claim?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub

AddStatus FileNumber, Now(), "Amended Proof of Claim"
POCReceived = Date
Call POCReceived_AfterUpdate
POC = Null
POCObj = Null
POCRespFiled = Null
POCStatus = Null


cbxPOCDisposition = Null
txtPOCDispositionDate = Null

Exit_cmdAmmendedPOC_Click:
    Exit Sub

Err_cmdAmmendedPOC_Click:
    MsgBox Err.Description
    Resume Exit_cmdAmmendedPOC_Click
    
End Sub

Private Function UpdateCalendar(calendarDateOldValue As Variant, calendarDate As Variant, calendarID As String, HearingType As String) As Variant

'UpdateCalendar = Null

'Exit Function

Dim emailGroup As String

UpdateCalendar = calendarID
' If existing date changed but we don't know the EntryID then user must update calendar manually
If (Not IsNull(calendarDateOldValue) And calendarID = "") Then
    MsgBox "Please update the Shared Calendar", vbExclamation
    Exit Function
End If

If (IsNull(calendarDate) And calendarID <> "") Then
    Call DeleteCalendarEvent(calendarID)
    UpdateCalendar = Null
    Exit Function
End If

emailGroup = "SharedCalRecipBK"
If (calendarID = "") Then     ' new event on calendar
    UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")" & " " & HearingType, "USBC - " & District.Column(1), 2, emailGroup)
Else                                    ' change existing event on calendar

   If (IsNull(calendarDateOldValue)) Then   ' new date
    UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")" & " " & HearingType, "USBC - " & District.Column(1), 2, emailGroup)
      
   Else
   '(DateDiff("d", calendarDateOldValue, calendarDate) > 0)  ' date in the future - create new calendar event
    UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")" & " " & HearingType, "USBC - " & District.Column(1), 2, emailGroup)
   'Else ' otherwise update calendar event
    'Call UpdateCalendarEvent(calendarID, CDate(calendarDate), False, DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![case list]!PrimaryDefName & " " & HearingType, "USBC - " & District.Column(1), 2)
   End If
End If
    
End Function

Private Sub Text255_DblClick(Cancel As Integer)

End Sub

Private Sub TitleReviewOrdered_AfterUpdate()
  AddStatus FileNumber, Me.TitleReviewOrdered, "Title Review Ordered"
  AddInvoiceItem FileNumber, "BK-POC", "Title Review", GetFeeAmount("Title Review"), 0, False, True, False, True
End Sub

Private Sub TitleReviewOrdered_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleReviewOrdered)
End Sub

Private Sub TitleReviewOrdered_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
  Me.TitleReviewOrdered = Date
  Call TitleReviewOrdered_AfterUpdate
End If
End Sub

Private Sub TitleReviewReceived_AfterUpdate()
  AddStatus FileNumber, Me.TitleReviewReceived, "Title Review Received"
End Sub

Private Sub TitleReviewReceived_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleReviewReceived)
End Sub

Private Sub TitleReviewReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
  Me.TitleReviewReceived = Date
  Call TitleReviewReceived_AfterUpdate
End If
End Sub


Private Sub TtlPayChgLtrReferral_AfterUpdate()
If Not IsNull(TtlPayChgLtrReferral) Then
AddStatus FileNumber, TtlPayChgLtrReferral, "Title Pay Change Letter Referral Received"
'New Fee Procedure
Dim ClientID As Integer
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
    FeeAmount = Nz(DLookup("TtlPayReferralFee", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "BK-PR", "BK/PayChange-Rec'd", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "BK-PR", "BK/PayChange-Rec'd", 1, 0, True, True, False, False
    End If
Else
AddStatus FileNumber, Date, "Removed Title Pay Change Letter Referral"
End If

End Sub

Private Sub TtlPayChgLtrReferral_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TtlPayChgLtrReferral)
End Sub

Private Sub TtlPayChgLtrReferral_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
TtlPayChgLtrReferral = Date
Call TtlPayChgLtrReferral_AfterUpdate
End If

End Sub

Private Sub TtlPayChgLtrSent_AfterUpdate()
If Not IsNull(TtlPayChgLtrSent) Then
AddStatus FileNumber, TtlPayChgLtrSent, "Pay Change Disposition Set"
Else
AddStatus FileNumber, TtlPayChgLtrSent, "Pay Change Disposition Removed"
End If

End Sub

Private Sub TtlPayChgLtrSent_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TtlPayChgLtrSent)
End Sub

Private Sub TtlPayChgLtrSent_DblClick(Cancel As Integer)
'TtlPayChgLtrSent = Date
'AddStatus FileNumber, TtlPayChgLtrSent, "Pay Change Disposition Set"
End Sub
Private Sub cmdFCClick_Click()
On Error GoTo Err_cmdNewFC_Click
'If MsgBox("Really do a new Fee and Cost?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
Me.txtNoticeDispositionDate = Null
Me.cbxNoticeDisposition = Null
txtFCReferral = Null
AddStatus FileNumber, txtFCReferral, "Removed Notice of Fees and Costs Referral"
txtFCFiled = Null
 'AddStatus FileNumber, txtFCFiled, "Removed Notice of Fees and Costs Referral"

'Call txtFCReferral_Change

Exit_cmdNewFC_Click:
    Exit Sub

Err_cmdNewFC_Click:
    MsgBox Err.Description
    Resume Exit_cmdNewFC_Click
    
End Sub

Private Sub txt362StatusDate_AfterUpdate()
'AddStatus FileNumber, Now(), "Disposition Status Entered (" & Combo315.Column(1) & ")"

End Sub

Private Sub txtEffectiveDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else

Me.txtEffectiveDate = Date
End If


End Sub

Private Sub txtFCFiled_AfterUpdate()
If Not IsNull(txtFCFiled) Then
 AddStatus FileNumber, txtFCFiled, "Filed Notice of Fees and Costs Referral"
 Else
  AddStatus FileNumber, txtFCFiled, "Removed Notice of Fees and Costs Referral"
End If

End Sub

Private Sub txtFCFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    txtFCFiled = Date
End If

End Sub

Private Sub txtFCReferral_AfterUpdate()
  If Not IsNull(txtFCReferral) Then
  ' 2012.02.27 DaveW
    Dim fCost  As Currency
    Dim iCount As Integer
    Dim s As String
    fCost = 0
    'If ckProSe Then

        iCount = 1 ' There is always a debtor.
        s = GetNames(0, 1, "BKCoDebtor=True")
        If "" <> s Then iCount = iCount + 1
        fCost = fCost + iCount * Nz(DLookup("IValue", "DB", "Name='PostageRegular'")) / 100
    'End If
    txtFCFiled = ""
    AddStatus FileNumber, txtFCReferral, "Received Notice of Fees and Costs Referral"
    AddInvoiceItem FileNumber, "BK-NFC", "NFC Referral Received", 50, 0, True, True, False, True
    ' Attorney fee for Fees and Costs
    AddInvoiceItem FileNumber, "BK-NFC", "NFC Postage", fCost, 76, False, True, False, True
    
Else
  AddStatus FileNumber, txtFCReferral, "Removed Notice of Fees and Costs Referral"
End If

End Sub

Private Sub txtFCReferral_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    txtFCReferral = Date
    Call txtFCReferral_Change
End If

End Sub
Private Sub txtFCReferral_Change()
'    ' 2012.02.27 DaveW
'    Dim fCost  As Currency
'    Dim iCount As Integer
'    Dim S As String
'    fCost = 0
'    'If ckProSe Then
'        iCount = 1 ' There is always a debtor.
'        S = GetNames(0, 1, "BKCoDebtor=True")
'        If "" <> S Then iCount = iCount + 1
'        fCost = fCost + iCount * Nz(DLookup("IValue", "DB", "Name='PostageRegular'")) / 100
'    'End If
'    txtFCFiled = ""
'    AddStatus FileNumber, txtFCReferral, "Received Notice of Fees and Costs Referral"
'    AddInvoiceItem FileNumber, "BK-NFC", "NFC Referral Received", 50, 0, True, True, False, True
'    ' Attorney fee for Fees and Costs
'    AddInvoiceItem FileNumber, "BK-NFC", "NFC Postage", fCost, 76, False, True, False, True
End Sub



Private Sub txtNOFCHearingDate_AfterUpdate()
AddStatus FileNumber, txtNOFCHearingDate, "NOFC Hearing Date Set"
End Sub

Private Sub txtNOFCHearingDate_BeforeUpdate(Cancel As Integer)

If Not IsNull(txtNOFCHearingDate) Then
    If (txtNOFCHearingDate < Date) Then
        Cancel = -1
        MsgBox "Hearing Date cannot be in the past.", vbCritical
        Exit Sub
    End If

Dim dteTimePortion As Date
dteTimePortion = TimeValue(txtNOFCHearingDate)

If Hour(dteTimePortion) < 8 Or Hour(dteTimePortion) > 18 Or (Hour(dteTimePortion) = 18 And Minute(dteTimePortion) > 0) Then
    Cancel = -1
    MsgBox "Hearing time must be between 8:00 AM and 6:00 PM"
End If

If HearingCheking(txtNOFCHearingDate, 2) = 1 Then
    Cancel = 1
End If
End If
End Sub

Private Sub txtNOFCHearingDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Me.txtNOFCHearingDate = Now()
    Call txtNOFCHearingDate_AfterUpdate
End If
End Sub

Private Sub txtNOFCObjectionFiled_AfterUpdate()
AddStatus FileNumber, txtNOFCObjectionFiled, "NOFC Objection Filed"

'3/30/15 adding NOFC Objection fee when Tr Objection Filed by

If Not IsNull(Me.NOFCObjectionFiled) Or Me.NOFCObjectionFiled <> "" Then
AddInvoiceItem FileNumber, "BK_NOFC", "NOFC Objection Fee", Nz(DLookup("NOFObjCFee", "ClientList", "ClientID=" & Forms![Case List]!ClientID)), 0, True, True, False, False
End If

End Sub

Private Sub txtNOFCObjectionFiled_BeforeUpdate(Cancel As Integer)
'Cancel = CheckFutureDate(txtNOFCObjectionFiled)
End Sub

Private Sub txtNOFCObjectionFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Me.txtNOFCObjectionFiled = Now()
    Call txtNOFCObjectionFiled_AfterUpdate
End If
End Sub

Private Sub txtNOFCReceived_AfterUpdate()
AddStatus FileNumber, txtNOFCReceived, "NOFC Received"

'3/30/15 adding NOFC fee when NOFC is Received
If Not IsNull(Me.NOFCReceived) Or Me.NOFCReceived <> "" Then
    AddInvoiceItem FileNumber, "BK_NOFC", "NOFC Fee", Nz(DLookup("NOFCFee", "ClientList", "ClientID=" & Forms![Case List]!ClientID)), 0, True, True, False, False
End If
End Sub

Private Sub txtNOFCReceived_BeforeUpdate(Cancel As Integer)
'Cancel = CheckFutureDate(txtNOFCReceived)
End Sub

Private Sub txtNOFCReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Me.txtNOFCReceived = Date
    Call txtNOFCReceived_AfterUpdate
End If
End Sub

Private Sub txtTOCFiled_Change()
    AddStatus FileNumber, txtTOCFiled, "Filed POC Transfer of Claim"
    'AddInvoiceItem FileNumber, "BK-NFC", "Referral Filed", 50, False, True, False, True
End Sub


Private Sub txtTOCFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    txtTOCFiled = Date
    Call txtTOCFiled_Change
End If

End Sub

Private Sub txtTOCReferral_Change()
    ' 2012.02.27 DaveW
    Dim fCost  As Currency
    Dim iCount As Integer
    Dim s As String
    fCost = 0
    'If ckProSe Then
        iCount = 1 ' There is always a debtor.
        s = GetNames(0, 1, "BKCoDebtor=True")
        If "" <> s Then iCount = iCount + 1
        fCost = fCost + iCount * Nz(DLookup("IValue", "DB", "Name='PostageRegular'")) / 100
    'End If
    txtFCFiled = ""
    AddStatus FileNumber, txtFCReferral, "Received Transfer of Claim Referral"
    AddInvoiceItem FileNumber, "BK-TOC", "TOC Referral Received", 75, 0, True, True, False, True
    ' Attorney fee for Fees and Costs
    AddInvoiceItem FileNumber, "BK-toc", "TOC Postage", fCost, 0, False, True, False, True
End Sub

Private Sub txtTOCReferral_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
txtTOCReferral = Date
Call txtTOCReferral_Change
End If
End Sub

Private Sub VAOrderTerminating_AfterUpdate()
AddStatus FileNumber, Me.VAOrderTerminating, "VA Order Terminating"
If Not IsNull(VAOrderTerminating) Then
If Me.District = 4 Or Me.District = 5 Then
AddInvoiceItem FileNumber, "BK/NOD-VAOrderTerm", "NOD VA Order Terminating", Nz(DLookup("VAOrder", "ClientList", "ClientID=" & Forms![Case List]!ClientID)), 0, True, True, False, False
End If
End If
End Sub

Private Sub VAOrderTerminating_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(VAOrderTerminating)
End Sub

Private Sub VAOrderTerminating_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Me.VAOrderTerminating = Date
    Call VAOrderTerminating_AfterUpdate
End If

End Sub


Private Sub SetDisposition(DispositionID As Long)
Dim StatusText As String

If DispositionID = 0 Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    SelectedDispositionID = 0
    DoCmd.OpenForm "SetDisposition", , , , , acDialog, "CD"
Else
    Disposition = DispositionID
End If

'This code will never run, not sure why it is in place
If SelectedDispositionID > 0 Then   ' if it was actually set
    cmdClose.SetFocus
    cmdSetDisposition.Enabled = False ' don't allow any changes
    
    CDDisposition = SelectedDispositionID
    CDDisposition.Requery
    Me.CDDispositionDate = Date
    
    
    If StaffID = 0 Then Call GetLoginName
    CDDispositionStaffID = StaffID
    DoCmd.RunCommand acCmdSaveRecord
    
    CDDispositionDesc.Requery
    CDDispositionInitials.Requery
    
    
    StatusText = "CD Disposition " & CDDispositionDesc
    AddStatus FileNumber, Now(), StatusText
    
    AddInvoiceItem FileNumber, "BK-MISC", "Cramdown", 350, 0, True, True, False, True

    
End If

End Sub





'Private Sub TitleAssignReceivedDate_AfterUpdate()
'If (Not IsNull(TitleAssignReceivedDate)) Then
' AddStatus FileNumber, TitleAssignReceivedDate, "Assignment Received from Client"
'End If
'
'End Sub
'
'Private Sub TitleAssignReceivedDate_DblClick(Cancel As Integer)
'TitleAssignReceivedDate = Now()
'Call TitleAssignReceivedDate_AfterUpdate
'End Sub
'
'Private Sub TitleAssignRecordedDate_AfterUpdate()
'If (Not IsNull(TitleAssignRecordedDate)) Then
'  AddStatus FileNumber, TitleAssignRecordedDate, "Assignment Sent to Record"
'
'End If
'
'
'End Sub
'
'Private Sub TitleAssignRecordedDate_DblClick(Cancel As Integer)
'TitleAssignRecordedDate = Now()
'Call TitleAssignRecordedDate_AfterUpdate
'End Sub
'
'Private Sub TitleAssignSentdate_AfterUpdate()
'If (Not IsNull(TitleAssignSentdate)) Then
' AddStatus FileNumber, TitleAssignSentdate, "Assignment Requested From Client"
'End If
'End Sub
'
'Private Sub TitleAssignSentdate_DblClick(Cancel As Integer)
'TitleAssignSentdate = Now()
'Call TitleAssignSentdate_AfterUpdate
'End Sub


