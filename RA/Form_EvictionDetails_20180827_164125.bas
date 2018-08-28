VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EvictionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub AppealExpires_AfterUpdate()
AddStatus FileNumber, Now(), "Appeal expires " & Format$(AppealExpires, "m/d/yyyy")
End Sub

Private Sub ShowBrokerInfo()
If IsNull(BrokerID) Then
    lblBrokerPhone.Caption = ""
    lblBrokerExt.Caption = ""
    lblBrokerEMail.Caption = ""
    
    lblBrokerEMail.HyperlinkAddress = ""
Else
    lblBrokerPhone.Caption = Format(Nz(DLookup("BrokerPhone", "Brokers", "BrokerID=" & [BrokerID])), "(###) ###-####")
    lblBrokerExt.Caption = Nz(DLookup("BrokerExt", "Brokers", "BrokerID=" & [BrokerID]))
    lblBrokerEMail.Caption = Nz(DLookup("BrokerEMail", "Brokers", "BrokerID=" & [BrokerID]))
    lblBrokerEMail.HyperlinkAddress = "mailto:" & lblBrokerEMail.Caption & "?Subject=" & sfrmForeclosure!PropertyAddress
End If
End Sub

Private Sub ShowClientContactInfo()
If IsNull(ClientContactID) Then
    lblClientPhone.Caption = ""
    lblClientEmail.Caption = ""
    lblClientEmail.HyperlinkAddress = ""
Else
    lblClientPhone.Caption = Format(Nz(DLookup("Phone", "ClientContacts", "ID=" & [ClientContactID])), "(###) ###-####") & " - " & Nz(DLookup("Ext", "ClientContacts", "ID=" & [ClientContactID]))
    lblClientEmail.Caption = Nz(DLookup("EMail", "ClientContacts", "ID=" & [ClientContactID]))
    lblClientEmail.HyperlinkAddress = "mailto:" & lblClientEmail.Caption & "?Subject=" & sfrmForeclosure!PropertyAddress
End If
End Sub
Private Sub ShowEVContactInfo()
If IsNull(ClientContactID) Then
    lblClientPhone.Caption = ""
    lblClientExt.Caption = ""
    lblClientEmail.Caption = ""
    lblClientEmail.HyperlinkAddress = ""
Else
    lblClientPhone.Caption = Format(Nz(DLookup("PhoneNumber", "EVContactsByClient", "ID=" & [ClientContactID])), "(###) ###-####")
    lblClientExt.Caption = Nz(DLookup("Extension", "EVContactsByClient", "ID=" & [ClientContactID]))
    lblClientEmail.Caption = Nz(DLookup("Email", "EVContactsByClient", "ID=" & [ClientContactID]))
    lblClientEmail.HyperlinkAddress = "mailto:" & lblClientEmail.Caption & "?Subject=" & sfrmForeclosure!PropertyAddress
End If
End Sub

Private Sub AppealHearingDate_AfterUpdate()
If Not IsNull(AppealHearingDate) Then
AddStatus FileNumber, Now(), "All Appeal Hearing Date & Time  " & AppealHearingDate
End If
End Sub

Private Sub BrokerID_AfterUpdate()
Call ShowBrokerInfo
End Sub

Private Sub CashForKeysDate_AfterUpdate()
'    AddStatus FileNumber, CashForKeysDate, "Cash for Keys Letter sent"
If Not IsNull(CashForKeysDate) Then
    AddStatus FileNumber, CashForKeysDate, "Cash for Keys Letter sent"
Else
    AddStatus FileNumber, Now(), "Removed Cash for Keys Letter sent"
    DoCmd.SetWarnings False
    strinfo = "Removed Cash for Keys Letter sent (" & CashForKeysDate & ") by " & GetFullName
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms![Case List]!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If

End Sub

Private Sub CashForKeysDate_BeforeUpdate(Cancel As Integer)
    Cancel = CheckFutureDate(CashForKeysDate)
End Sub

Private Sub CashForKeysDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    CashForKeysDate = Now()
    AddStatus FileNumber, CashForKeysDate, "Cash for Keys Letter sent"
End If

End Sub

Private Sub CFKCancelled_AfterUpdate()
'   AddStatus FileNumber, CFKCancelled, "Cash for Keys Canceled"
If Not IsNull(CFKCancelled) Then
    AddStatus FileNumber, CFKCancelled, "Cash for Keys Canceled"
Else
    AddStatus FileNumber, Now(), "Removed Cash for Keys Canceled"
    DoCmd.SetWarnings False
    strinfo = "Removed Cash for Keys Canceled (" & CFKCancelled & ") by " & GetFullName
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms![Case List]!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If

End Sub

Private Sub CFKCancelled_BeforeUpdate(Cancel As Integer)
    Cancel = CheckFutureDate(CFKCancelled)
End Sub

Private Sub CFKCancelled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    CFKCancelled = Now()
    AddStatus FileNumber, CFKCancelled, "Cash for Keys Cancelled"
End If

End Sub

Private Sub CFKCheckSent_AfterUpdate()
 '   AddStatus FileNumber, CFKCheckSent, "Cash for Keys Check sent"
If Not IsNull(CFKCheckSent) Then
    AddStatus FileNumber, CFKCheckSent, "Cash for Keys Check sent"
Else
    AddStatus FileNumber, Now(), "Removed Cash for Keys Check sent"
    DoCmd.SetWarnings False
    strinfo = "Removed Cash for Keys Check sent (" & CFKCheckSent & ") by " & GetFullName
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms![Case List]!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery

End If
End Sub

Private Sub CFKCheckSent_BeforeUpdate(Cancel As Integer)
    Cancel = CheckFutureDate(CFKCheckSent)
End Sub

Private Sub CFKCheckSent_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    CFKCheckSent = Now()
    AddStatus FileNumber, CFKCheckSent, "Cash for Keys Check sent"
End If

End Sub

Private Sub CFKCompleted_AfterUpdate()
'    AddStatus FileNumber, CFKCompleted, "Cash for Keys Completed"
If Not IsNull(CFKCompleted) Then
    AddStatus FileNumber, CFKCompleted, "Cash for Keys Completed"
Else
    AddStatus FileNumber, Now(), "Removed Cash for Keys Completed"
    DoCmd.SetWarnings False
    strinfo = "Removed Cash for Keys Completed (" & CFKCompleted & ") by " & GetFullName
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms![Case List]!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If
End Sub

Private Sub CFKCompleted_BeforeUpdate(Cancel As Integer)
    Cancel = CheckFutureDate(CFKCompleted)
End Sub

Private Sub CFKCompleted_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    CFKCompleted = Now()
    AddStatus FileNumber, CFKCompleted, "Cash for Keys Completed"
End If

End Sub

Private Sub CFKExecuted_AfterUpdate()
    'AddStatus FileNumber, CFKExecuted, "Cash for Keys Executed"
If Not IsNull(CFKExecuted) Then
    AddStatus FileNumber, CFKExecuted, "Cash for Keys Executed"
Else
    AddStatus FileNumber, Now(), "Removed Cash for Keys Executed"
    DoCmd.SetWarnings False
    strinfo = "Removed Cash for Keys Executed (" & CFKExecuted & ") by " & GetFullName
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms![Case List]!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If
End Sub

Private Sub CFKExecuted_BeforeUpdate(Cancel As Integer)
    Cancel = CheckFutureDate(CFKExecuted)
End Sub

Private Sub CFKExecuted_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    CFKExecuted = Now()
    AddStatus FileNumber, CFKExecuted, "Cash for Keys Executed"
End If

End Sub

Private Sub CFKVacateDate_AfterUpdate()
  Dim EV As Recordset
   Dim StEvent As String
   Dim EmailSub As String
   Dim Emailboudy As String
   Dim SendingDAte As Date
   
   
   
   
   StEvent = "EveNotice"
   EmailSub = "Reminder : CFK Check should be requested today. Due in 7 Days."
   Emailboudy = " File Nummber:  " & FileNumber & "   Clint: " & ClientShortName(Forms![Case List]!ClientID) & " " & "  Project Name:  " & Forms![Case List]!PrimaryDefName
   SendingDAte = Format(DateAdd("d", -7, [CFKVacateDate]), "mm/dd/yyyy")
   
   
   If Not IsNull(CFKVacateDate) Then
        AddStatus FileNumber, CFKVacateDate, "Cash for Keys Vacate Date set"
        Me.CFKVacateDate.Locked = True
        If Now < CFKVacateDate Then
       '     MsgBox (DateAdd("d", -7, [CFKVacateDate]))
      '  MsgBox (Format(DateAdd("d", -7, [CFKVacateDate]), "dd/mm/yyyy hh:mm:ss"))
            Set EV = CurrentDb.OpenRecordset("SELECT * FROM Sender WHERE [Event] = 'EveNotice'", dbOpenDynaset, dbSeeChanges)
            If Not EV.EOF Then
            
            
               DoCmd.SetWarnings False
               
                 DoCmd.RunSQL ("Insert into ScheduledEmail (Sender1, Sender2, Sender3, FileNumber, EmailSub, EmailBoudy, EmailTo, SendingDate, EntringDate, ActionDate, EventName, Department,SentEmail ) Values(" & EV!Sender1 & _
                 ", " & EV!Sender2 & " , " & EV!Sender3 & " , " & FileNumber & " ,'" & EmailSub & "', '" & Emailboudy & "', 48 , #" & SendingDAte & "#, #" & Now() & "# ,#" & CFKVacateDate & "#,'" & EV!Event & "','" & EV!Department & "', False )")
                 
                  
                 
                                                                                                                                                                        
              DoCmd.SetWarnings True
              End If
              Set EV = Nothing
        
        
    '    Call Email_Reminder(48, FileNumber, "Reminder : CFK Check should be requested today. Due in 7 Days.", " File Nummber:  " & FileNumber & "   Clint: " & ClientShortName(Forms![Case list]!ClientID) & " " & "  Project Name:  " & Forms![Case list]!PrimaryDefName, Format(DateAdd("d", -7, [CFKVacateDate]), "mm/dd/yyyy"))

          '  Call Email_Reminder(48, FileNumber, "Reminder : CFK Check should be requested today. Due in 7 Days.", " File Nummber:  " & FileNumber & "   Clint: " & ClientShortName(Forms![Case List]!ClientID) & " " & "  Project Name:  " & Forms![Case List]!PrimaryDefName, Format(DateAdd("d", -7, [CFKVacateDate]), "dd/mm/yyyy hh:mm:ss"))
        Me.CFKVacateDate.Locked = True
        Else
            MsgBox ("Date must be in the future")
            Me.Undo
        Exit Sub
        End If
    End If
  
    
End Sub





Private Sub ClientContactID_AfterUpdate()
Call ShowEVContactInfo
End Sub

'Private Sub cmdCFKexecuted_Click()
'
'If Not IsNull([CFKExecuted]) Then
'    If MsgBox(" You are about to remove cash for keys executed date ? ", vbOKCancel) = vbOK Then
'
'        AddStatus FileNumber, Now(), "Removed cash for keys executed date (" & [CFKExecuted] & ") by " & GetFullName
'
'        DoCmd.SetWarnings False
'        strInfo = "Removed Cash for Keys Executed date (" & [CFKExecuted] & ") by " & GetFullName
'        strInfo = Replace(strInfo, "'", "''")
'        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms![Case List]!FileNumber,Now,GetFullName(),'" & strInfo & "',1 )"
'        DoCmd.RunSQL strSQLJournal
'        DoCmd.SetWarnings True
'        Forms!Journal.Requery
'
'        Forms!EvictionDetails![CFKExecuted] = Null
'
'    End If
'Else
'    Exit Sub
'End If
'
'
'End Sub

Private Sub cmdCFKVacated_Click()

If Not IsNull([CFKVacateDate]) Then
    If MsgBox(" You are about to remove cash for keys vacated date ? ", vbOKCancel) = vbOK Then
        
        AddStatus FileNumber, Now(), "Removed cash for keys vacated date (" & [CFKVacateDate] & ") by " & GetFullName
    
        DoCmd.SetWarnings False
        strinfo = "Removed Cash for Keys Vacated date (" & [CFKVacateDate] & ") by " & GetFullName
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms![Case List]!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
     
        Forms!EvictionDetails![CFKVacateDate] = Null
        Me.CFKVacateDate.Locked = False
    End If
Else
    Exit Sub
End If

End Sub

Private Sub cmdNewEviction_Click()
On Error GoTo Err_cmdNewEviction_Click
If MsgBox("Are you sure you want to add another Eviction?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub

Forms![Case List]!ReferralDate = Date
Forms![Case List]!ReferralDocsReceived = Null

Me.AllowAdditions = True
DoCmd.GoToRecord , , acNewRec
Me.AllowAdditions = False

AddStatus FileNumber, Now(), "New Eviction Added "

    DoCmd.SetWarnings False
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms![case list]!FileNumber,Now,GetFullName(),'" & "New Eviction Added" & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    
  Forms!Journal.Requery
Exit_cmdNewEviction_Click:
    Exit Sub

Err_cmdNewEviction_Click:
    MsgBox Err.Description
    Resume Exit_cmdNewEviction_Click
End Sub

Private Sub CommdEdit_Click()
If Not CheckNameEdit() Then
Dim ctrl As Control
For Each ctrl In Me.sfrmNames.Form.Controls

If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Then ' TypeOf ctrl Is CommandButton Then
'If Not ctrl.Locked) Then
ctrl.Locked = False
'Else
'ctrl.Locked = True
End If
'End If
Next
With Me.sfrmNames.Form

.AllowAdditions = True
.AllowEdits = True
.AllowDeletions = True
.cmdCopyClient.Enabled = True
.cmdCopy.Enabled = True
.cmdTenant.Enabled = True
.cmdMERS.Enabled = True
.cmdEnterSSN.Enabled = True
.cmdNoNotice.Enabled = True
.cmdPrintNotice.Enabled = True
.cmdPrintLabel.Enabled = True
.cbxNotice.Enabled = True
.cmdDelete.Enabled = True
.cbxNotice.Enabled = True
.cbxNotice.Locked = False
End With
'Exit Sub
'Else

End If
End Sub

Private Sub ComAddName_Click()
DoCmd.OpenForm "sfrmNamesUpdate", , , , acFormAdd

End Sub

Private Sub CmbAppealDisposion_AfterUpdate()
If Nz(Forms![EvictionDetails]!CmbAppealDisposion.Column(1)) = 0 Then
AddStatus FileNumber, Now(), "Removed Hearing Appealed dispostion  "
AppealHearingDisposition = Null
Else
AddStatus FileNumber, Now(), "Add Hearing Appealed dispostion  " & Forms![EvictionDetails]!CmbAppealDisposion.Column(1)
If AppealHearingDate > Date Then
   If Not IsNull(AppealHearingCalendarEntryID) Then
    Call DeleteCalendarEvent(AppealHearingCalendarEntryID)
    AppealHearingCalendarEntryID = Null
    End If
End If
End If

End Sub



Private Sub CommEdit_Click()

DoCmd.OpenForm "sfrmNamesUpdate", , , WhereCondition:="ID= " & Forms!EvictionDetails!sfrmNames!ID

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
End Sub

Private Sub ComplaintServed_AfterUpdate()
AddStatus FileNumber, ComplaintServed, "Complaint Served"
If InPossession = 1 Then
    ResponseDeadlineOwner = DateAdd("d", 18, [ComplaintServed])
Else
    ResponseDeadlineOwner = DateAdd("d", 30, [ComplaintServed])
End If
End Sub

Private Sub ComplaintServed_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ComplaintServed)
End Sub

Private Sub ComplaintServed_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ComplaintServed = Now()
Call ComplaintServed_AfterUpdate
End If

End Sub

Private Sub Consent_AfterUpdate()
AddStatus FileNumber, Consent, "Consent / Motion Granted"
End Sub

Private Sub DocsReceived_AfterUpdate()
AddStatus FileNumber, DocsReceived, "Received Foreclosure Documents from Foreclosure Attorney"
End Sub

Private Sub DocsReceived_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DocsReceived)
End Sub

Private Sub DocsReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
DocsReceived = Date
Call DocsReceived_AfterUpdate
End If

End Sub

Private Sub DocsRequested_AfterUpdate()
AddStatus FileNumber, DocsRequested, "Requested Foreclosure Documents from Foreclosure Attorney"
End Sub

Private Sub DocsRequested_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DocsRequested)
End Sub

Private Sub DocsRequested_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
DocsRequested = Date
Call DocsRequested_AfterUpdate
End If

End Sub

Private Sub EvictionCancelled_AfterUpdate()

If IsNull(EvictionCancelled) Then Exit Sub


  AddStatus FileNumber, EvictionCancelled, "Eviction Cancelled"
  RentClosed = EvictionCancelled
  If (HearingDate > Now) Then
    HearingDate = Null
  End If
  
  If (LockoutDate > Now) Then
    LockoutDate = Null
  End If
End Sub

Private Sub EVTab_Change()

Select Case EVTab.Value
  Case 4
    sfrmRentMoney.Requery
  Case 6
    sfrmStatus.Requery
End Select

End Sub

Private Sub CheckFCAddress()
If IsNull(FCFileNumber) Then
    lblAddrWarning.Visible = False
Else
    lblAddrWarning.Visible = (sfrmForeclosure!PropertyAddress <> DLookup("PropertyAddress", "FCDetails", "Current=True AND FileNumber=" & FCFileNumber))
End If
End Sub

Private Sub FCFileNumber_AfterUpdate()
If Not IsNull(FCFileNumber) Then
    If sfrmForeclosure!LoanNumber <> DLookup("LoanNumber", "FCDetails", "Current=True AND FileNumber=" & FCFileNumber) Then
        FCFileNumber = Null
        MsgBox "The Loan Number in the Foreclosure file must match the Loan Number in the Eviction file.", vbCritical
    End If
End If
Call CheckFCAddress
End Sub

Private Sub FiledBankruptcy_AfterUpdate()
AddStatus FileNumber, FiledBankruptcy, "Filed Bankruptcy"
End Sub

Private Sub FiledBankruptcy_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(FiledBankruptcy)
End Sub

Private Sub FinalNoticeAffadavit_AfterUpdate()
AddStatus FileNumber, FinalNoticeAffadavit, "Final Notice Affadavit"
End Sub

Private Sub FinalNoticeAffadavit_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(FinalNoticeAffadavit)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
If (Nz(Me.HearingDate) <> Nz(Me.HearingDate.OldValue)) Then
  HearingCalendarEntryID = UpdateCalendar(HearingDate.OldValue, HearingDate, Nz(HearingCalendarEntryID), "Eviction Hearing")
End If

If (Nz(Me.LockoutDate) <> Nz(Me.LockoutDate.OldValue)) Then
  LockoutDateCalendarEntryID = UpdateCalendar(LockoutDate.OldValue, LockoutDate, Nz(LockoutDateCalendarEntryID), "LOCKOUT")
End If

If (Nz(Me.AppealHearingDate) <> Nz(Me.AppealHearingDate.OldValue)) Then
  AppealHearingCalendarEntryID = UpdateCalendar(AppealHearingDate.OldValue, AppealHearingDate, Nz(AppealHearingCalendarEntryID), "Eviction Hearing")
End If


End Sub

Private Sub Form_Current()
Me.Caption = "Eviction File " & [FileNumber] & " " & [PrimaryDefName]

Dim EV As Recordset

If FileReadOnly Or EditDispute Then
    Me.AllowEdits = False
    cmdPrint.Enabled = False
'    sfrmPropAddr.Form.AllowEdits = False
    sfrmForeclosure.Form.AllowEdits = False
    sfrmNames.Form.AllowEdits = False
    sfrmNames.Form.AllowAdditions = False
    sfrmNames.Form.AllowDeletions = False
    sfrmNames!cmdCopy.Enabled = False
    sfrmNames!cmdTenant.Enabled = False
    sfrmNames!cmdDelete.Enabled = False
    sfrmNames!cmdNoNotice.Enabled = False
    CommEdit.Enabled = False
    sfrmStatus.Form.AllowEdits = False
    sfrmStatus.Form.AllowAdditions = False
    sfrmStatus.Form.AllowDeletions = False
    Detail.BackColor = ReadOnlyColor
    sfrmForeclosure.Form.Detail.BackColor = ReadOnlyColor
    CommEdit.Enabled = False
    ComAddName.Enabled = False
    
Else
    Me.AllowEdits = True
    cmdPrint.Enabled = True
    'rmPropAddr.Form.AllowEdits = True
    sfrmForeclosure.Form.AllowEdits = True
    
'    If Not CheckNameEdit() Then SA10/05
'    sfrmNames.Form.AllowEdits = False
'    sfrmNames.Form.AllowAdditions = False
'    sfrmNames.Form.AllowDeletions = False
'    sfrmNames!cmdCopy.Enabled = False
'    sfrmNames!cmdTenant.Enabled = False
'    sfrmNames!cmdDelete.Enabled = False
'    sfrmNames!cmdNoNotice.Enabled = False
'    Else
'    sfrmNames.Form.AllowEdits = True
'    sfrmNames.Form.AllowAdditions = True
'    sfrmNames.Form.AllowDeletions = True
'    sfrmNames!cmdCopy.Enabled = True
'    sfrmNames!cmdTenant.Enabled = True
'    sfrmNames!cmdDelete.Enabled = True
'    sfrmNames!cmdNoNotice.Enabled = True
'    End If
'    sfrmStatus.Form.AllowEdits = True
'    sfrmStatus.Form.AllowAdditions = True
'    sfrmStatus.Form.AllowDeletions = True
       
   
    Detail.BackColor = -2147483633
    sfrmForeclosure.Form.Detail.BackColor = -2147483633
End If

If Me.NewRecord Then    ' fill in info from previous EV, if any
    Set EV = CurrentDb.OpenRecordset("SELECT * FROM EVDetails WHERE FileNumber = " & FileNumber & " AND Current = True", dbOpenDynaset, dbSeeChanges)
    If Not EV.EOF Then
        'PrimaryFirstName = fc("PrimaryFirstName")
        'PrimaryLastName = fc("PrimaryLastName")
        
        Do While Not EV.EOF     ' make all previously current records not current
            EV.Edit
            EV("Current") = False
            EV.Update
            EV.MoveNext
        Loop
    End If
    EV.Close
    Me!Current = True           ' and make this record current
End If

Call CheckFCAddress
Call ShowBrokerInfo
Call ShowEVContactInfo
Call SetCashForKeysDate

If Not IsNull(ServicerRelease) Then lblServicer.Visible = True

If (Not IsNull(ReportedVacant) And PrivReportedVacant = False) Then
  Call SetObjectAttributes(ReportedVacant, False)
Else
  Call SetObjectAttributes(ReportedVacant, True)
End If

'UpdateAttorneyList

If Me.State = "VA" Then
'  Attorney.RowSource = "SELECT Staff.ID, Staff.Name & ', ' & [CommonWealthTitle] AS CWRep " & _

  Me.NoticeToOccupant15days.Enabled = False
  Attorney.RowSource = "SELECT Staff.ID, Staff.Name & ' '  " & _
                       "FROM Staff " & _
                       "WHERE  ((staff.active = true) And (Staff.Attorney =True) And(staff.PracticeVA = true )) " & _
                       "ORDER BY Staff.CommonwealthTitle, Staff.Sort;"
                       
 '   "ORDER BY Staff.CommonwealthTitle, Staff.Sort;"
                   'It was  "WHERE (((Staff.CommonwealthTitle) Is Not Null)) and staff.active = true " S.A.
'staff.active=true
ElseIf Me.State = "MD" Then
  Me.NoticeToOccupant15days.Enabled = True
  Attorney.RowSource = "SELECT Staff.ID, Staff.Name & ', Esq.' FROM Staff WHERE ((Staff.active = true ) and (Staff.Attorney = True) and (Staff.PracticeMD = True)) ORDER BY Staff.Sort;"
Else
  Attorney.RowSource = "SELECT Staff.ID, Staff.Name FROM Staff WHERE ((staff.active = true) and (Staff.Attorney = True) and (staff.PracticeDC = true ))ORDER BY Staff.Sort;"

End If

If PrivEvictionCashForKeys Then
 '   Me.cmdCFKexecuted.Visible = True
    Me.cmdCFKVacated.Visible = True
End If

If Not IsNull(CFKVacateDate) = True Then Me.CFKVacateDate.Locked = True
    



End Sub


Public Sub SetCashForKeysDate()

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

'MSH - 7/14/11 - always enable Cash For Keys Date

'If (sfrmForeclosure.Form.LoanType = 4) Then
'  Call SetObjectAttributes(CashForKeysDate, True)
'Else
' Call SetObjectAttributes(CashForKeysDate, False)
'End If


End Sub

Public Sub UpdateAttorneyList()


'If Me.State = "VA" Then
'  Attorney.RowSource = "SELECT Staff.ID, Staff.Name & ', ' & [CommonWealthTitle] AS CWRep " & _
'                       "FROM Staff " & _
'                       "WHERE (((Staff.CommonwealthTitle) Is Not Null)) " & _
'                       "ORDER BY Staff.CommonwealthTitle, Staff.Sort;"
'ElseIf Me.State = "MD" Then
  Attorney.RowSource = "SELECT Staff.ID, Staff.Name & ', Esq.' FROM Staff WHERE (((Staff.Attorney)=True)) ORDER BY Staff.Sort;"
'Else
'  Attorney.RowSource = "SELECT Staff.ID, Staff.Name FROM Staff WHERE (((Staff.Attorney)=True)) ORDER BY Staff.Sort;"

'End If

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

'If IsNull(Attorney) Then
'    MsgBox "Select an attorney who will sign the document(s).", vbCritical
'    EVTab.Value = 0
'    Attorney.SetFocus
'    Exit Sub
'End If

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm "EvictionPrint", , , "[FileNumber]=" & Me![FileNumber]

Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrint_Click
    
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

Private Sub Form_Open(Cancel As Integer)
If FileReadOnly Or EditDispute Then

    Dim ctl As Control
    Dim lngI As Long
    Dim bSkip As Boolean

    For Each ctl In Form.Controls
    Select Case ctl.ControlType
    Case acTextBox, acComboBox, acListBox, acOptionGroup, acCheckBox, acSubform, acOptionButton
        
            If Not (ctl.Locked) Then ctl.Locked = True
            
    Case acCommandButton
        bSkip = False
            If ctl.Name = "cmdClose" Then bSkip = True
            If ctl.Name = "cmdSelectFile" Then bSkip = True
            If Not bSkip Then ctl.Enabled = False
         
       
    End Select
    Next
End If
End Sub

Private Sub HearingDate_AfterUpdate()
AddStatus FileNumber, Now(), "Hearing scheduled for " & Format$(HearingDate, "m/d/yyyy")
End Sub

Private Sub HearingDate_BeforeUpdate(Cancel As Integer)
If Not IsNull(HearingDate) Then

    If HearingCheking(HearingDate, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(HearingDate, 2) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(HearingDate, 3) = 1 Then
    Cancel = 1
    End If

End If
End Sub

Private Sub JudgementGranted_AfterUpdate()

If (HearingDate > Now) Then
  Me.HearingDate = Null
End If

AddStatus FileNumber, JudgementGranted, "Judgment for Possession granted"


'Old EV Invoice Methodology
Select Case State
    Case "DC"
        AddInvoiceItem FileNumber, "EV", "Writ Fee", 145, 0, False, True, False, True

    Case "MD"
        'If MsgBox("Was the Filing Fee paid up-front?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then AddInvoiceItem FileNumber, "EV", "Motion for Judgment of Possession County Up-Front Filing Fee", 25, 0, False, True, False, True
        'If MsgBox("Sheriff Fee?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then AddInvoiceItem FileNumber, "EV", "Sheriff Fee", 40, 0, False, True, False, True
    'added on 11_4_15
            If MsgBox("Was the Filing Fee paid up-front?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            AddInvoiceItem FileNumber, "EV", "Motion for Judgment of Possession County Up-Front Filing Fee", 25, 0, False, True, False, True
            Me.FillingFeePrepaid = False
            If MsgBox("Sheriff Fee?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then AddInvoiceItem FileNumber, "EV", "Sheriff Fee", 40, 0, False, True, False, True

            DoCmd.OpenReport "Eviction Writ and General Notice Cover", acViewPreview
        Else
            Me.FillingFeePrepaid = True
            If MsgBox("Sheriff Fee?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then AddInvoiceItem FileNumber, "EV", "Sheriff Fee", 40, 0, False, True, False, True

            DoCmd.OpenReport "Eviction Writ and General Notice Cover_yes", acViewPreview
        End If
         
        'If MsgBox("Sheriff Fee?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then AddInvoiceItem FileNumber, "EV", "Sheriff Fee", 40, 0, False, True, False, True

    Case "VA"
        AddInvoiceItem FileNumber, "EV", "Sheriff Fee", GetFeeAmount("Enter Sheriff fee"), 0, False, True, False, True
End Select

End Sub

Private Sub JudgementGranted_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(JudgementGranted)
End Sub

Private Sub JudgementGranted_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
JudgementGranted = Now()
Call JudgementGranted_AfterUpdate
End If

End Sub

Private Sub LockoutDate_AfterUpdate()
AddStatus FileNumber, Now(), "Lockout date: " & Format$(LockoutDate, "m/d/yyyy")
End Sub

Private Sub MotionFiled_AfterUpdate()

AddStatus FileNumber, MotionFiled, "Motion for Judgment of Possession filed"
Select Case State
    Case "DC"
        AddInvoiceItem FileNumber, "EV", "Motion for Judgment of Possession Served", 40, 0, False, True, False, False
        AddInvoiceItem FileNumber, "EV", "Motion for Judgment of Possession Courier Fee", 10, 0, False, True, False, False
        AddInvoiceItem FileNumber, "EV", "Complaint Filed: Court Fee", 15, 0, False, True, False, False
        AddInvoiceItem FileNumber, "EV", "Complaint Filed: Overnight Delivery", 10, 77, False, True, False, False
        AddInvoiceItem FileNumber, "EV", "Complaint Filed: Process Service", 65, 0, False, True, False, False

    Case "MD"
        DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter postage for mailing the Motion|EV|Motion for Judgment of Possession mailed"
        AddInvoiceItem FileNumber, "EV", "Motion for Judgment of Possession FedEx fee", 10, 77, False, True, False, True
        If MsgBox("Is there an up-front Filing Fee?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then AddInvoiceItem FileNumber, "EV", "Motion for Judgment of Possession County Up-Front Filing Fee", 31, 0, False, True, False, True
        If MsgBox("Is a Line of Appearance needed?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then AddInvoiceItem FileNumber, "EV", "Motion for Judgment of Possession Line of Appearance", 10, 0, False, True, False, True

    Case "VA"
        DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter postage for mailing the Motion|EV|Motion for Judgment of Possession mailed"
        AddInvoiceItem FileNumber, "EV", "Motion for Judgment of Possession FedEx fee", 10, 77, False, True, False, True
        AddInvoiceItem FileNumber, "EV", "Motion for Judgment of Possession Court fee", GetFeeAmount("Enter Court fee for filing Motion for Judgment of Possession"), 0, False, True, False, True
End Select
End Sub

Private Sub MotionFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(MotionFiled)
End Sub

Private Sub MotionFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
MotionFiled = Now()
Call MotionFiled_AfterUpdate
End If

End Sub

Private Sub NoticeToOccupant_AfterUpdate()
AddStatus FileNumber, NoticeToOccupant, "Final Notice to Occupants sent"
End Sub

Private Sub NoticeToOccupant_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
NoticeToOccupant = Date
Call NoticeToOccupant_AfterUpdate
End If

End Sub

Private Sub NoticeToOccupant15days_AfterUpdate()

If State = "MD" Then
    'No existing field
    'NoticeToOccupant15daysExpires = DateAdd("d", 15, NoticeToOccupant15days)
    AddStatus FileNumber, NoticeToOccupant15days, "MD Notice 1308 Filed, expires " & Format(Forms!EvictionDetails!txt15dayExpiration, "m/dd/yyyy")
Else
    AddStatus FileNumber, NoticeToOccupant15days, "MD Notice 1308 Filed"
End If

End Sub

Private Sub NoticeToOccupant15days_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(NoticeToOccupant15days)
End Sub

Private Sub NoticeToOccupant15days_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
NoticeToOccupant15days = Now()
Call NoticeToOccupant15days_AfterUpdate
End If

End Sub

Private Sub NoticeToQuitFiled_AfterUpdate()
If State = "VA" Then
    NoticeToQuitExpires = DateAdd("d", 10, NoticeToQuitFiled)
    AddStatus FileNumber, NoticeToQuitFiled, "Notice to Quit Filed, expires " & Format(NoticeToQuitExpires, "m/d/yyyy")
Else
    AddStatus FileNumber, NoticeToQuitFiled, "Notice to Quit Filed"
End If
End Sub

Private Sub NoticeToQuitFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(NoticeToQuitFiled)
End Sub

Private Sub NoticeToQuitFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
NoticeToQuitFiled = Now()
Call NoticeToQuitFiled_AfterUpdate
End If

End Sub

Private Sub NoticeToQuitServed_AfterUpdate()
If State = "DC" Then
    NoticeToQuitExpires = DateAdd("d", 30, NoticeToQuitServed)
    AddStatus FileNumber, NoticeToQuitServed, "Notice to Quit Served, expires " & Format$(NoticeToQuitExpires, "m/d/yyyy")
    AddInvoiceItem FileNumber, "EV", "Notice to Quit Served", 40, 0, False, True, False, False
    AddInvoiceItem FileNumber, "EV", "Courier Fee", 10, False, 0, True, False, False
Else
    AddStatus FileNumber, NoticeToQuitServed, "Notice to Quit Served"
End If
End Sub

Private Sub NoticeToQuitServed_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(NoticeToQuitServed)
End Sub

Private Sub NoticeToQuitServed_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
NoticeToQuitServed = Date
Call NoticeToQuitServed_AfterUpdate
End If

End Sub

Private Sub NoticeToTenants_AfterUpdate()
NoticeToTenantsExpires = DateAdd("d", 90, NoticeToTenants)
AddStatus FileNumber, NoticeToTenants, "90-Day notice sent to tenants, expires " & Format$(NoticeToTenantsExpires, "m/d/yyyy")
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter postage for the notices|EV|90-Day notice sent to tenants"

End Sub

Private Sub NoticeToTenants_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(NoticeToTenants)
End Sub

Private Sub NoticeToTenants_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
NoticeToTenants = Date
Call NoticeToTenants_AfterUpdate
End If

End Sub

Private Sub Referred_AfterUpdate()
Dim ClientID As Integer
AddStatus FileNumber, Referred, "Eviction referral received"
'Old Invoice method
'AddInvoiceItem FileNumber, "EV", "Eviction referral received", 350, True, True, False, False

'New EV Fee methodology
ClientID = DLookup("ClientID", "CaseList", "filenumber=" & FileNumber)
Select Case State

Case "DC"
    FeeAmount = Nz(DLookup("DCFee", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "EV", "Eviction referral received", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "EV", "Eviction referral received", 1, 0, True, True, False, False 'set unknown fee as $1, per Diane
    End If
Case "MD"  'Two Fees, need to resolve
    FeeAmount = Nz(DLookup("MDFee", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "EV", "Eviction referral received", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "EV", "Eviction referral received", 1, 0, True, True, False, False 'set unknown fee as $1, per Diane
    End If
    Case "VA"
    FeeAmount = Nz(DLookup("VAFee", "ClientList", "ClientID=" & ClientID))
    If FeeAmount > 0 Then
        AddInvoiceItem FileNumber, "EV", "Eviction referral received", FeeAmount, 0, True, True, False, False
    Else
        AddInvoiceItem FileNumber, "EV", "Eviction referral received", 1, 0, True, True, False, False 'set unknown fee as $1, per Diane
    End If


End Select
End Sub

Private Sub Referred_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Referred)
End Sub

Private Sub Referred_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
Referred = Now()
Call Referred_AfterUpdate
End If

End Sub

Private Sub RentAmount_AfterUpdate()
' If 1st payment date is null, then default to 1st day of next month
If IsNull(RentFirstPayment) Then RentFirstPayment = DateSerial(Year(DateAdd("m", 1, Date)), Month(DateAdd("m", 1, Date)), 1)
End Sub

Private Sub RentClosed_AfterUpdate()
AddStatus FileNumber, RentClosed, "Rent file closed"
End Sub

Private Sub RentClosed_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(RentClosed)
End Sub

Private Sub RentClosed_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
RentClosed = Date
Call RentClosed_AfterUpdate
End If

End Sub

Private Sub RentLeaseToBroker_AfterUpdate()
AddStatus FileNumber, RentLeaseToBroker, "Lease sent to broker"
End Sub

Private Sub RentLeaseToBroker_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(RentLeaseToBroker)
End Sub

Private Sub RentLeaseToBroker_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
RentLeaseToBroker = Date
Call RentLeaseToBroker_AfterUpdate
End If

End Sub

Private Sub RentReferral_AfterUpdate()
AddStatus FileNumber, RentReferral, "Rent Referral received"
End Sub

Private Sub RentReferral_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(RentReferral)
End Sub

Private Sub RentReferral_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
RentReferral = Date
Call RentReferral_AfterUpdate
End If

End Sub

Private Sub RentSignedLease_AfterUpdate()
AddStatus FileNumber, RentSignedLease, "Executed lease returned from broker"
End Sub

Private Sub RentSignedLease_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(RentSignedLease)
End Sub

Private Sub RentSignedLease_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
RentSignedLease = Date
Call RentSignedLease_AfterUpdate
End If

End Sub

Private Sub ReportedVacant_AfterUpdate()
If IsNull(ReportedVacant) Then Exit Sub
  
  
  AddStatus FileNumber, ReportedVacant, "Property reported vacant"
  RentClosed = ReportedVacant
  
  If (HearingDate > Now) Then
    HearingDate = Null
  End If
  
  If (LockoutDate > Now) Then
    LockoutDate = Null
  End If
  
End Sub

Private Sub ReportedVacant_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ReportedVacant = Now()
AddStatus FileNumber, ReportedVacant, "Property reported vacant"
End If

End Sub

Private Sub ResponseDeadlineOwner_BeforeUpdate(Cancel As Integer)

If Not IsNull(ResponseDeadlineOwner) Then

    If HearingCheking(ResponseDeadlineOwner, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(ResponseDeadlineOwner, 2) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(ResponseDeadlineOwner, 3) = 1 Then
    Cancel = 1
    End If

End If

End Sub

Private Sub ResponseFiled_AfterUpdate()
AddStatus FileNumber, ResponseFiled, "Response Filed"
End Sub

Private Sub ResponseFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ResponseFiled)
End Sub

Private Sub ResponseFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ResponseFiled = Date
AddStatus FileNumber, ResponseFiled, "Response Filed"
End If

End Sub

Private Sub ServicerRelease_AfterUpdate()
Dim Status As String, rstJnl As Recordset
If Not IsNull(ServicerRelease) Then servicereffective = InputBox("Please enter the effective date")
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
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EnterWriteOffReason!FileNumber,Now,GetFullName(),'" & strinfo & "',2 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
End Sub

Private Sub ShowCurrent_Click()
If ShowCurrent Then
    Me.Filter = "FileNumber = " & Me![FileNumber] & "AND Current = True"
Else
    Me.Filter = "FileNumber = " & Me![FileNumber]
End If
End Sub



Private Sub Withdrawn_AfterUpdate()

If (HearingDate > Now) Then
  Me.HearingDate = Null
End If


AddStatus FileNumber, Withdrawn, "Motion withdrawn"
End Sub

Private Sub Withdrawn_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Withdrawn)
End Sub

Private Sub WritRequested_AfterUpdate()
AddStatus FileNumber, WritRequested, "Writ requested"
End Sub

Private Sub WritRequested_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(WritRequested)
End Sub

Private Sub WritRequested_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
WritRequested = Now()
AddStatus FileNumber, WritRequested, "Writ requested"
End If

End Sub

Private Function UpdateCalendar(calendarDateOldValue As Variant, calendarDate As Variant, calendarID As String, Calendar_Type As String) As Variant
Dim emailGroup As String
Dim Subject As String

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

emailGroup = IIf(Calendar_Type = "Eviction Hearing", "SharedCalRecipEV", "SharedCalRecipLO")

If Calendar_Type = "Eviction Hearing" Then
Select Case State
Case "MD"
emailGroup = "SharedCalRecipEV-MD"
Case "DC"
emailGroup = "SharedCalRecipEV-DC"
Case "VA"
emailGroup = "SharedCalRecipEV-VA"
Case Else
emailGroup = "SharedCalRecip"
End Select
Else
emailGroup = "SharedCalRecipLO"
End If


Subject = Calendar_Type & ": " & Forms![Case List]!FileNumber & " - " & Forms![Case List]!PrimaryDefName
If (calendarID = "") Then     ' new event on calendar
    UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, Subject, Forms![Case List]!JurisdictionID.Column(1), 3, emailGroup)
Else                                    ' change existing event on calendar

   If (IsNull(calendarDateOldValue)) Then   ' new date
    UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, Subject, Forms![Case List]!JurisdictionID.Column(1), 3, emailGroup)
      
   Else  ' date in the future - create new calendar event
    UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, Subject, Forms![Case List]!JurisdictionID.Column(1), 3, emailGroup)
   'Else ' otherwise update calendar event
    'Call UpdateCalendarEvent(calendarID, CDate(calendarDate), False, Subject, Forms![case list]!JurisdictionID.Column(1), 3)
   End If
End If
    
End Function

