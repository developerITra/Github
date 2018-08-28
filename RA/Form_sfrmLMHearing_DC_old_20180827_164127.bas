VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmLMHearing_DC_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboLMConductedby_AfterUpdate()
'DoCmd.SetWarnings False
'strinfo = "Mediation conducted by: " & Me.cboLMConducted.Text
'strinfo = Replace(strinfo, "'", "''")
'strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info) Values(Forms![Case list]!FileNumber,Now,GetFullName(),'" & strinfo & "')"
'DoCmd.RunSQL strSQLJournal
'DoCmd.SetWarnings True


End Sub

Private Sub cbxCondactedTypeID_AfterUpdate()

Select Case cbxCondactedTypeID
    
    Case 1
        If Not IsNull(txtHearing) Then
            Call DeleteCalendarEvent(HearingCalendarEntryID)
            HearingCalendarEntryID = ""
            AddStatus FileNumber, Date, "Disposition entered for " & Trim(Me.HearingTypeID.Column(1)) & " "
            
            DoCmd.SetWarnings False
            strinfo = "Disposition of " & Me.cbxCondactedTypeID.Column(1) & " entered for " & Trim(Me.HearingTypeID.Column(1)) & ""
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery
        End If
    Case 2
        If Not IsNull(txtHearing) Then
            Call DeleteCalendarEvent(HearingCalendarEntryID)
            HearingCalendarEntryID = ""
            AddStatus FileNumber, Date, "Disposition entered for " & Trim(Me.HearingTypeID.Column(1)) & " "
            
            DoCmd.SetWarnings False
            strinfo = "Disposition of " & Me.cbxCondactedTypeID.Column(1) & " entered for " & Trim(Me.HearingTypeID.Column(1)) & ""
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery

        End If
    Case 3
        If Not IsNull(txtHearing) Then
            Call DeleteCalendarEvent(HearingCalendarEntryID)
            HearingCalendarEntryID = ""
            AddStatus FileNumber, Date, "Disposition entered for " & Trim(Me.HearingTypeID.Column(1)) & " "
            
            DoCmd.SetWarnings False
            strinfo = "Disposition of " & Me.cbxCondactedTypeID.Column(1) & " entered for " & Trim(Me.HearingTypeID.Column(1)) & ""
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery

        End If
        
    Case 4
        If Not IsNull(txtHearing) Then
            Call DeleteCalendarEvent(HearingCalendarEntryID)
            HearingCalendarEntryID = ""
            AddStatus FileNumber, Date, "Disposition entered for " & Trim(Me.HearingTypeID.Column(1)) & " "

            DoCmd.SetWarnings False
            strinfo = "Disposition of " & Me.cbxCondactedTypeID.Column(1) & " entered for " & Trim(Me.HearingTypeID.Column(1)) & ""
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery

        End If
    Case 5
        If Not IsNull(txtHearing) Then
            Call DeleteCalendarEvent(HearingCalendarEntryID)
            HearingCalendarEntryID = ""
            AddStatus FileNumber, Date, "Disposition entered for " & Trim(Me.HearingTypeID.Column(1)) & " "
            
            DoCmd.SetWarnings False
            strinfo = "Disposition of " & Me.cbxCondactedTypeID.Column(1) & " entered for " & Trim(Me.HearingTypeID.Column(1)) & ""
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery

        End If
    Case 6
    
        
        If txtHearing > Date And DateDiff("d", Date, [txtHearing]) > 1 Then
            Call DeleteCalendarEvent(HearingCalendarEntryID)
            HearingCalendarEntryID = ""
            AddStatus FileNumber, Date, "Disposition entered for " & Trim(Me.HearingTypeID.Column(1)) & " "
            
            DoCmd.SetWarnings False
            strinfo = "Disposition of " & Me.cbxCondactedTypeID.Column(1) & " entered for " & Trim(Me.HearingTypeID.Column(1)) & ""
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery

        End If
End Select



End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
  If (Nz(txtHearing) <> Nz(txtHearing.OldValue)) Then
    HearingCalendarEntryID = UpdateCalendar(txtHearing.OldValue, txtHearing, Nz(HearingCalendarEntryID))
  End If

End Sub


Private Sub Form_Current()
 '  If IsNull(FileNumber) Then FileNumber = [Forms]![ForeclosureDetails]![FileNumber]
 '  If IsNull(ForeclosureID) Then ForeclosureID = [Forms]![ForeclosureDetails]![ForeclosureID]
' If Hearing < Date Then Me.Combo18.Enabled = False
End Sub

Private Sub txtHearing_AfterUpdate()
 
AddStatus FileNumber, Date, "DC " & Trim(Me.HearingTypeID.Column(1)) & " Hearing scheduled for " & Hearing
   
If StaffID = 0 Then Call GetLoginName
HearingStaffID = StaffID
HearingInitials.Requery


    DoCmd.SetWarnings False
    strinfo = "DC " & Me.HearingTypeID.Column(1) & " hearing was scheduled for " & Hearing
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery

End Sub

Private Function UpdateCalendar(calendarDateOldValue As Variant, calendarDate As Variant, calendarID As String) As Variant

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

Select Case Forms!foreclosuredetails!State
Case "MD"
emailGroup = "SharedCalRecipFC-MD"
Case "DC"
emailGroup = "SharedCalRecipFC-DC"
Case "VA"
emailGroup = "SharedCalRecipFC-VA"
Case Else
emailGroup = "SharedCalRecip"
End Select

If (calendarID = "") Then     ' new event on calendar
    'UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")" & " DC " & Trim(Me.HearingTypeID.Column(1)) & " ", Parent!DCMedHearingLocation.Column(1), 8, emailGroup)
    UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")" & " DC " & Trim(Me.HearingTypeID.Column(1)) & " ", "District of Columbia, DC", 8, emailGroup)


'txtDCMedHearingLocation
Else                                    ' change existing event on calendar

   If (IsNull(calendarDateOldValue)) Or Format(calendarDateOldValue, "mm/dd/yyyy") = Date Then   ' new date, also keep existing date if today
    UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")" & " DC " & Trim(Me.HearingTypeID.Column(1)) & " ", "District of Columbia, DC", 8, emailGroup)
    
      
   Else ' otherwise update calendar event
    
 
    Call UpdateCalendarEvent(calendarID, CDate(calendarDate), False, DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")" & " DC " & Trim(Me.HearingTypeID.Column(1)) & "", "District of Columbia, DC", 8)
   End If
End If

    
End Function

Private Sub txtHearing_BeforeUpdate(Cancel As Integer)
  If Not IsNull(Hearing) Then
    If (Hearing < Date) Then
      Cancel = -1
      MsgBox "Hearing Date cannot be in the past.", vbCritical
      Exit Sub
    End If
    
    Dim dteTimePortion As Date
    dteTimePortion = TimeValue(Hearing)
    
    If Hour(dteTimePortion) < 8 Or Hour(dteTimePortion) > 18 Or (Hour(dteTimePortion) = 18 And Minute(dteTimePortion) > 0) Then
    
      Cancel = -1
      MsgBox "Hearing time must be between 8:00 AM and 6:00 PM"
    End If
  
     
    If HearingCheking(Hearing, 2) = 1 Then
    Cancel = 1
    End If
    
    
End If

  
End Sub



