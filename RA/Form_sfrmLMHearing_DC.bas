VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmLMHearing_DC"
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
'passHearingBeforeform = True

If IsNull(Me.txtHearing) Then
    MsgBox ("Hearing date is missing")
    cbxCondactedTypeID = Null

Exit Sub
End If

Select Case cbxCondactedTypeID
    
    Case 1
        If Now() < txtHearing And DateDiff("d", Date, [txtHearing]) > 1 Then
            If Not IsNull(HearingCalendarEntryID) Then

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
        End If
    Case 2
'        If Now() < txtHearing And DateDiff("d", Date, [txtHearing]) > 1 Then
'            If Not IsNull(HearingCalendarEntryID) Then
'                Call DeleteCalendarEvent(HearingCalendarEntryID)
'                HearingCalendarEntryID = ""
'                AddStatus FileNumber, Date, "Disposition entered for " & Trim(Me.HearingTypeID.Column(1)) & " "
         'added logic
        AddStatus FileNumber, Date, "Disposition entered for " & Trim(Me.HearingTypeID.Column(1)) & " "

         If Now() < txtHearing.OldValue Then
            If Not IsNull(HearingCalendarEntryID) Then
                Call DeleteCalendarEvent(HearingCalendarEntryID)
                HearingCalendarEntryID = ""
                txtHearing.Value = Null
                txtHearing.Enabled = False
                cmdAddNew.Enabled = True
                   Me.Requery
            Else
                txtHearing = Null
                txtHearing.Enabled = False
                cmdAddNew.Enabled = True
            End If
        Else
                txtHearing = Null
                txtHearing.Enabled = False
                cmdAddNew.Enabled = True
        End If
         '---------------
'    AddStatus FileNumber, Date, "Exception Hearing Continue "
'    If Now() < ExceptionsHearing.OldValue Then
'    If Not IsNull(ExceptionsHearingEntryID) Then
'    Call DeleteCalendarEvent(ExceptionsHearingEntryID)
'    ExceptionsHearingEntryID = Null
'
'    ExceptionsHearing.Value = Null
'    ExceptionsHearing.Enabled = False
'    ExceptionsHearingTime.Value = Null
'    ExceptionsHearingTime.Enabled = False
'    AddNewDateException.Enabled = True
'
'    Else
'    ExceptionsHearing.Value = Null
'    ExceptionsHearing.Enabled = False
'    ExceptionsHearingTime.Value = Null
'    ExceptionsHearingTime.Enabled = False
'    AddNewDateException.Enabled = True
'    End If
'    Else
'    ExceptionsHearing.Value = Null
'    ExceptionsHearing.Enabled = False
'    ExceptionsHearingTime.Value = Null
'    ExceptionsHearingTime.Enabled = False
'    AddNewDateException.Enabled = True
'    End If
'----------
        DoCmd.SetWarnings False
        strinfo = "Disposition of " & Me.cbxCondactedTypeID.Column(1) & " entered for " & Trim(Me.HearingTypeID.Column(1)) & ""
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
            'End If
        'End If
    Case 3
        If Now() < txtHearing And DateDiff("d", Date, [txtHearing]) > 1 Then
            If Not IsNull(HearingCalendarEntryID) Then
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
        End If
        
    Case 4
        If Now() < txtHearing And DateDiff("d", Date, [txtHearing]) > 1 Then
            If Not IsNull(HearingCalendarEntryID) Then
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
        End If
    Case 5
        If Now() < txtHearing Then 'And DateDiff("d", Now(), [txtHearing]) > 1 Then
            If Not IsNull(HearingCalendarEntryID) Then
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
        End If
    Case 6
    
        
        If Now() < txtHearing And DateDiff("d", Date, [txtHearing]) > 1 Then
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
        
        Case 7
    
        
        If Now() < txtHearing And DateDiff("d", Date, [txtHearing]) > 1 Then
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

Private Sub cmdAddNew_Click()

If cbxCondactedTypeID.Value = 2 Then
    txtHearing.Enabled = True
    txtHearing.Locked = False
    
    cbxCondactedTypeID.Value = Null
    cboLMConductedby.Value = Null
End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
If Not IsNull(txtHearing) Then
  If (Nz(txtHearing) <> Nz(txtHearing.OldValue)) Then
    HearingCalendarEntryID = UpdateCalendar(txtHearing.OldValue, txtHearing, Nz(HearingCalendarEntryID))
  End If
End If

End Sub

Private Sub Form_Current()
 '  If IsNull(FileNumber) Then FileNumber = [Forms]![ForeclosureDetails]![FileNumber]
 '  If IsNull(ForeclosureID) Then ForeclosureID = [Forms]![ForeclosureDetails]![ForeclosureID]
 '  If Hearing < Date Then Me.Combo18.Enabled = False

Dim rs As Recordset
Dim i As Integer

i = 0

Set rs = CurrentDb.OpenRecordset("Select * FROM LMHearings_DC where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
Do Until rs.EOF

    If Not IsNull(Hearing) Then
        txtHearing.Enabled = False
    Else
        txtHearing.Enabled = True
    End If
    
    If Not IsNull(HearingTypeID) Then
        i = i + 1
    End If

rs.MoveNext
Loop

rs.Close
Set rs = Nothing

If i = 3 Then
Me.AllowAdditions = False
End If


End Sub

Private Sub HearingTypeID_AfterUpdate()

Dim i As Integer
i = 0

If Not IsNull(HearingTypeID) Then
    Dim rs As Recordset
    i = HearingTypeID
    
    Set rs = CurrentDb.OpenRecordset("Select * FROM LMHearings_DC where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
    Do Until rs.EOF
    
    If rs!HearingTypeID = i Then
        MsgBox ("This hearing type has been selected")
    
       HearingTypeID = Null
    Exit Sub
    End If
    
    rs.MoveNext
    Loop

rs.Close
Set rs = Nothing
End If

End Sub

Private Sub txtHearing_AfterUpdate()
 
If Not IsNull(txtHearing) Then
    txtHearing.Locked = True
End If

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
'msgbox (calendarID)

'If (Not IsNull(calendarDateOldValue) And calendarID = "") Then
    'MsgBox "Please update the Shared Calendar", vbExclamation
    'Exit Function
'End If

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



