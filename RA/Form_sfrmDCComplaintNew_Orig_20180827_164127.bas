VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmDCComplaintNew_Orig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim DayNum As Variant
    Dim IsWeekend As Boolean

Private Sub bondPosted_AfterUpdate()
If Not IsNull(BondPosted) Then
 AddStatus FileNumber, BondPosted, "Bond Posted"
Else
 AddStatus FileNumber, Now(), "Removed Bond Posting Date"
End If
End Sub

Private Sub bondPosted_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    BondPosted = Now()
    Call bondPosted_AfterUpdate
End If

End Sub

Private Sub AllBorrowerServed_AfterUpdate()

If Not IsNull(AllBorrowerServed) Then
 AddStatus FileNumber, AllBorrowerServed, "All Borrower Served"

    'Forms!ForeclosureDetails!sfrmStatus.Requery

DoCmd.SetWarnings False
strinfo = "All Borrowers Served date Entered"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
Forms!Journal.Requery
Else

 AddStatus FileNumber, Now(), "Removed All Borrower Served"
End If

End Sub

Private Sub AllBorrowerServed_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    AllBorrowerServed = Now()
    Call AllBorrowerServed_AfterUpdate
End If
End Sub

Private Sub ApprovedDraft_AfterUpdate()
If Not IsNull(ApprovedDraft) Then
 AddStatus FileNumber, ApprovedDraft, "Approved Draft Complaint"


'Forms!ForeclosureDetails!sfrmStatus.Requery

DoCmd.SetWarnings False
strinfo = "Approved Draft Complaint"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
Forms!Journal.Requery

Else
 AddStatus FileNumber, Now(), "Removed Approved Draft Complaint"
End If

End Sub

Private Sub ApprovedDraft_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ApprovedDraft = Now()
    Call ApprovedDraft_AfterUpdate
End If
End Sub

Private Sub cbxDCHearingResults_AfterUpdate()

Select Case cbxDCHearingResults
    
    Case 1
    Dim i As Integer
    i = DateDiff("d", Date, InitialHearingConference)
        If Now() < InitialHearingConference And DateDiff("d", Date, InitialHearingConference) >= 1 Then
            If Not IsNull(DCHearingEntryID) Then

                Call DeleteCalendarEvent(DCHearingEntryID)
                DCHearingEntryID = ""
                AddStatus FileNumber, Date, "DC Hearing: Failure to appear"
                
                DoCmd.SetWarnings False
                strinfo = "DC Hearing: Failure to appear"
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery
            End If
        End If
        
Case 2
        If Now() < InitialHearingConference And DateDiff("d", Date, [InitialHearingConference]) >= 1 Then
            If Not IsNull(DCHearingEntryID) Then

                Call DeleteCalendarEvent(DCHearingEntryID)
                DCHearingEntryID = ""
                AddStatus FileNumber, Date, "DC Hearing:Mediation"
                
                DoCmd.SetWarnings False
                strinfo = "DC Hearing:Mediation"
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery
            End If
        End If
        
        Case 3
        If Now() < InitialHearingConference And DateDiff("d", Date, [InitialHearingConference]) >= 1 Then
            If Not IsNull(DCHearingEntryID) Then

                Call DeleteCalendarEvent(DCHearingEntryID)
                DCHearingEntryID = ""
                AddStatus FileNumber, Date, "DC Hearing: Trial"
                
                DoCmd.SetWarnings False
                strinfo = "DC Hearing: Trial"
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery
            End If
        End If
    Case 4
        If Now() < InitialHearingConference And DateDiff("d", Date, [InitialHearingConference]) >= 1 Then
            If Not IsNull(DCHearingEntryID) Then

                Call DeleteCalendarEvent(DCHearingEntryID)
                DCHearingEntryID = ""
                AddStatus FileNumber, Date, "DC Hearing: Withdraw/Cancelled"
                
                DoCmd.SetWarnings False
                strinfo = "DC Hearing: Withdraw/Cancelled"
                strinfo = Replace(strinfo, "'", "''")
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
                DoCmd.RunSQL strSQLJournal
                DoCmd.SetWarnings True
                Forms!Journal.Requery
            End If
        End If
   

    Case 5
'        If Now() < txtHearing And DateDiff("d", Date, [txtHearing]) > 1 Then
'            If Not IsNull(HearingCalendarEntryID) Then
'                Call DeleteCalendarEvent(HearingCalendarEntryID)
'                HearingCalendarEntryID = ""
'                AddStatus FileNumber, Date, "Disposition entered for " & Trim(Me.HearingTypeID.Column(1)) & " "
         'added logic
        

         'If Now() < InitialHearingConference.OldValue Then
         If Date <= InitialHearingConference Then
            If Not IsNull(DCHearingEntryID) Or DCHearingEntryID <> "" Then
                Call DeleteCalendarEvent(DCHearingEntryID)
                DCHearingEntryID = ""
                InitialHearingConference.Value = Null
                DCHearingTime = Null
                InitialHearingConference.Enabled = False
                DCHearingTime.Enabled = False
                cmdAddNew.Enabled = True
                   Me.Requery
            Else
                InitialHearingConference = Null
                DCHearingTime = Null
                InitialHearingConference.Enabled = False
                DCHearingTime.Enabled = False
                cmdAddNew.Enabled = True
            End If
        Else
                InitialHearingConference = Null
                InitialHearingConference.Enabled = False
                DCHearingTime = Null
                DCHearingTime.Enabled = False

                cmdAddNew.Enabled = True
        End If
      
            AddStatus FileNumber, Date, "DC Hearing:Continued"
            
            DoCmd.SetWarnings False
            strinfo = "DC Hearing:Continued"
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True
            Forms!Journal.Requery
End Select
End Sub

Private Sub cmdAddNew_Click()
'If MsgBox("Are you sure you want to re-schedule Hearing time?", vbYesNo + vbQuestion) = vbYes Then
If cbxDCHearingResults.Value = 5 Then
    InitialHearingConference.Enabled = True
    InitialHearingConference.Locked = False
    DCHearingTime.Enabled = True
    DCHearingTime.Locked = False
    cbxDCHearingResults.Value = Null
End If

End Sub

Private Sub cmdNew_Click()

If Not IsNull(ServiceDue) Then
    ServiceDeadline = ServiceDue
End If

DC_Order = Null
MotiontoExtendFiled = Null
ServiceDue = Null

MotiontoExtendFiled.Enabled = True
ServiceDue.Enabled = True
DC_Order.Enabled = True

If Not IsNull(ServiceDeadline) Then

AddStatus FileNumber, Now(), "Service Deadline by " & Me.ServiceDeadline & ""
AddStatus FileNumber, Now(), "Service Due Removed"


    DoCmd.SetWarnings False
    strinfo = "Service Deadline by " & Me.ServiceDeadline & ""
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    
    strinfo = "Service Due Removed "
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    
    DoCmd.SetWarnings True
    Forms!Journal.Requery
    
End If

End Sub

Private Sub ComplaintFiled_AfterUpdate()
If BHproject Then
    If Not IsNull(ComplaintFiled) Then
     AddStatus FileNumber, ComplaintFiled, "Complaint Filed"
    End If
Else

If Not IsNull(ComplaintFiled) Then
 AddStatus FileNumber, ComplaintFiled, "Complaint Filed"
 
 Dim newdate As Date
 newdate = ComplaintFiled + 60
 
'DayNum = Weekday(newdate)

 If Weekday(newdate) = 1 Or Weekday(newdate) = 7 Then
 
    If Weekday(newdate) = 1 Then 'Sunday
        ServiceDeadline = newdate + 1
    ElseIf Weekday(newdate) = 7 Then 'Saturday
        ServiceDeadline = newdate + 2
    End If
 Else
    ServiceDeadline = newdate
 End If
 
 
Dim cbxClient As Integer
cbxClient = Nz(DLookup("ClientID", "Caselist", "FileNumber=" & FileNumber))
Select Case Nz(DLookup("State", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
  Case "VA"
   Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
    Case 1 'Conventional
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
    Case 2 'VA or Veteran's Affairs
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
    Case 3 'FHA or HUD
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=570")) 'HUD/FHA
    Case 4
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=177")) 'Fannie Mae
    Case 5
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=263")) 'Freddie Mac
    Case Else
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
    End Select
    If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
      InvPct = DLookup("VAComplaintFiledPct", "clientlist", "clientid=" & cbxClient)
    Else
      InvPct = 1
    End If
    If FeeAmount > 0 Then
      If InvPct <= 1 Then
        AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when complaint filed received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      Else
        'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when complaint filed received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      End If
    End If
  Case "MD"
    Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
      Case 1 'Conventional
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & cbxClient))
      Case 2 'VA or Veteran's Affairs
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
      Case 3 'FHA or HUD
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=570")) 'HUD/FHA
      Case 4
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=177")) 'Fannie Mae
      Case 5
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=263")) 'Freddie Mac
      Case Else
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & cbxClient))
    End Select
    If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
      InvPct = DLookup("MDComplaintFiledPct", "clientlist", "clientid=" & cbxClient)
    Else
      InvPct = 1
    End If
    If FeeAmount > 0 Then
      If InvPct <= 1 Then
        AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when complaint filed received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      Else
        'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when complaint filed received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      End If
    End If
    Case "DC"
      Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
        Case 1 'Conventional
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=" & cbxClient))
        Case 2 'VA or Veteran's Affairs
          FeeAmount = Nz(DLookup("FeeDcReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
        Case 3 'FHA or HUD
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=570")) 'HUD/FHA
        Case 4
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=177")) 'Fannie Mae
        Case 5
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=263")) 'Freddie Mac
        Case Else
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=" & cbxClient))
      End Select
      'Debug.Print cbxClient & "," & (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient)))
      
      If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
        InvPct = DLookup("DCComplaintFiledPct", "clientlist", "clientid=" & cbxClient)
      Else
        InvPct = 1
      End If
      If FeeAmount > 0 Then
        If InvPct <= 1 Then
          AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when complaint filed received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
        Else
          'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when complaint filed received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
        End If
      End If
  End Select
Else
 AddStatus FileNumber, Now(), "Removed Complaint Filed date"
End If
End If

End Sub

Private Sub DC_Order_AfterUpdate()
If Not IsNull(DC_Order) Then
 AddStatus FileNumber, DC_Order, "Order Extending date"

'Forms!ForeclosureDetails!sfrmStatus.Requery
 DoCmd.SetWarnings False
    strinfo = "Order Extending date"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
    
Else
    AddStatus FileNumber, Now(), "Removed Order Extending date"
End If
End Sub

Private Sub DC_Order_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DC_Order = Now()
    Call DC_Order_AfterUpdate
End If
End Sub



Private Sub DCHearingTime_AfterUpdate()
''If Not BHproject Then
'    If IsNull(InitialHearingConference) Or IsNull(DCHearingTime) Then Exit Sub
'
'    If Hour(DCHearingTime) < 8 Or Hour(DCHearingTime) > 19 Then
'        DCHearingTime = DateAdd("h", 12, DCHearingTime)
'        If Hour(DCHearingTime) < 8 Or Hour(DCHearingTime) > 19 Then
'            'MsgBox "Invalid Hearing time: " & Format$(ExceptionsHearingTime, "h:nn am/pm")
'            DCHearingTime = Null
'            Cancel = -1
'            MsgBox "Hearing time must be between 8:00 AM and 7:00 PM"
'            Exit Sub
'        End If
'    End If
'
'
'    If Not IsNull(DCHearingTime) Then Call UpdateCalendarDCHearing
'    'Call Visuals
'    DCHearingTime.Locked = True
'    'DCHearingTime.Enabled = False
'
'    AddStatus FileNumber, Now(), "DC Initial Hearing scheduled time for " & Format$(InitialHearingConference, "m/d/yyyy") & " at " & Format$(DCHearingTime, "h:nn am/pm")
'
''End If

'Modified on 9_8_15

If Not BHproject Then

    If IsNull(InitialHearingConference) Or IsNull(DCHearingTime) Then
        DCHearingTime = Null
        InitialHearingConference = Null
    Exit Sub
    End If
    
    If Hour(DCHearingTime) >= 8 And Hour(DCHearingTime) < 13 Then
        DCHearingTime = Format$(DCHearingTime, "h:nn am/pm")
    ElseIf Hour(DCHearingTime) >= 1 And Hour(DCHearingTime) <= 7 Then
        DCHearingTime = DateAdd("h", 12, DCHearingTime)
    Else
        MsgBox "Hearing time must be between 8:00 AM and 7:00 PM" ': " & Format$(DCHearingTime, "h:nn am/pm")
        DCHearingTime = Null
        Exit Sub
    End If
    
    If Not IsNull(DCHearingTime) Then Call UpdateCalendarDCHearing
    'Call Visuals
    DCHearingTime.Locked = True
    'DCHearingTime.Enabled = False
    AddStatus FileNumber, Now(), "DC Initial Hearing scheduled time for " & Format$(InitialHearingConference, "m/d/yyyy") & " at " & Format$(DCHearingTime, "h:nn am/pm")

End If



End Sub

Private Sub Form_Current()

If FileReadOnly Then
   Me.AllowEdits = False
   Me.cmdNew.Enabled = False
   Exit Sub
End If

If DCTabView = False Then
 'Me.AllowEdits = False
 
    Dim ctl As Control
    Dim lngI As Long
    Dim bSkip As Boolean
        
    For Each ctl In Form.Controls
    Select Case ctl.ControlType
    'Case acTextBox, acComboBox, acListBox, acOptionGroup, acCheckBox, acSubform, acOptionButton
    Case acTextBox, acComboBox, acCheckBox
           
    If (ctl.Enabled) Then ctl.Enabled = False
                    
        Case acCommandButton
        bSkip = False
            If ctl.Name = "cmdNewDate" Then bSkip = True
            If ctl.Name = "cmdAddNew" Then bSkip = True
            If Not bSkip Then ctl.Enabled = False
                           
    End Select
    Next

End If

If Not IsNull(MotiontoExtendFiled) Then
    MotiontoExtendFiled.Enabled = False
End If

If Not IsNull(DC_Order) Then
    DC_Order.Enabled = False
End If

If Not IsNull(ServiceDue) Then
    ServiceDue.Enabled = False
End If

If Not IsNull(Me.InitialHearingConference) Then
    InitialHearingConference.Enabled = False
End If

If Not IsNull(Me.DCHearingTime) Then
    DCHearingTime.Enabled = False
End If




End Sub

Private Sub JudgmentEntered_AfterUpdate()
If BHproject Then
If Not IsNull(JudgmentEntered) Then
 AddStatus FileNumber, JudgmentEntered, "Judgment Entered"
 'Forms!ForeclosureDetails!sfrmStatus.Requery
End If
Else


If Not IsNull(JudgmentEntered) Then
 AddStatus FileNumber, JudgmentEntered, "Judgment Filed"
 Dim cbxClient As Integer
cbxClient = Nz(DLookup("ClientID", "Caselist", "FileNumber=" & FileNumber))
Select Case Nz(DLookup("State", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
  Case "VA"
   Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
    Case 1 'Conventional
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
    Case 2 'VA or Veteran's Affairs
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
    Case 3 'FHA or HUD
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=570")) 'HUD/FHA
    Case 4
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=177")) 'Fannie Mae
    Case 5
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=263")) 'Freddie Mac
    Case Else
      FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & cbxClient))
    End Select
    If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
      InvPct = DLookup("VAJudgmentEnteredPct", "clientlist", "clientid=" & cbxClient)
    Else
      InvPct = 1
    End If
    If FeeAmount > 0 Then
      If InvPct <= 1 Then
        AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when judgement entered received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      Else
        'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when complaint filed received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      End If
    End If
  Case "MD"
    Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
      Case 1 'Conventional
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & cbxClient))
      Case 2 'VA or Veteran's Affairs
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
      Case 3 'FHA or HUD
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=570")) 'HUD/FHA
      Case 4
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=177")) 'Fannie Mae
      Case 5
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=263")) 'Freddie Mac
      Case Else
        FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & cbxClient))
    End Select
    If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
      InvPct = DLookup("MDJudgmentEnteredPct", "clientlist", "clientid=" & cbxClient)
    Else
      InvPct = 1
    End If
    If FeeAmount > 0 Then
      If InvPct <= 1 Then
        AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when judgement entered received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      Else
        'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when complaint filed received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
      End If
    End If
    Case "DC"
      Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & FileNumber & " AND Current=1"))
        Case 1 'Conventional
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=" & cbxClient))
        Case 2 'VA or Veteran's Affairs
          FeeAmount = Nz(DLookup("FeeDcReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
        Case 3 'FHA or HUD
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=570")) 'HUD/FHA
        Case 4
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=177")) 'Fannie Mae
        Case 5
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=263")) 'Freddie Mac
        Case Else
          FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=" & cbxClient))
      End Select
      'Debug.Print cbxClient & "," & (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient)))
      
      If (Nz(DLookup("MilestoneBilling", "ClientList", "ClientID=" & cbxClient))) = -1 Then
        InvPct = DLookup("DCJudgmentEnteredPct", "clientlist", "clientid=" & cbxClient)
      Else
        InvPct = 1
      End If
      If FeeAmount > 0 Then
        If InvPct <= 1 Then
          AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when judgement entered received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
        Else
          'AddInvoiceItem FileNumber, "FC-REF", "Attorney Fee- " & Format(InvPct, "percent") & " due when complaint filed received of " & Format(FeeAmount, "currency"), InvPct * FeeAmount, 0, True, True, False, False
        End If
      End If
  End Select
  
    DoCmd.SetWarnings False
    strinfo = "Judgment Entered"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery

Else
 AddStatus FileNumber, Now(), "Removed Judgment Entered Date"
End If
End If
'Forms!ForeclosureDetails!sfrmStatus.Requery
End Sub

Private Sub JudgmentEntered_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    JudgmentEntered = Now()
    Call JudgmentEntered_AfterUpdate
End If
End Sub

Private Sub ComplaintFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ComplaintFiled = Now()
    Call ComplaintFiled_AfterUpdate
End If
End Sub

Private Sub InitialHearingConference_AfterUpdate()
If Not IsNull(InitialHearingConference) Then

 AddStatus FileNumber, InitialHearingConference, "Initial Hearing Conference"
 InitialHearingConference.Locked = True

Else
 AddStatus FileNumber, Now(), "Removed Initial Hearing Conference date"
End If

'InitialHearingConference.Locked = True
End Sub

Private Sub InitialHearingConference_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    InitialHearingConference = Now()
    Call InitialHearingConference_AfterUpdate
End If
End Sub



Private Sub LisPendensFiled_AfterUpdate()
If Not IsNull(LisPendensFiled) Then
 AddStatus FileNumber, LisPendensFiled, "Lis Pendens Filed"
Else
 AddStatus FileNumber, Now(), "Removed LisPendens Filed date"
End If
End Sub

Private Sub LisPendensFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    LisPendensFiled = Now()
    Call LisPendensFiled_AfterUpdate
End If
End Sub

Private Sub MotiontoExtendFiled_AfterUpdate()
If Not IsNull(MotiontoExtendFiled) Then
 AddStatus FileNumber, MotiontoExtendFiled, "Motion to Extend Filed"


'Forms!ForeclosureDetails!sfrmStatus.Requery

DoCmd.SetWarnings False
strinfo = "Motion to Extend Filed"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
Forms!Journal.Requery

Else
 AddStatus FileNumber, Now(), "Removed Motion to Extend Filed"
End If

Forms!foreclosuredetails!sfrmStatus.Requery

If Not IsNull(MotiontoExtendFiled) Then
    MotiontoExtendFiled.Locked = True
End If

End Sub

Private Sub MotiontoExtendFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    MotiontoExtendFiled = Now()
    Call MotiontoExtendFiled_AfterUpdate
End If
End Sub

Private Sub PrepareDraft_AfterUpdate()
If Not IsNull(PrepareDraft) Then
 AddStatus FileNumber, PrepareDraft, "Prepare Draft Complaint"


'Forms!ForeclosureDetails!sfrmStatus.Requery

DoCmd.SetWarnings False
strinfo = "Prepared Draft Complaint"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
Forms!Journal.Requery

Else
 AddStatus FileNumber, Now(), "Removed Prepare Draft Complaint"
End If


End Sub

Private Sub PrepareDraft_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    PrepareDraft = Now()
    Call PrepareDraft_AfterUpdate
End If
End Sub

Private Sub ReceivedClientComplaintSinged_AfterUpdate()
If Not IsNull(ReceivedClientComplaintSinged) Then
 AddStatus FileNumber, ReceivedClientComplaintSinged, "Received Client Complaint Signed"
Else
 AddStatus FileNumber, Now(), "Removed Client Complaint Signed date"
End If
End Sub

Private Sub ReceivedClientComplaintSinged_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ReceivedClientComplaintSinged = Now()
    Call ReceivedClientComplaintSinged_AfterUpdate
End If
    
End Sub

Private Sub SentClientComplaint_AfterUpdate()
If Not IsNull(SentClientComplaint) Then
 AddStatus FileNumber, SentClientComplaint, "Sent Client Complaint"
Else
 AddStatus FileNumber, Now(), "Removed Sent Client Complaint"
End If
End Sub

Private Sub SentClientComplaint_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    SentClientComplaint = Now()
    Call SentClientComplaint_AfterUpdate
End If

End Sub

Private Sub SentComplaintToCourt_AfterUpdate()
If Not IsNull(SentComplaintToCourt) Then
 AddStatus FileNumber, SentComplaintToCourt, "Sent Complaint To Court"
Else
 AddStatus FileNumber, Now(), "Removed Complaint To Court date"
End If
End Sub

Private Sub SentComplaintToCourt_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    SentComplaintToCourt = Now()
    Call SentComplaintToCourt_AfterUpdate
End If
End Sub

Private Sub ServiceDeadline_AfterUpdate()
If Not IsNull(ServiceDeadline) Then
 AddStatus FileNumber, ServiceDeadline, "Service Deadline"
 
 'Forms!ForeclosureDetails!sfrmStatus.Requery

 DoCmd.SetWarnings False
    strinfo = "Service Deadline Date Entered"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
Else
 AddStatus FileNumber, Now(), "Removed Service Deadline"
End If
End Sub

Private Sub ServiceDeadline_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ServiceDeadline = Now()
    Call ServiceDeadline_AfterUpdate
End If
End Sub

Private Sub ServiceDue_AfterUpdate()

If Not IsNull(ServiceDue) Then

    If IsNull(MotiontoExtendFiled) Or IsNull(DC_Order) Then
        MsgBox ("Missing Motion Extend Filed date or order Extending date")
    Exit Sub
    End If
    
    If Not IsNull(MotiontoExtendFiled) And Not IsNull(DC_Order) Then
        ServiceDeadline = ServiceDue
    End If
    
 'AddStatus FileNumber, ServiceDue, "Service Due"
 AddStatus FileNumber, Now(), "Service due by " & Me.ServiceDue & ""

'Forms!ForeclosureDetails!sfrmStatus.Requery

DoCmd.SetWarnings False
    strinfo = "Service Due Date Entered"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
    
Else
 AddStatus FileNumber, Now(), "Removed Service Due"
End If
End Sub

Private Sub ServiceDue_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ServiceDue = Now()
    Call ServiceDue_AfterUpdate
End If
End Sub

Private Sub ServiceSent_AfterUpdate()
If Not IsNull(ServiceSent) Then
 AddStatus FileNumber, ServiceSent, "Service Sent"


'Forms!ForeclosureDetails!sfrmStatus.Requery

DoCmd.SetWarnings False
    strinfo = "Service Sent Date Entered"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
    
Else
 AddStatus FileNumber, Now(), "Removed Service Sent"
End If
End Sub

Private Sub ServiceSent_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    ServiceSent = Now()
    Call ServiceSent_AfterUpdate
End If
End Sub

Private Sub SummonsReceived_AfterUpdate()
If Not IsNull(SummonsReceived) Then
 AddStatus FileNumber, SummonsReceived, "Summons Received "
Else
 AddStatus FileNumber, Now(), "Removed SummonsReceived date"
End If
End Sub

Private Sub SummonsReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    SummonsReceived = Now()
    Call SummonsReceived_AfterUpdate
End If

End Sub

Private Sub UpdateCalendarDCHearing()

Dim emailGroup As String

'If Nz(ExceptionsHearingEntryID) = "X" Then Exit Sub

If IsNull(InitialHearingConference) And Not IsNull(DCHearingEntryID) Then
    Call DeleteCalendarEvent(DCHearingEntryID)
    DCHearingEntryID = Null
    Exit Sub
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

If Nz(DCHearingEntryID) = "" Then   ' new event on calendar
    DCHearingEntryID = AddCalendarEvent(InitialHearingConference + Nz(DCHearingTime, 0), IsNull(DCHearingTime), DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " Initial Hearing Conference " & " (" & FileNumber & ")", "District of Columbia, DC", 8, emailGroup)
    'UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " (" & FileNumber & ")" & " DC " & Trim(Me.HearingTypeID.Column(1)) & " ", "District of Columbia, DC", 8, emailGroup)
Else                                    ' change existing event on calendar
    Call UpdateCalendarEvent(DCHearingEntryID, InitialHearingConference + Nz(DCHearingTime, 0), IsNull(DCHearingTime), DLookup("ShortClientName", "qryClientAddress") & " v. " & Forms![Case List]!PrimaryDefName & " Initial Hearing Conference " & " (" & FileNumber & ")", "District of Columbia, DC", 8)
End If
    
End Sub

'Function IsWeekendDay(MyDate As Variant) As Boolean
'
'    'Dim DayNum As Variant
'    'Dim IsWeekend As Boolean
'
'    DayNum = Application.Weekday(MyDate)
'    If Not IsError(DayNum) Then
'        Select Case DayNum
'        Case 2 To 6 ' Monday thru Friday
'            IsWeekend = False
'        Case Else
'            IsWeekend = True
'        End Select
'    Else
'        IsWeekend = False ' error
'    End If
'    IsWeekendDay = IsWeekend
'End Function

