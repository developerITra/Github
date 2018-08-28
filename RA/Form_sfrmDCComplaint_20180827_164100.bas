VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmDCComplaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub bondPosted_AfterUpdate()
If Not IsNull(BondPosted) Then
 AddStatus FileNumber, BondPosted, "Bond Filed"
    
    DoCmd.SetWarnings False
    strinfo = "Bond Filed"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery

 
 
Else
 AddStatus FileNumber, Now(), "Removed Bond Filed Date"
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

Private Sub ComplaintFiled_AfterUpdate()
If BHproject Then
    If Not IsNull(ComplaintFiled) Then
     AddStatus FileNumber, ComplaintFiled, "Complaint Filed"
    End If
Else

If Not IsNull(ComplaintFiled) Then
 AddStatus FileNumber, ComplaintFiled, "Complaint Filed"
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

Private Sub Form_Current()
If FileReadOnly Or EditDispute Then
   Me.AllowEdits = False
End If

End Sub

Private Sub JudgmentEntered_AfterUpdate()
If BHproject Then
If Not IsNull(JudgmentEntered) Then
 AddStatus FileNumber, JudgmentEntered, "Judgment Filed"
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
Else
 AddStatus FileNumber, Now(), "Removed Judgment Filed date"
End If
End If

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
Else
 AddStatus FileNumber, Now(), "Removed Initial Hearing Conference date"
End If
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
