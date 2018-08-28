VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmMediation_DC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database




Private Sub ckMediationRequested_AfterUpdate()
If ckMediationRequested Then
     
   AddStatus FileNumber, Now(), "mediation requested "
   AddInvoiceItem FileNumber, "FC-MED", "DC Mediation Fee", DLookup("DCMediationFee", "clientlist", "clientid=" & Forms![Case List]!ClientID), 0, True, True, False, False
     
  
    DoCmd.SetWarnings False
    strinfo = "mediation requested on " & Date & " by borrower"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery

    Call Showinfo
Else
    Call ShowinfoOFF
End If

End Sub

Private Sub DCBorrowerDocs_Complete_AfterUpdate()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent

Else
    'DCBorrowerDocs_Complete = Now()
End If

If Not IsNull(DCBorrowerDocs_Complete) Then

AddStatus FileNumber, Now(), "Borrower docs were completed on " & DCBorrowerDocs_Complete & " "
     
    DoCmd.SetWarnings False
    strinfo = "Borrower docs were completed on " & DCBorrowerDocs_Complete & " "
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If
End Sub

Private Sub DCBorrowerDocs_Complete_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DCBorrowerDocs_Complete = Now()
    Call DCBorrowerDocs_Complete_AfterUpdate
End If

End Sub

Private Sub DCBorrowerDocs_Due_AfterUpdate()

If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent

Else
    'DCBorrowerDocs_Due = Now()
End If

If Not IsNull(DCBorrowerDocs_Due) Then

AddStatus FileNumber, Now(), "Borrower docs are due on " & DCBorrowerDocs_Due & " "
     
    DoCmd.SetWarnings False
    strinfo = "Borrower docs are due on " & DCBorrowerDocs_Due & " "
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If


End Sub

Private Sub DCBorrowerDocs_Due_DblClick(Cancel As Integer)

If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DCBorrowerDocs_Due = Now()
    Call DCBorrowerDocs_Due_AfterUpdate
End If

End Sub

Private Sub DCCSS_Complete_AfterUpdate()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    'DCCSS_Complete = Now()
  
End If


If Not IsNull(DCCSS_Complete) Then

AddStatus FileNumber, Now(), "CSS completed on " & DCCSS_Complete & " "
     
    DoCmd.SetWarnings False
    strinfo = "CSS completed on " & DCCSS_Complete & " "
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If
End Sub

Private Sub DCCSS_Complete_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DCCSS_Complete = Now()
    Call DCCSS_Complete_AfterUpdate
End If
End Sub

Private Sub DCCSS_Due_AfterUpdate()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    
    'DCStmtInitialReview_Due = Now()
  
End If


If Not IsNull(DCCSS_Due) Then

AddStatus FileNumber, Now(), "CSS due on " & DCCSS_Due & " "
     
    DoCmd.SetWarnings False
    strinfo = "CSS due on " & DCCSS_Due & " "
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If
End Sub

Private Sub DCCSS_Due_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DCCSS_Due = Now()
    Call DCStmtInitialReview_Due_AfterUpdate
End If
End Sub

Private Sub DCStmtInitialReview_Complete_AfterUpdate()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    'DCStmtInitialReview_Complete = Now()
  
End If

If Not IsNull(DCStmtInitialReview_Complete) Then

AddStatus FileNumber, Now(), "Statement of Initial Review was completed on " & DCStmtInitialReview_Complete & " "
     
    DoCmd.SetWarnings False
    strinfo = "Statement of Initial Review as completed on " & DCStmtInitialReview_Complete & " "
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If

End Sub

Private Sub DCStmtInitialReview_Complete_DblClick(Cancel As Integer)

If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DCStmtInitialReview_Complete = Now()
Call DCStmtInitialReview_Complete_AfterUpdate
End If

End Sub

Private Sub DCStmtInitialReview_Due_AfterUpdate()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    'DCCSS_Due = Now()
  
End If

If Not IsNull(DCStmtInitialReview_Due) Then

AddStatus FileNumber, Now(), "Statement of Initial Review due on " & DCStmtInitialReview_Due & " "
     
    DoCmd.SetWarnings False
    strinfo = "Statement of Initial Review due on " & DCStmtInitialReview_Due & " "
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
End If
End Sub

Private Sub DCStmtInitialReview_Due_DblClick(Cancel As Integer)

If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DCStmtInitialReview_Due = Now
    Call DCStmtInitialReview_Due_AfterUpdate
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
If ckMediationRequested Then
    Call Showinfo
Else
    Call ShowinfoOFF
End If
End Sub

Private Sub rerecorded_AfterUpdate()

If Not IsNull(Rerecorded) Then
 AddStatus FileNumber, Rerecorded, "Deed re-recorded on " & Format(Rerecorded, "mm/dd/yyyy")
Else
 AddStatus FileNumber, Now(), "Removed Deed re-recorded date"
End If

End Sub

Private Sub rerecorded_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    Rerecorded = Now()
    Call rerecorded_AfterUpdate
End If
End Sub



Public Sub Showinfo()

Me.DCBorrowerDocs_Due.Enabled = True
Me.DCStmtInitialReview_Due.Enabled = True
Me.DCCSS_Due.Enabled = True
Me.DCBorrowerDocs_Complete.Enabled = True
Me.DCStmtInitialReview_Complete.Enabled = True
Me.DCCSS_Complete.Enabled = True


Me.lbBorrowerDocs.Enabled = True
Me.lbCSS.Enabled = True
Me.lbstmeInitialReview.Enabled = True

Forms!foreclosuredetails![sfrmLMHearing_DC].Enabled = True

End Sub

Public Sub ShowinfoOFF()

Me.DCBorrowerDocs_Due.Enabled = False
Me.DCStmtInitialReview_Due.Enabled = False
Me.DCCSS_Due.Enabled = False
Me.DCBorrowerDocs_Complete.Enabled = False
Me.DCStmtInitialReview_Complete.Enabled = False
Me.DCCSS_Complete.Enabled = False


Me.lbBorrowerDocs.Enabled = False
Me.lbCSS.Enabled = False
Me.lbstmeInitialReview.Enabled = False


Forms!foreclosuredetails![sfrmLMHearing_DC].Enabled = False

End Sub


