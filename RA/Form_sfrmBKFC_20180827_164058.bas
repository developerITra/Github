VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmBKFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TitleAssignSentToRecordDate_AfterUpdate()

  AddStatus FileNumber, TitleAssignSentToRecordDate, "Assignment Sent To Record Date"
  

End Sub

Private Sub TitleAssignSentToRecordDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
  TitleAssignSentToRecordDate = Now()
  Call TitleAssignSentToRecordDate_AfterUpdate
End If

  
End Sub

Private Sub TitleAssignNeededdate_AfterUpdate()
 AddStatus FileNumber, TitleAssignNeededDate, "Assignment Needed"
 If (Not IsNull(TitleAssignNeededDate)) Then
 AddInvoiceItem FileNumber, "BK/POC-Assign", "Assignment Drafted", GetFeeAmount("Assignment Drafted"), 0, False, True, False, True
End If
End Sub

Private Sub TitleAssignNeededdate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
TitleAssignNeededDate = Now()
Call TitleAssignNeededdate_AfterUpdate
End If

End Sub

Private Sub TitleAssignReceivedDate_AfterUpdate()

 AddStatus FileNumber, TitleAssignReceivedDate, "Assignment Received from Client"
 
 If Not IsNull(TitleAssignReceivedDate) Then
  AddInvoiceItem FileNumber, "BK/POC-Assi-Sent-Court", "Assignment Sent to Court", GetFeeAmount("Assignment Sent"), 0, False, True, False, True
End If


End Sub

Private Sub TitleAssignReceivedDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
TitleAssignReceivedDate = Now()
Call TitleAssignReceivedDate_AfterUpdate
End If

End Sub

Private Sub TitleAssignRecordedDate_AfterUpdate()
AddStatus FileNumber, TitleAssignRecordedDate, "Assignment Sent to Record"

End Sub

Private Sub TitleAssignRecordedDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
TitleAssignRecordedDate = Now()
Call TitleAssignRecordedDate_AfterUpdate
End If

End Sub

Private Sub TitleAssignSentdate_AfterUpdate()
If (Not IsNull(TitleAssignSentdate)) Then
 AddStatus FileNumber, TitleAssignSentdate, "Assignment Requested From Client"
End If
End Sub

Private Sub TitleAssignSentdate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
TitleAssignSentdate = Now()
Call TitleAssignSentdate_AfterUpdate
End If

End Sub
