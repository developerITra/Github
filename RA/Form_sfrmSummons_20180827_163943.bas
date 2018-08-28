VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmSummons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub AffidavitToCourt_AfterUpdate()
AddStatus FileNumber, AffidavitToCourt, "Affidavit to Court"
End Sub

Private Sub AffidavitToCourt_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
AffidavitToCourt = Date
AddStatus FileNumber, AffidavitToCourt, "Affidavit to Court"
End If

End Sub

Private Sub AnswerDue_AfterUpdate()
AddStatus FileNumber, Date, "Answer is due " & Format(AnswerDue, "m/d/yyyy")
End Sub

Private Sub AnswerFiled_AfterUpdate()
AddStatus FileNumber, AnswerFiled, "Answer Filed"
End Sub

Private Sub AnswerFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
AnswerFiled = Date
AddStatus FileNumber, AnswerFiled, "Answer Filed"
End If

End Sub

Private Sub HearingDate_AfterUpdate()
AddStatus FileNumber, Date, "Hearing scheduled for " & Format(HearingDate, "m/d/yyyy h:nn am/pm")
End Sub

Private Sub NoAnswerFiled_AfterUpdate()
If NoAnswerFiled Then AddStatus FileNumber, Date, "No answer filed"
End Sub

Private Sub ServiceDate_AfterUpdate()
AddStatus FileNumber, ServiceDate, "Served"
End Sub

Private Sub ServiceDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ServiceDate = Date
AddStatus FileNumber, ServiceDate, "Served"
End If

End Sub

Private Sub ServiceDeadline_AfterUpdate()
AddStatus FileNumber, Date, "Service Deadline is " & Format(ServiceDeadline, "m/d/yyyy")
End Sub

Private Sub SummonsReceived_AfterUpdate()
AddStatus FileNumber, SummonsReceived, "Summons Received"
End Sub

Private Sub SummonsReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
SummonsReceived = Date
AddStatus FileNumber, SummonsReceived, "Summons Received"
End If

End Sub

