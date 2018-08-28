VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmAuditorLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
If IsNull(Me.txtAuditorLetterReceived) = False Then
    Me.txtRespondDeadline.Enabled = True
    Me.txtResponseSent.Enabled = True
Else
    Me.txtRespondDeadline.Enabled = False
    Me.txtResponseSent.Enabled = False
End If

If IsLoaded("Case List") = True Then
    If Forms![Case List]!CaseTypeID = 8 Then
    Call SetObjectAttributes(txtRespondDeadline, True)
    Call SetObjectAttributes(txtResponseSent, True)
    End If
End If
End Sub

Private Sub txtAuditorLetterReceived_AfterUpdate()
    AddStatus FileNumber, Me.txtAuditorLetterReceived, "Auditor Letter Received"
    Me.txtRespondDeadline.Enabled = True
    Me.txtResponseSent.Enabled = True
End Sub

Private Sub txtAuditorLetterReceived_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Me.txtAuditorLetterReceived)
'if Not IsNull(Me.txtAuditorLetterReceived) Then
'    If (Me.txtAuditorLetterReceived < Date) Then
'      Cancel = -1
'      MsgBox "Date cannot be in the past.", vbCritical
'      Exit Sub
'    End If
'End If
End Sub

Private Sub txtAuditorLetterReceived_Click()
If Me.txtAuditorLetterReceived.Locked = True Then MsgBox ("You are not authorized to edit this field")
End Sub

Private Sub txtAuditorLetterReceived_DblClick(Cancel As Integer)

If FileReadOnly Or EditDispute Then
    DoCmd.CancelEvent
Else
    Me.txtAuditorLetterReceived = Now()
    Me.txtRespondDeadline.Enabled = True
    Me.txtResponseSent.Enabled = True
    AddStatus FileNumber, Me.txtAuditorLetterReceived, "Auditor Letter Received"
End If
End Sub

Private Sub txtRespondDeadline_AfterUpdate()
  AddStatus FileNumber, Me.txtRespondDeadline, "Deadline to respond to auditor set"
End Sub

Private Sub txtRespondDeadline_BeforeUpdate(Cancel As Integer)
If Not IsNull(Me.txtRespondDeadline) Then
    If (Me.txtRespondDeadline < Date) Then
      Cancel = -1
      MsgBox "Date cannot be in the past.", vbCritical
      Exit Sub
    End If
End If
End Sub

Private Sub txtRespondDeadline_Click()
If Me.txtRespondDeadline.Locked = True Then MsgBox ("You are not authorized to edit this field")
End Sub

Private Sub txtRespondDeadline_DblClick(Cancel As Integer)

If FileReadOnly Or EditDispute Then
    DoCmd.CancelEvent
Else
    Me.txtRespondDeadline = Now()
    AddStatus FileNumber, Me.txtRespondDeadline, "Deadline to respond to auditor set"
End If
End Sub

Private Sub txtResponseSent_AfterUpdate()
    AddStatus FileNumber, Me.txtResponseSent, "Auditor Response Sent"
End Sub

Private Sub txtResponseSent_BeforeUpdate(Cancel As Integer)
If Not IsNull(Me.txtResponseSent) Then
    If (Me.txtResponseSent < Date) Then
      Cancel = -1
      MsgBox "Date cannot be in the past.", vbCritical
      Exit Sub
    End If
End If
End Sub

Private Sub txtResponseSent_Click()
If Me.txtResponseSent.Locked = True Then MsgBox ("You are not authorized to edit this field")
End Sub

Private Sub txtResponseSent_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
    DoCmd.CancelEvent
Else
    Me.txtResponseSent = Now()
    AddStatus FileNumber, Me.txtResponseSent, "Auditor Response Sent"
End If
End Sub
