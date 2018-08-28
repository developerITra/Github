VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmFeesCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Current()
If FileReadOnly Or EditDispute Then
   Me.AllowEdits = False
End If

End Sub


Private Sub txtFeesCostRequested_AfterUpdate()

If txtFeesCostRequested > Date Then
MsgBox (" Date cannot be in the future")
txtFeesCostRequested = Null
Exit Sub
End If

If Not IsNull(txtFeesCostRequested) Then
 AddStatus FileNumber, txtFeesCostRequested, "Fees & Costs Requested"
 
 FeeCostRequestedStaffID = GetStaffID
 FeeCostRequestedStaffName = GetFullName()
 
 DoCmd.SetWarnings False
    strinfo = "Fees & Costs Requested"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
    
    
Else
 AddStatus FileNumber, Now(), "Removed Fees & Costs Requested date"
End If
End Sub

Private Sub txtFeesCostRequested_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    txtFeesCostRequested = Now()
    Call txtFeesCostRequested_AfterUpdate
End If

End Sub

Private Sub txtFeesCostSent_AfterUpdate()
If txtFeesCostSent > Date Then
MsgBox (" Date cannot be in the future")
txtFeesCostSent = Null
Exit Sub
End If

If Not IsNull(txtFeesCostSent) Then
 AddStatus FileNumber, txtFeesCostSent, "Fees & Costs Sent"
 
FeeCostSentStaffID = GetStaffID
FeeCostSentStaffName = GetFullName()

DoCmd.SetWarnings False
    strinfo = "Fees & Costs Sent"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
Else
 AddStatus FileNumber, Now(), "Removed Fees & Costs Sent date"
End If
End Sub

Private Sub txtFeesCostSent_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    txtFeesCostSent = Now()
    Call txtFeesCostSent_AfterUpdate
End If

End Sub
