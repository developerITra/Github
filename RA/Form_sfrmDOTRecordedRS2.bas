VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmDOTRecordedRS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub LoanMod_AfterUpdate()

If Not IsNull(LoanMod) Then
 AddStatus FileNumber, LoanMod, "Loan Mod recorded on " & Format(LoanMod, "mm/dd/yyyy")
Else
 AddStatus FileNumber, Now(), "Removed Loan Mod date"
End If

End Sub

Private Sub LoanMod_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    LoanMod = Now()
    Call LoanMod_AfterUpdate
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
