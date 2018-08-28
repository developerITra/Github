VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmCheckRequestCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub ChargeCleared_AfterUpdate()
  DoCmd.RunCommand acCmdSaveRecord
  
  Me.txtTotal.Requery
End Sub




Private Sub StatusID_AfterUpdate()
  DoCmd.RunCommand acCmdSaveRecord
  
  CompleteUpdate
End Sub

Private Sub CompleteUpdate()
 If (StatusID.Column(1) <> "Pending") Then
    CompletedDate = Date
    CompletedBy = GetStaffID()
 End If
  
End Sub

