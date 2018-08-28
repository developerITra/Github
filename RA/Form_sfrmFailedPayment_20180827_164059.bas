VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmFailedPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Escrow_AfterUpdate()
MontlyP.Value = PandI.Value + Escrow.Value
TotalP.Value = NumMissed * MontlyP.Value
End Sub

Private Sub MontlyP_AfterUpdate()
TotalP.Value = NumMissed * MontlyP.Value
End Sub

Private Sub NumMissed_AfterUpdate()
TotalP.Value = NumMissed * MontlyP.Value
End Sub

Private Sub PandI_AfterUpdate()
MontlyP.Value = PandI.Value + Escrow.Value
TotalP.Value = NumMissed * MontlyP.Value
End Sub

