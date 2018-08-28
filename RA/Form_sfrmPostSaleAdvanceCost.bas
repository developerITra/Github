VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmPostSaleAdvanceCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub txtCKRequest_Click()
  If Description = "" Or IsNull(Description) Or Amount = 0 Or IsNull(Amount) Or Amount = "" Or IsNull(txtVendor) Then
    MsgBox ("Please enter data")
  Exit Sub
  End If
  
  'DoCmd.OpenForm "Add Check Request"
  DoCmd.OpenForm "CheckRequest_PostSaleAdvanceCost"

End Sub

Private Sub txtCKRequest1_Click()

End Sub

Private Sub txtVendor_AfterUpdate()
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
Me.Requery
End Sub
