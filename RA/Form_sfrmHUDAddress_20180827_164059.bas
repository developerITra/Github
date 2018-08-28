VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmHUDAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub txtCCAddress_AfterUpdate()
HUDAddress1Line = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity.Value & ", " & txtCCCState.Value & ". " & txtCCCZip.Value
End Sub

Private Sub txtCCAddress2_AfterUpdate()
HUDAddress1Line = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity.Value & ", " & txtCCCState.Value & ". " & txtCCCZip.Value
End Sub

Private Sub txtCCATTN_AfterUpdate()
HUDAddress = txtCCATTN
HUDAddress1Line = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity.Value & ", " & txtCCCState.Value & ". " & txtCCCZip.Value
End Sub

Private Sub txtCCCCity_AfterUpdate()
HUDAddress1Line = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity.Value & ", " & txtCCCState.Value & ". " & txtCCCZip.Value
End Sub

Private Sub txtCCCState_AfterUpdate()
HUDAddress1Line = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity.Value & ", " & txtCCCState.Value & ". " & txtCCCZip.Value
End Sub

Private Sub txtCCCZip_AfterUpdate()
HUDAddress1Line = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity.Value & ", " & txtCCCState.Value & ". " & txtCCCZip.Value
End Sub
