VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmVAAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub txtCCAddress_AfterUpdate()
VAAddress1Line.Value = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity & ", " & txtCCCState.Value & ". " & txtCCCZip

End Sub

Private Sub txtCCAddress2_AfterUpdate()
VAAddress1Line.Value = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity & ", " & txtCCCState.Value & ". " & txtCCCZip

End Sub

Private Sub txtCCATTN_AfterUpdate()
VAAddress.Value = txtCCATTN
VAAddress1Line.Value = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity & ", " & txtCCCState.Value & ". " & txtCCCZip

End Sub

Private Sub txtCCCCity_AfterUpdate()
VAAddress1Line.Value = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity & ", " & txtCCCState.Value & ". " & txtCCCZip

End Sub

Private Sub txtCCCState_AfterUpdate()
VAAddress1Line.Value = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity & ", " & txtCCCState.Value & ". " & txtCCCZip

End Sub

Private Sub txtCCCZip_AfterUpdate()
VAAddress1Line.Value = txtCCATTN.Value & ", " & txtCCAddress.Value & ", " & txtCCAddress2.Value & ", " & txtCCCCity & ", " & txtCCCState.Value & ". " & txtCCCZip

End Sub
