VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmIRSAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub txtCCAddress_AfterUpdate()
IRSAddress1Line = IRSName.Value & " ATTN: " & txtCCATTN & ", " & txtCCAddress & ", " & txtCCAddress2 & ", " & txtCCCCity & ", " & txtCCCState & ". " & txtCCCZip

End Sub

Private Sub txtCCAddress2_AfterUpdate()
IRSAddress1Line = IRSName.Value & " ATTN: " & txtCCATTN & ", " & txtCCAddress & ", " & txtCCAddress2 & ", " & txtCCCCity & ", " & txtCCCState & ". " & txtCCCZip

End Sub

Private Sub txtCCATTN_AfterUpdate()
IRSAddress = "Internal Revenue Service"
IRSName = "Internal Revenue Service"
IRSAddress1Line = IRSName.Value & " ATTN: " & txtCCATTN & ", " & txtCCAddress & ", " & txtCCAddress2 & ", " & txtCCCCity & ", " & txtCCCState & ". " & txtCCCZip

End Sub

Private Sub txtCCCCity_AfterUpdate()
IRSAddress1Line = IRSName.Value & " ATTN: " & txtCCATTN & ", " & txtCCAddress & ", " & txtCCAddress2 & ", " & txtCCCCity & ", " & txtCCCState & ". " & txtCCCZip

End Sub

Private Sub txtCCCState_AfterUpdate()
IRSAddress1Line = IRSName.Value & " ATTN: " & txtCCATTN & ", " & txtCCAddress & ", " & txtCCAddress2 & ", " & txtCCCCity & ", " & txtCCCState & ". " & txtCCCZip

End Sub

Private Sub txtCCCZip_AfterUpdate()
IRSAddress1Line = IRSName.Value & " ATTN: " & txtCCATTN & ", " & txtCCAddress & ", " & txtCCAddress2 & ", " & txtCCCCity & ", " & txtCCCState & ". " & txtCCCZip

End Sub
