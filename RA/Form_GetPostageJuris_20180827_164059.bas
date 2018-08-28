VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetPostageJuris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
Static Vendor As Integer
On Error GoTo Err_cmdOK_Click

If Nz(txtTotal) <= 0 Then
    MsgBox "Total amount must be greater than zero", vbCritical
    
    Exit Sub
End If
AddInvoiceItem Forms![Case List]!FileNumber, txtProcess, txtDesc, txtTotal, Frame15, False, True, False, True

DoCmd.Close

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
PostageAmount = 0
'txtAmount = Nz(DLookup("IValue", "DB", "Name='" & Split(Me.OpenArgs, "|")(0) & "'")) / 100
lblPrompt.Caption = Split(Me.OpenArgs, "|")(0)
txtProcess = Split(Me.OpenArgs, "|")(1)
txtDesc = Split(Me.OpenArgs, "|")(2)
End Sub

Private Sub Option26_GotFocus()
MsgBox "Please confirm you are using UNITED PARCEL SERVICE, not the US Postal Service", vbExclamation
End Sub


