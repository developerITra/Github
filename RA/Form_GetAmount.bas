VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click

'If Nz(txtTotal) <= 0 Then
'    MsgBox "Amount must be greater than zero", vbCritical
'    txtTotal.SetFocus
'    Exit Sub
'End If
FeeAmount = txtTotal
DoCmd.Close

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
FeeAmount = 0
lblPrompt.Caption = Me.OpenArgs
txtTotal.SetFocus
End Sub

