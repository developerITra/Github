VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Withdraw Sale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdOK_Click()
Dim statusMsg As String

On Error GoTo Err_cmdOK_Click
Call DoReport("Withdraw Sale", Me.OpenArgs)
Call DoReport("Withdraw Sale Order", Me.OpenArgs)
If MsgBox("Add to status: ""Motion to Withdraw Report of Sale""?", vbYesNo) = vbYes Then
    AddStatus FileNumber, Now(), "Motion to Withdraw Report of Sale"
End If
cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub
