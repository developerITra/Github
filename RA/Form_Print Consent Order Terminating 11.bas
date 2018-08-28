VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Consent Order Terminating 11"
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
On Error GoTo Err_cmdOK_Click

Select Case Forms!BankruptcyDetails!District
    Case 8, 9, 10, 11, 18
        Call DoReport("ConsentTerminating11VAWD", Me.OpenArgs)
    Case Else
        Call DoReport("ConsentTerminating11" & IIf(DocumentFormat = 1, "VAED", ""), Me.OpenArgs)
End Select
'Call DoReport("Debt", Me.OpenArgs)
'If State = "VA" Then Call DoReport("Consent Order Virginia Endorsement", Me.OpenArgs)
cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

