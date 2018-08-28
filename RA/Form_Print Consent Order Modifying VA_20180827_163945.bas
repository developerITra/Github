VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Consent Order Modifying VA"
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

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
'Call DoReport("ConsentModifying13", Me.OpenArgs)

If (TestDocument("ConsentModifyingVA") = True) Then
  Call Doc_ConsentModifyingReliefVA(True)
End If

' Call DoReport("ConsentModifying" & Chapter & IIf(DocumentFormat = 1, "VAED", ""), Me.OpenArgs)
'If Judge = "RGM" Then Call DoReport("Consent Order Modifying RGM Addendum", Me.OpenArgs)
'If State = "VA" Then Call DoReport("Consent Order Virginia Endorsement", Me.OpenArgs)
'Call DoReport("Debt", Me.OpenArgs)

cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Current()
If IsNull(Me.ConsentOrderPaymentTo) Then
    If IsNull(BKPaymentsTo) Then
        ConsentOrderPaymentTo = Investor & vbNewLine & "Attn: Bankruptcy Department" & vbNewLine & InvestorAddress
    Else
        ConsentOrderPaymentTo = Investor & vbNewLine & BKPaymentsTo
    End If
End If
End Sub
