VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Assignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmd_GrantorInvestor_Click()
On Error GoTo Err_cmdGrantorInvestor_Click
Grantor = "Mortgage Electronic Registration Systems, Inc. as nominee for <beneficiary> its successors or assigns"
GrantorAddress = "P.O. Box 2026, Flint, MI 48501-2026"

Exit_cmdGrantorInvestor_Click:
    Exit Sub

Err_cmdGrantorInvestor_Click:
    MsgBox Err.Description
    Resume Exit_cmdGrantorInvestor_Click
    
End Sub

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


If Forms![Case List]!State = "VA" Then
    Call DoReport("Assignment VA", Me.OpenArgs)
Else
    Call DoReport("Assignment", Me.OpenArgs)
End If

If Forms![Case List]!ClientID = 446 Then Call DoReport("BOA Cover Sheet Assignment", Me.OpenArgs)



cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdGranteeInvestor_Click()

On Error GoTo Err_cmdGranteeInvestor_Click
Grantee = Investor
GranteeAddress = OneLine(InvestorAddress)

Exit_cmdGranteeInvestor_Click:
    Exit Sub

Err_cmdGranteeInvestor_Click:
    MsgBox Err.Description
    Resume Exit_cmdGranteeInvestor_Click
    
End Sub


Private Sub Form_Load()

If Forms![Case List]!ClientID = 446 Then Me.cboBOANoteLocation.Visible = True

End Sub
