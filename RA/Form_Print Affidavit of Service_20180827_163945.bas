VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Affidavit of Service"
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


If (IsNull(Me.txtPrintDate)) Then
  MsgBox "Enter Print Date before continuing.", vbCritical
  Exit Sub
End If

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

Call DoReport("Order to Docket Affidavit", Me.OpenArgs, , Me.txtPrintDate)
cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub


Private Sub txtPrintDate_DblClick(Cancel As Integer)
  txtPrintDate = Date
End Sub
