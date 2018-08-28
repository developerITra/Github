VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Exhibit 9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


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


Call DoReport("Declaration Wells", Me.OpenArgs)

'DoCmd.Close acForm, "Print Affidavit Collateral file"

'ShowForm ("Affidavit Collateral File")

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub




Private Sub Other_AfterUpdate()

If Other Then
OtherDocText = InputBox("Enter the Name of the document.", "Name of Other document")
Else
OtherDocText = "Other"
End If

End Sub
