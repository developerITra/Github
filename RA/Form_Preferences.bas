VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Preferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub Form_Current()
Me.Caption = "Preferences for " & txtName
End Sub

Private Sub cmdTestLabel_Click()
On Error GoTo Err_cmdTestLabel_Click

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

Call StartLabel
Print #6, "|FONTSIZE 14"
Print #6, "Here is your test label!"
Print #6, "|FONTSIZE 8"
Print #6, "|BOTTOM"
Print #6, "Here is the fine print at the bottom."
Call FinishLabel
Me.Requery
MsgBox "Test label has been printed", vbInformation

Exit_cmdTestLabel_Click:
    Exit Sub

Err_cmdTestLabel_Click:
    MsgBox Err.Description
    Resume Exit_cmdTestLabel_Click
    
End Sub
