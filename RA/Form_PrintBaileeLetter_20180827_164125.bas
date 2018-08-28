VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PrintBaileeLetter"
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


'MsgBox Me.OpenArgs
'
DoCmd.SetWarnings False
DoCmd.RunCommand acCmdSaveRecord
DoCmd.SetWarnings True

Call DoReport("Ditech Bailee Letter", Me.OpenArgs)


 DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Current()
If IsNull(FileNumber) Then FileNumber = [Forms]![foreclosuredetails]![FileNumber]
End Sub
