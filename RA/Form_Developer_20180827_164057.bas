VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Developer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click
DoCmd.Close

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdDeveloper_Click()
On Error GoTo Err_cmdDeveloper_Click

Call CreateDSN(True)
txtInfo = DeveloperInfo()

Exit_cmdDeveloper_Click:
    Exit Sub

Err_cmdDeveloper_Click:
    MsgBox Err.Description
    Resume Exit_cmdDeveloper_Click
    
End Sub

Private Sub cmdRegular_Click()
On Error GoTo Err_cmdRegular_Click

Call CreateDSN(False)
txtInfo = DeveloperInfo()

Exit_cmdRegular_Click:
    Exit Sub

Err_cmdRegular_Click:
    MsgBox Err.Description
    Resume Exit_cmdRegular_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
txtInfo = DeveloperInfo()
End Sub
