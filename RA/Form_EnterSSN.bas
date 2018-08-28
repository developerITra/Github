VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterSSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click

If Me.NewRecord And Not IsNull(Forms!EnterSSN!SSN) Then
SSNChange = True
SSNContainer = Forms!EnterSSN!SSN
Forms!sfrmNamesUpdate!SSN = SSNContainer
End If

If IsNull(SSN.OldValue) And Not IsNull(SSN) Then
SSNChange = True
End If

DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub Command3_Click()
Me.Undo
DoCmd.Close

End Sub

Private Sub Form_Current()
txtName = Forms!sfrmNamesUpdate!First & "     " & Forms!sfrmNamesUpdate!Last
End Sub
