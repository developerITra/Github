VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
On Error GoTo Err_cmdOK_Click

If IsNull(lstFiles) Then
    MsgBox "Select a file", vbCritical
    Exit Sub
End If

OpenCase lstFiles

On Error GoTo 0
Forms![Case List].SetFocus
'On Error GoTo Err_cmdOK_Click

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
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

Private Sub lstFiles_DblClick(Cancel As Integer)
Call cmdOK_Click
End Sub
