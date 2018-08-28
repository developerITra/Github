VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Districts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub Combo4_AfterUpdate()
    ' Find the record that matches the control.
    Me.RecordsetClone.FindFirst "[ID] = " & Me![Combo4]
    Me.Bookmark = Me.RecordsetClone.Bookmark
End Sub

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdAdd_Click()

On Error GoTo Err_cmdAdd_Click
DoCmd.GoToRecord , , acNewRec

Exit_cmdAdd_Click:
    Exit Sub

Err_cmdAdd_Click:
    MsgBox Err.Description
    Resume Exit_cmdAdd_Click
    
End Sub

Private Sub cmdDelete_Click()

On Error GoTo Err_cmdDelete_Click
DoCmd.RunCommand acCmdDeleteRecord

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
    
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
If Not PrivAdmin Then
    Cancel = 1
    Me.Undo
    Call Combo4_AfterUpdate
    MsgBox "You are not authorized to make changes", vbCritical
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Me.AllowAdditions = PrivAdmin
Me.AllowDeletions = PrivAdmin
cmdDelete.Enabled = PrivAdmin
cmdAdd.Enabled = PrivAdmin
End Sub
