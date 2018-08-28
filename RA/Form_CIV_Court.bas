VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CIV_Court"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub cbxSelect_AfterUpdate()
    ' Find the record that matches the control.
    Me.RecordsetClone.FindFirst "[CourtID] = " & Me![cbxSelect]
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

Private Sub cmdNew_Click()

On Error GoTo Err_cmdNew_Click
DoCmd.GoToRecord , , acNewRec

Exit_cmdNew_Click:
    Exit Sub

Err_cmdNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdNew_Click
    
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
If Not PrivAdmin Then
    Cancel = 1
    Me.Undo
    Call cbxSelect_AfterUpdate
    MsgBox "You are not authorized to make changes", vbCritical
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Me.AllowAdditions = PrivAdmin
Me.AllowDeletions = PrivAdmin
cmdNew.Enabled = PrivAdmin
End Sub
