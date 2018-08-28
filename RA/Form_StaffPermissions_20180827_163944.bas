VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_StaffPermissions"
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

Private Sub cmdStaff_Click()

On Error GoTo Err_cmdStaff_Click
DoCmd.OpenForm "Staff"

Exit_cmdStaff_Click:
    Exit Sub

Err_cmdStaff_Click:
    MsgBox Err.Description
    Resume Exit_cmdStaff_Click
    
End Sub

Private Sub Form_Current()
  If Not IsNull(Me.cbxSelect) Then
    Me.lblPermissionStaff.Caption = "* Changes to " & Me.cbxSelect.Column(1) & " permissions will take effect after " & Me.cbxSelect.Column(1) & " exits and logs back into system."
  Else
    Me.lblPermissionStaff.Caption = ""
  End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Call UpdateList

Dim C As Control
If DLookup("PrivPermission", "Staff", "name = '" & GetLoginName() & "'") Then

'Dim mm As Balloon
'mm = PrivPermission

For Each C In Me.Controls
    If C.ControlType = acCheckBox Or C.ControlType = acOptionGroup Then C.Locked = False
Next
Else
'Dim C As Control
For Each C In Me.Controls
    If C.ControlType = acCheckBox Or C.ControlType = acOptionGroup Then C.Locked = True
Next
End If


End Sub

Private Sub optActive_AfterUpdate()
Call UpdateList
End Sub

Private Sub UpdateList()
If optActive Then
    cbxSelect.RowSource = "SELECT Staff.* FROM Staff WHERE Active AND Initials Is Not Null ORDER BY Staff.Sort;"
Else
    cbxSelect.RowSource = "SELECT Staff.* FROM Staff WHERE Initials Is Not Null ORDER BY Staff.Sort;"
End If
Me.Requery
End Sub

Private Sub cbxSelect_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & str(Nz(Me![cbxSelect], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
End Sub

Private Sub cmdReport_Click()

On Error GoTo Err_cmdReport_Click
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenReport "Permissions", acViewPreview

Exit_cmdReport_Click:
    Exit Sub

Err_cmdReport_Click:
    MsgBox Err.Description
    Resume Exit_cmdReport_Click
    
End Sub
