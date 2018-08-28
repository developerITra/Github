VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdClose_Click()
On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdManageDepartments_Click()
On Error GoTo Err_cmdManageDepartments_Click

  Dim strStaffName As String
  
  If (IsNull([ID])) Then
    MsgBox "Please enter name details before continuing.", vbCritical, "Staff Details"
    Exit Sub
  End If
  
  strStaffName = txtName
  DoCmd.OpenForm "sfrmStaffDepartments", , , "[StaffID] = " & [ID], , , strStaffName
  

Exit_cmdManageDepartments_Click:
  Exit Sub
  
Err_cmdManageDepartments_Click:
  MsgBox Err.Description
  Resume Exit_cmdManageDepartments_Click
End Sub

Private Sub Name_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub cmdPassword_Click()
On Error GoTo Err_cmdPassword_Click

  If (IsNull([ID])) Then
    MsgBox "Please enter staff details before continuing.", vbCritical, "Staff Details"
    Exit Sub
  End If
  
  

  DoCmd.OpenForm "ChangePassword", , , "[ID] = " & [ID]

Exit_cmdPassword_Click:
  Exit Sub
  
Err_cmdPassword_Click:
  MsgBox Err.Description
  Resume Exit_cmdPassword_Click
End Sub
