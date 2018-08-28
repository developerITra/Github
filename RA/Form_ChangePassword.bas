VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Compare Database   ' incompatible with SELECT CASE in Subroutine ShowPasswordInfo
Option Explicit

Private Sub cmdCancel_Click()
On Error GoTo Err_cmdCancel_Click

If (getEntryPoint() = "Login") Then
  MsgBox "Cancel will close database.  Please re-open database to reset password.", vbOKOnly, "Cancel Password Reset"
  DoCmd.Quit
Else
  DoCmd.Close
End If

Exit_cmdCancel_Click:
  Exit Sub
  
Err_cmdCancel_Click:
  MsgBox Err.Description
  Resume Exit_cmdCancel_Click

End Sub

Private Sub cmdMorePasswords_Click()
    txtPasswordInfo = ""
    GeneratePasswords
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_cmdOK_Click

Dim rstUser As Recordset

If (IsNull(Me.lstPassword)) Then
  MsgBox "Please select a new password from the list.", vbCritical, "Select Password"
  Exit Sub
End If

If Me.lstPassword <> Nz(Me.txtNewPassword) Then
  MsgBox "New password doesn't match the one selected in the list.", vbCritical, "Select Password"
  txtNewPassword.SetFocus
  Exit Sub
End If

If (Me.txtOldPassword.Visible = True) Then   ' there was an old password
  If IsNull(Me.txtOldPassword) Then
    MsgBox "Please enter old password.", vbCritical, "Enter old password"
    txtOldPassword.SetFocus
    Exit Sub
  End If
End If

Set rstUser = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID=" & StaffID, dbOpenDynaset, dbSeeChanges)
If Not rstUser.EOF Then

  If (Me.txtOldPassword.Visible = True) Then
  
    If (Nz(rstUser!Password) <> DigestStrToHexStr(txtOldPassword & "Rosenberg " & rstUser!Username)) Then
      MsgBox "Old password is not correct.  Please re-enter.", vbOKOnly, "Old Password"
      rstUser.Close
      txtOldPassword.SetFocus
      Exit Sub
    End If
  End If
  
  rstUser.Edit
  
  rstUser!Password = DigestStrToHexStr(Me.lstPassword & "Rosenberg " & rstUser!Username)
  rstUser!PasswordExpires = DateAdd("m", 6, Now())
  
  rstUser.Update
  
  MsgBox "Password has been successfully updated.", vbOKOnly, "Password Updated"
    
Else
    rstUser.Close
    MsgBox "User context is lost, database is closing.", vbCritical
    DoCmd.Close acForm, Me.Name
    DoCmd.Quit
    
End If
rstUser.Close

If (getEntryPoint() = "Login") Then
  DoCmd.OpenForm "Main"
End If

DoCmd.Close acForm, Me.Name

Exit_cmdOK_Click:
  Exit Sub
  
Err_cmdOK_Click:
  MsgBox Err.Description
  Resume Exit_cmdOK_Click

End Sub

Private Sub Form_Open(Cancel As Integer)
Dim rstUser As Recordset

On Error GoTo Err_Form_Open

Set rstUser = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID=" & StaffID, dbOpenDynaset, dbSeeChanges)
If Not rstUser.EOF Then
  If (IsNull(rstUser!Password)) Then
    'lblInfo.Caption = "Please select new password below."
    lblOldPassword.Visible = False
    txtOldPassword.Visible = False
  Else
    'lblInfo.Caption = "Please enter old password and select new password."
    lblOldPassword.Visible = True
    txtOldPassword.Visible = True
    
  End If
    
Else
    rstUser.Close
    MsgBox "User context is lost, please exit the database", vbCritical
    DoCmd.Close acForm, Me.Name
    DoCmd.Quit
    
End If

rstUser.Close

GeneratePasswords

Exit_Form_Open:
  Exit Sub
  
Err_Form_Open:
  MsgBox Err.Description
  Resume Exit_Form_Open
End Sub

Public Sub GeneratePasswords()
On Error GoTo Err_GeneratePasswords

Dim i As Integer

Me.lstPassword.RowSource = ""

For i = 0 To 6
  Me.lstPassword.AddItem PasswordGenerator(6)
Next

Exit_GeneratePasswords:
  Exit Sub
  
Err_GeneratePasswords:
  MsgBox Err.Description
  Resume Exit_GeneratePasswords

End Sub

Public Function getEntryPoint()
  getEntryPoint = Me.OpenArgs
End Function

Private Sub lstPassword_AfterUpdate()
Call ShowPasswordInfo(Nz(lstPassword))
End Sub

Private Sub ShowPasswordInfo(Password As String)
Dim i As Integer, l As Integer, NewInfo As String

txtPasswordInfo = ""
l = Len(Password)
For i = 1 To l
    Select Case Mid$(Password, i, 1)
        Case "0" To "9"
            NewInfo = "Number " & Mid$(Password, i, 1)
        Case "A" To "Z"
            NewInfo = "Uppercase " & Mid$(Password, i, 1)
        Case "a" To "z"
            NewInfo = "Lowercase " & Mid$(Password, i, 1)
        Case Else
            NewInfo = "Invalid character: " & Mid$(Password, i, 1)
    End Select
    txtPasswordInfo = txtPasswordInfo & NewInfo & vbNewLine
Next i

End Sub
