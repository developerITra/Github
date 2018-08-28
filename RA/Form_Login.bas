VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdExit_Click()
On Error GoTo Err_cmdExit_Click
Call Unlockedfiles(GetStaffID())
Call StaffSignOut

  DoCmd.Quit

Exit_cmdExit_Click:
  Exit Sub
  
Err_cmdExit_Click:
  MsgBox Err.Description
  Resume Exit_cmdExit_Click
End Sub




Private Sub cmdLogin_Click()
On Error GoTo Err_cmdLogin_Click



If Nz(txtPassword) = "" Then
    MsgBox "Password can not be blank", vbCritical
    Exit Sub
End If

Dim rstUser As Recordset

Set rstUser = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID=" & StaffID, dbOpenDynaset, dbSeeChanges)
If Not rstUser.EOF Then
  If (Nz(rstUser!Password) <> DigestStrToHexStr(txtPassword & "Rosenberg " & rstUser!Username)) Then
    rstUser.Close
    MsgBox "Password is incorrect.  Please re-enter", vbCritical
    Exit Sub
  End If
End If

Dim daysLeft As Integer


StaffID = rstUser![ID]
daysLeft = DateDiff("d", Date, rstUser!PasswordExpires)

rstUser.Close

If (daysLeft <= 0) Then
  MsgBox "Your password has expired.", vbOKOnly, "Password Expired"
  DoCmd.OpenForm "ChangePassword", , , "[ID] = " & StaffID, , , Me.Name
  DoCmd.Close acForm, Me.Name
  Exit Sub
ElseIf (daysLeft < 7) Then
  If (MsgBox("Your password will expire in " & IIf(daysLeft = 1, "1 day.", daysLeft & " days.  Change password now?"), vbYesNo, "Password Expires") = vbYes) Then
    DoCmd.OpenForm "ChangePassword", , , "[ID] = " & StaffID, , , Me.Name
    DoCmd.Close acForm, Me.Name
    Exit Sub
  End If
End If


  DoCmd.OpenForm "Main"
   DoCmd.Close acForm, Me.Name
  
Exit_cmdLogin_Click:
  Exit Sub
  
Err_cmdLogin_Click:
  MsgBox Err.Description
  Resume Exit_cmdLogin_Click

End Sub

Private Sub Form_Open(Cancel As Integer)

Dim d As Recordset
'If CurrentProject.Name = "RA.accdb" Or CurrentProject.Name = "RA.accde" Then
txtVersion = DBVersion
'Else
'txtVersion = DBVersionTest
'Me.Detail.BackColor = -2147483615
'End If
Call CheckVersion(True)
Call CheckSQL

txtLoginName = GetLoginName()
If (txtLoginName = "") Then ' User is not in the database, exiting
  DoCmd.Quit
End If

txtDate = Now()

If StaffCheck() = True Then
MsgBox ("Please Exit the Other Rosie Copy")
DoCmd.Quit
Else
Call StaffSignIn
End If

Dim rstUser As Recordset

'If IsNull(DLookup("StaffID", "Locks", "StaffID=" & StaffID)) Then
Set rstUser = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID=" & StaffID, dbOpenDynaset, dbSeeChanges)
If Not rstUser.EOF Then



  If (IsNull(rstUser!Password)) Then
    rstUser.Close
    
    DoCmd.OpenForm "ChangePassword", , , , , , Me.Name
    DoCmd.Close acForm, Me.Name
  
  End If

Else
    rstUser.Close
    MsgBox "User context is lost, please exit the database", vbCritical
    DoCmd.Close acForm, Me.Name
    DoCmd.Quit
    
End If

'Else
'MsgBox ("You are already in other file")
'DoCmd.Quit
'End If


End Sub




Private Sub Form_Unload(Cancel As Integer)
If IsLoadedF("Main") = False Then
Call Unlockedfiles(GetStaffID())
Call StaffSignOut
End If
End Sub

Private Sub txtVersion_Click()
cmdExit.SetFocus
DoCmd.RunCommand acCmdAboutMicrosoftAccess
End Sub


