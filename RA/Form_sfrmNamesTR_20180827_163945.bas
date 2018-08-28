VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmNamesTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdDelete_Click()

On Error GoTo Err_cmdDelete_Click

DoCmd.RunCommand acCmdDeleteRecord

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click

End Sub

Private Sub cmdManageClients_Click()
On Error GoTo Err_cmdManageClients_Click

  Dim strAttorneyName As String
  
  If (Me.Dirty) Then DoCmd.RunCommand acCmdSaveRecord
  
  
  If (IsNull([ID])) Then
    MsgBox "Please enter name details before continuing.", vbCritical, "Name Details"
    Exit Sub
  End If
  
  If (Me.Borrower = True Or Me.Owner = True) Then
    MsgBox "Person cannot be a borrower or owner.", vbCritical, "Name Details"
    Exit Sub
  End If
  
  strAttorneyName = [First] & " " & [Last]
  DoCmd.OpenForm "sfrmContactsTRAttorneyRep", , , "[AttorneyNameID] = " & [ID], , , strAttorneyName
  

Exit_cmdManageClients_Click:
  Exit Sub
  
Err_cmdManageClients_Click:
  MsgBox Err.Description
  Resume Exit_cmdManageClients_Click
End Sub

Private Sub Other_AfterUpdate()
If Other Then
    Trustee = False
    SettlementCompany = False
    Borrower = False
    Owner = False
End If
End Sub

Private Sub Trustee_AfterUpdate()
If Trustee Then
    SettlementCompany = False
    Borrower = False
    Owner = False
    Other = False
    
    
End If
End Sub



