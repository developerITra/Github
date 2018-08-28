VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmContactsTRAttorneyRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo Err_cmdClose_Click

  DoCmd.Close
  Forms!TitleResolutionDetails!sfrmNamesTR.Form.AttorneyForSummary.Requery
  
Exit_cmdClose_Click:
  Exit Sub
  
Err_cmdClose_Click:
  MsgBox Err.Description
  Resume Exit_cmdClose_Click
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

Private Function GetAttorneyName()

  GetAttorneyName = Me.OpenArgs
End Function

Private Sub Form_Open(Cancel As Integer)

  Me.lblAttorney.Caption = "Attorney:  " & GetAttorneyName()
End Sub
