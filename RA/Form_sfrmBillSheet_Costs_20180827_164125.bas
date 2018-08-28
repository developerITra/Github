VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmBillSheet_Costs"
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


Private Sub Remove_AfterUpdate()
If Remove = True Then
    If (RemovedBy = "" Or IsNull(RemovedBy) = True) And IsNull(RemovedOn) = True Then
        RemovedBy = GetStaffInitials(GetStaffID())
        RemovedOn = Date
    End If
Else
    RemovedBy = ""
    RemovedOn = Null
End If
    Me.Requery
End Sub

