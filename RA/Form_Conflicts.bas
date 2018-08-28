VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Conflicts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboStatus_AfterUpdate()

Dim X As Boolean

X = (cboStatus = 2 Or cboStatus = 3)  ' completed
If (X) Then

  If StaffID = 0 Then Call GetLoginName
  CompletedID = StaffID
  CompletedByName.Requery
  
  CompletedDate = Date
End If
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



Private Sub Form_Open(Cancel As Integer)
Call UpdateList
End Sub



Private Sub optActive_AfterUpdate()
Call UpdateList
End Sub
  

Private Sub UpdateList()

Dim strFilter As String


If (optActive = 1) Then
  strFilter = "[Conflicts].[ConflictStatusID] = 1"  ' Pending
  Filter = strFilter
  FilterOn = True
Else
  FilterOn = False
End If
Me.Requery




End Sub




Private Sub optType_AfterUpdate()
  Call UpdateList
End Sub

