VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Brokers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cbxSelect_AfterUpdate()
 ' Find the record that matches the control.
 
    Me.RecordsetClone.FindFirst "[BrokerID] = " & Me![cbxSelect]
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

Private Sub ComDelete_Click()
If MsgBox("You are going to delete this contact, are you sure ? ", vbYesNo) = vbYes Then

DoCmd.SetWarnings False
DoCmd.RunSQL ("DELETE * FROM Brokers WHERE " & _
" Brokers.BrokerID = " & Me.BrokerID & ";")

Me.Form.Requery
DoCmd.SetWarnings True
Else
Exit Sub
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
If PrivJurisdic Then
Me.ComDelete.Enabled = True
End If


End Sub

Private Sub lstSelect_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[BrokerID] = " & str(Nz(Me![lstSelect], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
End Sub

Private Sub Form_AfterUpdate()
lstSelect.Requery
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
