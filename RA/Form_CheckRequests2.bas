VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CheckRequests2"
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

Private Function FetchBulkCheck()

  FetchBulkCheck = IIf([BulkCheck], "Bulk Check", "")
  
End Function


Private Sub cmdReconcile_Click()
On Error GoTo Err_cmdReconcile_Click
Dim R As Recordset

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

Set R = Forms![CheckRequests]!sfrmCheckRequest.Form.RecordsetClone
R.MoveFirst
Do While Not R.EOF
    If R!ChargeCleared Then
        R.Edit
        R!Reconciled = True
        R!StatusID = DLookup("[ID]", "[CheckRequestStatus]", "[Status] = 'Completed'")
        R!CompletedDate = Date
        R!CompletedBy = GetStaffID
        R.Update
    End If
    R.MoveNext
Loop
Set R = Nothing

Forms![CheckRequests]!sfrmCheckRequest.Requery


Exit_cmdReconcile_Click:
  Exit Sub
  
Err_cmdReconcile_Click:
  MsgBox Err.Description
  Resume Exit_cmdReconcile_Click

End Sub

Private Sub Form_Open(Cancel As Integer)
Call UpdateList
End Sub

Private Sub optActive_AfterUpdate()
Call UpdateList
End Sub
  
Private Sub UpdateList()

Dim strFilter As String

Select Case optType
    Case 1      ' check requests
        Me.sfrmCheckRequest.SourceObject = "sfrmCheckRequestCheck"
        Me.cmdReconcile.Visible = False
        
        strFilter = "[RequestTypeID] = 1"
  
    Case 2      ' credit card
        Me.sfrmCheckRequest.SourceObject = "sfrmCheckRequestCC"
        Me.cmdReconcile.Visible = True
        
        strFilter = "[RequestTypeID] = 2"
  
    Case 3      ' pre-cut check
        Me.sfrmCheckRequest.SourceObject = "sfrmCheckRequestPrecut"
        Me.cmdReconcile.Visible = False
        
        strFilter = "[RequestTypeID] = 3"
End Select

If (optActive = 2) Then strFilter = strFilter & " and [StatusID] = 1"  ' Pending
Forms![CheckRequests]!sfrmCheckRequest.Form.Filter = strFilter
Forms![CheckRequests]!sfrmCheckRequest.Form.FilterOn = True

End Sub

Private Sub optType_AfterUpdate()
  Call UpdateList
End Sub
