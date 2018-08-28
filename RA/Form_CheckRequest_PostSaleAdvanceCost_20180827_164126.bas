VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CheckRequest_PostSaleAdvanceCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdCancel_Click()
DoCmd.Close
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err_cmdDelete_Click
DoCmd.RunCommand acCmdDeleteRecord

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
Forms!AdvPostSaleCostPkg!sfrmPostSaleAdvanceCost.Requery

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
End Sub

Private Sub cmdOK_Click()
'If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
'Forms!AdvPostSaleCostPkg!sfrmPostSaleAdvanceCost.Requery
'Forms!AdvPostSaleCostPkg!sfrmPostSaleAdvanceCost.Requery


Dim rs As Recordset
Dim Des As String
Dim Vendor As String

Set rs = CurrentDb.OpenRecordset("PostSaleAdvancePkg", dbOpenDynaset, dbSeeChanges)
If rs.EOF Then
    MsgBox ("There is no data")
Exit Sub
End If

Do While Not rs.EOF
    
If rs!Description = "" Or IsNull(rs!Description) Or rs!Amount = 0 Or IsNull(rs!Amount) Or rs!Amount = "" Or IsNull(rs!Vendor) Then
    MsgBox ("Please enter all info before request a check")
Exit Sub
End If

Des = rs!Description

If Des = "Attorney Fee" Then

Else
   Call AddCheckRequest(Forms![Case List]!FileNumber, rs!Amount, Des, rs!Vendor, 1, "Other", False, 0, False, Forms![Case List]!State, "Post Sale")
End If
rs.MoveNext
Loop

rs.Close
Set rs = Nothing
  
Forms!AdvPostSaleCostPkg!sfrmPostSaleAdvanceCost!txtCKRequest.ForeColor = vbRed

DoCmd.Close
End Sub
