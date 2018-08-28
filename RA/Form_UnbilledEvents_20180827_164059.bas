VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_UnbilledEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click


DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub


Private Sub cmdClearReason_Click()
Dim rstBillReasons As Recordset

Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where ID=" & lstBillingReasons, dbOpenDynaset, dbSeeChanges)

With rstBillReasons
.Edit
!invoiced = Date
.Update
.Close
End With
lstBillingReasons.Requery
End Sub

Private Sub cmdComplete_Click()
If lstBillingReasons.ListCount = 0 Then
Forms![Case List]!BillCase = False
Forms![Case List]!lblBilling.Visible = False
MsgBox "File has been removed from the Need to Invoice Report", vbExclamation
End If
DoCmd.Close acForm, "unbilledevents"
End Sub
