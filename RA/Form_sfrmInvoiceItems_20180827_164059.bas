VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmInvoiceItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Approved_AfterUpdate()
ApprovedBy = GetStaffID()
End Sub

Private Sub Approved_DblClick(Cancel As Integer)
Approved = Date
ApprovedBy = GetStaffID()
End Sub

Private Sub cbxVendor_AfterUpdate()
'Me.Requery
'Forms![Case List].Requery
'Me.cbxVendor = ""
End Sub

Private Sub chUnbillable_Click()
cbxNonBillableReason.Enabled = True
DoNotBillBy = GetStaffID()
End Sub

Private Sub Form_Current()

If PrivBillingEdits = True Then
Me.Description.Locked = False
Me.InvoiceID.Locked = False
Me.Description.Locked = False
Me.cbxVendor.Locked = False
Me.EstimatedAmount.Locked = False
Me.ActualAmount.Locked = False
Me.txtVandor.Locked = False

Me.AllowAdditions = True
Me.AllowDeletions = True
Me.AllowEdits = True
Me.ScrollBars = 2
Else
Me.Description.Locked = True
Me.InvoiceID.Locked = True
Me.Description.Locked = True
Me.cbxVendor.Locked = True
Me.EstimatedAmount.Locked = True
Me.ActualAmount.Locked = True

Me.AllowAdditions = True
Me.AllowDeletions = True
Me.AllowEdits = True
Me.ScrollBars = 2
End If

End Sub
