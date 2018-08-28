VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmCosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CostIncurred_DblClick(Cancel As Integer)
CostIncurred = Date
End Sub

Private Sub InvoiceToClient_DblClick(Cancel As Integer)
InvoiceToClient = Date
End Sub

Private Sub PaymentReceived_DblClick(Cancel As Integer)
PaymentReceived = Date
End Sub

Private Sub VendorPaid_DblClick(Cancel As Integer)
VendorPaid = Date
End Sub
