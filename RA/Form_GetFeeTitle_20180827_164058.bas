VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetFeeTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click

FeeAmount = txtTotal
Forms!foreclosuredetails!chNoBillTitle = chNonBillable
DoCmd.Close

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
FeeAmount = 0
lblPrompt.Caption = Me.OpenArgs
txtTotal.SetFocus
cbxVendor.RowSource = "SELECT Vendors.ID, Vendors.VendorName FROM Vendors WHERE (((Vendors.Category)=Abstractor"
End Sub

