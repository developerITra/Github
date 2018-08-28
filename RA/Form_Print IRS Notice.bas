VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print IRS Notice"
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

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click
Call DoReport("IRS Notice", Me.OpenArgs)
If MsgBox("Update IRS Notice = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
    Forms!foreclosuredetails!IRSNotice = Now()
    AddInvoiceItem FileNumber, "FC-IRS", "IRS Notice Postage - Regular", Nz(DLookup("Value", "StandardCharges", "ID=" & 7)), 76, False, False, False, True
    AddInvoiceItem FileNumber, "FC-IRS", "IRS Notice Postage - Certified", Nz(DLookup("Value", "StandardCharges", "ID=" & 9)), 76, False, False, False, True
    AddStatus FileNumber, Now(), "IRS Notice sent"
End If
cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

