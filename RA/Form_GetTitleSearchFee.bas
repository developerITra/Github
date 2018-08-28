VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetTitleSearchFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
DoCmd.Close
End Sub

Private Sub cmdOK_Click()
Dim EstimateFlag As Boolean, Vendor As Integer
On Error GoTo Err_cmdOK_Click

'on 2_26-15

Vendor = cbxVendor

'Title search
AddInvoiceItem Forms![Case List]!FileNumber, txtProcess, txtDesc, txtTotal, Vendor, False, True, False, True

'Judgment cost
If txtJudAmount > 0 Then
    AddInvoiceItem Forms![Case List]!FileNumber, txtProcess, "Judgment Search", txtJudAmount, 229, False, True, False, True
End If

DoCmd.Close acForm, "EnterTitleOutReason"
DoCmd.Close


Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Close()

'Me.txtTotal.Visible = True

'DoCmd.Close

End Sub

Private Sub Form_Open(Cancel As Integer)
'lblPrompt.Caption = Split(Me.OpenArgs, "|")(0)
txtProcess = Split(Me.OpenArgs, "|")(1)
txtDesc = Split(Me.OpenArgs, "|")(2)
Select Case Split(Me.OpenArgs, "|")(3)
Case ""
cbxVendor.Enabled = False
Case "CoCounsel"
cbxVendor = DLookup("auctioneercocounsel", "jurisdictionlist", "jurisdictionid=" & Forms![Case List]!JurisdictionID)
cbxVendor.Enabled = False
Case "Advertising"
Me.txtTotal.Visible = False

cbxVendor.RowSource = "SELECT Vendors.ID, Vendors.VendorName FROM Vendors INNER JOIN JurisdictionNewspapers ON Vendors.ID = JurisdictionNewspapers.VendorID WHERE (((JurisdictionNewspapers.JurisdictionID)=" & Forms![Case List]!JurisdictionID & "));"
Case "EV-Service-Affidavit"
cbxVendor.RowSource = "SELECT ProcessServers.ID, ProcessServers.Name FROM ProcessServers;"
Case Else
cbxVendor.RowSource = "SELECT Vendors.ID, Vendors.VendorName FROM Vendors WHERE (((Vendors.Category)=""" & Split(Me.OpenArgs, "|")(3) & """))"


End Select
End Sub
