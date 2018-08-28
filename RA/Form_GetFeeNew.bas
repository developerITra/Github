VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetFeeNew"
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

If txtTotal.Visible = True Then
    If Nz(txtTotal) <= 0 Then
        If MsgBox("Please confirm you wish to enter a Zero amount", vbYesNo) = vbNo Then Exit Sub
    End If
End If

EstimateFlag = True 'False makes it an Estimate.  Quite the misnomer

Select Case Split(Me.OpenArgs, "|")(3)
    Case "Advertising"
        EstimateFlag = False
            
             If CurrentProject.AllForms("ForeclosureDetails").IsLoaded Then
                Forms!foreclosuredetails!NewspaperVendor = Int(cbxVendor)
                Forms![Case List]!txt_Vendor = Int(cbxVendor)
                Forms!foreclosuredetails.Requery

            End If
 
            
'If txtDesc = "Estimated Advertising Costs" Then EstimateFlag = False
If Not IsNull(cbxVendor) Then Vendor = cbxVendor
AddInvoiceItem Forms![Case List]!FileNumber, txtProcess, "Estimated Advertising Costs", txtTotal, Vendor, False, EstimateFlag, False, True

            
Forms!foreclosuredetails.SetFocus

'DoCmd.Close acForm, "ForeclosureDetails"
'txtTotal.Visible = True
'DoCmd.Close
 Exit Sub
       
    
    Case "CoCounsel"
        EstimateFlag = False
'    Case "Process Server"
'        EstimateFlag = False
    Case "EV-Service-Affidavit"
    Vendor = 0
    If Not IsNull(cbxVendor) Then Vendor = cbxVendor
    AddInvoiceItem Forms![Case List]!FileNumber, txtProcess, txtDesc, txtTotal, Vendor, False, EstimateFlag, False, True
    DoCmd.Close
   
    Exit Sub
         

    End Select

If txtDesc = "Estimated Process Server Costs" Then EstimateFlag = False
  

Vendor = 0
If Not IsNull(cbxVendor) Then Vendor = cbxVendor

'added on 2_26-15

If txtDesc = "Judgment Search" And (txtTotal = 0 Or IsNull(txtTotal)) Then
    Exit Sub
End If

AddInvoiceItem Forms![Case List]!FileNumber, txtProcess, txtDesc, txtTotal, Vendor, False, EstimateFlag, False, True


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
lblPrompt.Caption = Split(Me.OpenArgs, "|")(0)
txtProcess = Split(Me.OpenArgs, "|")(1)
txtDesc = Split(Me.OpenArgs, "|")(2)
Select Case Split(Me.OpenArgs, "|")(3)
Case ""
cbxVendor.Enabled = False
Case "CoCounsel"
cbxVendor = DLookup("auctioneercocounsel", "jurisdictionlist", "jurisdictionid=" & Forms![Case List]!JurisdictionID)
cbxVendor.Enabled = False
Case "Advertising"

    If Forms![Case List]!State = "DC" Then
        Me.txtTotal.Visible = True
    Else
        Me.txtTotal.Visible = False
    End If
    
cbxVendor.RowSource = "SELECT Vendors.ID, Vendors.VendorName FROM Vendors INNER JOIN JurisdictionNewspapers ON Vendors.ID = JurisdictionNewspapers.VendorID WHERE (((JurisdictionNewspapers.JurisdictionID)=" & Forms![Case List]!JurisdictionID & "));"
Case "EV-Service-Affidavit"
cbxVendor.RowSource = "SELECT ProcessServers.ID, ProcessServers.Name FROM ProcessServers;"
Case Else
cbxVendor.RowSource = "SELECT Vendors.ID, Vendors.VendorName FROM Vendors WHERE (((Vendors.Category)=""" & Split(Me.OpenArgs, "|")(3) & """))"


End Select
End Sub
