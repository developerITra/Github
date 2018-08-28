VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Order To Docket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim PrintTo As Integer, ContactType As String


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
Dim statusMsg As String, FeeAmount As Currency

On Error GoTo Err_cmdOK_Click

If IsNull(Me!DocketByDate) Then
    MsgBox "Enter Docket By Date.", vbCritical
    Exit Sub
End If


If IsNull(Me!optOccupancy) Then
    MsgBox "Choose occupancy.", vbCritical
    Exit Sub
End If

If IsNull(Me.optLossMitigation) Then
    MsgBox "Enter type of Loss Mitigation.", vbCritical
    Exit Sub
End If

Forms!foreclosuredetails!OccupancyStatusID = optOccupancy

Call DoReport("Order to Docket Cover", PrintTo)
If optLossMitigation = 1 Then
Call DoReport("Order to Docket", PrintTo, , optOccupancy & "|" & optLossMitigation)
Else
Call DoReport("Order to Docket Final", PrintTo, , optOccupancy & "|" & optLossMitigation)
End If
' Call DoReport("Certificate of Note", PrintTo)
'Call DoReport("Notice HB 365", PrintTo)
    If Forms!foreclosuredetails!State = "MD" Then
       FeeAmount = Nz(DLookup("Value", "StandardCharges", "ID=" & 1))
    Select Case Forms!foreclosuredetails!City
    Case "Annapolis"
    Call DoReport("RegAnnapolis", PrintTo)
    AddInvoiceItem Forms!foreclosuredetails!FileNumber, "FC-DKT", "Annapolis registration Mailing", FeeAmount, 0, False, False, False, True
    
    Case "Poolesville"
    Call DoReport("RegPoolesville", PrintTo)
     AddInvoiceItem Forms!foreclosuredetails!FileNumber, "FC-DKT", "Poolesville registration Mailing", FeeAmount, 0, False, False, False, True
    
    
    Case "College Park"
    Call DoReport("RegCollegPark", PrintTo)
    AddInvoiceItem Forms!foreclosuredetails!FileNumber, "FC-DKT", "College Park registration Mailing", FeeAmount, 0, False, False, False, True

    Case "Salisbury"
    Call DoReport("RegSalisbury", PrintTo)
    AddInvoiceItem Forms!foreclosuredetails!FileNumber, "FC-DKT", "Salisbury registration Mailing", FeeAmount, 0, False, False, False, True

    Case "Laurel"
    Call DoReport("RegLaurel", PrintTo)
    AddInvoiceItem Forms!foreclosuredetails!FileNumber, "FC-DKT", "Laurel registration Mailing", FeeAmount, 0, False, False, False, True

    End Select
    If Forms![Case List]!JurisdictionID = 18 Then Call DoReport("RegPrinceGeorge", PrintTo)
    AddInvoiceItem Forms!foreclosuredetails!FileNumber, "FC-DKT", "Prince George's registration Mailing", FeeAmount, 0, False, False, False, True

    End If
    
   If Forms![Case List]![ClientID] = 97 Then Call DoReport("Order to Docket Affidavit Chase", PrintTo)
   
    

cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub





Private Sub Form_Current()
PrintTo = Int(Split(Me.OpenArgs, "|")(0))

optOccupancy = Forms!foreclosuredetails!OccupancyStatusID
'If Forms!foreclosuredetails!WizardSource = "Docketing" Then
'ChPrMediation.Visible = True
'Label95.Visible = True
'End If

End Sub


