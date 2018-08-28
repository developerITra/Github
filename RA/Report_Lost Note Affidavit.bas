VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Lost Note Affidavit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Load()
'If (Forms![case list]!ClientID = 451 And Trim(UCase$(Me.Investor)) Like "LPP MORTGAGE*") Then
'   Me.txtInvestor.Visible = False
'   Me.txtDove.Visible = True
'Else
'   Me.txtDove.Visible = False
'End If
End Sub

Private Sub Report_Open(Cancel As Integer)
If (Forms![Case List]!ClientID = 451 And Trim(UCase$(Forms![Case List]!Investor)) Like "LPP MORTGAGE*") Then
   Me.txtInvestor.Visible = False
   Me.txtDove.Visible = True
Else
   Me.txtDove.Visible = False
End If
End Sub

Private Sub Report_Page()

If [FCdetails.State] = "VA" Then
    Call FirmMarginVA(Me, FileNumber, , PrimaryDefName, PropertyAddress)
Else
    Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress)
End If

If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 154, ""), 450, 6000, True)


'Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress)
End Sub

