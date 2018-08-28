VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Military Affidavit NoSSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim args() As String

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

args = Split(ReportArgs, "|")
txtName = args(0)
txtTitle = args(1)
txtDefendant = args(2)

End Sub

Private Sub Report_Open(Cancel As Integer)
If (Forms![Case List]!ClientID = 451 And Trim(UCase$(Forms![Case List]!Investor)) Like "LPP MORTGAGE*" And Forms!Foreclosureprint!txtDesignator <> 3) Then
   Me.txtInvestor.Visible = False
   Me.txtDove.Visible = True
Else
   Me.txtDove.Visible = False
End If
End Sub

Private Sub Report_Page()
If [FCdetails.State] = "VA" Then
Call FirmMarginVA(Me, FileNumber, , PrimaryDefName, PropertyAddress, Nz([Fair Debt]))
Else
Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress, , Nz([Fair Debt]))
End If
If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 152, ""), 450, 7000, True)

'Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress)
'If Page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 152, ""), 450, 7000, True)
End Sub

Private Function TrusteeSign() As Boolean
  TrusteeSign = IIf(args(3) = 3, True, False)

End Function
