VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Deed of Appointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
Dim args() As String

args = Split(ReportArgs, "|")
txtName = args(0)
txtTitle = args(1)


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
If PropertyState = "VA" Then
    Call FirmMarginVA(Me, FileNumber, , PrimaryDefName, PropertyAddress, Nz([Fair Debt]))
Else
    Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress, , Nz([Fair Debt]))
End If

If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 226, ""), 450, 7000, True)
End Sub

Private Sub Report_Load()
'If (Forms![Case List]!ClientID = 451 And Trim(UCase$(Me.Investor)) Like "LPP MORTGAGE*") Then
'   Me.txtInvestor.Visible = False
'   Me.txtDove.Visible = True
'Else
'   Me.txtDove.Visible = False
'End If
End Sub
