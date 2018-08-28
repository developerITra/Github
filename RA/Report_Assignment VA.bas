VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Assignment VA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
'If txtState <> "MD" Then
If [FCdetails.State] <> "MD" Then
    txtMD1.Visible = False
    txtMD2.Visible = False
End If
End Sub

Private Sub Report_Page()
Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress, , Nz([Fair Debt]))
If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 362, ""), 300, 7000, True)
End Sub
