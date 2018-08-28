VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Certification of Compliance PNC REVISIT"
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




End Sub

Private Sub Report_Page()
Call FirmMarginVA(Me, FileNumber, , PrimaryDefName, PropertyAddress)

If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 152, ""), 450, 7000, True)

End Sub

Private Function TrusteeSign() As Boolean
  TrusteeSign = IIf(args(2) = 3, True, False)

End Function
