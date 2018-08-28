VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Certification of Compliance"
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

'If Me.Text112 = "VA" Then
'    Call FirmMarginVANoLine(Me, FileNumber, , PrimaryDefName, Property)
'Else
'    Call FirmMarginVANoLine(Me, FileNumber, , PrimaryDefName, Property)
'End If
'If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 226, ""), 450, 7000, True)
End Sub

Private Function AttorneySign() As Boolean
AttorneySign = args(2)
End Function

