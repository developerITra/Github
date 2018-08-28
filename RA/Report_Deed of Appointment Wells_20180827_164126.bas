VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Deed of Appointment Wells"
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

 '   For i = 1 To cnt
  '          GetNames = GetNames & Names(i)
   '         If i < cnt Then
    '            If i < cnt - 1 Then
     '               GetNames = GetNames & ", "
      '          Else
       '             GetNames = GetNames & " and "
        '        End If
        '    End If
       ' Next i

Private Sub Report_Page()
If PropertyState = "VA" Then
   Call FirmMarginVA(Me, FileNumber, , PrimaryDefName, PropertyAddress, Nz([Fair Debt]))
Else
   Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress, , Nz([Fair Debt]))
End If

If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 226, ""), 450, 7000, True)


'Dim i As Integer
'Dim strTrustees As String
'strTrustees = ""
   ' For i = 0 To Forms!foreclosureDetails!lstTrustees.Listcount - 1
   '     If i < Forms!foreclosureDetails!lstTrustees.Listcount - 1 Then
   '         If i < Forms!foreclosureDetails!lstTrustees.Listcount - 2 Then
   '             strTrustees = strTrustees & Forms!foreclosureDetails!lstTrustees.Column(1, i) & ", "
   '         Else
   '             strTrustees = strTrustees & Forms!foreclosureDetails!lstTrustees.Column(1, i) & " and "
   '         End If
   '     Else
   '          strTrustees = strTrustees & Forms!foreclosureDetails!lstTrustees.Column(1, i)
   '     End If
   ' Next i
End Sub
