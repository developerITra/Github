VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Combined Affidavit of Compliance NationStar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
Dim args() As String

If Len(ReportArgs) <> 0 Then    'Mei 9/23/15
    args = Split(ReportArgs, "|")
    txtName = args(0)
    txtTitle = args(1)
End If



End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
Dim X As Single
Dim y1 As Single, y2 As Single

Me.ScaleMode = 1    ' twips (1440 twips per inch)

X = Me![LeftBox].Left + Me![LeftBox].Width + 1440 / 12
y1 = Me![LeftBox].Top
y2 = Me![LeftBox].Top + Me![LeftBox].Height + 1440 / 12

Me.DrawWidth = 4    ' in pixels
Me.Line (X, y1)-(X, y2)         ' vertical line
Me.Line (Me![LeftBox].Left, y2)-(Me.Width, y2)  ' horizontal line

End Sub

Private Sub Report_Load()

'Mei 9/25/15




'If (Forms![case list]!ClientID = 451 And Trim(UCase$(Forms![case list].Investor)) Like "LPP MORTGAGE*") Then
'   Me.txtInvestor.Visible = False
'   Me.txtDove.Visible = True
'Else
'   Me.txtDove.Visible = False
'End If

End Sub

Private Sub Report_Open(Cancel As Integer)

If Len(strPriorServicer) = 0 Then
    Text101.Visible = False
End If

If bReferee Then
    Text107.Visible = True
    Text126.Visible = False
    Text80.Visible = False
End If
If bLost Then
    Text126.Visible = True
    Text80.Visible = False
    Text107.Visible = False
End If
If bHolder Then
    Text80.Visible = True
    Text107.Visible = False
    Text126.Visible = False
End If
If Forms![Print Statement of Debt]!Check86 = True Then Me.srptSODNationStarADJRate.Visible = True


End Sub



Private Sub Report_Page()
Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress, , Nz([Fair Debt]))
If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 464, ""), 450, 4000, True)
End Sub
'2012.02.09 Davw: Changed:
'IIf(([RemainingPBal]>[OriginalPBal]),"Additional Interest","Paid on principal")
'=[OriginalPBal]-[RemainingPBal]
