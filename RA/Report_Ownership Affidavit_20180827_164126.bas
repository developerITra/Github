VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Ownership Affidavit"
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

Private Sub Report_Open(Cancel As Integer)
If (Forms![Case List]!ClientID = 451 And Trim(UCase$(Forms![Case List]!Investor)) Like "LPP MORTGAGE*") Then
   Me.txtInvestor.Visible = False
   Me.txtDove.Visible = True
Else
   Me.txtDove.Visible = False
End If
End Sub

Private Sub Report_Page()
Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress, , Nz([Fair Debt]))
If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 464, ""), 450, 4000, True)
End Sub

Private Function AgentSign() As Boolean
AgentSign = args(2)
End Function

Private Sub Report_Load()
'If (Forms![case list]!ClientID = 451 And Trim(UCase$(Me.Investor)) Like "LPP MORTGAGE*") Then
'   Me.txtInvestor.Visible = False
'   Me.txtDove.Visible = True
'Else
'   Me.txtDove.Visible = False
'End If
End Sub
