VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Loss Mitigation Final MDHCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
'Dim x As Single
'Dim y1 As Single, y2 As Single

'Me.ScaleMode = 1    ' twips (1440 twips per inch)

'x = Me![LeftBox].Left + Me![LeftBox].Width + 1440 / 12
'y1 = Me![LeftBox].Top
'y2 = Me![LeftBox].Top + Me![LeftBox].Height + 1440 / 12

'Me.DrawWidth = 4    ' in pixels
'Me.Line (x, y1)-(x, y2)         ' vertical line
'Me.Line (Me![LeftBox].Left, y2)-(Me.Width, y2)  ' horizontal line

End Sub

Private Sub Report_Page()
Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress)
If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 1345, ""), 450, 7000, True)


If Forms![Case List]![ClientID] = 456 Then
    Me.Caption = "Loss Mitigation Final M&T"
Else
Me.Caption = "Loss Mitigation Final MDHC"
End If
End Sub

