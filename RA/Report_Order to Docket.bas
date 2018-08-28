VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Order to Docket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim args() As String



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
  args = Split(Me.OpenArgs, "|")

End Sub

Private Sub Report_Page()
Call FirmMargin(Me, FileNumber)
End Sub

Private Function GetNonOwnerOccupied()
Select Case args(0)
  Case 1, 3
    GetNonOwnerOccupied = ""
  Case 2, 4
    GetNonOwnerOccupied = "non owner-occupied "
End Select
  
End Function

Private Function GetOccupancy()

Select Case args(0)
  Case 1
    GetOccupancy = "This is an owner-occupied residential property."
  Case 2, 4
    GetOccupancy = "This is a non owner-occupied residential property."
  Case 3
    GetOccupancy = "It is unknown if the property is owner-occupied."
End Select
  
End Function

Private Function GetLossMitigation()
  GetLossMitigation = IIf(args(1) = 1, "is not ", "is ")
End Function


