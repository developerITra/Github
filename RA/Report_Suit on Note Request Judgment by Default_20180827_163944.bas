VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Suit on Note Request Judgment by Default"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

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

Private Sub GroupFooter2_Print(Cancel As Integer, PrintCount As Integer)
Dim X As Single
Dim y1 As Single, y2 As Single

Me.ScaleMode = 1    ' twips (1440 twips per inch)

X = Me![LeftBox4].Left + Me![LeftBox4].Width + 1440 / 12
y1 = Me![LeftBox4].Top
y2 = Me![LeftBox4].Top + Me![LeftBox4].Height + 1440 / 12

Me.DrawWidth = 4    ' in pixels
Me.Line (X, y1)-(X, y2)         ' vertical line
Me.Line (Me![LeftBox4].Left, y2)-(Me.Width, y2)  ' horizontal line

End Sub

Private Sub GroupFooter3_Print(Cancel As Integer, PrintCount As Integer)
Dim X As Single
Dim y1 As Single, y2 As Single

Me.ScaleMode = 1    ' twips (1440 twips per inch)

X = Me![LeftBox2].Left + Me![LeftBox2].Width + 1440 / 12
y1 = Me![LeftBox2].Top
y2 = Me![LeftBox2].Top + Me![LeftBox2].Height + 1440 / 12

Me.DrawWidth = 4    ' in pixels
Me.Line (X, y1)-(X, y2)         ' vertical line
Me.Line (Me![LeftBox2].Left, y2)-(Me.Width, y2)  ' horizontal line

End Sub

Private Sub GroupFooter4_Print(Cancel As Integer, PrintCount As Integer)
Dim X As Single
Dim y1 As Single, y2 As Single

Me.ScaleMode = 1    ' twips (1440 twips per inch)

X = Me![LeftBox3].Left + Me![LeftBox3].Width + 1440 / 12
y1 = Me![LeftBox3].Top
y2 = Me![LeftBox3].Top + Me![LeftBox3].Height + 1440 / 12

Me.DrawWidth = 4    ' in pixels
Me.Line (X, y1)-(X, y2)         ' vertical line
Me.Line (Me![LeftBox3].Left, y2)-(Me.Width, y2)  ' horizontal line

End Sub

Private Sub GroupFooter0_Print(Cancel As Integer, PrintCount As Integer)
Dim X As Single
Dim y1 As Single, y2 As Single

Me.ScaleMode = 1    ' twips (1440 twips per inch)

X = Me![LeftBox1].Left + Me![LeftBox1].Width + 1440 / 12
y1 = Me![LeftBox1].Top
y2 = Me![LeftBox1].Top + Me![LeftBox1].Height + 1440 / 12

Me.DrawWidth = 4    ' in pixels
Me.Line (X, y1)-(X, y2)         ' vertical line
Me.Line (Me![LeftBox1].Left, y2)-(Me.Width, y2)  ' horizontal line

End Sub

Private Sub Report_Page()
Call FirmMargin(Me, FileNumber)
End Sub

