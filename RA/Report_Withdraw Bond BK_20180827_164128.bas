VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Withdraw Bond BK"
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

Private Sub Report_NoData(Cancel As Integer)
MsgBox "Cannot print Withdraw Bond: information is missing, most likely in Bankruptcy Details", vbCritical
End Sub

Private Sub Report_Page()
Call FirmMargin(Me, [CaseList.FileNumber])
End Sub

Private Sub ReportFooter_Print(Cancel As Integer, PrintCount As Integer)
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
