VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Affidavit of Default"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim NextNumber As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
chCoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And [BKdetails.Chapter] = 13)
End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

Dim X As Single, x0 As Single
Dim y1 As Single, y2 As Single, y3 As Single

Me.ScaleMode = 1    ' twips (1440 twips per inch)

X = Me![LeftBox].Left + Me![LeftBox].Width + 1440 / 12
x0 = Me![LeftBox].Left
y1 = Me![LeftBox].Top
y2 = Me![LeftBox].Top + Me![LeftBox].Height + 1440 / 12
y3 = Me![LeftBox2].Top + Me![LeftBox2].Height + 1440 / 12

Me.DrawWidth = 4    ' in pixels
Me.Line (X, y1)-(X, y3)         ' vertical line
Me.Line (x0, y2)-(Me.Width, y2)  ' horizontal line
Me.Line (x0, y3)-(Me.Width, y3)  ' horizontal line

End Sub


Private Sub Report_Open(Cancel As Integer)
NextNumber = 3
End Sub

Private Sub Report_Page()
NextNumber = 3
Call FirmMargin(Me, FileNumber, AttorneyInfo)

End Sub

Private Function GetNextNumber() As String
GetNextNumber = Space$(10) & NextNumber & ".  "
'Debug.Print NextNumber
NextNumber = NextNumber + 1
End Function
