VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Exempt Deed VA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub PageFooter_Format(Cancel As Integer, FormatCount As Integer)
If page = 1 Then
    txtConsideration.Visible = True
    txtAssessedValue.Visible = True
Else
    txtConsideration.Visible = False
    txtAssessedValue.Visible = False
End If
End Sub

Private Sub Report_Page()
Dim y1 As Single, y2 As Single, BIGFONT As Integer, SMALLFONT As Integer

BIGFONT = 8
SMALLFONT = 5

Const FONTSPACE = 30
'
' Simulate "redlines"
'
Me.ScaleMode = 5    ' measure in inches
Me.DrawWidth = 2    ' line will be 2 pixels wide
Me.Line (1.15, 0)-(1.15, 22), 0
Me.Line (1.18, 0)-(1.18, 22), 0
Me.Line (7.9, 0)-(7.9, 22), 0
'
' Add Firm's name and address to left margin
'
y1 = 1 * 1440
y2 = y1 + 60

    Me.ScaleMode = 1  ' twips
    Me.FontName = "Georgia"

        With Me
            .CurrentX = 0
            .FontSize = BIGFONT
            .CurrentY = y2
            .Print "Deed Prepared By:"
        End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 80
       
        
        With Me
            .CurrentX = 0
            .FontSize = BIGFONT
            .CurrentY = y2
            .FontBold = True
            .Print Me.NameVA
         End With
            
y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 60

        With Me
            .CurrentX = 0
            .FontSize = BIGFONT
            .CurrentY = y2
            .Print "VA BAR# " & Me!VABar
     End With

y1 = 7.5 * 1440
y2 = y1 + 20
BIGFONT = 6
SMALLFONT = 5

With Me
    .CurrentX = 260
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "C"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "OMMONWEALTH"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With Me
    .CurrentX = 320
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "T"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "RUSTEES"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print ", LLC"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With Me
    .CurrentX = 280
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "8601 W"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ESTWOOD"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With Me
    .CurrentX = 320
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " C"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ENTER"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " D"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "RIVE,"
    
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With Me
    .CurrentX = 540
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "S"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "UITE"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " 255"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With Me
    .CurrentX = 40
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "V"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "IENNA"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print ", V"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "IRGINIA"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " 22182"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With Me
    .CurrentX = 300
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "(703) 752-8500"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE * 5
y2 = y1 + 20
With Me
    .CurrentX = 190
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "F"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ILE"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " N"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "UMBER: "
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " " & FileNumber
End With



End Sub


