VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_MD Lease Termination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
If Me.page = 1 Then
    Me.srptLetterheadRosenberg.Visible = True
ElseIf Me.page = 3 Then
    Me.page = 1
    Me.srptLetterheadRosenberg.Visible = True
ElseIf Me.page = 5 Then
    Me.page = 1
    Me.srptLetterheadRosenberg.Visible = True
ElseIf Me.page = 7 Then
    Me.page = 1
    Me.srptLetterheadRosenberg.Visible = True
ElseIf Me.page = 9 Then
    Me.page = 1
    Me.srptLetterheadRosenberg.Visible = True
Else
    Me.page = 2
    Me.srptLetterheadRosenberg.Visible = False
End If
End Sub

