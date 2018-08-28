VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_BOA Cover Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(Cancel As Integer)
If Forms![Case List]!ClientID = 444 Then
Me.Caption = "PHH Cover Letter"
End If

End Sub
