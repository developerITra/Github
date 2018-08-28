VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Workflow Notice Problems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim NoData As Boolean

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
Dim mindays As Integer, maxdays As Integer, actualdays As Long

Select Case State
    Case "DC"
        mindays = 30
        maxdays = 35
    Case "MD"
        mindays = 10
        maxdays = 30
    Case "VA"
        mindays = 14
        maxdays = 30
End Select

If IsNull(Notices) Then
    actualdays = DateDiff("d", Date, Sale)
    If actualdays >= mindays Then Cancel = True
Else
    actualdays = DateDiff("d", Notices, Sale)
    If actualdays >= mindays And actualdays <= maxdays Then Cancel = True
End If

Days = actualdays

End Sub

Private Sub Report_NoData(Cancel As Integer)
NoData = True
End Sub

Private Function GetRecordCount() As String
If NoData Then
    GetRecordCount = "No files"
Else
    If txtRC = 1 Then
        GetRecordCount = "1 file"
    Else
        GetRecordCount = txtRC & " files"
    End If
End If
End Function


