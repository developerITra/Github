VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Workflow Service Deadline DC_PastDue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim NoData As Boolean

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

Private Sub Report_Open(Cancel As Integer)
'Dim rs As Recordset
'Dim i As Integer
'i = 0
'Set rs = CurrentDb.OpenRecordset("WkflServiceDeadline", dbOpenDynaset, dbSeeChanges)
'
'If rs.EOF Then
'    MsgBox ("There is no data")
'Exit Sub
'End If
'
'Do While Not rs.EOF
'
'i = DateDiff("d", Date, rs!ServiceDeadline)
'If i <= 15 Then
'  lbArrow.Visible = True
'Else
'  lbArrow.Visible = True
'End If
'
'If i = 16 Then
'ServiceDeadline.BackColor = 15
'End If
'
'i = 0
'rs.MoveNext
'Loop
'
'rs.Close
'Set rs = Nothing


End Sub
