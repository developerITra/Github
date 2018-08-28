VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Entered Appearance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim sql As String
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

Select Case Forms!ReportsMenu!optClient
    Case 1
        sql = ""
    Case 2
        sql = "ClientID=" & Forms!ReportsMenu!cbxClientID
End Select

If Not Forms!ReportsMenu!chMD Then sql = sql & " AND State <> 'MD'"
If Not Forms!ReportsMenu!chDC Then sql = sql & " AND State <> 'DC'"
If Not Forms!ReportsMenu!chVA Then sql = sql & " AND State <> 'VA'"

If sql = "" Then
    Me.FilterOn = False
Else
    If Left$(sql, 4) = " AND" Then sql = Mid$(sql, 5)
    Me.Filter = sql
    Me.FilterOn = True
End If

End Sub

Private Function GetClientName() As String

If sql = "" Then
    GetClientName = ""
Else
    If Forms!ReportsMenu!chMD Then GetClientName = GetClientName & "MD  "
    If Forms!ReportsMenu!chDC Then GetClientName = GetClientName & "DC  "
    If Forms!ReportsMenu!chVA Then GetClientName = GetClientName & "VA  "
    GetClientName = GetClientName & Nz(Forms!ReportsMenu!cbxClientID.Column(1))
End If

End Function

