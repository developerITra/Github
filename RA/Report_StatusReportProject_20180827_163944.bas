VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_StatusReportProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_NoData(Cancel As Integer)
Detail.Visible = False
GroupHeader0.Visible = False
GroupHeader1.Visible = False
GroupFooter2.Visible = False
MsgBox "No files found", vbInformation
End Sub

Private Sub Report_Open(Cancel As Integer)
Dim sql As String

Select Case Forms!ReportsMenu!optCriteria
    Case 1
        sql = ""
    Case 2
        sql = "ClientID=" & Forms!ReportsMenu!cbxClient
    Case 3
        sql = "FileNumber=" & Forms!ReportsMenu!txtFileNumber
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
