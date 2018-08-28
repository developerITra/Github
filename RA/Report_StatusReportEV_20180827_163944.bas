VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_StatusReportEV"
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

lblTitle.Caption = "Eviction Cases"
Select Case Forms!ReportsMenu!optCriteria
    Case 1
        sql = ""
    Case 2
        sql = "ClientID=" & Forms!ReportsMenu!cbxClient
    Case 3
        'SQL = "FileNumber=" & Forms!ReportsMenu!txtFileNumber
        sql = "" ' FileNumber IN (SELECT FileNumber FROM ReportFileNumbers)"
    
End Select

If Not Forms!ReportsMenu!chMD Then sql = sql & "State <> 'MD'"
If Not Forms!ReportsMenu!chDC Then sql = sql & "State <> 'DC'"
If Not Forms!ReportsMenu!chVA Then sql = sql & "State <> 'VA'"

'If sql = "" Then
'    Me.FilterOn = False
'Else
  '  If Left$(sql, 4) = " AND" Then sql = Mid$(sql, 5)
    Me.Filter = sql
    Me.FilterOn = True
'End If
End Sub

Private Sub Report_Unload(Cancel As Integer)
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM ChoronologyFileNumber"
DoCmd.SetWarnings True

Forms!ReportsMenu!sfrmChornologyFile.Requery

End Sub
