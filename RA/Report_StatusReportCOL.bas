VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_StatusReportCOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_NoData(Cancel As Integer)
txtNoData.Visible = True
End Sub

Private Sub Report_Open(Cancel As Integer)
Dim sql As String
Dim pj As Boolean

pj = Forms!ReportsMenu!chPostJudgment

If pj Then
    lblTitle.Caption = "Collection Status -- Post-Judgment"
    Select Case Forms!ReportsMenu!optCriteria
        Case 1
            sql = "FileNumber IN (SELECT FileNumber FROM rqryColJudgmentDate)"
        Case 2
            sql = "FileNumber IN (SELECT FileNumber FROM rqryColJudgmentDate) AND ClientID=" & Forms!ReportsMenu!cbxClient
        Case 3
            sql = "FileNumber=" & Forms!ReportsMenu!txtFileNumber
    End Select
Else
    lblTitle.Caption = "Collection Status -- Pre-Judgment"
    Select Case Forms!ReportsMenu!optCriteria
        Case 1
            sql = "FileNumber NOT IN (SELECT FileNumber FROM rqryColJudgmentDate)"
        Case 2
            sql = "FileNumber NOT IN (SELECT FileNumber FROM rqryColJudgmentDate) AND ClientID=" & Forms!ReportsMenu!cbxClient
        Case 3
            sql = "FileNumber=" & Forms!ReportsMenu!txtFileNumber
    End Select
End If

If Not Forms!ReportsMenu!chMD Then sql = sql & " AND State <> 'MD'"
If Not Forms!ReportsMenu!chDC Then sql = sql & " AND State <> 'DC'"
If Not Forms!ReportsMenu!chVA Then sql = sql & " AND State <> 'VA'"

Me.Filter = sql
Me.FilterOn = True

End Sub

Private Sub Report_Unload(Cancel As Integer)
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM ChoronologyFileNumber"
DoCmd.SetWarnings True

Forms!ReportsMenu!sfrmChornologyFile.Requery

End Sub
