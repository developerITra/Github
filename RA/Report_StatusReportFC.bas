VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_StatusReportFC"
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
Dim ps As Boolean
Dim sql As String

'ps = Forms!ReportsMenu!chPostSale

'If ps Then
'    'lblTitle.Caption = "Post-Sale Foreclosure Cases / REO" '3/12/2014 Per diane
'    Select Case Forms!ReportsMenu!optCriteria
'        Case 1
'            sql = "PostSale"
'        Case 2
'            sql = "PostSale AND ClientID=" & Forms!ReportsMenu!cbxClient
'        Case 3
'            'sql = "PostSale AND FileNumber=" & Forms!ReportsMenu!txtFileNumber
'           ' sql = "PostSale AND FileNumber IN (SELECT FileNumber FROM ReportFileNumbers)"
'             sql = "PostSale" ' AND FileNumber IN (SELECT FileNumber FROM ReportFileNumbers)"
'    End Select
'Else
'    'lblTitle.Caption = "Foreclosure Cases" '3/12/2014 Per diane
'    Select Case Forms!ReportsMenu!optCriteria
'        Case 1
'            sql = "NOT PostSale"
'        Case 2
'            sql = "NOT PostSale AND ClientID=" & Forms!ReportsMenu!cbxClient
'        Case 3
''            SQL = "NOT PostSale AND FileNumber=" & Forms!ReportsMenu!txtFileNumber
'           ' sql = "NOT PostSale AND FileNumber IN (SELECT FileNumber FROM ReportFileNumbers)"
'            sql = "NOT PostSale" ' AND FileNumber IN (SELECT FileNumber FROM ReportFileNumbers)"
'    End Select
'End If

'2012.01.30 DaveW Tis reportdoes not use FCDetails:
'If Not Forms!ReportsMenu!chMD Then sql = sql & " AND FCDetails.State <> 'MD'"
'If Not Forms!ReportsMenu!chDC Then sql = sql & " AND FCDetails.State <> 'DC'"
'If Not Forms!ReportsMenu!chVA Then sql = sql & " AND FCDetails.State <> 'VA'"

If Not Forms!ReportsMenu!chMD Then sql = "State <> 'MD'"
If Not Forms!ReportsMenu!chDC Then sql = "State <> 'DC'"
If Not Forms!ReportsMenu!chVA Then sql = "State <> 'VA'"

Me.Filter = sql
Me.FilterOn = True

'Contact.Visible = Not ps
'PostContact.Visible = ps
'PhoneNum.Visible = Not ps
'PostPhoneNum.Visible = ps
'FaxNum.Visible = Not ps
'PostFaxNum.Visible = ps

End Sub

Private Sub Report_Unload(Cancel As Integer)
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM ChoronologyFileNumber"
DoCmd.SetWarnings True

Forms!ReportsMenu!sfrmChornologyFile.Requery


End Sub
