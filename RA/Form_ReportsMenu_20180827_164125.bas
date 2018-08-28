VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ReportsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cbxClient_AfterUpdate()
If Not IsNull(cbxClient) Then
    optCriteria = 2
    cbxProject = Null
    cbxProject.RowSource = "SELECT DISTINCT Project FROM CaseList WHERE Project Is Not Null AND ClientID=" & cbxClient
End If
End Sub

Private Sub cbxClientID_AfterUpdate()

If Not IsNull(cbxClientID) Then optClient = 2

End Sub

Private Sub cmdCivil_Click()

End Sub


Private Sub cmdActiveCases_click()

On Error Resume Next
Kill "S:\ProductionReporting\ActiveCases" & Format$(Now(), "yyyymmdd") & ".xls"
On Error GoTo 0
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryActiveCases", "S:\ProductionReporting\ActiveCases" & Format$(Now(), "yyyymmdd") & ".xls"

Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModule.xlsm"
.Run "ActiveCases"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub


Private Sub BkReportPart2(strFileName As String, ReportNo As Integer)
On Error GoTo 0

Select Case ReportNo
Dim ExcelObj As Object
Case 0
   
    'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryBankruptcy", "S:\ProductionReporting\Bankruptcy" & Format$(Now(), "yyyymmdd") & ".xls"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryBankruptcy", strFileName
    
    Set ExcelObj = CreateObject("Excel.Application")
    With ExcelObj
    .Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleBKNew.xlsm"
    .Run "BKSpread"
    .ActiveWorkbook.Close
    '.Visible
    End With
    Set ExcelObj = Nothing
    MsgBox "The report is now ready to view in Excel"
Case 1
  
    'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryBankruptcy", "S:\ProductionReporting\Bankruptcy" & Format$(Now(), "yyyymmdd") & ".xls"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryBankruptcy", strFileName
    
    Set ExcelObj = CreateObject("Excel.Application")
    With ExcelObj
    .Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleBKNew1.xlsm"
    .Run "BKSpread"
    .ActiveWorkbook.Close
    '.Visible
    End With
    Set ExcelObj = Nothing
    MsgBox "The report is now ready to view in Excel"
Case 2
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryBankruptcy", strFileName
    
    Set ExcelObj = CreateObject("Excel.Application")
    With ExcelObj
    .Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleBKNew2.xlsm"
    .Run "BKSpread"
    .ActiveWorkbook.Close
    '.Visible
    End With
    Set ExcelObj = Nothing
    MsgBox "The report is now ready to view in Excel"
Case 3
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryBankruptcy", strFileName
    
    Set ExcelObj = CreateObject("Excel.Application")
    With ExcelObj
    .Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleBKNew3.xlsm"
    .Run "BKSpread"
    .ActiveWorkbook.Close
    '.Visible
    End With
    Set ExcelObj = Nothing
    MsgBox "The report is now ready to view in Excel"
Case 4
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryBankruptcy", strFileName
    
    Set ExcelObj = CreateObject("Excel.Application")
    With ExcelObj
    .Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleBKNew4.xlsm"
    .Run "BKSpread"
    .ActiveWorkbook.Close
    '.Visible
    End With
    Set ExcelObj = Nothing
    MsgBox "The report is now ready to view in Excel"
End Select
End Sub

Private Sub cmdBK_Click()
If FileLocked("S:\ProductionReporting\BankRuptcy" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
    
    If FileLocked("S:\ProductionReporting\BankRuptcy1" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
    
        If FileLocked("S:\ProductionReporting\BankRuptcy2" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
            
            If FileLocked("S:\ProductionReporting\BankRuptcy3" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
                    
                    If FileLocked("S:\ProductionReporting\BankRuptcy4" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
                    MsgBox ("Please Check with IT, there may be more than 5 sessions of this file open or a SQL problem")
                    Exit Sub
                    Else
                    Kill "S:\ProductionReporting\BankRuptcy4" & Format$(Now(), "yyyymmdd") & ".xls"
                    Call BkReportPart2("S:\ProductionReporting\BankRuptcy4" & Format$(Now(), "yyyymmdd") & ".xls", "4")
                    End If
                        
            Else
            Kill "S:\ProductionReporting\bankruptcy3" & Format$(Now(), "yyyymmdd") & ".xls"
            Call BkReportPart2("S:\ProductionReporting\bankruptcy3" & Format$(Now(), "yyyymmdd") & ".xls", "3")
            End If
                
                
         Else
          Kill "S:\ProductionReporting\bankruptcy2" & Format$(Now(), "yyyymmdd") & ".xls"
         Call BkReportPart2("S:\ProductionReporting\bankruptcy2" & Format$(Now(), "yyyymmdd") & ".xls", "2")
         End If
    Else
    Kill "S:\ProductionReporting\bankruptcy1" & Format$(Now(), "yyyymmdd") & ".xls"
    Call BkReportPart2("S:\ProductionReporting\bankruptcy1" & Format$(Now(), "yyyymmdd") & ".xls", "1")
    End If
    
Else
Kill "S:\ProductionReporting\bankruptcy" & Format$(Now(), "yyyymmdd") & ".xls"
Call BkReportPart2("S:\ProductionReporting\bankruptcy" & Format$(Now(), "yyyymmdd") & ".xls", "0")
End If
End Sub

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdFNMA_Click()
On Error GoTo Err_cmdFNMA_Click
If Not CheckDates() Then Exit Sub

DoCmd.OpenReport "FNMA: Loans-Purchasers", PrintTo

Exit_cmdFNMA_Click:
  Exit Sub

Err_cmdFNMA_Click:
  MsgBox Err.Description
  Resume Exit_cmdFNMA_Click
End Sub

Private Sub cmdInvoiced_Click()
On Error GoTo Err_cmdInvoiced_Click
If Not CheckDates() Then Exit Sub
DoCmd.OpenReport "Invoiced", PrintTo

Exit_cmdInvoiced_Click:
    Exit Sub

Err_cmdInvoiced_Click:
    MsgBox Err.Description
    Resume Exit_cmdInvoiced_Click

End Sub

Private Sub cmdLastMonth_Click()

On Error GoTo Err_cmdLastMonth_Click
DateFrom = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), 1)
DateThru = DateAdd("d", (Day(Date)) * -1, Date)

Exit_cmdLastMonth_Click:
    Exit Sub

Err_cmdLastMonth_Click:
    MsgBox Err.Description
    Resume Exit_cmdLastMonth_Click
    
End Sub

Private Sub cmdLastYear_Click()
On Error GoTo Err_cmdLastYear_Click
DateFrom = DateSerial(Year(DateAdd("yyyy", -1, Date)), 1, 1)
Dim nextYear As Variant
nextYear = DateAdd("yyyy", 1, DateFrom)
DateThru = DateAdd("d", -1, nextYear)


Exit_cmdLastYear_Click:
    Exit Sub

Err_cmdLastYear_Click:
    MsgBox Err.Description
    Resume Exit_cmdLastYear_Click
    
End Sub

Private Sub cmdNewCases_Click()
On Error Resume Next
Kill "S:\ProductionReporting\NewCases" & Format$(Now(), "yyyymmdd") & ".xls"
On Error GoTo 0
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryProductionExport", "S:\ProductionReporting\NewCases" & Format$(Now(), "yyyymmdd") & ".xls"

Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModule.xlsm"
.Run "NewCases"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"

End Sub

Private Sub cmdNewRestarts_Click()
On Error GoTo Err_cmdNewRestarts_Click
If Not CheckDates() Then Exit Sub


' nasty way of doing this but with the way subforms open and determine recordsource before the main form, this works

Dim lqry As QueryDef
Dim lqryName As String
Dim sql As String


Select Case optClient
    Case 1
        sql = ""
    Case 2
        sql = "ClientID=" & cbxClientID
End Select

If Not chMD Then sql = sql & IIf(sql = "", "", " AND ") & "State <> 'MD'"
If Not chDC Then sql = sql & IIf(sql = "", "", " AND ") & "State <> 'DC'"
If Not chVA Then sql = sql & IIf(sql = "", "", " AND ") & "State <> 'VA'"

lqryName = "rqryNewRestartsRespCnt"
On Error Resume Next
DoCmd.DeleteObject acQuery, lqryName


On Error GoTo Err_cmdNewRestarts_Click
Set lqry = CurrentDb.CreateQueryDef(lqryName, "SELECT rqryNewRestarts.Initials, Count(rqryNewRestarts.FileNumber) AS RespCnt " & _
                  "FROM rqryNewRestarts " & _
                  IIf(sql = "", "", "WHERE " & sql & " ") & _
                  "GROUP BY rqryNewRestarts.Initials")
                  
lqry.Close
Set lqry = Nothing

DoCmd.OpenReport "New Restarts", PrintTo

Exit_cmdNewRestarts_Click:
    Exit Sub

Err_cmdNewRestarts_Click:
    MsgBox Err.Description
    Resume Exit_cmdNewRestarts_Click
End Sub

Private Sub cmdOpenAmtReport_Click()

'If chVA.Value = -1 Then
'    Dim StrFileName, MacName As String
'
'
'If FileLocked("S:\ProductionReporting\FCReportVA" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
'
'    If FileLocked("S:\ProductionReporting\FCReportVA1" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
'
'        If FileLocked("S:\ProductionReporting\FCReportVA2" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
'
'            If FileLocked("S:\ProductionReporting\FCReportVA3" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
'
'                    If FileLocked("S:\ProductionReporting\FCReportVA4" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
'                    MsgBox ("Please Check with IT , Might be more than 5 opens this file or SQL problem")
'                    Exit Sub
'                    Else
'                    Kill "S:\ProductionReporting\FCReportVA4" & Format$(Now(), "yyyymmdd") & ".xls"
'                    Call DoExcelFC("S:\ProductionReporting\FCReportVA4" & Format$(Now(), "yyyymmdd") & ".xls", "FCReportVA4")
'                    End If
'
'            Else
'            Kill "S:\ProductionReporting\FCReportVA3" & Format$(Now(), "yyyymmdd") & ".xls"
'            Call DoExcelFC("S:\ProductionReporting\FCReportVA3" & Format$(Now(), "yyyymmdd") & ".xls", "FCReportVA3")
'            End If
'
'
'         Else
'          Kill "S:\ProductionReporting\FCReportVA2" & Format$(Now(), "yyyymmdd") & ".xls"
'         Call DoExcelFC2("S:\ProductionReporting\FCReportVA2" & Format$(Now(), "yyyymmdd") & ".xls", "FCReportVA2")
'         End If
'    Else
'    Kill "S:\ProductionReporting\FCReportVA1" & Format$(Now(), "yyyymmdd") & ".xls"
'    Call DoExcelFC1("S:\ProductionReporting\FCReportVA1" & Format$(Now(), "yyyymmdd") & ".xls", "FCReportVA1")
'    End If
'
'Else
'Kill "S:\ProductionReporting\FCReportVA" & Format$(Now(), "yyyymmdd") & ".xls"
'Call DoExcelFC("S:\ProductionReporting\FCReportVA" & Format$(Now(), "yyyymmdd") & ".xls", "FCReporVA")
'End If
'Else
'
'If chMD.Value = -1 Then
  '  Dim StrFileName, MacName As String


If FileLocked("S:\ProductionReporting\FCReport" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
    
    If FileLocked("S:\ProductionReporting\FCReport1" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
    
        If FileLocked("S:\ProductionReporting\FCReport2" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
            
            If FileLocked("S:\ProductionReporting\FCReport3" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
                    
                    If FileLocked("S:\ProductionReporting\FCReport4" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
                    MsgBox ("Please Check with IT, there may be more than 5 sessions of this file open or a SQL problem")
                    Exit Sub
                    Else
                    Kill "S:\ProductionReporting\FCReport4" & Format$(Now(), "yyyymmdd") & ".xls"
                    Call DoExcelMDFC4("S:\ProductionReporting\FCReport4" & Format$(Now(), "yyyymmdd") & ".xls", "FCRepor4")
                    End If
                        
            Else
            Kill "S:\ProductionReporting\FCReport3" & Format$(Now(), "yyyymmdd") & ".xls"
            Call DoExcelMDFC3("S:\ProductionReporting\FCReport3" & Format$(Now(), "yyyymmdd") & ".xls", "FCRepor3")
            End If
                
                
         Else
          Kill "S:\ProductionReporting\FCReport2" & Format$(Now(), "yyyymmdd") & ".xls"
         Call DoExcelMDFC2("S:\ProductionReporting\FCReport2" & Format$(Now(), "yyyymmdd") & ".xls", "FCRepor2")
         End If
    Else
    Kill "S:\ProductionReporting\FCReport1" & Format$(Now(), "yyyymmdd") & ".xls"
    Call DoExcelMDFC1("S:\ProductionReporting\FCReport1" & Format$(Now(), "yyyymmdd") & ".xls", "FCRepor1")
    End If
    
Else
Kill "S:\ProductionReporting\FCReport" & Format$(Now(), "yyyymmdd") & ".xls"
Call DoExcelMDFC("S:\ProductionReporting\FCReport" & Format$(Now(), "yyyymmdd") & ".xls", "FCRepor")
End If

'End If
'End If


   



'If chVA.Value = -1 Then
'    Dim A As String
'    Dim strFileName, MacName As String
'     A = ""
'   strFileName = "S:\ProductionReporting\FCReportVA" & Format$(Now(), "yyyymmdd") & ".xls"
'   MacName = "FCReportVA" & A
'
''   If Not FileExcelLocked(strFileName) Then
''   MsgBox (" this is closed file")
''   Else
''   MsgBox ("this is open by another")
'' '  DoExcelFC strFileName, MacName
''   End If
'
'   End If
   



''On Error GoTo Err_comdOpenRepot
'If chVA.Value = -1 Then
'On Error Resume Next
''Format(Time, "hh:mm")
'
'Kill "S:\ProductionReporting\FCReportVA" & Format$(Now(), "yyyymmdd") & ".xls"
'
'On Error GoTo 0
'
'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCVA", "S:\ProductionReporting\FCReportVA" & Format$(Now(), "yyyymmdd") & ".xls"
'Dim ExcelObj As Object
'Set ExcelObj = CreateObject("Excel.Application")
'With ExcelObj
'.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModule.xlsm"
'.Run "FCReporVA"
'.ActiveWorkbook.Close
''.Visible
'End With
'Set ExcelObj = Nothing
'MsgBox "The report of FC VA is now ready to view in Excel"
'End If
'
'If chMD.Value = -1 Then
'On Error Resume Next
'
'Kill "S:\ProductionReporting\FCReportMD" & Format$(Now(), "yyyymmdd") & ".xls"
'On Error GoTo 0
'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCMD", "S:\ProductionReporting\FCReportMD" & Format$(Now(), "yyyymmdd") & ".xls"
'Dim ExcelObj1 As Object
'Set ExcelObj1 = CreateObject("Excel.Application")
'With ExcelObj1
'.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModule.xlsm"
'.Run "FCReporMD"
'.ActiveWorkbook.Close
''.Visible
'End With
'Set ExcelObj1 = Nothing
'MsgBox "The report FC MD is now ready to view in Excel"
'End If
'
''Err_comdOpenRepot:
''MsgBox ("Another employee is in the report. Try again later.")
''Exit_cmdOpenAmtReport_Click:
''Exit Sub
''
''
''Resume Exit_cmdOpenAmtReport_Click

End Sub

Private Sub cmdPlanRefRecd_Click()
On Error GoTo Err_cmdPlanRefRecd_Click
If Not CheckDates() Then Exit Sub
DoCmd.OpenReport "Plan Referral Received", PrintTo

Exit_cmdPlanRefRecd_Click:
    Exit Sub

Err_cmdPlanRefRecd_Click:
    MsgBox Err.Description
    Resume Exit_cmdPlanRefRecd_Click

End Sub

Private Sub cmdPOCReceived_Click()
On Error GoTo Err_cmdPOCReceived_Click
If Not CheckDates() Then Exit Sub
DoCmd.OpenReport "POC Received", PrintTo

Exit_cmdPOCReceived_Click:
    Exit Sub

Err_cmdPOCReceived_Click:
    MsgBox Err.Description
    Resume Exit_cmdPOCReceived_Click
End Sub

Private Sub cmdRentPaymentRecd_Click()
On Error GoTo Err_cmdRentPaymentRecd_Click
If Not CheckDates() Then Exit Sub
DoCmd.OpenReport "Rent Payments", PrintTo

Exit_cmdRentPaymentRecd_Click:
    Exit Sub

Err_cmdRentPaymentRecd_Click:
    MsgBox Err.Description
    Resume Exit_cmdRentPaymentRecd_Click

End Sub

Private Sub cmdSalesOccurred_Click()
On Error GoTo Err_cmdSalesOccurred_Click
If Not CheckDates() Then Exit Sub
DoReport "Sales Occurred", PrintTo

Exit_cmdSalesOccurred_Click:
    Exit Sub

Err_cmdSalesOccurred_Click:
    MsgBox Err.Description
    Resume Exit_cmdSalesOccurred_Click
End Sub

Private Sub cmdStatusCivil_Click()
On Error GoTo Err_cmdStatusCivil_Click
DoReport "StatusReportCivil", PrintTo

Exit_cmdStatusCivil_Click:
    Exit Sub

Err_cmdStatusCivil_Click:
    MsgBox Err.Description
    Resume Exit_cmdStatusCivil_Click

End Sub

Private Sub cmdThisMonth_Click()

On Error GoTo Err_cmdThisMonth_Click
DateFrom = DateAdd("d", (Day(Date) - 1) * -1, Date)
DateThru = DateAdd("d", -1, DateAdd("m", 1, DateFrom))

Exit_cmdThisMonth_Click:
    Exit Sub

Err_cmdThisMonth_Click:
    MsgBox Err.Description
    Resume Exit_cmdThisMonth_Click
    
End Sub

Private Sub cmdNextMonth_Click()

On Error GoTo Err_cmdNextMonth_Click
DateFrom = DateSerial(Year(DateAdd("m", 1, Date)), Month(DateAdd("m", 1, Date)), 1)
DateThru = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 2, Date)), Month(DateAdd("m", 2, Date)), 1))

Exit_cmdNextMonth_Click:
    Exit Sub

Err_cmdNextMonth_Click:
    MsgBox Err.Description
    Resume Exit_cmdNextMonth_Click
    
End Sub

Private Function CheckDates() As Boolean
If Not IsDate(Nz(DateFrom)) Or Not IsDate(Nz(DateThru)) Then
    MsgBox "Unrecognized dates", vbCritical
    CheckDates = False
    Exit Function
End If
If DateThru < DateFrom Then DateThru = DateFrom
CheckDates = True
End Function

Private Sub cmdSalesScheduled_Click()

On Error GoTo Err_cmdSalesScheduled_Click
If Not CheckDates() Then Exit Sub
DoReport "Sales Scheduled", PrintTo

Exit_cmdSalesScheduled_Click:
    Exit Sub

Err_cmdSalesScheduled_Click:
    MsgBox Err.Description
    Resume Exit_cmdSalesScheduled_Click
    
End Sub

Private Sub cmdThisYear_Click()
On Error GoTo Err_cmdThisYear_Click
DateFrom = DateSerial(Year(Date), 1, 1)
Dim nextYear As Variant
nextYear = DateAdd("yyyy", 1, DateFrom)
DateThru = DateAdd("d", -1, nextYear)


Exit_cmdThisYear_Click:
    Exit Sub

Err_cmdThisYear_Click:
    MsgBox Err.Description
    Resume Exit_cmdThisYear_Click
End Sub

Private Sub cmdTitleInvestors_Click()
On Error GoTo Err_cmdTitleInvestors_Click

If Not CheckClient() Then Exit Sub

DoReport "Title Investors", PrintTo


Exit_cmdTitleInvestors_Click:
  Exit Sub
  
Err_cmdTitleInvestors_Click:
  MsgBox Err.Description
  Resume Exit_cmdTitleInvestors_Click
  
End Sub

Private Sub cmdTitleResolution_Click()
On Error GoTo Err_cmdTitleResolution_Click

DoReport "StatusReportTitleRes", PrintTo


Exit_cmdTitleResolution_Click:
  Exit Sub
  
Err_cmdTitleResolution_Click:
  MsgBox Err.Description
  Resume Exit_cmdTitleResolution_Click
  
End Sub

Private Sub CommEVReport_Click()
If FileLocked("S:\ProductionReporting\EVReport0" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
    
    If FileLocked("S:\ProductionReporting\EVReport1" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
    
        If FileLocked("S:\ProductionReporting\EVReport2" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
            
            If FileLocked("S:\ProductionReporting\EVReport3" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
                    
                    If FileLocked("S:\ProductionReporting\EVReport4" & Format$(Now(), "yyyymmdd") & ".xls") = True Then
                    MsgBox ("Please Check with IT, there may be more than 5 sessions of this file open or a SQL problem")
                    Exit Sub
                    Else
                    Kill "S:\ProductionReporting\EVReport4" & Format$(Now(), "yyyymmdd") & ".xls"
                    Call EVReportExcel("S:\ProductionReporting\EVReport4" & Format$(Now(), "yyyymmdd") & ".xls", 4)
                    End If
                        
            Else
            Kill "S:\ProductionReporting\EVReport3" & Format$(Now(), "yyyymmdd") & ".xls"
            Call EVReportExcel("S:\ProductionReporting\EVReport3" & Format$(Now(), "yyyymmdd") & ".xls", 3)
            End If
                
                
         Else
          Kill "S:\ProductionReporting\EVReport2" & Format$(Now(), "yyyymmdd") & ".xls"
         Call EVReportExcel("S:\ProductionReporting\EVReport2" & Format$(Now(), "yyyymmdd") & ".xls", 2)
         End If
    Else
    Kill "S:\ProductionReporting\EVReport1" & Format$(Now(), "yyyymmdd") & ".xls"
    Call EVReportExcel("S:\ProductionReporting\EVReport1" & Format$(Now(), "yyyymmdd") & ".xls", 1)
    End If
    
Else
Kill "S:\ProductionReporting\EVReport0" & Format$(Now(), "yyyymmdd") & ".xls"
Call EVReportExcel("S:\ProductionReporting\EVReport0" & Format$(Now(), "yyyymmdd") & ".xls", 0)
End If

End Sub

Private Sub DateFrom_DblClick(Cancel As Integer)
DateFrom = Date
End Sub

Private Sub DateThru_DblClick(Cancel As Integer)
DateThru = Date
End Sub

Private Sub cmdFCStatusPre_Click()

On Error GoTo Err_cmdFCStatusPre_Click
If Not CheckClient() Then Exit Sub

chPostSale = False
DoReport "StatusReportFC", PrintTo, "StatusReport"

Exit_cmdFCStatusPre_Click:
    Exit Sub

Err_cmdFCStatusPre_Click:
    MsgBox Err.Description
    Resume Exit_cmdFCStatusPre_Click
    
End Sub

Private Sub cmdFCStatusPost_Click()

On Error GoTo Err_cmdFCStatusPost_Click

If Not CheckClient() Then Exit Sub

chPostSale = True
DoReport "StatusReportFC", PrintTo, "StatusReportPostSale"

Exit_cmdFCStatusPost_Click:
    Exit Sub

Err_cmdFCStatusPost_Click:
    MsgBox Err.Description
    Resume Exit_cmdFCStatusPost_Click
    
End Sub

Private Function CheckClient() As Boolean
Dim rstFN As Recordset, NumberList As String, i As Integer, FileNumber As Long, cnt As Integer

CheckClient = True
Select Case optCriteria
    Case 2
        If IsNull(cbxClient) Then
            MsgBox "Select a client", vbCritical
            CheckClient = False
        End If
'    Case 3
'        DoCmd.SetWarnings False
'        DoCmd.RunSQL "DELETE * FROM ReportFileNumbers"
'        DoCmd.SetWarnings True
'        NumberList = txtFileNumbers
'        For i = 1 To Len(NumberList)
'            If Not IsNumeric(Mid$(NumberList, i, 1)) Then Mid$(NumberList, i, 1) = ","
'        Next i
'        Set rstFN = CurrentDb.OpenRecordset("ReportFileNumbers", dbOpenDynaset)
'        For i = 0 To UBound(Split(NumberList, ","))
'            FileNumber = Val(Split(NumberList, ",")(i))
'            If FileNumber > 0 Then
'                rstFN.AddNew
'                rstFN!FileNumber = FileNumber
'                rstFN.Update
'                cnt = cnt + 1
'            End If
'        Next i
'        rstFN.Close
'        If cnt = 0 Then MsgBox "No file numbers were specified", vbCritical
'        CheckClient = (cnt > 0)
'        If IsNull(txtFileNumber) Or Not IsNumeric(txtFileNumber) Then
'            MsgBox "Enter a file number", vbCritical
'            CheckClient = False
'            Exit Function
'        End If
'        If IsNull(DLookup("FileNumber", "CaseList", "FileNumber=" & txtFileNumber)) Then
'            MsgBox "No such file number: " & txtFileNumber, vbCritical
'            CheckClient = False
'        End If
End Select
End Function

Private Sub cmdClosedCases_Click()

On Error GoTo Err_cmdClosedCases_Click
If Not CheckDates() Then Exit Sub
DoReport "Closed Files", PrintTo

Exit_cmdClosedCases_Click:
    Exit Sub

Err_cmdClosedCases_Click:
    MsgBox Err.Description
    Resume Exit_cmdClosedCases_Click
    
End Sub

Private Sub cmdBKStatus_Click()
On Error GoTo Err_cmdBKStatus_Click

If Not CheckClient() Then Exit Sub
DoReport "StatusReportBK", PrintTo

Exit_cmdBKStatus_Click:
    Exit Sub

Err_cmdBKStatus_Click:
    MsgBox Err.Description
    Resume Exit_cmdBKStatus_Click
    
End Sub

Private Sub cmd362Filed_Click()

On Error GoTo Err_cmd362Filed_Click

If Not CheckDates() Then Exit Sub
DoReport "362 Filed", PrintTo

Exit_cmd362Filed_Click:
    Exit Sub

Err_cmd362Filed_Click:
    MsgBox Err.Description
    Resume Exit_cmd362Filed_Click
    
End Sub

Private Sub cmdColStatusPre_Click()

On Error GoTo Err_cmdColStatusPre_Click

If Not CheckClient() Then Exit Sub
chPostJudgment = False
DoReport "StatusReportCOL", PrintTo, "StatusReportPreJudgment"

Exit_cmdColStatusPre_Click:
    Exit Sub

Err_cmdColStatusPre_Click:
    MsgBox Err.Description
    Resume Exit_cmdColStatusPre_Click
    
End Sub

Private Sub cmdColStatusPost_Click()

On Error GoTo Err_cmdColStatusPost_Click

If Not CheckClient() Then Exit Sub
chPostJudgment = True
DoReport "StatusReportCOL", PrintTo, "StatusReportPostJudgment"

Exit_cmdColStatusPost_Click:
    Exit Sub

Err_cmdColStatusPost_Click:
    MsgBox Err.Description
    Resume Exit_cmdColStatusPost_Click
    
End Sub

Private Sub cmdReferrals_Click()

On Error GoTo Err_cmdReferrals_Click

If Not CheckDates() Then Exit Sub
DoReport "Referrals", PrintTo

Exit_cmdReferrals_Click:
    Exit Sub

Err_cmdReferrals_Click:
    MsgBox Err.Description
    Resume Exit_cmdReferrals_Click
    
End Sub

Private Sub Form_Close()
If EMailStatus = 1 Then MsgBox "Reminder: EMail is still active", vbInformation
End Sub

Private Sub Form_Current()
If Not PrivChrono Then
Me.Option27.Enabled = False
Me.Option69.Enabled = False
Me.cbxClient.Enabled = False
Me.txtFileNumbers.Enabled = False
Me.cmdTitleInvestors.Enabled = False
Me.cmdFCStatusPre.Enabled = False
Me.cmdFCStatusPost.Enabled = False
Me.cmdEVstatus.Enabled = False
Me.cmdBKStatus.Enabled = False
Me.cmdMonitorSales.Enabled = False
Me.cmdStatusREO.Enabled = False
Me.cmdStatusCivil.Enabled = False
Me.cmdTitleResolution.Enabled = False
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
tglEMail.Enabled = DLookup("iValue", "DB", "Name='EMailPDF'")
If EMailStatus = 1 Then
    tglEMail = True
    MsgBox "Reminder: EMail is still active", vbInformation
End If


End Sub

Private Sub tglEMail_Click()
If tglEMail Then        ' click on
    If Not EMailInit() Then tglEMail = False
Else                    ' click off
    If Not EMailEnd() Then tglEMail = True
End If
End Sub

Private Sub cmdEVstatus_Click()

On Error GoTo Err_cmdEVstatus_Click
DoReport "StatusReportEV", PrintTo

Exit_cmdEVstatus_Click:
    Exit Sub

Err_cmdEVstatus_Click:
    MsgBox Err.Description
    Resume Exit_cmdEVstatus_Click
    
End Sub

Private Sub cmdPOCFiled_Click()

On Error GoTo Err_cmdPOCFiled_Click
If Not CheckDates() Then Exit Sub
DoCmd.OpenReport "POC Filed", PrintTo

Exit_cmdPOCFiled_Click:
    Exit Sub

Err_cmdPOCFiled_Click:
    MsgBox Err.Description
    Resume Exit_cmdPOCFiled_Click
    
End Sub

Private Sub cmdAffidavits_Click()

On Error GoTo Err_cmdAffidavits_Click
If Not CheckDates() Then Exit Sub
DoCmd.OpenReport "AffDef Filed", PrintTo

Exit_cmdAffidavits_Click:
    Exit Sub

Err_cmdAffidavits_Click:
    MsgBox Err.Description
    Resume Exit_cmdAffidavits_Click
    
End Sub

Private Sub cmdNewFC_Click()

On Error GoTo Err_cmdNewFC_Click
If Not CheckDates() Then Exit Sub


' nasty way of doing this but with the way subforms open and determine recordsource before the main form, this works

Dim lqry As QueryDef
Dim lqryName As String
Dim sql As String


Select Case optClient
    Case 1
        sql = ""
    Case 2
        sql = "ClientID=" & cbxClientID
End Select

If Not chMD Then sql = sql & IIf(sql = "", "", " AND ") & "State <> 'MD'"
If Not chDC Then sql = sql & IIf(sql = "", "", " AND ") & "State <> 'DC'"
If Not chVA Then sql = sql & IIf(sql = "", "", " AND ") & "State <> 'VA'"

lqryName = "rqryNewForeclosuresRespCnt"
On Error Resume Next
DoCmd.DeleteObject acQuery, lqryName


On Error GoTo Err_cmdNewFC_Click
Set lqry = CurrentDb.CreateQueryDef(lqryName, "SELECT rqryNewForeclosures.Initials, Count(rqryNewForeclosures.FileNumber) AS RespCnt " & _
                  "FROM rqryNewForeclosures " & _
                  IIf(sql = "", "", "WHERE " & sql & " ") & _
                  "GROUP BY rqryNewForeclosures.Initials")
                  
lqry.Close
Set lqry = Nothing

DoCmd.OpenReport "New Foreclosures", PrintTo

Exit_cmdNewFC_Click:
    Exit Sub

Err_cmdNewFC_Click:
    MsgBox Err.Description
    Resume Exit_cmdNewFC_Click
    
End Sub

Private Sub cmdNewEV_Click()

On Error GoTo Err_cmdNewEV_Click
If Not CheckDates() Then Exit Sub
DoCmd.OpenReport "New Evictions", PrintTo

Exit_cmdNewEV_Click:
    Exit Sub

Err_cmdNewEV_Click:
    MsgBox Err.Description
    Resume Exit_cmdNewEV_Click
    
End Sub

Private Sub cmdSalesInvestors_Click()
On Error GoTo Err_cmdSalesInvestors_Click

If Not CheckDates() Then Exit Sub
DoReport "Sales for Investors", PrintTo

Exit_cmdSalesInvestors_Click:
    Exit Sub

Err_cmdSalesInvestors_Click:
    MsgBox Err.Description
    Resume Exit_cmdSalesInvestors_Click
    
End Sub

Private Sub cmdEnteredAppearance_Click()

On Error GoTo Err_cmdEnteredAppearance_Click
If Not CheckDates() Then Exit Sub
DoReport "Entered Appearance", PrintTo

Exit_cmdEnteredAppearance_Click:
    Exit Sub

Err_cmdEnteredAppearance_Click:
    MsgBox Err.Description
    Resume Exit_cmdEnteredAppearance_Click
    
End Sub

Private Sub cmdMonitorSales_Click()

On Error GoTo Err_cmdMonitorSales_Click
If Not CheckClient() Then Exit Sub
DoReport "StatusReportMON", PrintTo, "StatusReportMonitor"

Exit_cmdMonitorSales_Click:
    Exit Sub

Err_cmdMonitorSales_Click:
    MsgBox Err.Description
    Resume Exit_cmdMonitorSales_Click
    
End Sub

Private Sub cmdProjectReport_Click()

On Error GoTo Err_cmdProjectReport_Click

If IsNull(cbxProject) Then
    MsgBox "Select a project", vbCritical
    Exit Sub
End If

If Not CheckClient() Then Exit Sub
DoReport "StatusReportProject", PrintTo, "StatusReportProject"

Exit_cmdProjectReport_Click:
    Exit Sub

Err_cmdProjectReport_Click:
    MsgBox Err.Description
    Resume Exit_cmdProjectReport_Click
    
End Sub

Private Sub cmdDaysToSale_Click()
Dim rstfiles As Recordset, rstDTS As Recordset, FileNumber As Long, startDate As Date, SaleDate As Date

On Error GoTo Err_cmdDaysToSale_Click

If Not CheckDates() Then Exit Sub

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM DaysToSale"
DoCmd.SetWarnings True

Set rstfiles = CurrentDb.OpenRecordset("qryDaysToSale", dbOpenSnapshot)
Set rstDTS = CurrentDb.OpenRecordset("DaysToSale", dbOpenDynaset, dbSeeChanges)
Do While Not rstfiles.EOF
    If rstfiles!FileNumber <> FileNumber Then
        If FileNumber <> 0 Then
            If startDate <> 0 And SaleDate <> 0 Then
                rstDTS.AddNew
                rstDTS!FileNumber = FileNumber
                rstDTS!DaysToSale = DateDiff("d", startDate, SaleDate)
                rstDTS!SaleDate = SaleDate
                rstDTS.Update
            End If
        End If
        FileNumber = rstfiles!FileNumber
    End If
    Select Case rstfiles!Event
        Case 1, 2       ' referral or restart
            startDate = rstfiles!ActionDate
        Case 5          ' sale
            SaleDate = rstfiles!ActionDate
    End Select
    rstfiles.MoveNext
Loop
rstDTS.Close
rstfiles.Close

DoReport "Days To Sale", PrintTo, "DaysToSale"
DoReport "Days To Sale Detail", PrintTo, "DaysToSaleDetail"

Exit_cmdDaysToSale_Click:
    Exit Sub

Err_cmdDaysToSale_Click:
    MsgBox Err.Description
    Resume Exit_cmdDaysToSale_Click
    
End Sub

Private Sub txtFileNumber_Change()
optCriteria = 3
End Sub

Private Sub cmdDaysToRat_Click()

On Error GoTo Err_cmdDaysToRat_Click

If Not CheckDates() Then Exit Sub
DoReport "Days Sale to Ratification", PrintTo, "DaysToSale"

Exit_cmdDaysToRat_Click:
    Exit Sub

Err_cmdDaysToRat_Click:
    MsgBox Err.Description
    Resume Exit_cmdDaysToRat_Click
    
End Sub

Private Sub cmdObjectionsFiled_Click()

On Error GoTo Err_cmdObjectionsFiled_Click

If Not CheckDates() Then Exit Sub
DoReport "ObjectionsFiled", PrintTo, "ObjectionsFiled"

Exit_cmdObjectionsFiled_Click:
    Exit Sub

Err_cmdObjectionsFiled_Click:
    MsgBox Err.Description
    Resume Exit_cmdObjectionsFiled_Click
    
End Sub

Private Sub cmdRespPoC_Click()
On Error GoTo Err_cmdRespPoC_Click

If Not CheckDates() Then Exit Sub
DoReport "POCObjResp", PrintTo, "Responses to PoC"

Exit_cmdRespPoC_Click:
    Exit Sub

Err_cmdRespPoC_Click:
    MsgBox Err.Description
    Resume Exit_cmdRespPoC_Click
    
End Sub

Private Sub cmdNewREO_Click()

On Error GoTo Err_cmdNewREO_Click

If Not CheckDates() Then Exit Sub
DoReport "New REO", PrintTo

Exit_cmdNewREO_Click:
    Exit Sub

Err_cmdNewREO_Click:
    MsgBox Err.Description
    Resume Exit_cmdNewREO_Click
    
End Sub

Private Sub cmdStatusREO_Click()

On Error GoTo Err_cmdStatusREO_Click
DoReport "StatusReportREO", PrintTo

Exit_cmdStatusREO_Click:
    Exit Sub

Err_cmdStatusREO_Click:
    MsgBox Err.Description
    Resume Exit_cmdStatusREO_Click
    
End Sub

Private Sub cmdREOList_Click()

On Error GoTo Err_cmdREOList_Click
DoReport "REO List", PrintTo

Exit_cmdREOList_Click:
    Exit Sub

Err_cmdREOList_Click:
    MsgBox Err.Description
    Resume Exit_cmdREOList_Click
    
End Sub

Private Sub cmdDaysToDocket_Click()

On Error GoTo Err_cmdDaysToDocket_Click
If Not CheckDates() Then Exit Sub
DoReport "Days To Docket", PrintTo

Exit_cmdDaysToDocket_Click:
    Exit Sub

Err_cmdDaysToDocket_Click:
    MsgBox Err.Description
    Resume Exit_cmdDaysToDocket_Click
    
End Sub

Private Sub cmdDaysToFirstPub_Click()

On Error GoTo Err_cmdDaysToFirstPub_Click
If Not CheckDates() Then Exit Sub
DoReport "Days to First Pub", PrintTo

Exit_cmdDaysToFirstPub_Click:
    Exit Sub

Err_cmdDaysToFirstPub_Click:
    MsgBox Err.Description
    Resume Exit_cmdDaysToFirstPub_Click
    
End Sub

Private Sub cmdDaysToDeedRecorded_Click()
On Error GoTo Err_cmdDaysToDeedRecorded_Click

If Not CheckDates() Then Exit Sub
DoReport "Days Sale to Deed Recorded", PrintTo, "DaysToDeedRecorded"

Exit_cmdDaysToDeedRecorded_Click:
    Exit Sub

Err_cmdDaysToDeedRecorded_Click:
    MsgBox Err.Description
    Resume Exit_cmdDaysToDeedRecorded_Click
    
End Sub

Private Sub cmdDocketToSale_Click()

On Error GoTo Err_cmdDocketToSale_Click

If Not CheckDates() Then Exit Sub
DoReport "Docket To Sale", PrintTo

Exit_cmdDocketToSale_Click:
    Exit Sub

Err_cmdDocketToSale_Click:
    MsgBox Err.Description
    Resume Exit_cmdDocketToSale_Click
    
End Sub

Private Sub cmdReferralToMotion_Click()

On Error GoTo Err_cmdReferralToMotion_Click

If Not CheckDates() Then Exit Sub
DoReport "BK Referral to 362", PrintTo, "ReferralToMotion"

Exit_cmdReferralToMotion_Click:
    Exit Sub

Err_cmdReferralToMotion_Click:
    MsgBox Err.Description
    Resume Exit_cmdReferralToMotion_Click
    
End Sub

Private Sub cmdMotionToOrder_Click()

On Error GoTo Err_cmdMotionToOrder_Click

If Not CheckDates() Then Exit Sub
DoReport "BK 362 to Order", PrintTo, "MotionToOrder"

Exit_cmdMotionToOrder_Click:
    Exit Sub

Err_cmdMotionToOrder_Click:
    MsgBox Err.Description
    Resume Exit_cmdMotionToOrder_Click
    
End Sub

Private Function FileLocked(strFileName As String) As Boolean
   On Error Resume Next
   ' If the file is already opened by another process,
   ' and the specified type of access is not allowed,
   ' the Open operation fails and an error occurs.
   Open strFileName For Binary Access Read Write Lock Read Write As #1
   Close #1
   ' If an error occurs, the document is currently open.
   If Err.Number <> 0 Then
      ' Display the error number and description.
  '    MsgBox "Error #" & str(Err.Number) & " - " & Err.Description
      FileLocked = True
      Err.Clear
    Else
    FileLocked = False
    
   End If
End Function

Private Sub DoExcelFC(strFileName As String, MacName As String)


'Kill StrFileName

On Error GoTo 0

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCVA", strFileName '"S:\ProductionReporting\FCReportVA" & A & Format$(Now(), "yyyymmdd") & ".xls"
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleFC0.xlsm"
.Run MacName
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub

Private Sub DoExcelFC1(strFileName As String, MacName As String)


'Kill StrFileName

On Error GoTo 0

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCVA", strFileName '"S:\ProductionReporting\FCReportVA" & A & Format$(Now(), "yyyymmdd") & ".xls"
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleFC1.xlsm"
.Run "FCReporVA1"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub

Private Sub DoExcelFC2(strFileName As String, MacName As String)


'Kill StrFileName

On Error GoTo 0

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCVA", strFileName '"S:\ProductionReporting\FCReportVA" & A & Format$(Now(), "yyyymmdd") & ".xls"
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleFC2.xlsm"
.Run "FCReportVA2"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub

Private Sub DoExcelFC3(strFileName As String, MacName As String)


'Kill StrFileName

On Error GoTo 0

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCVA", strFileName '"S:\ProductionReporting\FCReportVA" & A & Format$(Now(), "yyyymmdd") & ".xls"
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleFC3.xlsm"
.Run "FCReportVA3"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub
Private Sub DoExcelFC4(strFileName As String, MacName As String)


'Kill StrFileName

On Error GoTo 0

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCVA", strFileName '"S:\ProductionReporting\FCReportVA" & A & Format$(Now(), "yyyymmdd") & ".xls"
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleFC4.xlsm"
.Run "FCReportVA4"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub

Private Sub DoExcelMDFC(strFileName As String, MacName As String)


'Kill StrFileName

On Error GoTo 0

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCAll", strFileName
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCDC", strFileName

'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "Table4", "test.xlsx", True
'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "Query6", "test.xlsx", True
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleFC0.xlsm"
.Run "FCReporMD"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub

Private Sub DoExcelMDFC1(strFileName As String, MacName As String)


'Kill StrFileName

On Error GoTo 0

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCAll", strFileName
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCDC", strFileName

Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleFC1.xlsm"
.Run "FCReporMD1"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub

Private Sub DoExcelMDFC2(strFileName As String, MacName As String)


'Kill StrFileName

On Error GoTo 0

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCAll", strFileName
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCDC", strFileName

Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleFC2.xlsm"
.Run "FCReporMD2"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub

Private Sub DoExcelMDFC3(strFileName As String, MacName As String)


'Kill StrFileName

On Error GoTo 0

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCAll", strFileName
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCDC", strFileName

Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleFC3.xlsm"
.Run "FCReporMD3"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub

Private Sub DoExcelMDFC4(strFileName As String, MacName As String)


'Kill StrFileName

On Error GoTo 0

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCAll", strFileName
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wkflFCDC", strFileName

Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleFC4.xlsm"
.Run "FCReporMD4"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"
End Sub


Private Sub EVReportExcel(strFileName As String, ReportNo As Integer)

'Kill StrFileName

On Error GoTo 0

Select Case ReportNo
Dim ExcelObj As Object
Case 0
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wfEVReport", strFileName

Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleEV0.xlsm"
.Run "EVReporAll0"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"

Case 1
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wfEVReport", strFileName
'Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleEV1.xlsm"
.Run "EVReporAll1"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"

Case 2
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wfEVReport", strFileName
'Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleEV2.xlsm"
.Run "EVReporAll2"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"

Case 3
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wfEVReport", strFileName
'Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleEV3.xlsm"
.Run "EVReporAll3"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"

Case 4
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "wfEVReport", strFileName
'Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Workbooks.Open "\\fileserver\Applications\Database\ExcelAutomationModuleEV4.xlsm"
.Run "EVReporAll4"
.ActiveWorkbook.Close
'.Visible
End With
Set ExcelObj = Nothing
MsgBox "The report is now ready to view in Excel"

End Select

End Sub

