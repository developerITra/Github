Attribute VB_Name = "ExcelUtilities"
Option Compare Database
Option Explicit

Public vStatusBar As Variant

Public Sub OpenExcel(ReportName As String, QueryName) ' Currently forcing saves C:\Database
Dim outputFileName As String

'outputFileName = CurrentProject.Path & "\Export_" & Format(Date, "yyyymmmdd") & ReportName & ".xls"

'outputFileName = "C:\Users\" & Environ$("Username") & _   'For SENDING TO DESKTOP
'    "\Desktop" & "\Export_" & Format(Date, "yyyymmmdd") & ReportName & ".xls"

outputFileName = "C:\Database" & "\Export_" & Format(Date, "yyyymmmdd") & ReportName & ".xls"

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, QueryName, outputFileName, True
        
Dim xlApp As Excel.Application
Dim xlWB As Excel.Workbook
    Set xlApp = New Excel.Application
    'Set xlWB = xlApp.Workbooks.Open
    With xlApp
        .Visible = True
    'Set xlWB = .Workbooks.Open(CurrentProject.Path & "\Export_" & Format(Date, "yyyymmmdd") & ReportName & ".xls", , False)
    
'    Set xlWB = .Workbooks.Open("C:\Users\" & Environ$("Username") & _   'FOR SENDING TO DESKTOP
'    "\Desktop" & "\Export_" & Format(Date, "yyyymmmdd") & ReportName & ".xls", , False)
    
    'May want to change this to a CASE Select Later
       
    If ReportName = "Workflow Civil All Litigation" Then
        xlApp.Quit
        Call ModifyExportedExcelFileFormats("C:\Database" & "\Export_" & Format(Date, "yyyymmmdd") & ReportName & ".xls")
    ElseIf ReportName = "Workflow Docs Out" Then
        Call ModifyExportedExcelFileFormats_new("C:\Database" & "\Export_" & Format(Date, "yyyymmmdd") & ReportName & ".xls")

    Else
     Set xlWB = xlApp.Workbooks.Open("C:\Database" & "\Export_" & Format(Date, "yyyymmmdd") & ReportName & ".xls", , False)
 
    End If
    
   ' xlApp.Quit
    End With


End Sub

Public Sub OutputExcel(ReportName As String, QueryName)
On Error GoTo Err_OutputExcel_Click

Dim Filespec As String
Dim FileName As String
Dim strFilter As String
Dim lngFlags As Long
        
    
FileName = ReportName & ".xls"
    
Select Case ReportName
        'VOLUME REPORTS and whatever other files we want to open directly
        Case "WizardRSICompleted", "WizardRSIICompleted", "WizardFairDebtCompleted", "WizardDemandCompleted", _
            "WizardDocketingCompleted", "WizardServiceCompleted", "WizardborrowerservedCompleted", "WizardServiceMailedCompleted", _
            "WizardFLMACompleted", "WizardSaleSettingCompleted", "WizardSCRAFCCompleted", "WizardRestartCompleted", "WizardIntakeCompleted", _
            "WizardSAICompleted", "WizardVASaleSettingCompleted", "LexisNexisCompleted", "WizardLNNCompleted", "WizardTiteOrderCompleted", _
            "WizardTitleOutCompleted", "WizardTitleReviewCompleted", "WizardNOICompleted", "WizardNOIUpload", "WizardNOIUploadMissing", "Workflow DIL ALL", _
            "Workflow CIVIL All Litigation", "WOrkflow Deeds not Sent", "Workflow Cases to be Closed", "WOrkflow DOA NOt Recorded", "VolumeLitigationBilling", "VolumePSAdvancedBilling", "TrackingLitigationBilling", "TrackingPSAdvancedBilling", _
            "Workflow Files in Pending Status", "TrackingPayOffSent", "TrackingFcDispositio", "TrackingDebtVerified", "TrackingReinstatementSent", "Workflow Docs Out", _
            "WOrkflow DOA Out", "Workflow Assignment Not Received from Client", "VolumeESC", "Workflow Need To Invoice BK", _
            "Workflow Need To Invoice EV", "Workflow Need To Invoice Rent", "Workflow Need To Invoice Servicer Released", "Workflow Need To Invoice TR", "Workflow Need To Invoice DIL", _
            "Workflow Need To Invoice Title", "Workflow Receivables_DIL", "Workflow Need To Invoice FC New", "Workflow Hearing Scheduled FC", "Workflow Hearing Scheduled DC", "Workflow Judgment Entered Need Set Sale", "Workflow Service Deadline DC", "Workflow Cancel Service Due to Disposition"
            
            
            Call OpenExcel(ReportName, QueryName)
                


Case Else
    strFilter = ahtAddFilterItem(strFilter, "Excel Files (*.xls)", "*.XLS")
    Filespec = ahtCommonFileOpenSave(FileName:=FileName, OpenFile:=False, _
    Filter:=strFilter, FilterIndex:=1, DialogTitle:="Export File Name")
    If Len(Filespec) > 0 Then
        DoEvents
        
        
        Select Case ReportName
          Case "Workflow FNMA BK"
            
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, QueryName, Filespec, , "Bankruptcy Inventory"
          
          
          Case "Workflow FNMA FC"
              
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, QueryName, Filespec, , "Foreclosure Inventory"
    
          Case "Workflow FNMA Holds"
              
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, QueryName, Filespec, , "Holds"
              
          Case "Workflow FNMA Missing Docs"
              
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, QueryName, Filespec, , "Missing Docs"
    
          Case "Workflow FNMA Postponements"
              
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, QueryName, Filespec, , "Postponements Cancellations"
    
    
          Case "Workflow FNMA Combined"  ' need to hard-code the query stuff since this report is combined and only prompt for filename once
              
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "rqryFNMAFCInventory", Filespec, , "Foreclosure Inventory"
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "rqryFNMABKInventory", Filespec, , "Bankruptcy Inventory"
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "rqryFNMAHolds", Filespec, , "Holds"
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "rqryFNMAMissingDocs", Filespec, , "Missing Docs"
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "rqryFNMAPostponements", Filespec, , "Postponements Cancellations"
              
          Case "Eviction Status"
          
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "rqryEvictionDC", Filespec, , "DC Eviction Status"
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "rqryEvictionVA", Filespec, , "VA Eviction Status"
              DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "rqryEvictionMDCircuit", Filespec, , "MD Circuit Eviction Status"
          
         ' Case "Workflow Files in Pending Status"
         '       DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, QueryName, Filespec, , "Files in Pending Status"
    
          Case "Fannemae"
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "FannieMaeQuery", Filespec, , "Fannie Mae Data Status"
          Case Else
         
           DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, QueryName, Filespec, , QueryName
    
        End Select
        
    
        
        MsgBox "Data transfer to spreadsheet is finished."
    End If
End Select
  
Exit_OutputExcel_Click:
    Exit Sub

Err_OutputExcel_Click:
    MsgBox Err.Description
    Resume Exit_OutputExcel_Click
    
End Sub

Public Sub OutputSupTest()
  'Call OutputFNMASupportingData("c:\mikki\Workflow FNMA Combined.xls")
  Call OutputFNMASupportingData("c:")
End Sub

Public Sub OutputFNMASupportingData(strFileName As String)

Dim xl
Dim xlSheet
Dim xlwbook
Dim rng

    Dim rst As DAO.Recordset


Set xl = CreateObject("Excel.Application")
Set xlwbook = xl.Workbooks.Open(strFileName)

On Error Resume Next
Set xlSheet = xlwbook.Sheets("Supporting Data")
On Error GoTo 0
If IsEmpty(xlSheet) Then
  Set xlSheet = xlwbook.Sheets.Add
  xlSheet.Name = "Supporting Data"
End If

  xlSheet.Cells(1, 1).Value = "Servicer Name"
  Set rng = xlSheet.Range("A2:I4001")


   
    Set rst = CurrentDb.OpenRecordset("Select UCASE$(LongClientName) from ClientList order by LongClientName")
 '   If (rst.RecordCount > 0) Then
 '       cnt = 1
 '       For Each fld In rst.Fields
 '           wks.Cells(1, cnt).Value = fld.Name
 '           cnt = cnt + 1
 '       Next fld
        Call rng.CopyFromRecordset(rst, 4000, 26)
 '   End If
 
     rst.Close
    Set rst = Nothing
    
    xlSheet.Cells(1, 3).Value = "Last Action Complete"
    xlSheet.Cells(2, 3).Value = "Referral"
    xlSheet.Cells(3, 3).Value = "1st Legal Filed"
    xlSheet.Cells(4, 3).Value = "Service Complete"
    xlSheet.Cells(5, 3).Value = "Judgement Entered"
    xlSheet.Cells(6, 3).Value = "Sale Scheduled"
    xlSheet.Cells(7, 3).Value = "Sale Held"
    
    
    xlSheet.Cells(1, 5).Value = "Occupancy"
    
    
    Set rng = Nothing
    Set xlSheet = Nothing
    
    xlwbook.Save
    xlwbook.Close
    Set xlwbook = Nothing
    Set xl = Nothing


End Sub

Public Sub ModifyExportedExcelFileFormats(sFile As String)
On Error GoTo Err_ModifyExportedExcelFileFormats

    Application.SetOption "Show Status Bar", True

    vStatusBar = SysCmd(acSysCmdSetStatus, "Formatting export file... please wait.")

    Dim xlApp As Object
    Dim xlSheet As Object

    Set xlApp = CreateObject("Excel.Application")
    Set xlSheet = xlApp.Workbooks.Open(sFile).Sheets(1)
    
    With xlApp
            .Application.Sheets("wkflCivLitigation").Select
            '.Application.Cells.Select
            '.Application.Selection.ClearFormats
            .Application.Rows("1:1").Select
            .Application.Selection.Font.Bold = True
            .Application.Cells.Select
            .Application.Selection.RowHeight = 12.75
            .Application.Selection.Columns.AutoFit
            .Application.Range("A2").Select
            .Application.ActiveWindow.FreezePanes = True
            .Application.Range("A1").Select
            .Application.Selection.AutoFilter

            .Application.ActiveWorkbook.Save
            .Application.ActiveWorkbook.Close
           .Quit
    End With
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Visible = True
.Workbooks.Open sFile
End With

  '  Set xlApp = Nothing
  '  Set xlSheet = Nothing

    vStatusBar = SysCmd(acSysCmdClearStatus)

Exit_ModifyExportedExcelFileFormats:
    Exit Sub

Err_ModifyExportedExcelFileFormats:
    vStatusBar = SysCmd(acSysCmdClearStatus)
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_ModifyExportedExcelFileFormats
End Sub

'
'On Error Resume Next
'Kill "S:\ProductionReporting\FinReports\FeesCostsbyVendorDaily" & Format$(Now(), "yyyymmdd") & ".xls"
'Kill "S:\ProductionReporting\FinReports\FeesCostsDaily" & Format$(Now(), "yyyymmdd") & ".xls"
'On Error GoTo 0
'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryDailyFeesCostsbyVendor", "S:\ProductionReporting\FinReports\FeesCostsbyVendorDaily" & Format$(Now(), "yyyymmdd") & ".xls"
'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryDailyFeesCosts", "S:\ProductionReporting\FinReports\FeesCostsDaily" & Format$(Now(), "yyyymmdd") & ".xls"
'Dim ExcelObj As Object
'Set ExcelObj = CreateObject("Excel.Application")
'With ExcelObj
'.Visible = True
'.Workbooks.Open "S:\ProductionReporting\FinReports\FinancialReportingMenu.xlsm"
Public Sub ModifyExportedExcelFileFormats_new(sFile As String)
On Error GoTo Err_ModifyExportedExcelFileFormats_new

    Application.SetOption "Show Status Bar", True

    vStatusBar = SysCmd(acSysCmdSetStatus, "Formatting export file... please wait.")

    Dim xlApp As Object
    Dim xlSheet As Object

    Set xlApp = CreateObject("Excel.Application")
    Set xlSheet = xlApp.Workbooks.Open(sFile).Sheets(1)
    
    With xlApp
            .Application.Sheets("rqryDocsOutNew_Excel").Select
            '.Application.Cells.Select
            '.Application.Selection.ClearFormats
            .Application.Rows("1:1").Select
            .Application.Selection.Font.Bold = True
            .Application.Cells.Select
            .Application.Cells.Font.Name = "Calbri"
            .Application.Cells.Font.Size = 11
            .Application.Selection.RowHeight = 13.75
            .Application.Columns("A:Z").Select
            .Application.Selection.Columns.AutoFit
            '.Application.Range("A2").Select
            '.Application.ActiveWindow.FreezePanes = True
            '.Application.Range("A1").Select
            '.Application.Selection.AutoFilter

            .Application.ActiveWorkbook.Save
            .Application.ActiveWorkbook.Close
           .Quit
    End With
Dim ExcelObj As Object
Set ExcelObj = CreateObject("Excel.Application")
With ExcelObj
.Visible = True
.Workbooks.Open sFile
End With

  '  Set xlApp = Nothing
  '  Set xlSheet = Nothing

    vStatusBar = SysCmd(acSysCmdClearStatus)

Exit_ModifyExportedExcelFileFormats_new:
    Exit Sub

Err_ModifyExportedExcelFileFormats_new:
    vStatusBar = SysCmd(acSysCmdClearStatus)
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_ModifyExportedExcelFileFormats_new
End Sub

