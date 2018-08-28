VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ReportsMenu_Volume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdBorrowerServed_Click()
 If Not CheckDatesOK Then Exit Sub
    DoReport "WizardborrowerservedCompleted", PrintTo
End Sub

Private Sub cmdDaysToDocket_Click()
 If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSAICompleted", PrintTo
End Sub

Private Sub cmdDocketing_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "WizardDocketingCompleted", PrintTo
End Sub

Private Sub cmdFLMA_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "WizardFLMACompleted", PrintTo
End Sub

Private Sub cmdIntake_Click()
 If Not CheckDatesOK Then Exit Sub
    DoReport "WizardIntakeCompleted", PrintTo
End Sub

Private Sub cmdLNN_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "WizardLNNCompleted", PrintTo
End Sub

Private Sub cmdSaleSetting_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSaleSettingCompleted", PrintTo
End Sub

Private Sub cmdSCRABK_Click()
    If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRABKCompleted", PrintTo
End Sub

Private Sub cmdService_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "WizardServiceCompleted", PrintTo
End Sub

Private Sub cmdServiceMailed_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "WizardServiceMailedCompleted", PrintTo
End Sub

Private Sub CmdtitleOut_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "WizardTitleOutCompleted", PrintTo
End Sub

Private Sub CmdtitleReview_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "WizardTitleReviewCompleted", PrintTo
End Sub

Private Sub cmdVASaleSetting_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "WizardVAsalesettingCompleted", PrintTo
End Sub

Private Sub ComAccoun_Click()
DoCmd.OpenForm "ValumeAccountingMenue"
End Sub

Private Sub ComLIMBO_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "VolumeLimboReport", PrintTo
End Sub

Private Sub Form_Open(Cancel As Integer)
tglEMail.Enabled = DLookup("iValue", "DB", "Name='EMailPDF'")
If EMailStatus = 1 Then
    tglEMail = True
    MsgBox "Reminder: EMail is still active", vbInformation
End If
End Sub

Private Sub Form_Close()
If EMailStatus = 1 Then MsgBox "Reminder: EMail is still active", vbInformation
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

Private Sub DateFrom_DblClick(Cancel As Integer)
DateFrom = Date
End Sub

Private Sub DateThru_DblClick(Cancel As Integer)
DateThru = Date
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

Private Sub cmdLastWeek_Click()
Dim mon As Date

On Error GoTo Err_cmdLastWeek_Click
'What was this past Monday?
mon = DateAdd("d", 1 - Weekday(Now(), vbMonday), Now())

DateFrom = DateAdd("d", -7, mon)
DateThru = DateAdd("d", -3, mon)

Exit_cmdLastWeek_Click:
    Exit Sub

Err_cmdLastWeek_Click:
    MsgBox Err.Description
    Resume Exit_cmdLastWeek_Click
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
Private Sub cmdYesterday_Click()

On Error GoTo Err_cmdYesterday_Click
DateFrom = Date - 1
DateThru = Date - 1


Exit_cmdYesterday_Click:
    Exit Sub

Err_cmdYesterday_Click:
    MsgBox Err.Description
    Resume Exit_cmdYesterday_Click
    
End Sub
Private Sub cmdToday_Click()

On Error GoTo Err_cmdToday_Click
DateFrom = Date
DateThru = Date


Exit_cmdToday_Click:
    Exit Sub

Err_cmdToday_Click:
    MsgBox Err.Description
    Resume Exit_cmdToday_Click
    
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


Private Sub cmdBulkNOImissing_Click()
    'Does Today
    DoReport "WizardNOIuploadmissing", PrintTo
End Sub

Private Sub cmdBulkNOI_Click()
    'Does today
    DoReport "WizardNOIUpload", PrintTo
End Sub

Private Sub CmdDemand_Click()
    If Not CheckDatesOK Then Exit Sub
    DoReport "WizardDemandCompleted", PrintTo
End Sub

Private Sub cmdFairDebt_Click()
    If Not CheckDatesOK Then Exit Sub
    DoReport "WizardFairDebtCompleted", PrintTo
End Sub

Private Sub cmdNOI_Click()
    If Not CheckDatesOK Then Exit Sub
    DoReport "WizardNOICompleted", PrintTo
End Sub

Private Sub cmdReferrals_Click()
    If Not CheckDatesOK Then Exit Sub
    DoReport "WizardRestartCompleted", PrintTo
End Sub

Private Sub cmdRSII_Click()
    If Not CheckDatesOK Then Exit Sub
    DoReport "WizardRSIICompleted", PrintTo
End Sub

Private Sub cmdRSI_Click()
    If Not CheckDatesOK Then Exit Sub
    DoReport "WizardRSICompleted", PrintTo
End Sub

Private Sub cmdSCRA1_Click()
    If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRA1Completed", PrintTo
End Sub

Private Sub cmdSCRA2_Click()
 If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRA2Completed", PrintTo
End Sub

Private Sub cmdSCRA3_Click()
    If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRA3Completed", PrintTo
End Sub

Private Sub cmdSCRA4a_Click()
     If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRA4aCompleted", PrintTo
End Sub

Private Sub cmdSCRA4b_Click()
     If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRA4bCompleted", PrintTo
End Sub

Private Sub cmdSCRA5_Click()
     If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRA5Completed", PrintTo
End Sub

Private Sub cmdSCRA6_Click()
     If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRA6Completed", PrintTo
End Sub

Private Sub cmdSCRA7_Click()
     If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRA7Completed", PrintTo
End Sub

Private Sub cmdSCRA8_Click()
     If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRA8Completed", PrintTo
End Sub

Private Sub cmdSCRAFC_Click()
    If Not CheckDatesOK Then Exit Sub
    DoReport "WizardSCRAFCCompleted", PrintTo
End Sub

Private Sub LexisNexis_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "LexisNexisCompleted", PrintTo
End Sub

Private Sub tglEMail_Click()
If tglEMail Then        ' click on
    If Not EMailInit() Then tglEMail = False
Else                    ' click off
    If Not EMailEnd() Then tglEMail = True
End If
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

Private Function CheckDatesOK() As Boolean
    Dim dt1 As Date, dt2 As Date
    Dim eMsg As String
    eMsg = ""
    On Error Resume Next
    dt1 = Forms!ReportsMenu_Volume!DateFrom
    dt2 = Forms!ReportsMenu_Volume!DateThru
    On Error GoTo 0

    If (1899 = Year(dt1)) And (1899 = Year(dt2)) Then
        eMsg = "Please fill-in dates or select date range."
    ElseIf (1899 = Year(dt1)) Then
        eMsg = "Please fill-in From Date, or select date range."
    ElseIf (1899 = Year(dt2)) Then
        eMsg = "Please fill-in Through Date, or select date range."
    ElseIf (dt1 > dt2) Then
        eMsg = "From Date must not be after Through Date."
    End If
    
    If "" <> eMsg _
    Then
        MsgBox eMsg, vbExclamation, "Valid date range must be supplied"
        CheckDatesOK = False
        Exit Function
    End If
    CheckDatesOK = True
End Function
Private Sub Command148_Click()

On Error Resume Next
Kill "S:\TitlesOrdered\TitleOrder" & ".xls"
On Error GoTo 0
'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "TitleOrderQ", "S:\TitlesOrdered\TitleOrder" & ".xls"
DoCmd.OutputTo acOutputQuery, "TitleOrderQ", acFormatXLS, TemplatePath & "TitleOrderT.xlt", 1, , True



'Dim ExcelObj As Object
'Set ExcelObj = CreateObject("Excel.Application")
'With ExcelObj
'.Workbooks.Open "\\fileserver\Applications\Database\TitleOrderT.xlsm"
'.Run "Mod_TitleOrder"
'.ActiveWorkbook.Close
''.Visible
'End With
'Set ExcelObj = Nothing
'MsgBox "The report is now ready to view in Excel"
    
End Sub

Private Sub TitleOrder_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "WizardTiteOrderCompleted", PrintTo
End Sub
