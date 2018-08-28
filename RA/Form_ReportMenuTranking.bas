VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ReportMenuTranking"
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



Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub


Private Sub cmdFeesCostSent_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "TrackingFeeCostSent", PrintTo
End Sub

Private Sub CmdLitit_Click()


If Not CheckDatesOK Then Exit Sub
    DoReport "TrackingLitigationBilling", PrintTo

End Sub

Private Sub ComDebtVerified_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "TrackingDebtVerified", PrintTo

End Sub

Private Sub ComFCDisposition_Click()
If Not CheckDatesOK Then Exit Sub
    DoReport "TrackingFcDisposition", PrintTo
End Sub


Private Sub ComPayOffSent_Click()

If Not CheckDatesOK Then Exit Sub
    DoReport "TrackingPayOffSent", PrintTo
End Sub

Private Sub ComPS_Click()


If Not CheckDatesOK Then Exit Sub
    DoReport "TrackingPSAdvancedBilling", PrintTo
End Sub

Private Sub ComReinstSent_Click()
If Not CheckDatesOK Then Exit Sub
DoReport "TrackingReinstatementSent", PrintTo


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
    dt1 = Forms!ReportMenuTranking!DateFrom
    dt2 = Forms!ReportMenuTranking!DateThru
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


