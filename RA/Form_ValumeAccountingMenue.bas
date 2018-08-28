VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ValumeAccountingMenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub CmdAccouLitigBill_Click()

If Not CheckDatesOK() Then Exit Sub
Call DoReport("VolumeLitigationBilling", Forms!ReportsMenu_Volume!PrintTo)



End Sub
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
Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdSCRA1_Click()

    Dim stDocName As String
  
    stDocName = "queSCRA1"
    DoCmd.OpenForm stDocName

End Sub



Private Sub cmdSCRA2_Click()
Dim stDocName As String

    stDocName = "queSCRA2"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA3_Click()
Dim stDocName As String

    stDocName = "queSCRA3"
    DoCmd.OpenForm stDocName
End Sub


Private Sub cmdSCRA4a_Click()
Dim stDocName As String

    stDocName = "queSCRA4a"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA4b_Click()
Dim stDocName As String

    stDocName = "queSCRA4b"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA5_Click()
Dim stDocName As String

    stDocName = "queSCRA5"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA6_Click()
Dim stDocName As String

    stDocName = "queSCRA6"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA7_Click()
Dim stDocName As String

    stDocName = "queSCRA7"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA8_Click()
Dim stDocName As String

    stDocName = "queSCRA8"
    DoCmd.OpenForm stDocName
End Sub

Private Sub cmdSCRA9_Click()
Dim stDocName As String

    stDocName = "queSCRA9"
    DoCmd.OpenForm stDocName
End Sub

Private Sub cmdSCRA9Waiting_Click()
Dim stDocName As String

    stDocName = "queSCRA9Waiting"
    DoCmd.OpenForm stDocName
End Sub


Private Sub cmdSCRAUnionNew_Click()
 stDocName = "queSCRAFCNew"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Command75_Click()
stDocName = "queSCRABK"
    DoCmd.OpenForm stDocName
End Sub



Private Sub ComESc_Click()
If Not CheckDatesOK() Then Exit Sub
Call DoReport("VolumeESC", Forms!ReportsMenu_Volume!PrintTo)
End Sub

Private Sub ComValumPS_Click()
If Not CheckDatesOK() Then Exit Sub
Call DoReport("VolumePSAdvancedBilling", Forms!ReportsMenu_Volume!PrintTo)

End Sub
