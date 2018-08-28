VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Queues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub cmdAttyReview_Click()
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queuesatty"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End Sub



Private Sub cmdBorrowerServed_Click()
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queBorrowerServed"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End Sub

Private Sub cmdClose_Click()

DoCmd.Close

    
End Sub



Private Sub cmdDeceaQueue_Click()
If Not LexisNexis Then
MsgBox ("You are not Authorized to Access this Queue")
Exit Sub
Else

On Error GoTo Err_cmdDeceaQueue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queDecea"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdDeceaQueue_Click:
    Exit Sub

Err_cmdDeceaQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdDeceaQueue_Click
End If

End Sub

Private Sub cmdDemandQueue_Click()
On Error GoTo Err_cmdFairDebtQueue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queDemand"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdFairDebtQueue_Click:
    Exit Sub

Err_cmdFairDebtQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdFairDebtQueue_Click
End Sub

Private Sub cmdDemandWaiting_Click()
On Error GoTo Err_cmdFairDebtQueue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queDemandWaiting"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdFairDebtQueue_Click:
    Exit Sub

Err_cmdFairDebtQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdFairDebtQueue_Click
End Sub

Private Sub cmdDocketing_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queDocketing"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End Sub

Private Sub cmdDocketingWaiting_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queDocketingWaiting"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End Sub

Private Sub cmdFairDebt_Click()
On Error GoTo Err_cmdFairDebtQueue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queFairDebt"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdFairDebtQueue_Click:
    Exit Sub

Err_cmdFairDebtQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdFairDebtQueue_Click
End Sub
Private Sub cmdFairDebtWaiting_Click()
On Error GoTo Err_cmdFairDebtQueue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queFairDebtWaiting"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdFairDebtQueue_Click:
    Exit Sub

Err_cmdFairDebtQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdFairDebtQueue_Click
End Sub

Private Sub cmdFLMA_Click()
DoCmd.OpenForm "queFLMA"
End Sub

Private Sub cmdIntakeQueue_Click()
DoCmd.OpenForm "queIntake"
End Sub

Private Sub cmdIntakeWaiting_Click()
  

'Dim rs As Recordset
'Set rs = CurrentDb.OpenRecordset("Select * FROM qryqueueIntakeWaitinglst_P", dbOpenDynaset, dbSeeChanges)
'rs.Close
'Set rs = Nothing


DoCmd.OpenForm "queIntakeWaiting"
End Sub

Private Sub CmdLimb_Click()
DoCmd.OpenForm "QueueLimboForm"

End Sub

Private Sub cmdRestart_Click()
Dim stDocName As String
    

    stDocName = "queRestart"
    DoCmd.OpenForm stDocName
End Sub

Private Sub cmdRestartRSI_Click()
Dim stDocName As String

    stDocName = "queRSIReview"
    DoCmd.OpenForm stDocName
End Sub

Private Sub cmdRestartWaiting_Click()
Dim stDocName As String

    stDocName = "queRestartWaiting"
    DoCmd.OpenForm stDocName
End Sub

Sub cmdrsIIqueue_click()


On Error GoTo Err_cmdRSIIQueue_Click

    Dim stDocName As String
    

    stDocName = "queRSII"
    DoCmd.OpenForm stDocName

Exit_cmdRSIIQueue_Click:
    Exit Sub

Err_cmdRSIIQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdRSIIQueue_Click
End Sub
Sub cmdNOInewqueue_click()


On Error GoTo Err_cmdNOInewQueue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queNOInew"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdNOInewQueue_Click:
    Exit Sub

Err_cmdNOInewQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdNOInewQueue_Click
End Sub
Sub cmdNOIdocsqueue_click()


On Error GoTo Err_cmdNOIdocsQueue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queNOIdocs"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdNOIdocsQueue_Click:
    Exit Sub

Err_cmdNOIdocsQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdNOIdocsQueue_Click
End Sub


Private Sub cmdSAIqueue_Click()
DoCmd.OpenForm "QueSAI"
End Sub

Private Sub cmdSaleSetting_Click()
DoCmd.OpenForm "QueSaleSetting"
End Sub

Private Sub cmdSaleWaiting_Click()
DoCmd.OpenForm "QueSaleSettingwaiting"
End Sub

Private Sub cmdSCRA_Click()
DoCmd.OpenForm "QueuesSCRA"
End Sub

Private Sub cmdService_Click()
DoCmd.OpenForm "QueService"
End Sub

Private Sub cmdServiceMailed_Click()
DoCmd.OpenForm "queServiceMailed"
End Sub

Private Sub cmdTitleOrder_Click()
If Not TitleOrder Then
MsgBox ("You are not Authorized to Access this Queue")
Exit Sub
Else
With DoCmd
.SetWarnings False
.OpenQuery "TitleOrderUpdatInsert"
.SetWarnings True
End With
  DoCmd.OpenForm "queTitelOrder"
End If
End Sub

Private Sub CmdtitleOut_Click()
DoCmd.OpenForm "queTitleOut"
End Sub

Private Sub CmdtitleReview_Click()
DoCmd.OpenForm "queTitleReview"
End Sub

Private Sub cmdVAappriasal_Click()
DoCmd.OpenForm "Quevaappraisal"
End Sub

Private Sub cmdVALNN_Click()
DoCmd.OpenForm "queVASaleSettingSub"
End Sub

Private Sub cmdVASaleSetting_Click()
    DoCmd.OpenForm "queVAsalesetting"
    
End Sub

Private Sub cmdVAsalewaiting_Click()
DoCmd.OpenForm "queVAsalesettingwaiting"
End Sub

Private Sub Command83_Click()
DoCmd.OpenForm "queTitelOrder"
  
End Sub





Private Sub ComdAccou_Click()
DoCmd.OpenForm "QueueAccounting"

End Sub

Private Sub Form_Open(Cancel As Integer)
If PrivAttyQueue Then cmdAttyReview.Visible = True
If PrivSCRA Then cmdSCRA.Visible = True
If GetStaffID = 1 Or GetStaffID = 103 Or GetStaffID = 458 Then
cmdDocketingWaiting.Enabled = True
cmdDocketing.Enabled = True
End If

End Sub


Private Sub Command86_Click()
On Error GoTo Err_Command86_Click

    Dim stDocName As String

    stDocName = "TitleOrderQueue"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_Command86_Click:
    Exit Sub

Err_Command86_Click:
    MsgBox Err.Description
    Resume Exit_Command86_Click
    
End Sub
