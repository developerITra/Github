VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QueuesAtty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdAttyReview_Click()
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "que"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End Sub

Private Sub cmdAtty5_Click()
DoCmd.OpenForm "queAttyMilestone6"
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

Private Sub CmdDemand_Click()
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queAttyMilestone1_25"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End Sub

Private Sub cmdFairDebt_Click()
On Error GoTo Err_cmdFairDebtQueue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queAttyMilestone1"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdFairDebtQueue_Click:
    Exit Sub

Err_cmdFairDebtQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdFairDebtQueue_Click
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

Private Sub CmdMang125_Click()
DoCmd.OpenForm "queAttyMilestone1_25Mgr"

End Sub

Private Sub cmdMgr1_5_Click()
DoCmd.OpenForm "queAttyMilestone1_5Mgr"
End Sub

Private Sub cmdMgr1a_Click()
DoCmd.OpenForm "queAttyMilestone1Mgr"
End Sub
Private Sub cmdMgr1_Click()
DoCmd.OpenForm "queAttyMilestone2Mgr"
End Sub
Private Sub cmdMgr2_Click()
DoCmd.OpenForm "queAttyMilestone3Mgr"
End Sub
Private Sub cmdMgr3_Click()
DoCmd.OpenForm "queAttyMilestone4Mgr"
End Sub
Private Sub cmdMgr4_Click()
DoCmd.OpenForm "queAttyMilestone5Mgr"
End Sub
Private Sub cmdMgr5_Click()
DoCmd.OpenForm "queAttyMilestone6Mgr"
End Sub

Private Sub cmdNOI_Click()
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queAttyMilestone1_5"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End Sub

Sub cmdrsIIqueue_click()


On Error GoTo Err_cmdRSIIQueue_Click

    Dim stDocName As String
    

    stDocName = "queAttyMilestone2"
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

    stDocName = "queAttyMilestone4"
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

    stDocName = "queAttyMilestone5"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdNOIdocsQueue_Click:
    Exit Sub

Err_cmdNOIdocsQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdNOIdocsQueue_Click
End Sub



Private Sub cmdSaleSet_Click()
   DoCmd.OpenForm "queAttyMilestone3"
End Sub


