VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CollectionPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClear_Click()
On Error GoTo Err_cmdClear_Click

chFairDebtDeficiency = 0
chFairDebtSuitNote = 0
chMilitaryAffidavit = 0
chSOD = 0
chMotionDeficiency = 0
chSuitNote = 0
chReqEntryOrderDefault = 0
chReqJudgmentDefault = 0

Exit_cmdClear_Click:
    Exit Sub

Err_cmdClear_Click:
    MsgBox Err.Description
    Resume Exit_cmdClear_Click
    
End Sub

Private Sub PrintDocs(PrintTo As Integer)
Dim ReportName As String

On Error GoTo Err_PrintDocs

If chFairDebtDeficiency Then
    Call DoReport("Fair Debt Letter Deficiency", PrintTo)
    If MsgBox("Update Fair Debt Letter Sent = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
        Forms!CollectionDetails!FairDebtLetter = Now()
        AddStatus [CaseList.FileNumber], Now(), "Fair Debt Letter sent"
    End If
End If

If chFairDebtSuitNote Then
    Call DoReport("Fair Debt Letter Suit on Note", PrintTo)
    If MsgBox("Update Fair Debt Letter Sent = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
        Forms!CollectionDetails!FairDebtLetter = Now()
        AddStatus [CaseList.FileNumber], Now(), "Fair Debt Letter sent"
    End If
End If

If chMilitaryAffidavit Then
    If Me!State = "MD" Then
        Call DoReport("Military Affidavit Collection MD", PrintTo)
    Else
        Call DoReport("Military Affidavit Collection", PrintTo)
    End If
End If

If chSOD Then Call DoReport("Statement of Debt", PrintTo)

If chMotionDeficiency Then Call DoReport("Motion for Deficiency", PrintTo)

If chSuitNote Then Call DoReport("Suit on Note", PrintTo)

If chReqEntryOrderDefault Then DoReport "Suit on Note Request Entry of Order", PrintTo

If chReqJudgmentDefault Then DoReport "Suit on Note Request Judgment by Default", PrintTo

Exit Sub

Err_PrintDocs:
    MsgBox Err.Description
    Exit Sub

End Sub

Private Sub Form_Current()
Me.Caption = "Print Collection " & [CaseList.FileNumber] & " " & [PrimaryDefName]
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

Private Sub cmdWord_Click()
Call PrintDocs(-1)
End Sub

Private Sub cmdPrint_Click()
Call PrintDocs(acViewNormal)
End Sub

Private Sub cmdView_Click()
Call PrintDocs(acPreview)
End Sub
