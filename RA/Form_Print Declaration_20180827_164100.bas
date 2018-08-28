VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Declaration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub chExhibit_AfterUpdate()
If chExhibit Then
If Forms![Case List]![ClientID] = 6 Then
DoCmd.OpenForm " Print Exhibit 9", , , , , acDialog, Me.OpenArgs
End If
End If

End Sub

Private Sub cmdCancel_Click()
DoCmd.Close

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdOK_Click()
'If Me.Dirty Then
'DoCmd.RunCommand acCmdSaveRecord
'If Forms![Case list]![ClientID] = 97 Then
'    Call DoReport("Declaration", Me.OpenArgs)
'Else
'    If Forms![Case list]![ClientID] = 6 Or Forms![Case list]![ClientID] = 556 Then
'        If IsNull(Forms![Print Declaration]!chExhibit) Then
'            Call DoReport("Declaration Wells", Me.OpenArgs)
'        End If
'    End If
'End If

'8/31/2015 need to put back below

DoCmd.SetWarnings False
DoCmd.OpenQuery ("DeletePostPetitionFeesTax")
DoCmd.OpenQuery ("Append_Post-PetitionFees_Tax_Insurance_1")
DoCmd.OpenQuery ("Append_Post-PetitionFees_Tax_Insurance_2")
DoCmd.SetWarnings True

If Forms![Case List]![ClientID] = 97 Then
    Call DoReport("Declaration", Me.OpenArgs)
ElseIf Forms![Case List]![ClientID] = 6 Or Forms![Case List]![ClientID] = 556 Then
    If IsNull(Forms![Print Declaration]!chExhibit) Then
        Call DoReport("Declaration Wells", Me.OpenArgs)
    End If
ElseIf Forms![Case List]![ClientID] = 385 Then
    Call DoReport("Declaration Nationstar", Me.OpenArgs)
End If
   
End Sub

Private Sub Form_Current()
If Forms![Case List]![ClientID] = 97 Or Forms![Case List]![ClientID] = 385 Then Frame94.Visible = True
End Sub

Private Sub Frame94_AfterUpdate()

'8/31/15 need to put back below line
If Forms![Case List]![ClientID] = 385 Then Exit Sub


If Frame94.Value = 2 Then
If MsgBox("Is the Note properly endorsed?", vbYesNo, "Note Properly Endorsed") = vbYes Then
Endorsed = "The Movant's agent has possession of the original promissory note, and the note is payable to the Movant."
Else
Endorsed = "The Movant's agent has possession of the original promissory note, and the note is endorsed in blank."
End If
Else
If Frame94.Value = 1 Then Endorsed = "The original promissory note is lost or missing. A copy of an affidavit of lost or missing note, [with a copy of the note], is attached as Exhibit B."
End If

End Sub
