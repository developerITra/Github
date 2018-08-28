VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Motion for Relief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim strRptName As String
Dim PrintTo As Integer
Dim Debtor_CoDebtor As Integer


Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click

If Forms![Case List]![ClientID] = 97 And Forms![BankruptcyDetails]![Chapter] = 13 Then
    Text33.Value = "Yes"
    If MsgBox("   Is the Note Lost?   ", vbYesNo, "Note Lost") = vbYes Then
    Text31.Value = "Yes"
    Else
        If MsgBox("  Is the Note properly endorsed?  ", vbYesNo, "Note Properly Endorsed") = vbYes Then
        Text34.Value = "Yes"
        Else
        Text34.Value = "No"
        End If
    Text31.Value = "No"
    End If
Else
    Text33.Value = "No"
End If


'added on 4/29/15

If Forms![Case List]![ClientID] = 328 Then 'And Forms![BankruptcyDetails]![Chapter] = 13 Then
    Text33.Value = "Yes"

    If MsgBox("   Is the Note Lost?   ", vbYesNo, "Note Lost") = vbYes Then
    Text31.Value = "Yes"
    Else
        If MsgBox("  Is the Note properly endorsed?  ", vbYesNo, "Note Properly Endorsed") = vbYes Then
        Text34.Value = "Yes"
        Else
        Text34.Value = "No"
        End If
    End If

ElseIf Forms![Case List]![ClientID] <> 97 And Forms![BankruptcyDetails]![Chapter] <> 13 Then
    Text33.Value = "No"
End If


Dim strStatus As String

Call DoReport(strRptName, PrintTo)
Call DoReport("Debt", PrintTo)
If MsgBox("Update timeline: 362 = " & Format$(Date, "mm/dd/yyyy") & "?", vbYesNo + vbQuestion) = vbYes Then
    If (Debtor_CoDebtor = 0) Then  ' debtor
      Forms!BankruptcyDetails![362] = Date
      strStatus = "Filed debtor motion for relief from automatic stay"
    Else
      Forms!BankruptcyDetails![362_CoDebtor] = Date
      strStatus = "Filed codebtor motion for relief from automatic stay"
    End If
    AddStatus FileNumber, Date, strStatus
End If

cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Current()

If Forms![Case List]![ClientID] = 385 Then
    Frame36.Visible = True
    Label35.Visible = True
Else
    Frame36.Visible = False
    Label35.Visible = False
End If

DueDate = Format$(Date, "m/d/yyyy")
Select Case Chapter
    Case 7
        Payment.Enabled = False
    Case 11, 13
        Payment.Enabled = True
End Select

Debtor_CoDebtor = Int(Split(Me.OpenArgs, "|")(2))
PrintTo = Int(Split(Me.OpenArgs, "|")(1))
strRptName = Split(Me.OpenArgs, "|")(0)
End Sub
