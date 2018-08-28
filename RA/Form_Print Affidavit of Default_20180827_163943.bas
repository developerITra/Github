VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Affidavit of Default"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub By_AfterUpdate()
txtName.Enabled = (By = 0)
txtTitle.Enabled = (By = 0)
End Sub

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
If IsNull([NoDays]) Then
MsgBox (" Please Add Number of days")
Exit Sub
End If

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
GoSub UpdateTimeline
If Judge = "RGM" Then
    Call DoReport("Affidavit of Default RGM", Me.OpenArgs)
Else
  If ([District] = 4 Or [District] = 5) Then
    Call DoReport("Affidavit of Default Rich-Alex", Me.OpenArgs)
  
  Else
  
    Call DoReport("Affidavit of Default", Me.OpenArgs)
  End If
End If
Call DoReport("Debt", Me.OpenArgs)
cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click

UpdateTimeline:
    If MsgBox("Update Affidavit = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
        If IsNull(Forms!BankruptcyDetails!Affidavit) Then
            Forms!BankruptcyDetails!Affidavit = Now()
            AddStatus [FileNumber], Now(), "Filed Affidavit of Default"
            Return
        End If
        If IsNull(Forms!BankruptcyDetails![2ndAff]) Then
            Forms!BankruptcyDetails![2ndAff] = Now()
            AddStatus [FileNumber], Now(), "Filed 2nd Affidavit of Default"
            Return
        End If
        If IsNull(Forms!BankruptcyDetails![3rdAff]) Then
            Forms!BankruptcyDetails![3rdAff] = Now()
            AddStatus [FileNumber], Now(), "Filed 3rd Affidavit of Default"
            Return
        End If
    End If
    Return
    
End Sub

Private Sub cmdBKSpec_Click()
On Error GoTo Err_cmdBKSpec_Click

txtName = "_________________________"
txtTitle = "Bankruptcy Specialist"
By = 0

Exit_cmdBKSpec_Click:
    Exit Sub

Err_cmdBKSpec_Click:
    MsgBox Err.Description
    Resume Exit_cmdBKSpec_Click
    
End Sub

Private Sub Form_Current()
lblAttorney.Caption = [Forms]![BankruptcyPrint]![cbxAttorney]
End Sub
