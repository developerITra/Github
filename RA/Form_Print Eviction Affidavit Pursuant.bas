VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Eviction Affidavit Pursuant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim PrintTo As Integer, ContactType As String


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
Dim statusMsg As String

On Error GoTo Err_cmdOK_Click

If IsNull(Me!optDeterminedBy) Then
    MsgBox "Choose determined by", vbCritical
    Exit Sub
End If

If (Me!optDeterminedBy < 4) Then
  If IsNull(Me!DateDetermined) Then
    MsgBox "Enter date determined", vbCritical
    Exit Sub
  End If
Else
  If IsNull(Me!txtOtherDetermination) Then
    MsgBox "Enter other determination", vbCritical
    Exit Sub
  End If
End If
    
Dim strDetermination As String

Select Case optDeterminedBy

    Case 1
        strDetermination = "the purchaser/lender’s agent personally visited the property on " & Format$(Me.DateDetermined, "short Date") & "."
    Case 2
        strDetermination = "the personal process server personally went to the property on " & Format$(Me.DateDetermined, "short Date") & "."
    Case 3
        strDetermination = "an inspection of the property conducted on " & Format$(Me.DateDetermined, "short Date") & "."
    Case 4
        strDetermination = Me.txtOtherDetermination
        
End Select

Call DoReport("Eviction Affidavit 14-102b", PrintTo, , strDetermination)

If MsgBox("Add to status: " & Format$(Date, "mm/dd/yyyy") & " " & statusMsg, vbYesNo + vbQuestion) = vbYes Then
    AddStatus Me!FileNumber, Date, "Eviction Affidavit Pursuant Sent"
End If
cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub


Private Sub Form_Current()
PrintTo = Int(Split(Me.OpenArgs, "|")(0))
Me.DateDetermined.Enabled = False

End Sub





Private Sub optDeterminedBy_AfterUpdate()
  If (Me!optDeterminedBy < 4) Then
    Me.DateDetermined.Enabled = True
  Else
    Me.DateDetermined.Enabled = False
  End If
End Sub

