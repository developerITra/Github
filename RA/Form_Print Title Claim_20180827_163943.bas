VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Title Claim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

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

If optClaim Then    ' if title claim
    If MsgBox("Update Title Claim Sent = " & Format$(Date, "m/d/yyyy") & vbNewLine & "and clear Title Claim Resolved" & vbNewLine & "and add to status?", vbYesNo) = vbYes Then
        TitleClaim = True
        TitleClaimSent = Date
        TitleClaimResolved = Null
        AddStatus FileNumber, Now(), "Title Claim sent"
    End If
End If
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
Call DoReport("Title Claim", Me.OpenArgs)

cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub optClaim_AfterUpdate()
If optClaim Then
    lblClaim.Caption = "Claim Number:"
Else
    lblClaim.Caption = "Commitment Number:"
End If
End Sub
