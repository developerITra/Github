VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QueuesSCRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


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
