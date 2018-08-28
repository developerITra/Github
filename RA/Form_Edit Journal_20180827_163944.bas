VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Edit Journal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
If Not PrivJournalFlags Then Cancel = 1
End Sub

Private Sub Warning_AfterUpdate()

Dim TextMsg As String
TextMsg = "You are not authorized to make changes"
Select Case Warning

Case 50
If Not PrivWaitingForBill Then
MsgBox (TextMsg)
Me.Undo
Exit Sub
End If

Case 100
If Not PrivWaitingForDoc Then
MsgBox (TextMsg)
Me.Undo
Exit Sub
End If

Case 200
If Not PrivTitleIssue Then
MsgBox (TextMsg)
Me.Undo
Exit Sub
End If

Case 300
If Not PriveCaution Then
MsgBox (TextMsg)
Me.Undo
Exit Sub
End If

Case 400
If Not PrivStop Then
MsgBox (TextMsg)
Me.Undo
Exit Sub
End If

End Select






'Select Case WarningLevel
'    Case 50
'        imgWarning.Picture = dbLocation & "dollar.emf"
'        imgWarning.Visible = True
'    Case 100
'        imgWarning.Picture = dbLocation & "papertray.emf"
'        imgWarning.Visible = True
'    Case 200
'        imgWarning.Picture = dbLocation & "house.emf"
'        imgWarning.Visible = True
'    Case 300
'        imgWarning.Picture = dbLocation & "caution.bmp"
'        imgWarning.Visible = True
'    Case 400
'        imgWarning.Picture = dbLocation & "stop.emf"
'        imgWarning.Visible = True
'    Case Else
'        imgWarning.Visible = False
'End Select
'



End Sub
