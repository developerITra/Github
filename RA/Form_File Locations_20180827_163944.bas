VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_File Locations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cbxFileLocation_AfterUpdate()
If Not IsNull(cbxFileLocation) Then optWhich = 2
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

Private Sub Form_Open(Cancel As Integer)
lblMine.Caption = "Mine (" & Forms!Main!txtLoginName & ")"
End Sub

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click

Select Case optWhich
    Case 1      ' mine
        DoCmd.OpenReport "File Locations", optPrintTo, , "FileLocation=""" & Forms!Main!txtLoginName & """"
    
    Case 2      ' specified
        If IsNull(cbxFileLocation) Then
            MsgBox "Select a file location", vbCritical
            Exit Sub
        End If
        DoCmd.OpenReport "File Locations", optPrintTo, , "FileLocation=""" & cbxFileLocation.Value & """"
    
    Case 3      ' all
        DoCmd.OpenReport "File Locations", optPrintTo
End Select

DoCmd.Close acForm, Me.Name

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub
