VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Affidavit Collateral fileold"
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


On Error GoTo Err_cmdOK_Click


Call DoReport("Affidavit Collateral File", Me.OpenArgs)

'DoCmd.Close acForm, "Print Affidavit Collateral file"

'ShowForm ("Affidavit Collateral File")

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Close()
DoCmd.Maximize
End Sub

Private Sub Form_Open(Cancel As Integer)
OtherDocText = "Other"

End Sub

Private Sub Other_AfterUpdate()

If Other Then
OtherDocText = InputBox("Enter the Name of the document.", "Name of Other document")
Else
OtherDocText = "Other"
End If
End Sub



'Private Sub ShowForm(formName As String, Optional args As String = "")
'On Error GoTo Err_ShowForm
'  ' DoCmd.OptenForm formName, acNormal, , , , acWindowNormal, args
'    DoCmd.OpenReport formName, acViewNormal
'   Dim frm As Form
'   Set frm = Forms.Item(formName)
'   If Not frm Is Nothing Then
'      frm.Modal = True
'      While (SysCmd(acSysCmdGetObjectState, acForm, formName) And acObjStateOpen) = acObjStateOpen
'         DoEvents
'      Wend
'   End If
'Err_ShowForm:
'Resume

'End Sub


