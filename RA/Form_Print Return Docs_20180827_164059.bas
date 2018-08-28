VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Return Docs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdCancel_Click()
On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
End Sub

Private Sub cmdClear_Click()

On Error GoTo Err_cmdClear_Click

DoCmd.RunSQL "DELETE * FROM ReturnedDocs WHERE FileNumber=" & Forms!foreclosuredetails!FileNumber
sfrmReturnDocs.Requery

Exit_cmdClear_Click:
    Exit Sub

Err_cmdClear_Click:
    MsgBox Err.Description
    Resume Exit_cmdClear_Click
    
End Sub

Private Sub cmdOK_Click()
    
    Call DoReport("return doc cover ltr", Me.OpenArgs)
       
End Sub
