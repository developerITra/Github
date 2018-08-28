VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Add Document Request"
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

If IsNull(Me.cboDocType) Then
    MsgBox "Enter a document type.", vbExclamation
    Exit Sub
End If

If IsNull(Me.cboDocLocation) Then
    MsgBox "Enter a document location.", vbExclamation
    Exit Sub
End If
  
Dim FileNumber As Long

FileNumber = Forms![Case List]!FileNumber
Call AddDocumentRequest(FileNumber, cboDocType, cboDocLocation)
   

DoCmd.Close


Forms![Case List]!sfrmFileDocRequest.Requery

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
    
End Sub


