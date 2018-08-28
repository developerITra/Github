VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetSkipTraceCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdOK_Click()
Dim strinfo As String

If IsNumeric(Me.txtDesc) = False Or Me.txtDesc = "" Or IsNull(Me.txtDesc) = True Or Me.txtDesc <= 0 Then
MsgBox ("Please enter dollar amount")
Me.txtDesc.SetFocus
Exit Sub
End If

'On Error GoTo Err_cmdOK_Click
DoCmd.SetWarnings False
AddInvoiceItem Forms![Case List].FileNumber, "FC-SKP", "Skip Trace actual costs.", Format$(Me.txtDesc, "Currency"), 0, False, True, False, False
DoCmd.SetWarnings True

MsgBox "Skip Trace Cost Entered"
DoCmd.Close
End Sub



