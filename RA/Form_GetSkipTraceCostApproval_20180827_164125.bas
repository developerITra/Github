VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetSkipTraceCostApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdOK_Click()
Dim strinfo As String

'On Error GoTo Err_cmdOK_Click

AddInvoiceItem Forms![Case List].FileNumber, "FC-SKP", "Skip Trace estimated costs.", Format$(Me.SkipTraceCost, "Currency"), 0, False, False, False, False

DoCmd.SetWarnings False

strinfo = txtDesc
strinfo = Format$(Me.SkipTraceCost, "Currency") & " Skip Trace estimated costs. " & strinfo

strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms![Case List].FileNumber & ",Now,GetFullName(),'" & strinfo & "',2 )"
DoCmd.RunSQL strSQLJournal
strinfo = ""

DoCmd.SetWarnings True
Forms!Journal.Requery

MsgBox "Cost accepted"
'DoCmd.Close acForm, "GetSkipTraceCostApproval"
DoCmd.Close
End Sub



