VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_REOdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ClientPaid_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ClientPaid)
End Sub

Private Sub ClientPaid_DblClick(Cancel As Integer)
ClientPaid = Date
End Sub

Private Sub Closing_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Closing)
End Sub

Private Sub CommitmentSent_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(CommitmentSent)
End Sub

Private Sub CommitmentSent_DblClick(Cancel As Integer)
CommitmentSent = Date
End Sub

Private Sub Complete_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Complete)
End Sub

Private Sub Complete_DblClick(Cancel As Integer)
Complete = Date
End Sub

Private Sub Form_Close()
DoCmd.Restore
End Sub

Private Sub Form_Current()


Me.Caption = "REO File " & Me![FileNumber] & " " & [PrimaryDefName]

If FileReadOnly Or EditDispute Then
    Me.AllowEdits = False
    cmdNew.Enabled = False
    cmdPrint.Enabled = False
    sfrmPropAddr.Form.AllowEdits = False
    sfrmComments.Form.AllowEdits = False
    sfrmStatus.Form.AllowEdits = False
    sfrmStatus.Form.AllowAdditions = False
    sfrmStatus.Form.AllowDeletions = False
    Detail.BackColor = ReadOnlyColor
Else
    Me.AllowEdits = True
    cmdNew.Enabled = True
    cmdPrint.Enabled = True
    sfrmPropAddr.Form.AllowEdits = True
    sfrmComments.Form.AllowEdits = True
    sfrmStatus.Form.AllowEdits = True
    sfrmStatus.Form.AllowAdditions = True
    sfrmStatus.Form.AllowDeletions = True
    Detail.BackColor = -2147483633
End If

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

Private Sub cmdPrint_Click()
On Error GoTo Err_cmdPrint_Click

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
'DoCmd.OpenForm "BankruptcyPrint", , , "BankruptcyID=" & Me!BankruptcyID

Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrint_Click
    
End Sub

Private Sub cmdNew_Click()

On Error GoTo Err_cmdNew_Click
If MsgBox("Are you sure you want to add another REO?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub

Me.AllowAdditions = True
DoCmd.GoToRecord , , acNewRec
Me.AllowAdditions = False

Exit_cmdNew_Click:
    Exit Sub

Err_cmdNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdNew_Click
    
End Sub

Private Sub cmdSelectFile_Click()

On Error GoTo Err_cmdSelectFile_Click

DoCmd.Close
DoCmd.OpenForm "Select File"

Exit_cmdSelectFile_Click:
    Exit Sub

Err_cmdSelectFile_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectFile_Click
    
End Sub

Private Sub OrderFC_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(OrderFC)
End Sub

Private Sub OrderFC_DblClick(Cancel As Integer)
OrderFC = Date
End Sub

Private Sub OrderTitle_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(OrderTitle)
End Sub

Private Sub OrderTitle_DblClick(Cancel As Integer)
OrderTitle = Date
End Sub

Private Sub ReceivedContract_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ReceivedContract)
End Sub

Private Sub ReceivedContract_Click()
ReceivedContract = Date
End Sub

Private Sub ReceivedFC_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ReceivedFC)
End Sub

Private Sub ReceivedFC_DblClick(Cancel As Integer)
ReceivedFC = Date
End Sub

Private Sub Referred_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(Referred)
End Sub

Private Sub Referred_DblClick(Cancel As Integer)
Referred = Date
End Sub

Private Sub tabREO_Change()
If tabREO.Value = 2 Then sfrmStatus.Requery
End Sub

Private Sub TitleClear_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleClear)
End Sub

Private Sub TitleClear_DblClick(Cancel As Integer)
TitleClear = Date
End Sub

Private Sub TitleReceived_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleReceived)
End Sub

Private Sub TitleReceived_DblClick(Cancel As Integer)
TitleReceived = Date
End Sub
