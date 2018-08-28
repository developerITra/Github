VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CollectionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
DoCmd.Close
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Err_cmdPrint_Click

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm "CollectionPrint", , , "[CaseList].[FileNumber]=" & Me![FileNumber]

Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrint_Click
End Sub

Private Sub cmdSelectFile_Click()
DoCmd.Close
DoCmd.OpenForm "Select File"
End Sub

Private Sub ComplaintFiled_AfterUpdate()
AddStatus FileNumber, ComplaintFiled, "Complaint Filed"
End Sub

Private Sub ComplaintFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ComplaintFiled)
End Sub

Private Sub ComplaintFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ComplaintFiled = Date
AddStatus FileNumber, ComplaintFiled, "Complaint Filed"
End If
End Sub

Private Sub FairDebtLetter_AfterUpdate()
AddStatus FileNumber, FairDebtLetter, "Fair Debt Letter sent"
End Sub

Private Sub FairDebtLetter_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(FairDebtLetter)
End Sub

Private Sub FairDebtLetter_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
FairDebtLetter = Date
AddStatus FileNumber, FairDebtLetter, "Fair Debt Letter sent"
End If

End Sub

Private Sub Form_Current()

If FileReadOnly Or EditDispute Then
    Me.AllowEdits = False
    cmdPrint.Enabled = False
    sfrmNames.Form.AllowEdits = False
    sfrmNames.Form.AllowAdditions = False
    sfrmNames.Form.AllowDeletions = False
    sfrmNames!cmdCopy.Enabled = False
    sfrmNames!cmdTenant.Enabled = False
    sfrmNames!cmdDelete.Enabled = False
    sfrmNames!cmdNoNotice.Enabled = False
    sfrmSummons.Form.AllowEdits = False
    sfrmSummons.Form.AllowAdditions = False
    sfrmSummons.Form.AllowDeletions = False
    sfrmCollectionActions.Form.AllowEdits = False
    sfrmCollectionActions.Form.AllowAdditions = False
    sfrmCollectionActions.Form.AllowDeletions = False
    sfrmStatus.Form.AllowEdits = False
    sfrmStatus.Form.AllowAdditions = False
    sfrmStatus.Form.AllowDeletions = False
    Detail.BackColor = ReadOnlyColor
Else
    Me.AllowEdits = True
    cmdPrint.Enabled = True
    If Not CheckNameEdit() Then
    sfrmNames.Form.AllowEdits = False
    sfrmNames.Form.AllowAdditions = False
    sfrmNames.Form.AllowDeletions = False
    sfrmNames!cmdCopy.Enabled = False
    sfrmNames!cmdTenant.Enabled = False
    sfrmNames!cmdDelete.Enabled = False
    sfrmNames!cmdNoNotice.Enabled = False
    Else
    sfrmNames.Form.AllowEdits = True
    sfrmNames.Form.AllowAdditions = True
    sfrmNames.Form.AllowDeletions = True
    sfrmNames!cmdCopy.Enabled = True
    sfrmNames!cmdTenant.Enabled = True
    sfrmNames!cmdDelete.Enabled = True
    sfrmNames!cmdNoNotice.Enabled = True
    End If
    sfrmSummons.Form.AllowEdits = True
    sfrmSummons.Form.AllowAdditions = True
    sfrmSummons.Form.AllowDeletions = True
    sfrmCollectionActions.Form.AllowEdits = True
    sfrmCollectionActions.Form.AllowAdditions = True
    sfrmCollectionActions.Form.AllowDeletions = True
    sfrmStatus.Form.AllowEdits = True
    sfrmStatus.Form.AllowAdditions = True
    sfrmStatus.Form.AllowDeletions = True
    Detail.BackColor = -2147483633
End If

Me.Caption = "Collection File " & Me![FileNumber] & " " & [PrimaryDefName]
End Sub

Private Sub Form_Open(Cancel As Integer)
If FileReadOnly Or EditDispute Then

    Dim ctl As Control
    Dim lngI As Long
    Dim bSkip As Boolean

    For Each ctl In Form.Controls
    Select Case ctl.ControlType
    Case acTextBox, acListBox, acOptionGroup, acCheckBox, acSubform, acOptionButton
         bSkip = False
            If ctl.Name = "lstDocs" Then bSkip = True
            If Not bSkip Then ctl.Locked = True
            
            
    Case acCommandButton
        bSkip = False
            If ctl.Name = "cbxDetails" Then bSkip = True
            If ctl.Name = "cmdDetails" Then bSkip = True
            If ctl.Name = "cmdGoToFile" Then bSkip = True
            If ctl.Name = "cmdClose" Then bSkip = True
            If ctl.Name = "cmdSelectFile" Then bSkip = True
            If Not bSkip Then ctl.Enabled = False
         
    Case acComboBox
        bSkip = False
        If ctl.Name = "cbxDetails" Then bSkip = True
        If Not bSkip Then ctl.Locked = True
    
    
    End Select
    Next
End If
End Sub

Private Sub InterestRate_AfterUpdate()
If InterestRate > 1 Then InterestRate = InterestRate / 100#
End Sub

Private Sub tabCollection_Change()
If tabCollection.Value = 5 Then sfrmStatus.Requery  ' requery when switching to status tab
End Sub
