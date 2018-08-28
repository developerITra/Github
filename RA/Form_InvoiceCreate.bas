VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_InvoiceCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ActualAmount_AfterUpdate()
Call UpdateTotals
End Sub

Private Sub Approved_AfterUpdate()
ApprovedBy = GetStaffID()
End Sub

Private Sub Approved_DblClick(Cancel As Integer)
Approved = Date
ApprovedBy = GetStaffID()
End Sub

Private Sub cmdSelectAll_Click()
Dim R As Recordset

On Error GoTo Err_cmdSelectAll_Click

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

Set R = Me.RecordsetClone
R.MoveFirst
Do While Not R.EOF
    If Not R!Selected Then
        R.Edit
        R!Selected = True
        R.Update
    End If
    R.MoveNext
Loop
Set R = Nothing

Call UpdateTotals

Exit_cmdSelectAll_Click:
    Exit Sub

Err_cmdSelectAll_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectAll_Click
    
End Sub

Private Sub cmdInvertSelection_Click()
Dim R As Recordset

On Error GoTo Err_cmdInvertSelection_Click

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

Set R = Me.RecordsetClone
R.MoveFirst
Do While Not R.EOF
    R.Edit
    R!Selected = Not R!Selected
    R.Update
    R.MoveNext
Loop
Set R = Nothing

Call UpdateTotals

Exit_cmdInvertSelection_Click:
    Exit Sub

Err_cmdInvertSelection_Click:
    MsgBox Err.Description
    Resume Exit_cmdInvertSelection_Click
    
End Sub

Private Sub UpdateTotals()
Dim SelectCount As Integer

If IsNull(FileNumber) Then
    txtTotal = Null
    lblTotal.Caption = "0 lines"
Else
    txtTotal = DSum("ActualAmount", "InvoiceItems", "Selected AND FileNumber=" & FileNumber)
    SelectCount = DCount("*", "InvoiceItems", "Selected AND FileNumber=" & FileNumber)
    lblTotal.Caption = SelectCount & " line" & IIf(SelectCount = 1, "", "s") & " selected, total amount:"
End If
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
Call UpdateTotals
End Sub

Private Sub Form_AfterInsert()
Call UpdateTotals
End Sub

Private Sub Form_AfterUpdate()
Call UpdateTotals
End Sub

Private Sub Form_Current()
Call UpdateTotals
Select Case Forms![Case List]!SCRAID
    Case "AccPSAdvanced"
     Me.cbxInvType.DefaultValue = 711
    Case "AccLitig"
     Me.cbxInvType.DefaultValue = 710
    Case Else
     Me.cbxInvType.DefaultValue = 99
End Select

'Me.cbxInvType.DefaultValue = 99
End Sub

Private Sub Selected_AfterUpdate()
If IsNull(ActualAmount) Then Selected = False
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
Call UpdateTotals
End Sub

Private Sub cmdCreate_Click()
Dim rstInvItems As Recordset, rstInv As Recordset, InvID As String

On Error GoTo Err_cmdCreate_Click
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

If Val(Nz(txtTotal)) <= 0 Then
    MsgBox "You cannot create an invoice because no items are selected, or the total is not greater than zero.", vbCritical
    Exit Sub
End If

If IsNull(cbxInvType) Then
    MsgBox "Select the type of invoice to create", vbCritical
    Exit Sub
End If

If MsgBox("Really create an invoice with the selected items?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
If Utility.StaffID = 0 Then Call GetLoginName

Set rstInv = CurrentDb.OpenRecordset("Invoices", dbOpenDynaset, dbSeeChanges)
With rstInv
    .AddNew
    !FileNumber = FileNumber
    !InvoiceType = cbxInvType
    !InvoiceNumber = Trim$(str(FileNumber) & "-" & Format$(Date, "mmddyy"))
    !InvoiceAmount = Val(txtTotal)
    !AdditionalInvoiceNeeded = 0
    !DateSent = Now()
    !CreatedBy = Utility.StaffID
    !CreateMethod = 2
    .Update
    .Bookmark = .LastModified
    InvID = Trim$(str(FileNumber) & "-" & Format$(Date, "mmddyy"))
    .Close
End With

Set rstInvItems = Me.RecordsetClone
rstInvItems.MoveFirst
Do While Not rstInvItems.EOF
    If rstInvItems!Selected Then
        rstInvItems.Edit
        rstInvItems!InvoiceID = InvID
        rstInvItems!Selected = False
        rstInvItems.Update
    End If
    rstInvItems.MoveNext
Loop
Set rstInvItems = Nothing

MsgBox "Invoice has been created", vbInformation

Forms![Case List]!sfrmInvoices.Requery
'DoCmd.OpenReport "Invoice", acViewPreview, , "Invoices.InvoiceID=""" & InvID & """"

Select Case Forms![Case List]!SCRAID
    Case "AccPSAdvanced"
    Case "AccLitig"
    Case Else

DoCmd.OpenForm "UnbilledEvents"
Forms!unbilledevents!FileNumber = Forms![Case List]!FileNumber
Forms!unbilledevents!lstBillingReasons.RowSource = "SELECT BillingReasonsFCarchive.ID, BillingReasonsFC.Reason, BillingReasonsFCarchive.Date AS [Date] FROM BillingReasonsFCarchive INNER JOIN BillingReasonsFC ON BillingReasonsFCarchive.billingreasonID=BillingReasonsFC.ID WHERE (((BillingReasonsFCarchive.FileNumber)=" & Forms!unbilledevents!FileNumber & " AND Invoiced is null));"
End Select

DoCmd.Close acForm, "InvoiceCreate"
Exit_cmdCreate_Click:
    Exit Sub

Err_cmdCreate_Click:
    MsgBox Err.Description
    Resume Exit_cmdCreate_Click
    
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub
