VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmMonitorReferRecd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAddMonitor_Click()
Dim rs As Recordset

If (IsNull(Me!Monitor_Refer_reced) Or (Me!Monitor_Refer_reced = "")) And (Forms![Case List]!CaseTypeID = 8) Or (IsNull(Me!Monitor_Refer_reced) Or (Me!Monitor_Refer_reced = "")) And (Forms![Case List]!CaseTypeID = 1) Then
    Me!Monitor_Refer_reced = Date
End If


'Set rs = CurrentDb.OpenRecordset("Select * FROM FCDetails where filenumber=" & Forms!Foreclosuredetails!FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)

   ' rs.Edit
    Forms!foreclosuredetails!Deposit = Null
    Forms!foreclosuredetails!SaleSet = Null
    Forms!foreclosuredetails!Disposition = Null
    Forms!foreclosuredetails!DispositionDate = Null
    'Forms!Foreclosuredetails!DispositionInitials = Null
    'rs.Update
    'rs.Close
    'Set rs = Nothing
    
    
If Not IsNull(Forms!foreclosuredetails!Sale) Then
    Forms!foreclosuredetails!Sale = Null
    Forms!foreclosuredetails!SaleTime = Null
    'Forms!Foreclosuredetails!.Requery
End If
Forms!foreclosuredetails.Sale.Locked = False
Forms!foreclosuredetails.SaleTime.Locked = False
'Forms!Foreclosuredetails!.Requery

If Not IsNull(Monitor_Refer_reced) Then
    AddStatus FileNumber, Monitor_Refer_reced, "Monitor Referral Recieved"
 
    Dim FeeAmt As Currency
        FeeAmt = Nz(DLookup("MonitorFee", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
        AddInvoiceItem FileNumber, "FC-MON", "Monitor sale fee", Format$(FeeAmt, "Currency"), 0, True, True, False, False
Else
    AddStatus FileNumber, Now(), "Moitor Referral Removed"
End If


End Sub

Private Sub Form_Current()
If IsNull(Forms!foreclosuredetails!DispositionDesc) Or Forms!foreclosuredetails!DispositionDesc = "" Then Me.Monitor_Refer_reced.Locked = False
End Sub

Private Sub Monitor_Refer_reced_AfterUpdate()
If Not IsNull(Monitor_Refer_reced) Then
    AddStatus FileNumber, Monitor_Refer_reced, "Monitor Referral Recieved"
 
    Dim FeeAmt As Currency
        FeeAmt = Nz(DLookup("MonitorFee", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
        AddInvoiceItem FileNumber, "FC-MON", "Monitor sale fee", Format$(FeeAmt, "Currency"), 0, True, True, False, False
Else
    AddStatus FileNumber, Now(), "Moitor Referral Removed"
End If
'Call cmdAddMonitor_Click
End Sub

Private Sub Monitor_Refer_reced_DblClick(Cancel As Integer)

If FileReadOnly Or EditDispute Then

   DoCmd.CancelEvent
Else
    If IsNull(Forms!foreclosuredetails!DispositionDesc) Then
    Me.Monitor_Refer_reced.Locked = False
    Monitor_Refer_reced = Now()
    Call Monitor_Refer_reced_AfterUpdate
    'Call cmdAddMonitor_Click
    End If
End If


End Sub
