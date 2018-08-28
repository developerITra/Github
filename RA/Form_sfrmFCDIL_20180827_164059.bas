VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmFCDIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub DILcancelled_AfterUpdate()
If BHproject Then
    If Not IsNull(DILCancelled) Then
    AddStatus FileNumber, DILCancelled, "DIL Cancelled"
    End If
Else

If Not IsNull(DILCancelled) Then
AddStatus FileNumber, DILCancelled, "DIL Cancelled"
Forms![Case List]!BillCase = True
Forms![Case List]!BillCaseUpdateUser = GetStaffID()
Forms![Case List]!BillCaseUpdateDate = Date
Forms![Case List]![BillCaseUpdateReasonID] = 30
Forms![Case List]!lblBilling.Visible = True
'lstBillingReasons
End If

Dim rstBillReasons As Recordset
Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstBillReasons
.AddNew
!FileNumber = FileNumber
!billingreasonid = 30
!UserID = GetStaffID
!Date = Date
.Update
End With

End If

End Sub

Private Sub DILcancelled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DILCancelled = Date
    Call DILcancelled_AfterUpdate
End If
End Sub

Private Sub DILReferralReceived_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DILReferralReceived = Date
    Call DILReferralReceived_AfterUpdate
End If
End Sub
Private Sub DILReceivedBorrowers_AfterUpdate()
If BHproject Then
AddStatus FileNumber, DILReceivedBorrowers, "DIL Received from Borrowers"


Else
        
        AddStatus FileNumber, DILReceivedBorrowers, "DIL Received from Borrowers"
        
        ' Moved from DILSentBorrowers_AfterUpdate 2012.01.19 DaveW
        AddInvoiceItem FileNumber, "DIL", "DIL Title Search", GetFeeAmount("Title Search"), 0, False, True, False, True
        AddInvoiceItem FileNumber, "DIL", "DIL Title Update", 50, 0, False, True, False, True
        
        Select Case Forms!foreclosuredetails!State
          Case "MD"
              AddInvoiceItem FileNumber, "DIL", "DIL Deed Recording Cost", 40, 0, False, True, False, True
          Case "VA"
              AddInvoiceItem FileNumber, "DIL", "DIL Deed Recording Cost", 21, 0, False, True, False, True
          Case "DC"
              AddInvoiceItem FileNumber, "DIL", "DIL Deed Recording Cost", 26.5, 0, False, True, False, True
        End Select
        
        AddInvoiceItem FileNumber, "DIL", "DIL Transfer Tax", GetFeeAmount("Transfer Tax"), 0, False, True, False, True
        AddInvoiceItem FileNumber, "DIL", "DIL Recordation Fee", GetFeeAmount("Recordation Fee"), 0, False, True, False, True
End If

End Sub
 
Private Sub DILReceivedBorrowers_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DILReceivedBorrowers)
End Sub

Private Sub DILReceivedBorrowers_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DILReceivedBorrowers = Date
    Call DILReceivedBorrowers_AfterUpdate
End If
End Sub

Private Sub DILReceivedClient_AfterUpdate()
AddStatus FileNumber, DILReceivedClient, "DIL Received from Client"
End Sub

Private Sub DILReceivedClient_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DILReceivedClient)
End Sub

Private Sub DILReceivedClient_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DILReceivedClient = Date
    Call DILReceivedClient_AfterUpdate
End If
End Sub

Private Sub DILRecorded_AfterUpdate()
If BHproject Then
If Not IsNull(DILRecorded) Then
AddStatus FileNumber, DILRecorded, "DIL Recorded"
End If
Else

    If Not IsNull(DILRecorded) Then
    AddInvoiceItem FileNumber, "FC-DIL", "Return Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, True, False, False
    AddStatus FileNumber, DILRecorded, "DIL Recorded"
    End If
    AddStatus FileNumber, DILRecorded, "DIL Recorded"
End If

End Sub

Private Sub Form_AfterUpdate()
Me.Parent.Requery
End Sub

Private Sub TitleClearForDIL_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(TitleClearForDIL)
End Sub
Private Sub TitleClearForDIL_AfterUpdate()
AddStatus FileNumber, TitleClearForDIL, "Title Clear for DIL"
End Sub
Private Sub TitleClearForDIL_DblClick(Cancel As Integer)
TitleClearForDIL = Date
Call TitleClearForDIL_AfterUpdate
End Sub
Private Sub DILRecorded_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DILRecorded)
End Sub

Private Sub DILRecorded_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DILRecorded = Date
    Call DILRecorded_AfterUpdate
End If
End Sub

Private Sub DILReferralReceived_AfterUpdate()
'If Not IsNull(DILReferralReceived) Then
'AddStatus FileNumber, DILReferralReceived, "DIL Referral Received"
'AddInvoiceItem FileNumber, "DIL", "DIL referral received", DLookup("DeedinLieu", "clientlist", "clientid=" & Forms![Case list]!ClientID), 0, True, True, False, False
If Not BHproject Then


'End If
        If Not IsNull(DILReferralReceived) Then
        
            Dim amt As Double
        
            If Not IsNull(DLookup("DeedinLieu", "clientlist", "clientid=" & Forms![Case List]!ClientID)) Then
                amt = DLookup("DeedinLieu", "clientlist", "clientid=" & Forms![Case List]!ClientID)
            Else
                amt = 0
            End If
            AddStatus FileNumber, DILReferralReceived, "DIL Referral Received"
        
            If amt = 0 Then
                MsgBox ("Please see Manager to put appropriate DIL Referral Fee amount")
            Else
                AddInvoiceItem FileNumber, "DIL", "Attorney Fee - DIL", DLookup("DeedinLieu", "clientlist", "clientid=" & Forms![Case List]!ClientID), 0, True, True, False, False
            End If
            
        End If

End If

End Sub

Private Sub DILReferralReceived_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DILReferralReceived)
End Sub



Private Sub DILSentBorrowers_AfterUpdate()
If BHproject Then
    If Not IsNull(DILSentBorrowers) Then
    AddStatus FileNumber, DILSentBorrowers, "DIL Sent to Borrowers"
    Exit Sub
    End If
Else


    If Not IsNull(DILSentBorrowers) Then
    AddStatus FileNumber, DILSentBorrowers, "DIL Sent to Borrowers"
    DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter overnight costs to the borrower AND return (2X) for the DIL|FC-DIL|DIL Overnight Costs"
    'AddInvoiceItem FileNumber, "DIL", "DIL Title Review Fee", 100, 0, True, True, False, False
    End If

End If

End Sub

Private Sub DILSentBorrowers_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DILSentBorrowers)
End Sub

Private Sub DILSentBorrowers_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DILSentBorrowers = Date
    Call DILSentBorrowers_AfterUpdate
End If

End Sub

Private Sub DILSentClient_AfterUpdate()
If BHproject Then
AddStatus FileNumber, DILSentClient, "DIL Sent to Client"
Else
AddStatus FileNumber, DILSentClient, "DIL Sent to Client"

If Not IsNull(DILSentClient) Then
Forms![Case List]!BillCase = True
Forms![Case List]!BillCaseUpdateUser = GetStaffID()
Forms![Case List]!BillCaseUpdateDate = Date
Forms![Case List]![BillCaseUpdateReasonID] = 30
Forms![Case List]!lblBilling.Visible = True
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter overnight costs for sending the DIL to client|FC-DIL|DIL Overnight Costs"
AddStatus FileNumber, DILSentClient, "DIL Sent to Client"
End If
End If

End Sub

Private Sub DILSentClient_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DILSentClient)
End Sub

Private Sub DILSentClient_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    DILSentClient = Date
    Call DILSentClient_AfterUpdate
End If
End Sub

Private Sub DILSentRecord_AfterUpdate()
If BHproject Then
    If Not IsNull(DILSentRecord) Then
    AddStatus FileNumber, DILSentRecord, "DIL Sent to Record"
    End If
Else


If Not IsNull(DILSentRecord) Then
DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter overnight costs for sending the DIL to Land Records|FC-DIL|DIL Overnight Costs"
AddStatus FileNumber, DILSentRecord, "DIL Sent to Record"
End If

End If

End Sub

Private Sub DILSentRecord_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(DILSentRecord)
End Sub

