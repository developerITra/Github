VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmFCtitle_Orig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub TitleAssignNeededdate_AfterUpdate()
If (Not IsNull(TitleAssignNeededDate)) Then
 AddStatus FileNumber, TitleAssignNeededDate, "Assignment Needed"
End If
End Sub

Private Sub TitleAssignNeededdate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    TitleAssignNeededDate = Now()
    Call TitleAssignNeededdate_AfterUpdate
End If

End Sub

Private Sub TitleAssignReceivedDate_AfterUpdate()
If (Not IsNull(TitleAssignReceivedDate)) Then
Dim qty As Integer
'Assignment fee
qty = InputBox("Enter the number of assigments")
If qty = 0 Then
MsgBox "Please enter a number greater than zero"
Exit Sub
End If
AddInvoiceItem Forms!foreclosuredetails!FileNumber, "FC-Title", "Attorney Fee for Assignment Prep", DLookup("assignmentprepfee", "clientlist", "clientid=" & Forms![Case List]!ClientID), 0, True, True, False, False
 AddStatus FileNumber, TitleAssignReceivedDate, "Assignment Received from Client"
End If

End Sub

Private Sub TitleAssignReceivedDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    TitleAssignReceivedDate = Now()
    Call TitleAssignReceivedDate_AfterUpdate
End If
End Sub

Private Sub TitleAssignRecordedDate_AfterUpdate()
If (Not IsNull(TitleAssignRecordedDate)) Then
  AddStatus FileNumber, TitleAssignRecordedDate, "Assignment Sent to Record"

End If

If (IsNull(TitleAssignRecordedDate)) Then

    RecordedLiber.Enabled = False
    RecordedLiber.Locked = True
    RecordedLiber.BackColor = 16777215
    RecordedLiber.BackStyle = 0
    RecordedFolio.Enabled = False
    RecordedFolio.Locked = True
    RecordedFolio.BackColor = 16777215
    RecordedFolio.BackStyle = 0

Else

    RecordedLiber.Enabled = True
    RecordedLiber.Locked = False
    RecordedLiber.BackColor = -2147483643
    RecordedLiber.BackStyle = 1
    RecordedFolio.Enabled = True
    RecordedFolio.Locked = False
    RecordedFolio.BackColor = -2147483643
    RecordedFolio.BackStyle = 1

End If
 
End Sub

Private Sub TitleAssignRecordedDate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    TitleAssignRecordedDate = Now()
    Call TitleAssignRecordedDate_AfterUpdate
End If
End Sub

Private Sub TitleAssignSentdate_AfterUpdate()
If (Not IsNull(TitleAssignSentdate)) Then
 AddStatus FileNumber, TitleAssignSentdate, "Assignment Requested From Client"
End If
End Sub

Private Sub TitleAssignSentdate_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    TitleAssignSentdate = Now()
    Call TitleAssignSentdate_AfterUpdate
End If
End Sub

Private Sub Form_Current()

       If Not IsNull(Forms!foreclosuredetails!OriginalPBal) Then
            If IsNull(TitleReviewNameOf) Then TitleReviewNameOf = GetNames(FileNumber, 2, "Owner=True") & _
                    " Owner" & IIf(CountNames(FileNumber, "Owner=True") > 1, "s", "") & " of a " & _
                    IIf(Forms!foreclosuredetails!optLeasehold = 1, "leasehold property with an annual ground rent of " & _
                    Format$(Forms!foreclosuredetails!GroundRentAmount, "Currency") & " payable " & Forms!foreclosuredetails!GroundRentPayable, _
                    "fee simple property") & " by Deed dated"
            If IsNull(TitleReviewLiens) Then TitleReviewLiens = "1.  " & DOTWord(Forms!foreclosuredetails!DOT) & " dated " & _
                Format$(Forms!foreclosuredetails!DOTdate, "mmmm d, yyyy") & " securing " & Forms!foreclosuredetails!OriginalBeneficiary & " in the original amount of " & _
                Format$(Forms!foreclosuredetails!OriginalPBal, "Currency") & " and recorded on " & Format$(Forms!foreclosuredetails!DOTrecorded, "mmmm d, yyyy") & _
                " " & LiberFolio(Forms!foreclosuredetails!Liber, Forms!foreclosuredetails!Folio, Forms!foreclosuredetails!State)
            If IsNull(TitleReviewStatus) Then If Not IsNull(Forms!foreclosuredetails!LienPosition) Then TitleReviewStatus = Forms!foreclosuredetails!Investor & " is foreclosing in " & Ordinal(Forms!foreclosuredetails!LienPosition) & " position."
        End If

End Sub

Private Sub TitleClearFC_AfterUpdate()
If (Not IsNull(TitleClearFC)) Then
 AddStatus FileNumber, TitleClearFC, "Title Clear FC"
End If
End Sub

Private Sub TitleClearFC_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    TitleClearFC = Now()
    Call TitleClearFC_AfterUpdate
End If
End Sub
