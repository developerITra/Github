VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmCheckRequestPrecut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CheckCut_AfterUpdate()

  If (CheckCut) Then  ' checks
    StatusID = DLookup("[ID]", "[CheckRequestStatus]", "[Status] = 'Completed'")
    CompleteUpdate

    AddInvoiceItem FileNumber, FeeType, Description, Amount, 0, False, True, False, True

  End If

End Sub

Private Sub CompleteUpdate()
 If (StatusID.Column(1) <> "Pending") Then
    CompletedDate = Date
    CompletedBy = GetStaffID()
 End If
  
End Sub

Private Sub StatusID_AfterUpdate()
  CompleteUpdate
End Sub

