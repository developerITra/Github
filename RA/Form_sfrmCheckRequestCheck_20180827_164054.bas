VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmCheckRequestCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CheckCut_AfterUpdate()

  If (CheckCut) Then  ' checks
    StatusID = DLookup("[ID]", "[CheckRequestStatus]", "[Status] = 'Completed'")
    CompleteUpdate
    
    'If (Not PreviousBilled) Then
      'AddInvoiceItem FileNumber, FeeType, Description, Amount, 0, False, True, False, True
    'End If
    
    
    AddInvoiceItem FileNumber, "FC-PSA", Description, Amount, 0, False, True, False, True
                    
    Dim rs1, rs2 As Recordset
                    
    Set rs1 = CurrentDb.OpenRecordset("PostSaleCostAdvancePKG", dbOpenDynaset, dbSeeChanges)
    Set rs2 = CurrentDb.OpenRecordset("Select max(InvoiceItemID)as [InvoiceitemNo] FROM InvoiceItems where filenumber=" & Me.FileNumber, dbOpenDynaset, dbSeeChanges)
                
            rs1.AddNew
                rs1!FileNumber = Me.FileNumber
                rs1!Description = Me.Description
                rs1!Amount = Me.Amount
                rs1!Timestamp = Now()
                rs1!Vendor = Me.PayableTo
                rs1!Fees = 0
                rs1!Username = GetFullName()
                rs1!InvoiceItemID = rs2!InvoiceitemNo
                rs1.Update
                                
            rs1.Close
            Set rs1 = Nothing
            rs2.Close
            Set rs2 = Nothing

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

