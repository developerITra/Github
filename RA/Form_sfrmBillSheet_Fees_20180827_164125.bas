VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmBillSheet_Fees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Check225_AfterUpdate()
'If Check225 Then Me.Text254 = 1
    'Me.RecordsetClone.FindFirst "[InvoiceItemID] = " & Me.FileNumber
   ' Me.Bookmark = Me.RecordsetClone.Bookmark
        
   ' Me.RemovedBy.Enabled = False
    'Me.RemovedBy = GetStaffInitials(GetStaffID())
    'Me.RemovedOn = Date
    
    'Me.RemovedBy.Enabled = False
    'Me.RemovedOn.Enabled = False
    
    'Me.Remove!Enabled = False
    'Me.RemovedBy.Enabled = False
    'Me.RemovedOn.Enabled = False
    
    'Me.Description.Enabled = False
    'Me.ActFees.Enabled = False
    'Me.VendorName.Enabled = False
    
    
    Dim rs As Recordset

Set rs = CurrentDb.OpenRecordset("select * from InvoiceItems where Adjust = true and FileNumber = " & Me.FileNumber, dbOpenDynaset, dbSeeChanges)
    Do Until rs.EOF
        rs.Edit
        If rs!Adjust = True And (IsNull(rs!Adjustedby) Or rs!Adjustedby = "") Then
            rs!Adjustedby = GetStaffInitials(GetStaffID())
            rs!AdjustedOn = Date
            rs.Update
        'ElseIf (rs!Adjust = False And Not IsNull(rs!Adjustedby)) Then
            
        ElseIf (rs!Adjust = True And Not IsNull(rs!Adjustedby)) Or rs!Adjust = False Then
            Me.Undo
        
        End If
        
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!Timestamp.ForeColor = vbRed
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!Description.ForeColor = vbRed
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!ActFees.ForeColor = vbRed
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!VendorName.ForeColor = vbRed
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!Adjustedby.ForeColor = vbRed
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!AdjustedOn.ForeColor = vbRed
        
        rs.MoveNext
        Loop
        
     rs.Close
    Set rs = Nothing

Me.Requery
End Sub

Private Sub ckRemove_AfterUpdate()
 
 If Remove = True Then
    If (RemovedBy = "" Or IsNull(RemovedBy) = True) And IsNull(RemovedOn) = True Then
        RemovedBy = GetStaffInitials(GetStaffID())
        RemovedOn = Date
    End If
Else
    RemovedBy = ""
    RemovedOn = Null
End If
    Me.Requery
 

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


Private Sub cmdRefrash_Click()

Dim rs As Recordset

'Set rs = CurrentDb.OpenRecordset("select * from InvoiceItems where Adjust = true and FileNumber = " & Me.FileNumber, dbOpenDynaset, dbSeeChanges)
'Set rs = CurrentDb.OpenRecordset("select * from InvoiceItems where (Adjust = true or isnull(adjustedby) = false) and FileNumber = " & Me.FileNumber, dbOpenDynaset, dbSeeChanges)
Set rs = CurrentDb.OpenRecordset("select * from InvoiceItems where Fee = True and FileNumber = " & Me.FileNumber, dbOpenDynaset, dbSeeChanges)

     Do Until rs.EOF
        rs.Edit
        
        Dim Fee As String
        Fee = rs!ActualAmount
         Fee = rs!Adjustedby
        If rs!Adjust = -1 And IsNull(rs!Adjustedby) = True Then
            rs!Adjustedby = GetStaffInitials(GetStaffID())
            rs!AdjustedOn = Date
            'rs.Update
               
        ElseIf rs!Adjust = -1 And IsNull(rs!Adjustedby) = False Then
            Me.Undo
        
        ElseIf rs!Adjust = 0 Then
            rs!Adjustedby = ""
            rs!AdjustedOn = Null
        
        End If
             rs.Update
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!Timestamp.ForeColor = vbRed
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!Description.ForeColor = vbRed
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!ActFees.ForeColor = vbRed
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!VendorName.ForeColor = vbRed
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!Adjustedby.ForeColor = vbRed
        'Forms![Bill sheet_Last]![sfrmBillSheet_Fees]!AdjustedOn.ForeColor = vbRed
        
        rs.MoveNext
        Loop
        
     rs.Close
    Set rs = Nothing
        
Me.Requery

End Sub



