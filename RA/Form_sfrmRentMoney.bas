VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmRentMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterUpdate()
Dim FeeAmount As Currency, recCnt As Integer

  If ([RentPayment] = -1) Then  ' potential invoice item
  
    If (Not IsNull([Amount]) And Not IsNull([TransDate])) Then
      recCnt = DCount("[FileNumber]", "InvoiceItems", "FileNumber=" & Forms![Case List]![CaseList.FileNumber] & " and Process='EV-RC' and datediff('m',[TimeStamp],#" & [TransDate] & "#) = 0")
      If (recCnt = 0) Then    ' no invoice item for rent collection for month/year
        FeeAmount = Nz(DLookup("FeeRentCollection", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
  
        If FeeAmount > 0 Then
          AddInvoiceItem FileNumber, "EV-RC", "Rent Collection Fee " & MonthName(Month([TransDate])) & ", " & Year([TransDate]), FeeAmount, 0, True, True, False, False
        Else
          AddInvoiceItem FileNumber, "EV-RC", "Rent Collection Fee " & MonthName(Month([TransDate])) & ", " & Year([TransDate]), GetFeeAmount("Rent Collection Fee"), 0, True, True, False, False
        End If
      End If
    End If
  End If
  
End Sub

Private Sub Form_Current()
If Me.NewRecord Then
    Me.AllowEdits = True
Else
    ' allow edits if made the same day
    Me.AllowEdits = DateSerial(Year(EntryDate), Month(EntryDate), Day(EntryDate)) = Date
End If
End Sub
