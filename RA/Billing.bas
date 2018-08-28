Attribute VB_Name = "Billing"
Option Compare Database
Option Explicit

Public Function FetchPropertyTaxBasis() As Currency

  If (Forms![Case List]!JurisdictionID = 58 Or Forms![Case List]!JurisdictionID = 36) Then   ' Newport News or VA Beach - always use sale price
    FetchPropertyTaxBasis = Nz(Forms![foreclosuredetails]!SalePrice)
    Exit Function
  End If
  
  If (Nz(Forms![foreclosuredetails]!AssessedValue) > (Nz(Forms![foreclosuredetails]!SalePrice))) Then
    FetchPropertyTaxBasis = Nz(Forms![foreclosuredetails]!AssessedValue)
  Else
    FetchPropertyTaxBasis = Nz(Forms![foreclosuredetails]!SalePrice)
  End If
    
  
End Function

Public Sub test()
  MsgBox "calculate auditor fee: " & Format(CalculateVAAuditorFee(300000.25), "currency")
  
  
End Sub

Public Function CalculateVAGrantorTax(PropertyValue As Currency) As Currency

Dim X As Currency
Dim n As Integer
Dim roundValue As Currency


X = PropertyValue
n = 500

roundValue = Round(X / n) * n
CalculateVAGrantorTax = (roundValue / 1000) * 1

End Function

Public Function CalculateVAStateTransferTax(PropertyValue As Currency) As Currency
  CalculateVAStateTransferTax = (PropertyValue / 1000) * 2.5

End Function


Public Function CalculateVAAuditorFee(PropertyValue As Currency) As Currency

If (PropertyValue >= 0 And PropertyValue <= 100000) Then
  CalculateVAAuditorFee = 266#
ElseIf (PropertyValue > 100000 And PropertyValue <= 300000) Then
  CalculateVAAuditorFee = 316#
ElseIf (PropertyValue > 300000 And PropertyValue <= 450000) Then
  CalculateVAAuditorFee = 466#
ElseIf (PropertyValue > 450000 And PropertyValue <= 600000) Then
  CalculateVAAuditorFee = 616#
ElseIf (PropertyValue > 600000 And PropertyValue <= 750000) Then
  CalculateVAAuditorFee = 766#
ElseIf (PropertyValue > 750000 And PropertyValue <= 900000) Then
  CalculateVAAuditorFee = 916#
Else
  CalculateVAAuditorFee = 1016#
End If
  
End Function
