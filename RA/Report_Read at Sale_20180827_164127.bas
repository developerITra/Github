VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Read at Sale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function Due() As String
Select Case txtState
    Case "MD", "DC"
        Due = "ten (10) days after ratification of sale"
    Case "VA"
        Due = "fifteen (15) days of sale"
End Select
End Function

Private Function tax() As String
Select Case txtState
    Case "MD", "DC"
        tax = "all transfer taxes"
    Case "VA"
        tax = "Grantor's tax"
End Select
End Function
