VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Bid Instructions VA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim BidType As String

Private Function GetBidType() As String
Dim UserInput As String

If BidType <> "" Then
    GetBidType = BidType
Else
    UserInput = InputBox$("Enter 's' for Specified" & vbNewLine & "or 't' for Total Debt" & vbNewLine & "or enter other type of bid:", "Bid Type")
    If UserInput = "" Then DoCmd.Close acReport, Me.Name
    Select Case UCase$(UserInput)
        Case "S"
            BidType = "Specified"
        Case "T", "D"
            BidType = "Total Debt"
        Case Else
            BidType = UserInput
    End Select
    GetBidType = BidType
End If
End Function

