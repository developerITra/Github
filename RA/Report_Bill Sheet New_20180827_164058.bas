VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Bill Sheet New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Function GetLoanType()
Dim FileNum As Long, LoanType As Integer

FileNum = Forms![Case List]!FileNumber
LoanType = DLookup("LoanType", "FCdetails", "FileNumber=" & FileNum)

Select Case LoanType
Case 4
GetLoanType = "FNMA"
Case 5
GetLoanType = "FHLMC"
Case Else
GetLoanType = ""
End Select

End Function


Function getmilestonebilling()
If DLookup("milestonebilling", "clientlist", "clientid=" & Forms![Case List]!ClientID) = 0 Then
getmilestonebilling = ""
Else
getmilestonebilling = "Milestone Billing Client"
End If


End Function
