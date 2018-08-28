VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_45 Day Cover Letter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Function FetchLoanModAgent()

  FetchLoanModAgent = DLookup("[LoanModAgent]", "[ClientList]", "ClientID = " & Forms![Case List]![ClientID])

End Function

Private Function FetchLoanModPhone()

  FetchLoanModPhone = DLookup("[LoanModPhone]", "[ClientList]", "ClientID = " & Forms![Case List]![ClientID])

End Function

