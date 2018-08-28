VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_45 Day Cover Letter Wiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Function FetchLoanModAgent()

  FetchLoanModAgent = DLookup("[LoanModAgent]", "[ClientList]", "ClientID = " & Forms![wizNOI]![ClientID])

End Function

Private Function FetchLoanModPhone()

  FetchLoanModPhone = DLookup("[LoanModPhone]", "[ClientList]", "ClientID = " & Forms![wizNOI]![ClientID])

End Function

