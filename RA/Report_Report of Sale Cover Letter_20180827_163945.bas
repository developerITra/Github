VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Report of Sale Cover Letter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function P1() As String
Select Case [FCdetails.State]
    Case "MD"
        P1 = "Enclosed you will find a copy of the final order of ratification for the above referenced " & _
            "property that you purchased at foreclosure sale on " & Format$(Sale, "mmmm d, yyyy") & _
            ".  You should now prepare to settle within 10 days from the date of this letter.  " & _
            "If you fail to go to settlement pursuant to the Contract we will resell the property at your " & _
            "risk and expense.  This is the only notice you will receive from this office."
    Case "DC"
        P1 = "Enclosed you will find a copy of the executed contract of sale for the above referenced property " & _
            "that you purchased at foreclosure sale on " & Format$(Sale, "mmmm d, yyyy") & ". Settlement is to " & _
            "occur within thirty (30) days of sale.  You should have your settlement company contact us " & _
            "to coordinate settlement.  If you fail to go to settlement pursuant to the Contract we will resell " & _
            "the property at your risk and expense.  This is the only notice you will receive from this office." & "Hi!"
    Case "VA"
        P1 = "This office represents Commonwealth Trustees, LLC, seller of the above referenced property.  " & _
            "Enclosed you will find a copy of the executed contract of sale from the courthouse steps for the " & _
            "above referenced property that you purchased at foreclosure sale on " & Format$(Sale, "mmmm d, yyyy") & _
            ".  Settlement is to occur within fifteen (15) days of sale.  You should have your settlement " & _
            "company contact us to coordinate settlement.  If you fail to go to settlement pursuant to the " & _
            "Contract we will resell the property at your risk and expense.   This is the only notice you will receive from this office."
End Select
End Function
