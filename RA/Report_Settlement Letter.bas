VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Settlement Letter"
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
            "If you fail to go to settlement pursuant to the Contract we will resell "
    Case "DC"
        P1 = "Enclosed you will find a copy of the executed contract of sale for the above referenced property " & _
            "that you purchased at foreclosure sale on " & Format$(Sale, "mmmm d, yyyy") & ".  ""Settlement"" is to " & _
            "occur within Sixty (60) days of ratification.  You should have your settlement company contact us " & _
            "to coordinate settlement.  If you fail to go to settlement pursuant to the Contract we will resell " & _
            "the property at your risk and expense.  This is the only notice you will receive from this office."
    Case "VA"
        P1 = "This office represents Commonwealth Trustees, LLC, seller of the above referenced property.  " & _
            "Enclosed you will find a copy of the executed contract of sale from the courthouse steps for the " & _
            "above referenced property that you purchased at foreclosure sale on " & Format$(Sale, "mmmm d, yyyy") & _
            ".  Settlement is to occur within fifteen (15) days of sale.  You should have your settlement " & _
            "company contact us to coordinate settlement.  If you fail to go to settlement pursuant to the " & _
            "Contract we will resell"
End Select
End Function


Private Function P2() As String
Select Case [FCdetails.State]
    Case "MD"
    
        P2 = "Settlement must occur by " & Format$(DateAdd("d", 10, SaleRat), "mmmm d, yyyy") & _
             ".  If settlement does not occur we will proceed with filing a motion to resell the property.  No further notice will be sent after this letter."

    Case "DC"
        P2 = ""
        
    Case "VA"
        P2 = "Settlement must occur by " & Format$(DateAdd("d", 15, Sale), "mmmm d, yyyy") & _
            ".  If settlement does not occur we will proceed with the process of reselling the property.  This is the only notice you will receive from this office."
            
End Select
End Function


Private Function P3() As String
Select Case [FCdetails.State]
    Case "MD"
    
        P3 = "does not occur we will proceed with filing a motion to resell the property.  No further notice will be sent after this letter."

    Case "DC"
        P3 = ""
        
    Case "VA"
        P3 = "does not occur we will proceed with the process of reselling the property.  This is the only notice you will receive from this office."
            
End Select
End Function


