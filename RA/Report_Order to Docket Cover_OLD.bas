VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Order to Docket Cover_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Function GetDocketByDate()
  GetDocketByDate = Format$(Forms![Print Order To Docket].DocketByDate, "mmmm d"", ""yyyy")
End Function

Public Function GetDeadlineDate()
  GetDeadlineDate = Format$(NextWeekDay(DateAdd("d", 1, Forms![Print Order To Docket].DocketByDate)), "mmmm d"", ""yyyy")
End Function

