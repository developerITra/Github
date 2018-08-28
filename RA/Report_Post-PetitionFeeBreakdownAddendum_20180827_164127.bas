VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Post-PetitionFeeBreakdownAddendum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(Cancel As Integer)
If Forms![Case List]!ClientID = 385 Then
    lbFooter.Visible = True
    lb1.Visible = True
    Text31.Visible = False
Else
    lbFooter.Visible = False
    lb1.Visible = False
    Text31.Visible = True

End If
End Sub
