VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Workflow Collection Post-Judgment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
txtInfo = ""

If Not IsNull(JudgmentPrinAmount) Then txtInfo = txtInfo & "Judgment Principal Amount: " & Format$(JudgmentPrinAmount, "Currency") & vbNewLine
If Not IsNull(JudgmentInterest) Then txtInfo = txtInfo & "Judgment Interest Amount: " & Format$(JudgmentInterest, "Currency") & vbNewLine
If Not IsNull(JudgmentFees) Then txtInfo = txtInfo & "Judgment Fees: " & Format$(JudgmentFees, "Currency") & vbNewLine
If Not IsNull(NotifyClient) Then txtInfo = txtInfo & "Notified Client: " & Format$(NotifyClient, "mmmm d, yyyy") & vbNewLine
If Not IsNull(ReceivedInstructions) Then txtInfo = txtInfo & "Received Instructions: " & Format$(ReceivedInstructions, "mmmm d, yyyy") & vbNewLine
If GarnishWages Then txtInfo = txtInfo & "Garnish Wages" & vbNewLine
If AttachPersonalProperty Then txtInfo = txtInfo & "Attach Personal Property" & vbNewLine
If AttachRealProperty Then txtInfo = txtInfo & "Attach Real Property" & vbNewLine
If PostJudgmentDiscovery Then txtInfo = txtInfo & "Post Judgment Discovery" & vbNewLine
If Not IsNull(SettlementDate) Then txtInfo = txtInfo & "SettlementDate: " & Format$(SettlementDate, "mmmm d, yyyy") & vbNewLine
If Not IsNull(SettlementAmount) Then txtInfo = txtInfo & "Settlement Amount: " & Format$(SettlementAmount, "Currency") & vbNewLine
If Not IsNull(SettlementDetails) Then txtInfo = txtInfo & "Settlement Details: " & SettlementDetails & vbNewLine
End Sub

Private Sub Report_NoData(Cancel As Integer)
txtNoData.Visible = True
End Sub
