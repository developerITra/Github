VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ClientRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btn_Auction_Click()
Me.DetailsR.ControlSource = "AuctionDotCom"
Me.RuleName = "Auction.com RULES"
Me.DetailsR.Visible = True
'Me.DetailsR.Enabled = PrivAdmin
'Me.DetailsR.Locked = Not (PrivAdmin)
End Sub

Private Sub ComFannieMae_Click()
Me.DetailsR.ControlSource = "FannieMae"
Me.RuleName = "Fannie Mae Rules"
Me.DetailsR.Visible = True
End Sub

Private Sub Command4_Click()
Me.DetailsR.ControlSource = "SCRAR"
Me.RuleName = "SCRA RULES"
Me.DetailsR.Visible = True
'Me.DetailsR.Enabled = PrivAdmin
End Sub

Private Sub Command7_Click()
Me.DetailsR.ControlSource = "Negative"
Me.RuleName = "Negative hit RULES"
Me.DetailsR.Visible = True
'Me.DetailsR.Enabled = PrivAdmin

End Sub

Private Sub Form_Current()
'Text2.ControlSource  =
'DetailsR.ControlSource = "Select Rules.SCRAR FROM Rules Where ClientNo =22 "
End Sub
