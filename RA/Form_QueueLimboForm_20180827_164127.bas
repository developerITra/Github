VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QueueLimboForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdClose_Click()
DoCmd.Close acForm, "QueueLimboForm"
End Sub

Private Sub ComDC_Click()
DoCmd.OpenForm "LIMBO_DC"
End Sub

Private Sub ComMD_Click()
DoCmd.OpenForm "LIMBO_MD"
End Sub

Private Sub ComVA_Click()
DoCmd.OpenForm "LIMBO_VA"
End Sub
