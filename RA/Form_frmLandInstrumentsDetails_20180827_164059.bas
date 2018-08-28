VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLandInstrumentsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub btnCancel_Click()
Me.Undo
DoCmd.Close
End Sub


Private Sub btnClose_Click()
DoCmd.Close
End Sub

Private Sub btnPrint_Click()

DoCmd.Close acForm, "frmLandInstrumentsDetails", acSaveYes
Call DoReport("MD Land Instruments", -1)
'call that printdocs jazz music
End Sub
