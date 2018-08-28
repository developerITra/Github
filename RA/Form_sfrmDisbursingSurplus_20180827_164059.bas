VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmDisbursingSurplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub chDisbursingSurplus_AfterUpdate()

If Me.chDisbursingSurplus = True Then
    Forms!foreclosuredetails!btnNewDSurplus.Visible = True
Else
    Forms!foreclosuredetails!btnNewDSurplus.Visible = False
End If
End Sub
