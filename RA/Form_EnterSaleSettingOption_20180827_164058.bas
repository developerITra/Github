VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterSaleSettingOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOK_Click()

If Frame0 = 1 Then
Forms!foreclosuredetails!Autovalue = "Autovalue"
Else
Forms!foreclosuredetails!Autovalue = "BPO"
End If
DoCmd.Close
End Sub
