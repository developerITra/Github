VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cbxPrint_AfterUpdate()
Select Case cbxPrint
Case "Acrobat"
Forms!foreclosuredetails!PrintOutput = -2
Case "Preview"
Forms!foreclosuredetails!PrintOutput = acPreview
Case "Printer"
Forms!foreclosuredetails!PrintOutput = acViewNormal
End Select
DoCmd.Close
End Sub
