VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_StaffConflictMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Command41_Click()
DoCmd.SetWarnings False
 DoCmd.OutputTo acOutputQuery, "StaffConflictID", acFormatXLS, TemplatePath & "Staff Conflict Data.xlt", True
DoCmd.SetWarnings True

End Sub

Private Sub ComPrint_Click()

'DoCmd.OpenReport "StaffConflictIdReport", acViewPreview, , , acHidden
'DoCmd.RunCommand acCmdPrint
'DoCmd.Close acReport, "StaffConflictIdReport"
DoCmd.OpenReport "StaffConflictIdReport", acViewNormal
End Sub
