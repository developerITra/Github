VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditProjectName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdupdate_Click()
'If (Not IsNull(txtLegalDesc.Value) And IsNull(Forms!ForeclosureDetails.LegalDescription)) Or txtLegalDesc.Value <> Forms!ForeclosureDetails.LegalDescription Or (IsNull(txtLegalDesc.Value) And Not IsNull(Forms!ForeclosureDetails.LegalDescription)) Then
    DoCmd.SetWarnings False
    strinfo = "Project Name has been updated"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EditProjectName!txtFileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

    Forms!Journal.Requery
    Forms![Case List].PrimaryDefName = txtProjectName
    MsgBox "Project Name has been updated"
'End If
DoCmd.Close acForm, Me.Name
Forms![Case List].Requery

End Sub


