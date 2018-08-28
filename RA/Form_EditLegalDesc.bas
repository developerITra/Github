VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditLegalDesc"
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

Dim sLegalDesc As String
 If Not IsNull(Forms!foreclosuredetails.LegalDescription) Then
    sLegalDesc = Forms!foreclosuredetails.LegalDescription
Else
    sLegalDesc = ""
End If

'If (Not IsNull(txtLegalDesc.Value) And IsNull(Forms!ForeclosureDetails.LegalDescription)) Or txtLegalDesc.Value <> Forms!ForeclosureDetails.LegalDescription Or (IsNull(txtLegalDesc.Value) And Not IsNull(Forms!ForeclosureDetails.LegalDescription)) Then
    DoCmd.SetWarnings False
    strinfo = "Legal Description has been updated"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EditLegalDesc!txtFileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

    Forms!Journal.Requery
    Forms!foreclosuredetails.LegalDescription = txtLegalDesc
    MsgBox "Legal Description has been updated"
'End If
'DoCmd.Close acForm, Me.Name
'Forms!ForeclosureDetails.Requery

'added track on 3/4/15
Dim rs As Recordset

If (IsNull(txtLegalDesc) And Not IsNull(sLegalDesc)) Or (Not IsNull(txtLegalDesc) And IsNull(sLegalDesc)) Then GoTo tracking:
If (Not IsNull(sLegalDesc) Or sLegalDesc <> "") And Not IsNull(txtLegalDesc) And sLegalDesc <> txtLegalDesc Then GoTo tracking:

tracking:
Set rs = CurrentDb.OpenRecordset("SELECT * FROM Audit_4", dbOpenDynaset, dbSeeChanges)

rs.AddNew
rs!FileNumber = Forms!foreclosuredetails.FileNumber
rs!TableName = "FCdetails"
rs!FieldName = "LegalDescription"
rs!Username = GetFullName
rs!ChangeDate = Now()
rs!ChangeType = "UPDATE"
rs!OldValue = sLegalDesc
rs!NewValue = Me.txtLegalDesc
rs.Update

rs.Close
Set rs = Nothing

'End If
'End If
'End If
'End If

DoCmd.Close acForm, Me.Name
Forms!foreclosuredetails.Requery
End Sub


