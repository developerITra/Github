VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmLineStayingCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub Form_Current()
'If IsLoaded("Case List") = True Then
'    If Forms![Case List]!CaseTypeID = 8 Then
'    Call SetObjectAttributes(CertOfPubField, False)
''    Me.CertOfPubField.Enabled = False
''    Me.CertOfPubField.Locked = True
'    End If
'End If
'
'If DCTabView = False Then
'    CertOfPubField.Enabled = False
'End If
'
'End Sub


Private Sub LineStayingcase_AfterUpdate()
 If Not IsNull(LineStayingcase) Then
 AddStatus FileNumber, LineStayingcase, "Line staying case sent to court"
 
 Forms!foreclosuredetails!sfrmStatus.Requery

 DoCmd.SetWarnings False
    strinfo = "Line sent to court, staying case"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    Forms!Journal.Requery
'Else
 'AddStatus FileNumber, Now(), "Removed Service Deadline"
End If

End Sub

Private Sub LineStayingcase_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    LineStayingcase = Now()
    Call LineStayingcase_AfterUpdate
End If
End Sub
