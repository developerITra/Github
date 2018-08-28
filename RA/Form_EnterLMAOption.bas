VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterLMAOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOK_Click()
Dim rstJnl As Recordset, jnltxt As String
If Frame0 = 2 Then
'Forms!foreclosuredetails!Docket = Date
Forms!foreclosuredetails!FLMASenttoCourt = Date
jnltxt = "Docketed with Final LMA"
Else
jnltxt = "Docketed with Preliminary LMA"
End If
'2/11/14

    DoCmd.SetWarnings False
    strinfo = jnltxt
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    DoCmd.Close

'Set rstJnl = CurrentDb.OpenRecordset("select * from journal", dbOpenDynaset, dbSeeChanges)
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = Forms!ForeclosureDetails!FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = jnltxt
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing
'DoCmd.Close
End Sub
