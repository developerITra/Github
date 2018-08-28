VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterWriteOffReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdComplete_Click()
Dim ctr As Integer, rstJnl As Recordset

If btn1 = True Then
ctr = ctr + 1
End If
If btn2 = True Then
ctr = ctr + 1
End If
If btn3 = True Then
ctr = ctr + 1
End If
If btn4 = True Then
ctr = ctr + 1
End If

If ctr > 1 Then
MsgBox "Please select only 1 reason", vbCritical
Exit Sub
End If

If ctr = 0 Then
MsgBox "Please select a reason", vbCritical
Exit Sub
End If

'Set rstJnl = CurrentDb.OpenRecordset("select * from journal where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
'2/11/14
'lisa


'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
If btn1 = True Then
    DoCmd.SetWarnings False
    strinfo = "Invoice written off because client will not pay for item"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EnterWriteOffReason!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Invoice written off because client will not pay for item"
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing
End If

If btn2 = True Then
'2/11/14

    DoCmd.SetWarnings False
    strinfo = "Invoice written off because of a customary discount"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EnterWriteOffReason!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Invoice written off because of a customary discount"
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing
End If
If btn3 = True Then
'2/11/14
    DoCmd.SetWarnings False
    strinfo = "Invoice written off because items past date billable to client"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EnterWriteOffReason!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Invoice written off because items past date billable to client"
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing
End If
If btn4 = True Then
'2/11/14
    DoCmd.SetWarnings False
    strinfo = "Invoice written off because of other reason"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EnterWriteOffReason!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True


'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Invoice written off because of other reason"
'!Color = 1
'.Update
'End With
'Set rstJnl = Nothing
End If
DoCmd.Close

End Sub
