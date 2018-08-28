VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MissingDocsListIntakeViewOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
DoCmd.Close
End Sub

Private Sub cmdCancel_Click()
Dim rstdocs As Recordset, DocsFlag As Boolean, rstJnl As Recordset, DocName As String
Dim DocStillMissing As String
Dim InfoJournal As String

'Remove document record from table once received
DocsFlag = True

Set rstdocs = CurrentDb.OpenRecordset("Select * FROM IntakeDocsNeeded where filenumber=" & FileNbr & " AND DOCID=" & lstFiles.Value, dbOpenDynaset, dbSeeChanges)

With rstdocs
.Edit
!DocReceived = Now
!docreceivedby = GetStaffID
DocName = !DocName
InfoJournal = rstdocs!DocName & " was manually removed from the intake waiting list of outstanding items"
.Update

'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & FileNbr, dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNbr
'!JournalDate = Now
'!Who = GetFullName
'!Info = rstDocs!DocName & " was manually removed from the intake waiting list of outstanding items"
'.Update
'.Close
'End With
.Close
End With

Set rstdocs = CurrentDb.OpenRecordset("Select * FROM IntakeDocsNeeded where filenumber=" & FileNbr, dbOpenDynaset, dbSeeChanges)

With rstdocs
DocStillMissing = ""
Do Until .EOF
If IsNull(!DocReceived) Then
DocStillMissing = DocStillMissing & IIf(DocStillMissing <> "", ", ", "") & !DocName
DocsFlag = False
End If
.MoveNext
Loop
End With
Set rstdocs = Nothing

'Add manual removed journal note
'Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNbr
'!JournalDate = Now
'!Who = GetFullName
'!Info = InfoJournal & IIf(DocStillMissing <> "", ". Waiting on ", "") & DocStillMissing & "."
'!Warning = IIf(DocStillMissing <> "", 100, Null)
'.Update
'.Close
'End With

DoCmd.SetWarnings False
strinfo = InfoJournal & IIf(DocStillMissing <> "", ". Waiting on ", "") & DocStillMissing & "."
strinfo = Replace(strinfo, "'", "''")
Dim strWarning: strWarning = ""
strWarning = IIf(DocStillMissing <> "", 100, Null)
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Warning,Info) Values(" & FileNbr & ",Now(),GetFullName(),'" & strWarning & "','" & strinfo & "')"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

If DocsFlag = True Then
'Note file in DocsMissing queue as Received
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNbr & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!IntakeDocsRecdFlag = True

!IntakeWaitingby = GetStaffID
!IntakeWaitingLastEdited = Now

.Update
End With

End If

Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("Select * FROM qryqueueIntakeWaitinglst_P", dbOpenDynaset, dbSeeChanges)
rs.Close
Set rs = Nothing

lstFiles.Requery
'DoCmd.Close
    
    
End Sub


