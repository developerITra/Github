VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MissingTitleOrderViewOnly"
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
'Remove document record from table once received
DocsFlag = True
If IsNull(lstFiles) Then
MsgBox ("Please select missing doc")
Exit Sub
End If


Set rstdocs = CurrentDb.OpenRecordset("Select * FROM TitleDocumentMissing where filenumber=" & FileNbr & " AND DOCID=" & lstFiles.Value, dbOpenDynaset, dbSeeChanges)

With rstdocs
.Edit
!DocReceived = Now
!docreceivedby = GetStaffID
DocName = !DocName
.Update
.Close
End With

Set rstdocs = CurrentDb.OpenRecordset("Select * FROM TitleDocumentMissing where filenumber=" & FileNbr, dbOpenDynaset, dbSeeChanges)

With rstdocs
Do Until .EOF
If IsNull(!DocReceived) Then
DocsFlag = False
End If
.MoveNext
Loop
End With
Set rstdocs = Nothing

'Add manual removed journal note

'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
''Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal where filenumber=" & lstFiles, dbOpenDynaset, dbSeeChanges)02/05/14
'With rstJnl
'.AddNew
'!FileNumber = lstFiles
'!JournalDate = Now
'!Who = GetFullName
'!Info = DocName & "was manually removed from the Title Order list of outstanding items"
'.Update
'.Close
'End With

DoCmd.SetWarnings False
strinfo = DocName & "was manually removed from the Title Order list of outstanding items"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info) Values(" & lstFiles & ",Now,GetFullName(),'" & strinfo & "')"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

If DocsFlag = True Then
'Note file in DocsMissing queue as Received
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNbr & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!TitelMissingReson = True
.Update
End With

End If
lstFiles.Requery
'DoCmd.Close
    
End Sub


