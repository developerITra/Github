VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEmailImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub UpdateEmailList()
Dim strOutlookFolder As String

On Error GoTo UpdateEmailListErr



Select Case Me.optOutlookFolder
  Case 1
    strOutlookFolder = "Inbox"
  Case 2
    strOutlookFolder = "Sent Items"
End Select

Call GetOutlookItemList(strOutlookFolder)

RefreshList


Exit Sub

UpdateEmailListErr:
    MsgBox Err.Description, vbCritical
    Exit Sub
    
End Sub

Private Sub RefreshList()

lstEmail.RowSource = "SELECT ID, ReceivedDate as [Received Date], Who, Subject, Content from Emails order by " & Me.cboSortby
lstEmail.Requery

End Sub

Private Sub cboSortby_AfterUpdate()
  RefreshList
End Sub


Private Sub cmdCancel_Click()
On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click

End Sub

Private Sub cmdDeleteMsg_Click()

Dim i As Long

With lstEmail
  For i = 0 To .ListCount - 1
    If .Selected(i) Then
      DeleteItem (.Column(0, i))
    End If
  Next i
End With

Call UpdateEmailList

End Sub

Private Sub cmdImport_Click()
Dim i As Long

'Dim lrs As Recordset
'Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)

With lstEmail
  For i = 0 To .ListCount - 1
    If .Selected(i) Then
'      lrs.AddNew
'
'      lrs![FileNumber] = Forms![Case list]![FileNumber]
'      lrs![JournalDate] = Now
'      lrs![Who] = GetFullName()
'      lrs![Info] = "EMAIL - Received: " & DLookup("[ReceivedDate]", "[Emails]", "[ID] = '" & .Column(0, i) & "'") & " " & _
'                   "From: " & DLookup("[Who]", "[Emails]", "[ID] = '" & .Column(0, i) & "'") & " " & _
'                   "Subject: " & DLookup("[Subject]", "[Emails]", "[ID] = '" & .Column(0, i) & "'") & " " & _
'                   DLookup("[Content]", "[Emails]", "[ID] = '" & .Column(0, i) & "'")
'
'      lrs![Color] = 1
'      lrs.Update
    DoCmd.SetWarnings False
    strinfo = "EMAIL - Received: " & DLookup("[ReceivedDate]", "[Emails]", "[ID] = '" & .Column(0, i) & "'") & " " & _
                   "From: " & DLookup("[Who]", "[Emails]", "[ID] = '" & .Column(0, i) & "'") & " " & _
                   "Subject: " & DLookup("[Subject]", "[Emails]", "[ID] = '" & .Column(0, i) & "'") & " " & _
                   DLookup("[Content]", "[Emails]", "[ID] = '" & .Column(0, i) & "'")
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms![Case list]![FileNumber],Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
     
    End If
  Next i
End With

lrs.Close
Set lrs = Nothing

If IsLoaded("Journal") Then
   Forms![Journal].ViewJournal
End If


End Sub

Private Sub Form_Open(Cancel As Integer)
  Call UpdateEmailList
End Sub

Private Sub optOutlookFolder_AfterUpdate()
  Call UpdateEmailList
End Sub

