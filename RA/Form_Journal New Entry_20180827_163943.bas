VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Journal New Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()


If IsLoadedF("sfrmNamesUpdate") = True Then
If IsNull(Len(Me.Info)) Then
MsgBox ("Please add Data to the Journal")
Exit Sub
End If
End If



Dim J As Recordset, FN As Integer

Me.cmdCancel.SetFocus
Me.cmdOK.Enabled = False  'To stop the annoying Duplicate Journal Entries

On Error GoTo Err_cmdOK_Click


DoCmd.SetWarnings False
strinfo = Info
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",JournalDate,User,'" & strinfo & "'," & IIf(chAccounting, 2, 1) & ")"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

'Set J = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'J.AddNew
'J!FileNumber = FileNumber
'J!JournalDate = JournalDate
'J!Who = User
'J!Info = Info
'J!Color = IIf(chAccounting, 2, 1)
'J.Update
'J.Close

'
' Keep a plain text file as a backup in case the journal records get corrupted.
' There are too many cases to keep a separate file for each one in the same folder.
' Divide them into 'buckets', using CaseID MOD 100 as the bucket number.
'
FN = FreeFile(1)
Open JournalPath & "\" & Format$(FileNumber Mod 100, "00") & "\" & FileNumber & ".txt" For Append As FN
Print #FN, User
Print #FN, JournalDate
Print #FN, Info
Print #FN,
Print #FN,
Close FN



DoEvents
Me.cmdOK.Enabled = True
DoCmd.Close acForm, Me.Name

If IsLoaded("Journal") Then
   Forms![Journal].ViewJournal
End If

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    If Err.Number = 76 Then     ' path not found
        MkDir JournalPath & "\" & Format$(FileNumber Mod 100, "00") & "\"
        Resume
    End If
    MsgBox "Error encountered attempting to add to journal: " & Err.Description, vbExclamation
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
If IsLoadedF("sfrmNamesUpdate") = True Then


UpdateName = True
End If


DoCmd.Close acForm, Me.Name

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdOK_DblClick(Cancel As Integer)
MsgBox "Please only click the OK button once", vbCritical
End Sub




Private Sub Form_Open(Cancel As Integer)
FileNumber = Me.OpenArgs

End Sub

Private Sub Form_Unload(Cancel As Integer)
If IsLoadedF("sfrmNamesUpdate") = True Then
'MsgBox ("asdfasdf  " & Len(Me.Info))

If IsNull(Len(Me.Info)) Then
UpdateName = True
End If
End If


End Sub
