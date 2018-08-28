VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_StaffConflict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CmAppr_Click()
Dim strinfo As String
Dim rstqueueW As Recordset

If MsgBox("Are you sure you want to approve the conflict?", vbYesNo) = vbYes Then
ConflictStatus.Value = "Approved"
ConflictStatusDate.Value = Now()
ConflictStatusBy.Value = GetLoginName()

    DoCmd.SetWarnings False
    strinfo = "Approved Employee conflict."
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Me.FileNumber & ",Now, GetLoginName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    AddStatus Me.FileNumber, Now(), "Approved Employee conflict."


Set rstqueueW = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & Me.FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueueW
.Edit
!StaffConflict = "Approved"
.Update
End With
Set rstqueueW = Nothing



Else
Exit Sub
End If
End Sub



Private Sub ComFileRej_Click()
Dim rsfC As Recordset
Dim rstCase As Recordset
Dim strinfo As String
Dim rstqueueW As Recordset
If MsgBox("Are you sure you want to rejected the File? ", vbYesNo) = vbYes Then
    If IsNull(Me.FileNumber) Then
    MsgBox ("there is no file number")
    Exit Sub
    Else
    

        Set rsfC = CurrentDb.OpenRecordset("Select * from FCdetails where filenumber = " & Me.FileNumber & " and current=true ", dbOpenDynaset, dbSeeChanges)
            With rsfC
            .Edit
            rsfC!Disposition = 32
            rsfC!DispositionDate = Now()
            rsfC!DispositionStaffID = StaffID
            .Update
            End With
        Set rsfC = Nothing
        
        Set rstCase = CurrentDb.OpenRecordset("SELECT * FROM CaseList WHERE FileNumber=" & Me.FileNumber, dbOpenDynaset, dbSeeChanges)
        With rstCase
            .Edit
            !Active = False
              .Update
            .Close
        End With
        Set rstCase = Nothing
    
    
    
    DoCmd.SetWarnings False
    strinfo = "Employee conflict. Rejecting File and Must advise Client."
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Me.FileNumber & ",Now, GetLoginName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    
    Me.ConflictStatus.Value = "File Rejected"
    Me.ConflictStatusDate.Value = Now()
    Me.ConflictStatusBy.Value = GetLoginName()
    
    AddStatus Me.FileNumber, Now(), "Employee conflict. Rejecting File"
    
    Set rstqueueW = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & Me.FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
    With rstqueueW
    .Edit
    !StaffConflict = "File Rejected"
    .Update
    End With
    Set rstqueueW = Nothing

    End If
    
 End If
 

End Sub

Private Sub ComRej_Click()
Dim strinfo As String
Dim rstqueueW As Recordset
If MsgBox("Are you sure you want to Reject the conflict?", vbYesNo) = vbYes Then
ConflictStatus.Value = "No Conflict"
ConflictStatusDate.Value = Now()
ConflictStatusBy.Value = GetLoginName()

    DoCmd.SetWarnings False
    strinfo = "No Employee conflict."
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Me.FileNumber & ",Now, GetLoginName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

    AddStatus Me.FileNumber, Now(), "No Employee conflict"

Set rstqueueW = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & Me.FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueueW
.Edit
!StaffConflict = "No Conflict"
.Update
End With
Set rstqueueW = Nothing


Else
Exit Sub
End If

End Sub

Private Sub FirstName_Click()
Call ChangeBackColour
End Sub

Private Sub Form_Current()
Forms!StaffConflict!NameStaff = DLookup("Name", "Staff", "ID= " & Me.StaffID)
If Not IsNull(Me.FileNumber) Then Me.FileNumber.Locked = True


End Sub

Private Sub Form_GotFocus()
ChangeBackColour
End Sub

Private Sub Form_Load()
'Me.Filter = "((staffid =forms!staffconflictmain!textid))"
'Me.FilterOn = True

End Sub

Private Function ChangeBackColour()
   On Error Resume Next
   Screen.ActiveControl.BackColor = "#FFF200"
   
End Function
