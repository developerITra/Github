VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FileLocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdRelease_Click()
Dim rstLocks As Recordset

On Error GoTo Err_cmdRelease_Click

If MsgBox("Do you really want to release all the locks?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

Set rstLocks = CurrentDb.OpenRecordset("SELECT * FROM Locks WHERE StaffID Is Not Null AND StaffID<>0", dbOpenDynaset, dbSeeChanges)
Do While Not rstLocks.EOF
    rstLocks.Edit
    rstLocks!StaffID = 0
    rstLocks.Update
    rstLocks.MoveNext
Loop
rstLocks.Close
MsgBox "All file locks have been released.", vbInformation
DoCmd.Close

Exit_cmdRelease_Click:
    Exit Sub

Err_cmdRelease_Click:
    MsgBox Err.Description
    Resume Exit_cmdRelease_Click
    
End Sub

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)

If Me.Recordset.RecordCount = 0 Then
    lblHeader.Caption = "There are no files locked at this time."
    cmdRelease.Enabled = False
Else
    lblHeader.Caption = "The following files are locked:"
End If

End Sub

Private Sub cmdRelLock_Click()
Dim rstLocks As Recordset

On Error GoTo Err_cmdRelLock_Click

If MsgBox("Do you really want to release this lock?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

Set rstLocks = CurrentDb.OpenRecordset("SELECT * FROM Locks WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
Do While Not rstLocks.EOF
    rstLocks.Edit
    rstLocks!StaffID = 0
    rstLocks.Update
    rstLocks.MoveNext
Loop
rstLocks.Close
Me.Requery
DoEvents
If Me.Recordset.RecordCount = 0 Then DoCmd.Close
MsgBox "Lock released.", vbInformation

Exit_cmdRelLock_Click:
    Exit Sub

Err_cmdRelLock_Click:
    MsgBox Err.Description
    Resume Exit_cmdRelLock_Click
    
End Sub
