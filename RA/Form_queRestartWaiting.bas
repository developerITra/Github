VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queRestartWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub cmdAddExceptions_click()
DoCmd.OpenForm "wizRestartCaseList1", , , "FileNumber = " & lstFiles
End Sub

Private Sub cmdExcel_Click()
    
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputTable, "qryQueueRestartWaiting", acFormatXLS, TemplatePath & "Restart Waiting Queue info.xlt", True
    DoCmd.SetWarnings True

End Sub

Private Sub Fnumber_AfterUpdate()
If Not IsNull(Fnumber) Then
Dim Value As String
Dim blnFound As Boolean
blnFound = False
Dim J As Integer
Dim A As Integer
For J = 0 To lstFiles.ListCount - 1
   Value = lstFiles.Column(0, J)
   If InStr(Value, Fnumber.Value) Then
        blnFound = True
         A = J
        Me.lstFiles.Selected(A) = True
        lstFiles = Fnumber

    Exit For
    End If
Next J

If Not blnFound Then MsgBox ("File not in the queue.")
lstFiles.SetFocus

End If
End Sub

Private Sub Fnumber_DblClick(Cancel As Integer)
Fnumber.Value = Null

End Sub

Private Sub lstFiles_GotFocus()
cmdViewItems.Enabled = True
End Sub

Private Sub cmdViewItems_Click()

DoCmd.OpenForm "MissingDocsListrestartviewonly"
Forms!MissingDocsListrestartviewonly!FileNbr = lstFiles
End Sub
Private Sub cmdOK_Click()
Dim rstqueue As Recordset
On Error GoTo Err_cmdOK_Click

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!RestartWaitingLastEdited = Date
!RestartWaitingUser = GetStaffID
.Update
End With

Set rstqueue = Nothing

RestartCallFromQueue lstFiles

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
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

Private Sub cmdRefresh_Click()
Me!lstFiles.Requery
Me.Requery
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueRestartWaiting", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)
Dim rstqueue As Recordset
AddToList (lstFiles)


Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!RestartWaitingLastEdited = Date
!RestartWaitingUser = GetStaffID
.Update
End With

Set rstqueue = Nothing

RestartCallFromQueue lstFiles
Forms!foreclosuredetails!cmdWizComplete.Visible = False
Forms!foreclosuredetails!cmdWaiting1.Visible = True

If Forms!querestartWaiting!lstFiles.Column(8) = "1. Approved" Then
    Forms!foreclosuredetails!cmdWizComplete.Visible = True
End If

End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueRestartWaiting", dbOpenDynaset, dbSeeChanges)

'changed on 7/28/2014
If rstqueue.EOF Then
    QueueCount = 0
Else
    QueueCount = 0
    rstqueue.MoveLast
QueueCount = rstqueue.RecordCount
End If
'Do Until rstqueue.EOF
'cntr = cntr + 1
'rstqueue.MoveNext
'Loop
'QueueCount = cntr
Set rstqueue = Nothing


End Sub

Private Sub lstFilesred_DblClick(Cancel As Integer)
Dim rstqueue As Recordset
AddToList (lstfilesRed)


Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstfilesRed & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!RestartWaitingLastEdited = Date
!RestartWaitingUser = GetStaffID
.Update
End With

Set rstqueue = Nothing

RestartCallFromQueue lstfilesRed
Forms!foreclosuredetails!cmdWizComplete.Visible = False
Forms!foreclosuredetails!cmdWaiting1.Visible = True
End Sub
