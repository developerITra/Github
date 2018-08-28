VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queIntake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdExcel_Click()

    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputQuery, "qryQueueIntake", acFormatXLS, TemplatePath & "Intake Queue info.xlt", True
    DoCmd.SetWarnings True


End Sub

Private Sub cmdOK_Click()
'Dim rstQueue As Recordset, i As Integer, FileNum As Long
'On Error GoTo Err_cmdOK_Click
'
'For i = 0 To Me.lstFiles.ListCount - 1
'    If Me.lstFiles.Selected(i) Then
'        IntakeCallFromQueue lstFiles.ItemData(i)
'        FileNum = lstFiles.ItemData(i)
'    Else
'    End If
'Next i
'
'Set rstQueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNum & " AND current=true", dbOpenDynaset, dbSeeChanges)
'
'With rstQueue
'.Edit
'If IsNull(!IntakeCompleteby) Then !IntakeCompleteby = GetStaffID
'!IntakeLastEdited = Date
'.Update
'End With
'
'Set rstQueue = Nothing
'
'Exit_cmdOK_Click:
'    Exit Sub
'
'Err_cmdOK_Click:
'    MsgBox Err.Description
'    Resume Exit_cmdOK_Click
    
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
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueintake", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
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

Private Sub lstFiles_DblClick(Cancel As Integer)

Dim rstqueue As Recordset
AddToList (lstFiles)
IntakeCallFromQueue lstFiles
Forms!foreclosuredetails!cmdWizComplete.Visible = False
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
'If IsNull(!IntakeCompleteby) Then !IntakeCompleteby = GetStaffID
!IntakeLastEdited = Date
!IntakeIuser = GetStaffID
.Update
End With
Set rstqueue = Nothing
End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueintake", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
