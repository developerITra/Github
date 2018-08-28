VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queIntakeWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdExcel_Click()
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputQuery, "IntakeWaitingQueue", acFormatXLS, TemplatePath & "Intake Waiting Queue info.xlt", True
    DoCmd.SetWarnings True

End Sub

Private Sub cmdOK_Click()
'Dim rstQueue As Recordset, i As Integer
'Dim FileNum As Long
'On Error GoTo Err_cmdOK_Click
'
'For i = 0 To Me.lstFiles.ListCount - 1
'    If Me.lstFiles.Selected(i) Then
'        IntakeCallFromQueue lstFiles.ItemData(i)
'        FileNum = lstFiles.ItemData(i)
'    Else
'    '    MsgBox "File not found"
'    End If
'Next i
''Marko
'
'
'
'Set rstQueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNum & " AND current=true", dbOpenDynaset, dbSeeChanges)
'
'With rstQueue
'.Edit
'!IntakeWaitingby = StaffID
'!IntakeWaitingLastEdited = Date
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
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select count(*)as ct FROM qryQueueIntakeWaiting_P")

If rstqueue.EOF Then
    QueueCount = 0
Else
    rstqueue.MoveLast
    QueueCount = rstqueue!ct

End If

rstqueue.Close
Set rstqueue = Nothing

Me!lstFiles.Requery
Me.Requery

End Sub

Private Sub cmdViewItems_Click()
Dim i As Integer
'Dim Filenum As Long

For i = 0 To Me.lstFiles.ListCount - 1
    If Me.lstFiles.Selected(i) Then
      '  IntakeCallFromQueue lstFiles.ItemData(i)
        FileNum = lstFiles.ItemData(i)
    Else
    '    MsgBox "File not found"
    End If
Next i
Me.FileNumb = FileNum
DoCmd.OpenForm "MissingDocsListIntakeviewonly"
Forms!MissingDocsListIntakeviewonly!FileNbr = FileNum
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
        'lstFiles = Fnumber
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

Private Sub Form_Load()
Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("select count(*)as ct from qryqueueIntakeWaitinglst_P")

If rs.EOF Then
    QueueCount = 0
Else
    rs.MoveLast
    QueueCount = 0
    
QueueCount = rs!ct
End If

rs.Close
Set rs = Nothing

'Me.lstFiles.RowSource = "select * from IntakeWaitingQueue order by [Docs Rec'd] DESC"

End Sub



Private Sub lstFiles_DblClick(Cancel As Integer)


Dim rstqueue As Recordset
AddToList (lstFiles)
IntakeCallFromQueue lstFiles
If Forms!queIntakeWaiting!lstFiles.Column(10) <> "1. Approved" Then Forms!foreclosuredetails!cmdWizComplete.Visible = False

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!IntakeWaitingby = StaffID
!IntakeWaitingLastEdited = Date
.Update
End With
Set rstqueue = Nothing
'7/28/14
'Dim rs As Recordset
'Set rs = CurrentDb.OpenRecordset("Select * FROM qryqueueIntakeWaitinglst_P")
'rs.Close
'Set rs = Nothing
'Me.Requery

End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer
'Set rstqueue = CurrentDb.OpenRecordset("Select * FROM IntakeWaitingQueue", dbOpenDynaset, dbSeeChanges)
Set rstqueue = CurrentDb.OpenRecordset("Select count(*) as ct FROM IntakeWaitingQueue")

If rstqueue.EOF Then
    QueueCount = 0
Else
    QueueCount = 0
    rstqueue.MoveLast
'QueueCount = rstqueue.RecordCount
QueueCount = rstqueue!ct

End If

'Do Until rstqueue.EOF
'cntr = cntr + 1
'rstqueue.MoveNext
'Loop
'QueueCount = cntr
rstqueue.Close
Set rstqueue = Nothing

End Sub

Private Sub lstFiles_GotFocus()
cmdViewItems.Enabled = True
End Sub

Private Sub lstFilesred_DblClick(Cancel As Integer)
Dim rstqueue As Recordset

AddToList (lstfilesRed)
IntakeCallFromQueue lstfilesRed
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!IntakeWaitingby = StaffID
!IntakeWaitingLastEdited = Date
.Update
End With
Set rstqueue = Nothing
End Sub
