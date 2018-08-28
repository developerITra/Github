VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queNOIdocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdExcel_Click()
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputQuery, "qryQueueNOIdocs", acFormatXLS, TemplatePath & "NOI Waiting Queue info.xlt", True
    DoCmd.SetWarnings True

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
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueNOIdocs", dbOpenDynaset)
If rstqueue.EOF Then
    cntr = 0
    Else
    rstqueue.MoveLast
    cntr = rstqueue.RecordCount
End If
QueueCount = cntr
Set rstqueue = Nothing


End Sub
Private Sub cmdViewItems_Click()
DoCmd.OpenForm "MissingDocsListNOIviewonly"
Forms!MissingDocsListNOIviewonly!FileNbr = lstFiles
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

'If blnFound Then
' Me.lstFiles.Selected(a) = True
' Exit Sub
If Not blnFound Then MsgBox ("File not in the queue.")
lstFiles.SetFocus
End If
'Else
'lstFiles.SetFocus
'End If
End Sub

Private Sub Fnumber_DblClick(Cancel As Integer)
Fnumber.Value = Null

End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)
Dim rstqueue As Recordset
Dim rsfC As Recordset
AddToList (lstFiles)
NOICallFromQueue lstFiles
If Forms!queNOIdocs!lstFiles.Column(8) = "1. Approved" Then Forms!wizNOI!cmdOK.Visible = True
Forms!wizNOI!ComLable.Visible = True

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!NOIuser = StaffID
!NOIWaitingLastEdited = Date
.Update
End With
Set rstqueue = Nothing

Set rsfC = CurrentDb.OpenRecordset("Select * from FCdetails where filenumber = " & lstFiles & " and current=true ", dbOpenDynaset, dbSeeChanges)
    With rsfC
        If rsfC!ClientSentNOI = "C" Then
        Forms!wizNOI!ComLable.Visible = False
        End If
    End With
Set rsfC = Nothing


End Sub
Private Sub Form_Open(Cancel As Integer)

Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueNOIdocs", dbOpenDynaset)
If rstqueue.EOF Then
    cntr = 0
    Else
    rstqueue.MoveLast
    cntr = rstqueue.RecordCount
End If
QueueCount = cntr
Set rstqueue = Nothing





End Sub
Private Sub lstFiles_GotFocus()
cmdViewItems.Enabled = True
End Sub

Private Sub lstFilesred_DblClick(Cancel As Integer)
Dim rstqueue As Recordset
AddToList (lstfilesRed)
NOICallFromQueue lstfilesRed
If Forms!queNOIdocs!lstfilesRed.Column(8) = "1. Approved" Then Forms!wizNOI!cmdOK.Visible = True

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstfilesRed & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!NOIuser = StaffID
!NOIWaitingLastEdited = Date
.Update
End With
Set rstqueue = Nothing
End Sub
