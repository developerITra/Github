VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queDemandWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdExcel_Click()
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputTable, "qryQueueDemandWaiting", acFormatXLS, TemplatePath & "Demand Waiting Queue info.xlt", True
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

Private Sub lstFiles_GotFocus()
cmdViewItems.Enabled = True
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
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDemanddocswithrestart", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
End Sub

Private Sub cmdViewItems_Click()
DoCmd.OpenForm "MissingDocsListDemandviewonly"
Forms!MissingDocsListDemandviewonly!FileNbr = lstFiles
End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)
Dim rsfC As Recordset
Dim rstqueue As Recordset
AddToList (lstFiles)
DemandCallFromQueue lstFiles
Forms!WizDemand!ComLable.Visible = True
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!DemandWaitingUser = StaffID
!DemandWaitingLastEdited = Date
.Update
End With
Set rstqueue = Nothing


Set rsfC = CurrentDb.OpenRecordset("Select * from FCdetails where filenumber = " & lstFiles & " and current=true ", dbOpenDynaset, dbSeeChanges)
    With rsfC
        If rsfC!ClientSentAcceleration = "C" Then
        Forms!WizDemand!ComLable.Visible = False
        Forms!WizDemand!AccelerationIssued.Locked = False
        Forms!WizDemand!AccelerationExpires.Locked = False
        End If
    End With
Set rsfC = Nothing

If Forms!queDemandWaiting!lstFiles.Column(9) = "1. Approved" Then
    Forms!WizDemand!cmdOKd.Visible = True
End If


End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDemanddocswithrestart", dbOpenDynaset)
If rstqueue.EOF Then
    cntr = 0
    Else
    rstqueue.MoveLast
    cntr = rstqueue.RecordCount
End If
QueueCount = cntr
Set rstqueue = Nothing

End Sub

Private Sub lstFilesred_DblClick(Cancel As Integer)

Dim rstqueue As Recordset
AddToList (lstfilesRed)
DemandCallFromQueue lstfilesRed
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstfilesRed & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!DemandWaitingUser = StaffID
!DemandWaitingLastEdited = Date
.Update
End With
Set rstqueue = Nothing
End Sub
