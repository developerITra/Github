VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queFLMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
Dim rstqueue As Recordset
On Error GoTo Err_cmdOK_Click
        
FLMACallFromQueue lstFiles
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!FLMAUser = StaffID
!FLMALastEdited = Date
.Update
End With

Set rstqueue = Nothing

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
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueFLMA", dbOpenDynaset, dbSeeChanges)
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
FLMACallFromQueue lstFiles
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!FLMAUser = GetStaffID
!FLMALastEdited = Date
.Update
End With
Set rstqueue = Nothing
End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer, rstwiz As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueFLMA", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
Set rstwiz = CurrentDb.OpenRecordset("Select FLMAqueue FROM wizardqueuestats where filenumber=9999999", dbOpenDynaset, dbSeeChanges)
If rstwiz!FLMAQueue <> Date Then
rstwiz.Edit
rstwiz!FLMAQueue = Date
rstwiz.Update
rstwiz.Close
DoCmd.SetWarnings False
DoCmd.OpenQuery "rqryFLMAMakeTable"
DoCmd.OpenQuery "rqryFLMAUpdate"
DoCmd.SetWarnings True
End If
End Sub
