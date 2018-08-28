VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queVASaleSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
Dim rstqueue As Recordset
On Error GoTo Err_cmdOK_Click

        
VAsalesettingCallFromQueue lstFiles
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!VASaleSettingUser = StaffID
!VASaleSettingLastEdited = Date
.Update
End With

If Not rstqueue!VASaleSettingReason < 3 Then
Forms!foreclosuredetails!Sale.Enabled = False
End If


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
Dim rstqueue As Recordset, cntr As Integer

Set rstqueue = CurrentDb.OpenRecordset("Select count(*) as ct FROM qryqueuevasalesettinggroupby")

If rstqueue.EOF Then
    QueueCount = 0
Else
    rstqueue.MoveLast
    QueueCount = rstqueue!ct

rstqueue.Close
Set rstqueue = Nothing
End If

Me!lstFiles.Requery
Me.Requery

End Sub

Private Sub Form_Current()
''Me!lstFiles.Requery
''Me.Requery

'Dim rstqueue As Recordset, cntr As Integer
'Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuevasalesettinggroupby", dbOpenDynaset, dbSeeChanges)
'Do Until rstqueue.EOF
'cntr = cntr + 1
'rstqueue.MoveNext
'Loop
'QueueCount = cntr
'Set rstqueue = Nothing

'10/14/14
'Set rstqueue = CurrentDb.OpenRecordset("qryqueuevasalesettinggroupby")
'If Not rstqueue.EOF Then
    'rstqueue.MoveLast
    'QueueCount = rstqueue.RecordCount
'Else
    'QueueCount = 0
'End If

'rstqueue.Close
'Set rstqueue = Nothing
End Sub

Private Sub Form_Load()
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("select count(*) as ct from qryqueuevasalesettinggroupby")

If rstqueue.EOF Then
    QueueCount = 0
Else
    rstqueue.MoveLast
QueueCount = rstqueue!ct
End If

'Me.lstFiles.RowSource = "SELECT * from qryqueuevasalesettinggroupby"

rstqueue.Close
Set rstqueue = Nothing

End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)

Dim rstqueue As Recordset
AddToList (lstFiles)
VAsalesettingCallFromQueue lstFiles
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!VASaleSettingUser = GetStaffID
!VASaleSettingLastEdited = Date
.Update
End With
Set rstqueue = Nothing
End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer, rstwiz As Recordset

'Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuevasalesettinggroupby", dbOpenDynaset, dbSeeChanges)

'Do Until rstqueue.EOF
'cntr = cntr + 1
'rstqueue.MoveNext
'Loop
'QueueCount = cntr
'Set rstqueue = Nothing


'10/14/14

'Set rstqueue = CurrentDb.OpenRecordset("select count(*) as ct from qryqueuevasalesettinggroupby", dbOpenDynaset, dbSeeChanges)

'If Not rstqueue.EOF Then
    'rstqueue.MoveLast
   ' QueueCount = rstqueue!ct

'Else
    'QueueCount = 0
    'End If


'rstqueue.Close
'Set rstqueue = Nothing

'Set rstWiz = CurrentDb.OpenRecordset("Select vasalesettingqueue FROM wizardqueuestats where filenumber=9999999", dbOpenDynaset, dbSeeChanges)
'If rstWiz!vasalesettingQueue <> Date Then
'rstWiz.Edit
'rstWiz!vasalesettingQueue = Date
'rstWiz.Update
'rstWiz.Close
DoCmd.SetWarnings False
DoCmd.OpenQuery "rqryVAsalesettingUpdatefinal"
'DoCmd.OpenQuery "qryQueueVASaleSetting_P"
'DoCmd.OpenQuery "rqryvasalesettingMakeTable"
'DoCmd.OpenQuery "rqryvasalesettingUpdate"
DoCmd.SetWarnings True
'End If
End Sub

Private Sub lstFilesred_DblClick(Cancel As Integer)
Dim rstqueue As Recordset
VAsalesettingCallFromQueue lstFiles
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!VASaleSettingUser = GetStaffID
!vasalesettingLastReviewed = Date
.Update
End With
Set rstqueue = Nothing
End Sub

