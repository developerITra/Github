VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queSCRAFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

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
lstFiles.Requery
lstfiles2.Requery
Dim rstqueue As Integer
rstqueue = DCount("filenumber", "scraqueuefiles", "completed=no")

QueueCount = rstqueue

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

Private Sub Form_Current()

Dim rstqueue As Integer
rstqueue = DCount("filenumber", "scraqueuefiles", "completed=no")

QueueCount = rstqueue


End Sub

Private Sub cmdRefresh1_Click()
DoCmd.SetWarnings False


DoCmd.RunSQL "delete * from SCRAQueueFiles"
DoCmd.OpenQuery "SCRAALLNEW"
DoCmd.OpenQuery "SCRAALLNEW_D"

DoCmd.SetWarnings True
lstFiles.Requery
lstfiles2.Requery
Dim rstqueue As Integer
rstqueue = DCount("filenumber", "scraqueuefiles", "Completed = 0")
lstFiles.Requery
lstfiles2.Requery
QueueCount = rstqueue

MsgBox "SCRA Queue Refresh Complete"

End Sub



Private Sub lstFiles_DblClick(Cancel As Integer)
Dim SCRAID As String

DoCmd.OpenForm "SCRA Search Info"
Forms![SCRA Search Info]!FileNumber = lstFiles
Call SCRAnames(lstFiles)

SCRAID = DLookup("SCRAstageID", "SCRAqueuefiles", "filenumber=" & lstFiles)
OpenCase lstFiles
Select Case SCRAID
Case 10, 20, 30
SCRAID = Left(SCRAID, 1)
Case 11, 12
SCRAID = 111
Case 31
SCRAID = 31
Case 32
SCRAID = 32
Case 33
SCRAID = 33
Case 41
SCRAID = 4
Case 42
SCRAID = 42
Case 43
SCRAID = 43
Case 44
SCRAID = 44
Case 45
SCRAID = 45
Case 39
SCRAID = 39
Case 38
SCRAID = 38

Case 50, 60, 70, 80
SCRAID = Left(SCRAID, 1) + 1
Case 51 ' JPM Rat
SCRAID = 61
Case 55 ' JPM Nisi
SCRAID = 65
Case 90
SCRAID = 90
Case 34
SCRAID = 34

End Select

Forms![Case List]!SCRAID = SCRAID
Forms![Case List]!Page97.SetFocus
Forms![SCRA Search Info].SetFocus
End Sub

Private Sub lstfiles2_DblClick(Cancel As Integer)
Dim SCRAID As String

DoCmd.OpenForm "SCRA Search Info"
Forms![SCRA Search Info]!FileNumber = lstfiles2
Call SCRAnames(lstfiles2)

SCRAID = DLookup("SCRAstageID", "qryqueueSCRAfc", "filenumber=" & lstFiles)
OpenCase lstfiles2
Select Case SCRAID
Case 10, 20, 30
SCRAID = Left(SCRAID, 1)
Case 11
SCRAID = 111
Case 31
SCRAID = 31
Case 32
SCRAID = 32
Case 33
SCRAID = 33
Case 41
SCRAID = 4
Case 42
SCRAID = 5
Case 43
SCRAID = 43
Case 44
SCRAID = 44
Case 45
SCRAID = 45
Case 38
SCRAID = 38
Case 39
SCRAID = 39


Case 50, 60, 70, 80
SCRAID = Left(SCRAID, 1) + 1
Case 51 ' JPM Rat
SCRAID = 61
Case 55 ' JPM Nisi
SCRAID = 65
Case 90
SCRAID = 90
End Select

Forms![Case List]!SCRAID = SCRAID
Forms![Case List]!Page97.SetFocus
Forms![SCRA Search Info].SetFocus
End Sub
