VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QueueAccounLitigationBill"
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
Me!lstFiles.Requery
Me.Requery
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM QueueLitigationBilling", dbOpenDynaset, dbSeeChanges)
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
   End If
Next J
lstFiles.SetFocus
If blnFound Then
Me.lstFiles.Selected(A) = True
Else: MsgBox ("File not in the queue.")
lstFiles.SetFocus
End If
Else
lstFiles.SetFocus
End If
End Sub

Private Sub Fnumber_DblClick(Cancel As Integer)
Fnumber.Value = Null
End Sub



Private Sub lstFiles_DblClick(Cancel As Integer)
AddToList (lstFiles)
LitigationBillingCallFromQueue lstFiles
End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer

Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueueLitigationBilling ")
If Not rstqueue.EOF Then
    rstqueue.MoveLast
    QueueCount = rstqueue!ct
Else
    QueueCount = 0
End If

rstqueue.Close
Set rstqueue = Nothing


End Sub
