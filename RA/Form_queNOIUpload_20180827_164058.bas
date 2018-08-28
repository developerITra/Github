VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queNOIUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
Dim rstFCdetails As Recordset, rstJnl As Recordset, FileArray() As Variant, i As Integer, A As Integer
'On Error GoTo Err_cmdOK_Click

DoCmd.TransferText acExportDelim, "NOIexportSpec", "qryNOIexport", "s:\Bulk NOI Upload\NOIupload" & Month(Date) & Day(Date) & Year(Date) & ".txt", False
DoCmd.SetWarnings False
DoCmd.OpenQuery "qryqueueNOIUpdate"
DoCmd.SetWarnings True
DoCmd.Close

Set rstFCdetails = CurrentDb.OpenRecordset("SELECT FileNumber FROM FCdetails WHERE (((NOI)=Date() AND ClientSentNOI is null))", dbOpenDynaset, dbSeeChanges)
Set rstJnl = CurrentDb.OpenRecordset("Select * FROM journal", dbOpenDynaset, dbSeeChanges)

'Add filter for sent by client

With rstFCdetails
.MoveLast
A = .RecordCount
ReDim FileArray(A, 1)
.MoveFirst
i = 1
Do Until .EOF
FileArray(i, 1) = !FileNumber
i = i + 1
.MoveNext
Loop
End With

Set rstFCdetails = Nothing

With rstJnl
For i = 1 To A
.AddNew
!FileNumber = FileArray(i, 1)
!JournalDate = Now
!Who = GetFullName
!Info = "File submitted to Commissioner on " & Date
!Color = 1
.Update
Next i
End With
Set rstJnl = Nothing

MsgBox "The bulk upload file has been successfully created"

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
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueNOIupload", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueNOIupload", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

End Sub
