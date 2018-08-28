VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QueueAccounESC"
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
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM QueueESC", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
End Sub



Private Sub ComApp_Click()

Dim JurText As String
Dim shortClint As String
Dim FileNu As Long
FileNu = Forms!QueueAccounESC.lstFiles.Column(0)
shortClint = Forms!QueueAccounESC.lstFiles.Column(2)
JurText = "Escrow Audit Approved"

If Forms!QueueAccounESC!lstFiles.Column(9) <> "Yes" Then
MsgBox ("You can not approved with Escrow aduit document Not done")
Exit Sub
Else
DoCmd.SetWarnings False
DoCmd.RunSQL ("UPDATE DocIndex set Hold = '' WHERE DocID = " & Forms!QueueAccounESC.lstFiles.Column(11))

DoCmd.RunSQL ("UPDATE Accou_EscQueue set ApproveStatus = '1-Approved',ApproveDate =#" & Now() & "#,ApproveBy = '" & GetFullName() & "',  Hold = '',Dismissed = True , MangerQ = False  WHERE InvoiceId = " & Forms!QueueAccounESC.lstFiles.Column(17) & " And Dismissed = False")

DoCmd.RunSQL ("Insert into ValumeESC (CaseFile,InvoiceNumber,ClientName,Name,CompleteBill,CompleteBillCount,CaseType) Values('" & FileNu & "','" & Forms!QueueAccounESC.lstFiles.Column(19) & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1,'" & Forms!QueueAccounESC.lstFiles.Column(3) & "' )")

DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & FileNu & "', #" & Now & "#,'" & GetFullName() & "','" & JurText & "',2 )")

DoCmd.SetWarnings True
End If

Me!lstFiles.Requery
Me.Requery
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM QueueESC", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing


End Sub



Private Sub ComApprwithCorre_Click()

Dim JurText As String
Dim shortClint As String
Dim FileNu As Long
FileNu = Forms!QueueAccounESC.lstFiles.Column(0)
shortClint = Forms!QueueAccounESC.lstFiles.Column(2)

If Forms!QueueAccounESC!lstFiles.Column(9) <> "Yes" Then
MsgBox ("You can not Approve with correction without Escrow aduit document done")
Exit Sub
Else
JurText = InputBox("Please Add the reason for Approve with Correction")
JurText = Replace(JurText, "'", "''")

DoCmd.SetWarnings False
DoCmd.RunSQL ("UPDATE DocIndex set Hold = '' WHERE DocID = " & Forms!QueueAccounESC.lstFiles.Column(11))

DoCmd.RunSQL ("UPDATE Accou_EscQueue set ApproveStatus = 'correction with Appr',ApproveDate =#" & Now() & "#,ApproveBy = '" & GetFullName() & "',  Hold = '',Dismissed = True , MangerQ = False  WHERE InvoiceId = " & Forms!QueueAccounESC.lstFiles.Column(17) & " And Dismissed = False")

DoCmd.RunSQL ("Insert into ValumeESC (CaseFile,InvoiceNumber,ClientName,Name,ApprWithCorre,ApprWithCorreCount,CaseType) Values('" & FileNu & "','" & Forms!QueueAccounESC.lstFiles.Column(19) & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1,'" & Forms!QueueAccounESC.lstFiles.Column(3) & "' )")


DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & FileNu & "', #" & Now & "#,'" & GetFullName() & "','" & JurText & "',2 )")

DoCmd.SetWarnings True
End If

Me!lstFiles.Requery
Me.Requery
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM QueueESC", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
End Sub

Private Sub ComHold_Click()

Dim JurText As String
Dim shortClint As String
Dim FileNu As Long
FileNu = Forms!QueueAccounESC.lstFiles.Column(0)
shortClint = Forms!QueueAccounESC.lstFiles.Column(2)

If Forms!QueueAccounESC!lstFiles.Column(9) <> "Yes" Then
MsgBox ("You can not be Hold without Escrow aduit document done")
Exit Sub
Else
JurText = InputBox("Please Add the reason of Hold")
JurText = Replace(JurText, "'", "''")

DoCmd.SetWarnings False
DoCmd.RunSQL ("UPDATE DocIndex set Hold = 'H' WHERE DocID = " & Forms!QueueAccounESC.lstFiles.Column(11))

DoCmd.RunSQL ("UPDATE Accou_EscQueue set ApproveStatus = 'Hold',ApproveDate =#" & Now() & "#,ApproveBy = '" & GetFullName() & "',  Hold = 'H',Dismissed = True , MangerQ = False  WHERE InvoiceId = " & Forms!QueueAccounESC.lstFiles.Column(17) & " And Dismissed = False")

DoCmd.RunSQL ("Insert into ValumeESC (CaseFile,InvoiceNumber,ClientName,Name,CompleteHold,CompleteHoldCount,CaseType) Values('" & FileNu & "','" & Forms!QueueAccounESC.lstFiles.Column(19) & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1,'" & Forms!QueueAccounESC.lstFiles.Column(3) & "')")

DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & FileNu & "', #" & Now & "#,'" & GetFullName() & "','" & JurText & "',2 )")

DoCmd.SetWarnings True
End If

Me!lstFiles.Requery
Me.Requery
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM QueueESC", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
End Sub



Private Sub ComToManager_Click()

Dim JurText As String
Dim shortClint As String
Dim FileNu As Long
FileNu = Forms!QueueAccounESC.lstFiles.Column(0)
shortClint = Forms!QueueAccounESC.lstFiles.Column(2)

If Forms!QueueAccounESC!lstFiles.Column(9) <> "Yes" Then
MsgBox ("You can not Send to Manager without Escrow aduit document done")
Exit Sub
Else
JurText = InputBox("Please Add the reason of send to Manager")
JurText = Replace(JurText, "'", "''")

DoCmd.SetWarnings False
'DoCmd.RunSQL ("UPDATE DocIndex set Hold = '' WHERE DocID = " & Forms!QueueAccounESC.lstFiles.Column(11))

DoCmd.RunSQL ("UPDATE Accou_EscQueue set ApproveStatus = 'Send to manager',ApproveDate =#" & Now() & "#,ApproveBy = '" & GetFullName() & "', MangerQ = True ,MangNotic ='" & JurText & "' ,Dismissed = True   WHERE InvoiceId = " & Forms!QueueAccounESC.lstFiles.Column(17) & " And Dismissed = False")

DoCmd.RunSQL ("Insert into ValumeESC (CaseFile,InvoiceNumber,ClientName,Name,ToManagerQ,ToManagerQCount,CaseType) Values('" & FileNu & "','" & Forms!QueueAccounESC.lstFiles.Column(19) & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1,'" & Forms!QueueAccounESC.lstFiles.Column(3) & "')")

DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & FileNu & "', #" & Now & "#,'" & GetFullName() & "','" & JurText & "',2 )")

DoCmd.SetWarnings True
End If
' MangerQ = True ,MangNotic ='" & JrlTxt & "'
Me!lstFiles.Requery
Me.Requery
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM QueueESC", dbOpenDynaset, dbSeeChanges)
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

EscrowCallFromQueue lstFiles
End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer

Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueueESC")
If Not rstqueue.EOF Then
    rstqueue.MoveLast
    QueueCount = rstqueue!ct
Else
    QueueCount = 0
End If

rstqueue.Close
Set rstqueue = Nothing


End Sub

Private Sub lstFilesR_DblClick(Cancel As Integer)

AddToList (lstFilesR)
EscrowCallFromQueueR lstFilesR

End Sub
