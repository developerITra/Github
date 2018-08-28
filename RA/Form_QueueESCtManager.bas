VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QueueESCtManager"
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
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM QueuePSAdvancedCostMnager", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing
End Sub



Private Sub ComApp_Click()
If IsNull(lstFiles) Then
MsgBox ("Please select File")
Exit Sub
End If


Dim JurText As String
Dim shortClint As String
Dim FileNu As Long
FileNu = Forms!QueueESCtManager.lstFiles.Column(0)
shortClint = Forms!QueueESCtManager.lstFiles.Column(2)

If Forms!QueueESCtManager!lstFiles.Column(9) <> "Yes" Then
MsgBox ("You can not Send to Manager without Escrow aduit document done")
Exit Sub
Else
JurText = InputBox("Please Add the explaination for Approvel")
JurText = Replace(JurText, "'", "''")

DoCmd.SetWarnings False
DoCmd.RunSQL ("UPDATE DocIndex set Hold = '' WHERE DocID = " & Forms!QueueESCtManager.lstFiles.Column(11))

DoCmd.RunSQL ("UPDATE Accou_EscQueue set ApproveStatus = 'Approved by Manager',ApproveDate =#" & Now() & "#,ApproveBy = '" & GetFullName() & "', MangerQ = False ,MangNotic ='" & JurText & "' ,Dismissed = False   WHERE InvoiceId = " & Forms!QueueESCtManager.lstFiles.Column(17))

'DoCmd.RunSQL ("Insert into ValumeESC (CaseFile,InvoiceNumber,ClientName,Name,ToManagerQ,ToManagerQCount) Values('" & FileNu & "','" & Forms!QueueAccounESC.lstFiles.Column(19) & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1)")

DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & FileNu & "', #" & Now & "#,'" & GetFullName() & "','" & JurText & "',2 )")

DoCmd.SetWarnings True

End If
' MangerQ = True ,MangNotic ='" & JrlTxt & "'
Me!lstFiles.Requery
Me.Requery
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM QueeESCManager", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing


Call closeformL








End Sub
Private Sub closeformL()
Dim F As Form
Dim FormClosed As Boolean

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "QueueESCtManager"  '  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

End Sub

Private Sub ComReg_Click()
If IsNull(lstFiles) Then
MsgBox ("Please select File")
Exit Sub
End If


Dim JurText As String
Dim shortClint As String
Dim FileNu As Long
FileNu = Forms!QueueESCtManager.lstFiles.Column(0)
shortClint = Forms!QueueESCtManager.lstFiles.Column(2)

If Forms!QueueESCtManager!lstFiles.Column(9) <> "Yes" Then
MsgBox ("You can not Send to Manager without Escrow aduit document done")
Exit Sub
Else
JurText = InputBox("Please Add the reason of rejected")
JurText = Replace(JurText, "'", "''")

DoCmd.SetWarnings False
'DoCmd.RunSQL ("UPDATE DocIndex set Hold = '' WHERE DocID = " & Forms!QueueAccounESC.lstFiles.Column(11))

DoCmd.RunSQL ("UPDATE Accou_EscQueue set ApproveStatus = 'Reg',ApproveDate =#" & Now() & "#,ApproveBy = '" & GetFullName() & "', MangerQ = False ,MangNotic ='" & JurText & "' ,Dismissed = False   WHERE InvoiceId = " & Forms!QueueESCtManager.lstFiles.Column(17))


'DoCmd.RunSQL ("Insert into ValumeESC (CaseFile,InvoiceNumber,ClientName,Name,ToManagerQ,ToManagerQCount) Values('" & FileNu & "','" & Forms!QueueAccounESC.lstFiles.Column(19) & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1)")

DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & FileNu & "', #" & Now & "#,'" & GetFullName() & "','" & JurText & "',2 )")

DoCmd.SetWarnings True
End If
' MangerQ = True ,MangNotic ='" & JrlTxt & "'
Me!lstFiles.Requery
Me.Requery
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM QueeESCManager", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
QueueCount = cntr
Set rstqueue = Nothing

Call closeformL





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
If IsNull(lstFiles) Then
MsgBox ("Please select File")
Exit Sub
End If
AddToList (lstFiles)
PSAdvancedCostsCallFromQueue lstFiles
End Sub
Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, cntr As Integer

Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueeESCManager ")
If Not rstqueue.EOF Then
    rstqueue.MoveLast
    QueueCount = rstqueue!ct
Else
    QueueCount = 0
End If

rstqueue.Close
Set rstqueue = Nothing


End Sub
