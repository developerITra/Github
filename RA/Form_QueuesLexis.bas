VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QueuesLexis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdBankoQueue_Click()
If Not LexisNexis Then
MsgBox ("You are not Authorized to Access this Queue")
Exit Sub
Else

On Error GoTo Err_cmdBankoQueue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queBanko"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdBankoQueue_Click:
    Exit Sub

Err_cmdBankoQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdBankoQueue_Click
End If
End Sub



Private Sub CmdBankoQueueNoSSN_Click()
If Not LexisNexis Then
MsgBox ("You are not Authorized to Access this Queue")
Exit Sub
Else

Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queBankoNoSS"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End If
End Sub



Private Sub cmdClose_Click()

DoCmd.Close

    
End Sub



Private Sub cmdDeceaQueue_Click()
If Not LexisNexis Then
MsgBox ("You are not Authorized to Access this Queue")
Exit Sub
Else

On Error GoTo Err_cmdDeceaQueue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queDecea"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdDeceaQueue_Click:
    Exit Sub

Err_cmdDeceaQueue_Click:
    MsgBox Err.Description
    Resume Exit_cmdDeceaQueue_Click
End If

End Sub







Private Sub ComNegative_Click()

If Not LexisNexis Then
MsgBox ("You are not Authorized to Access this Queue")
Exit Sub
Else


DoCmd.SetWarnings False

    
    DoCmd.OpenQuery "AppendIntoBankoDelay", acNormal, acEdit

Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queBankoNegitve"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
DoCmd.SetWarnings True

Exit_cmdBankoQueue_Click:
    Exit Sub
End If

End Sub

Private Sub DeceNoSSNQueue_Click()
If Not LexisNexis Then
MsgBox ("You are not Authorized to Access this Queue")
Exit Sub
Else
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "queDeceaNoSSN"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End If
End Sub




Private Sub Form_Open(Cancel As Integer)


Dim rstbankoqueue, rstbankoqueueNoSSN As Recordset
Dim cntr, cntrB, BankoQueueCount As Integer
Dim BankoQueueNoSSN As Integer
Set rstbankoqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueBanko", dbOpenDynaset, dbSeeChanges)
Do Until rstbankoqueue.EOF
cntr = cntr + 1
rstbankoqueue.MoveNext
Loop
BankoQueueCount = cntr
Set rstbankoqueue = Nothing

cmdBankoQueue.Caption = "Banko SSN (" & BankoQueueCount & ")"

Set rstbankoqueueNoSSN = CurrentDb.OpenRecordset("Select * FROM qryQueueBankoNoSSN", dbOpenDynaset, dbSeeChanges)
Do Until rstbankoqueueNoSSN.EOF
cntrB = cntrB + 1
rstbankoqueueNoSSN.MoveNext
Loop
BankoQueueNoSSN = cntrB
Set rstbankoqueueNoSSN = Nothing

CmdBankoQueueNoSSN.Caption = "Banko No SSN (" & BankoQueueNoSSN & ")"

Dim rstDecqueue As Recordset, cntrD As Integer
Set rstDecqueue = CurrentDb.OpenRecordset("Select * FROM qryQueueDecea", dbOpenDynaset, dbSeeChanges)
Do Until rstDecqueue.EOF
cntrD = cntrD + 1
rstDecqueue.MoveNext
Loop
DecQueueCount = cntrD
cmdDeceaQueue.Caption = "Dece. SSN (" & DecQueueCount & ")"

Dim rstDecNoSSN As Recordset, cntrN As Integer
Set rstDecNoSSN = CurrentDb.OpenRecordset("Select * FROM qryQueueDeceaNoSSN", dbOpenDynaset, dbSeeChanges)
Do Until rstDecNoSSN.EOF
cntrN = cntrN + 1
rstDecNoSSN.MoveNext
Loop
DecNoSSNQueueCount = cntrN
DeceNoSSNQueue.Caption = "Dece. NoSSN (" & DecNoSSNQueueCount & ")"

Dim rstNegative As Recordset, cntrNeg As Integer
Set rstNegative = CurrentDb.OpenRecordset("Select * FROM BankoQueueNegativeHitsQ")
Do Until rstNegative.EOF
cntrNeg = cntrNeg + 1
rstNegative.MoveNext
Loop
NegQueue = cntrNeg
ComNegative.Caption = "Negative Hit (" & NegQueue & ")"


 

End Sub


