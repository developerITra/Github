VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QueueAccountPSAdvancedCostManager"
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
Dim FileN As Long
FileN = Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(0)

JurText = ""
If Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(4) = "Yes" Then JurText = "Approved Void By Manager "
If Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(5) = "H" Then
    If JurText = "" Then
        JurText = "Approved Hold by manager"
        Else
        JurText = JurText + "and Approved Hold by manager"
    End If
End If
If Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(6) = "Yes" Then
    If JurText = "" Then
    JurText = "Approved Offset, Manager to transfer funds from Retainer"
        Else
        JurText = JurText + "and Approved Offset, Manager to transfer funds from Retaine "
    End If
End If

If Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(7) = "Yes" Then
    If JurText = "" Then
    JurText = "Approved Other by Manager"
        Else
        JurText = JurText + "and Approved Other by Manager "
    End If
End If


DoCmd.SetWarnings False
'If IsLoadedF("QueueAccountLitigationBillManager") = True Then

    If (Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(9)) <> 0 Then
      '  DoCmd.RunSQL ("UPDATE DocIndex set Hold = '' WHERE DocID = " & Forms!QueueAccountLitigationBillManager.lstFiles.Column(9))
        DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Void = False , Offset = false , Other = false, Dismissed = True , MangerQ = False  WHERE DocIndexID = " & Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(9))
    Else
        DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Dismissed = True, Void = False , Offset = false , Other = false , MangerQ = False  WHERE CaseFile = " & Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(0))
    End If
    
    Forms!QueueAccountPSAdvancedCostManager!lstFiles.Requery
    Forms!QueueAccountPSAdvancedCostManager.Requery
        Dim rstqueue As Recordset, cntr As Integer
        
        Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueuePSAdvancedCostMnager ")
        If Not rstqueue.EOF Then
            rstqueue.MoveLast
            Forms!QueueAccountPSAdvancedCostManager!QueueCount = rstqueue!ct
        Else
            Forms!QueueAccountPSAdvancedCostManager!QueueCount = 0
        End If
        rstqueue.Close
        Set rstqueue = Nothing
            
        
    
    
'Else
  '  DoCmd.RunSQL ("UPDATE Accou_LitigationBillingQueue set Dismissed = True , MangerQ = False  WHERE DocIndexID <> 0 And CaseFile = " & Forms![Case List]!FileNumber)
'End If
Dim strInsert As String

DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & FileN & "', #" & Now & "#,'" & GetFullName() & "','" & JurText & "',2 )")

DoCmd.SetWarnings True



Call closeformL



End Sub
Private Sub closeformL()
Dim F As Form
Dim FormClosed As Boolean

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "QueueAccountPSAdvancedCostManager"  '  leave these forms open
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

Dim Rege As String
Dim JurText As String

Dim FileN As Long
FileN = Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(0)

JurText = ""
Rege = ""
If Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(4) = "Yes" Then JurText = "Rej. Void:"
If Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(5) = "H" Then JurText = "Rej. Hold:"
If Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(6) = "Yes" Then JurText = "Rej. Offset:"
If Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(7) = "Yes" Then JurText = "Rej. Other:"


Rege = InputBox("Please Add the reason for rejected")
Rege = JurText & " " & Rege
    Rege = Replace(Rege, "'", "''")


DoCmd.SetWarnings False
    'If IsLoadedF("QueueAccountLitigationBillManager") = True Then
    
        If (Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(9)) <> 0 Then
            DoCmd.RunSQL ("UPDATE DocIndex set Hold = '' WHERE DocID = " & Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(9))
            DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Hold = '', Dismissed = False ,Void = false, Offset = false, Other = false, MangerQ = False, MangNotic = '" & Rege & "' WHERE DocIndexID = " & Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(9))
        Else
            DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Hold = '', Dismissed = False , Void = false, Offset = false, Other = false, MangerQ = False, MangNotic = '" & Rege & "' WHERE CaseFile = " & Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(0))
        End If
        
        Forms!QueueAccountPSAdvancedCostManager!lstFiles.Requery
        Forms!QueueAccountPSAdvancedCostManager.Requery
        Dim rstqueue As Recordset, cntr As Integer
        
        Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueuePSAdvancedCostMnager ")
        If Not rstqueue.EOF Then
            rstqueue.MoveLast
            Forms!QueueAccountPSAdvancedCostManager!QueueCount = rstqueue!ct
        Else
            Forms!QueueAccountPSAdvancedCostManager!QueueCount = 0
        End If
        rstqueue.Close
        Set rstqueue = Nothing
            
        
        
        
'    Else
'        DoCmd.RunSQL ("UPDATE Accou_LitigationBillingQueue set Dismissed = False , MangerQ = False, MangNotic = '" & Rege & "'  WHERE DocIndexID <>0 And CaseFile = " & FileN)
'    End If

DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & FileN & "', #" & Now & "#,'" & GetFullName() & "','" & Rege & "',2 )")

DoCmd.SetWarnings True

JurText = ""
Rege = ""

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

Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueuePSAdvancedCostMnager ")
If Not rstqueue.EOF Then
    rstqueue.MoveLast
    QueueCount = rstqueue!ct
Else
    QueueCount = 0
End If

rstqueue.Close
Set rstqueue = Nothing


End Sub
