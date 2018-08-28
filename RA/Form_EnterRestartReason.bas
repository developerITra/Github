VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterRestartReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btn4_AfterUpdate()
If btn4 = True Then
Other.Enabled = True
Else
Other.Enabled = False
End If
End Sub
Private Sub cmdComplete_Click()
Dim ctr As Integer, rstwizqueue As Recordset, Other As String, rstdocs As Recordset, JrlTxt As String, rstFCdetailsCurrent As Recordset, rstFCdetailsPrior As Recordset
Dim rstqueue As Recordset
If btn1 = True Then
ctr = ctr + 1
End If
If btn2 = True Then
ctr = ctr + 1
End If
If btn4 = True Then
ctr = ctr + 1
End If

If btn5 = True Then
ctr = ctr + 1
End If

If btn6 = True Then
ctr = ctr + 1
End If

If btn7 = True Then
ctr = ctr + 1
End If

If btn8 = True Then  ' title and dose not work
ctr = ctr + 1
End If

If btn9 = True Then '
ctr = ctr + 1
End If

If btn10 = True Then
ctr = ctr + 1
End If

If btn11 = True Then
ctr = ctr + 1
End If


If ctr = 0 Then
MsgBox "Please select a reason", vbCritical
Exit Sub
End If

If btn1 = True Then
JrlTxt = " Demand Letter "
    Set rstdocs = CurrentDb.OpenRecordset("select * from Restartdocumentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Demand Letter"
    !DocNeededby = GetStaffID
    .Update
    End With
    
  
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartDocsRecdFlag = False
    .Update
    End With
    Set rstqueue = Nothing
      
End If

If btn2 = True Then
    If JrlTxt = "" Then
    JrlTxt = " NOI "
    Else: JrlTxt = JrlTxt & ", NOI"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Restartdocumentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "NOI"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartDocsRecdFlag = False
    .Update
    End With
    Set rstqueue = Nothing
    
End If


'----

If btn5 = True Then
    If JrlTxt = "" Then
    JrlTxt = " Fair debt "
    Else: JrlTxt = JrlTxt & ", Fair debt"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Restartdocumentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "FD"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartDocsRecdFlag = False
    .Update
    End With
    Set rstqueue = Nothing
    
End If

'----

If btn6 = True Then
    If JrlTxt = "" Then
    JrlTxt = " Judgment figures "
    Else: JrlTxt = JrlTxt & ", Judgment figures"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Restartdocumentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Jfigs"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartDocsRecdFlag = False
    .Update
    End With
    Set rstqueue = Nothing
    
End If
'----


If btn7 = True Then
    If JrlTxt = "" Then
    JrlTxt = " SOT "
    Else: JrlTxt = JrlTxt & ", SOT"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Restartdocumentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "SOT"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartDocsRecdFlag = False
    .Update
    End With
    Set rstqueue = Nothing
    
End If

'----

If btn8 = True Then
    If JrlTxt = "" Then
    JrlTxt = " Title "
    Else: JrlTxt = JrlTxt & ", Title"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Restartdocumentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Title"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartDocsRecdFlag = False
    .Update
    End With
    Set rstqueue = Nothing
    
End If


'-----
If btn9 = True Then
    If JrlTxt = "" Then
    JrlTxt = " Note/Allonge "
    Else: JrlTxt = JrlTxt & ", Note/Allonge"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Restartdocumentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Note"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartDocsRecdFlag = False
    .Update
    End With
    Set rstqueue = Nothing
    
End If
'-----

If btn10 = True Then
    If JrlTxt = "" Then
    JrlTxt = " AOM "
    Else: JrlTxt = JrlTxt & ", AOM"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Restartdocumentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "AOM"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartDocsRecdFlag = False
    .Update
    End With
    Set rstqueue = Nothing
    
End If



'-----

If btn11 = True Then
    If JrlTxt = "" Then
    JrlTxt = " SSN "
    Else: JrlTxt = JrlTxt & ", SSN"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Restartdocumentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "SSN"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartDocsRecdFlag = False
    .Update
    End With
    Set rstqueue = Nothing
    
End If

'-----


If btn4 = True Then
If JrlTxt = "" Then
    JrlTxt = Me!Other
    Else: JrlTxt = JrlTxt & ", " & Me!Other
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Restartdocumentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Other"
    !DocNeededby = GetStaffID
    .Update
    End With
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !RestartDocsRecdFlag = False
    .Update
    End With
    Set rstqueue = Nothing
    
End If
'''''''''''
Set rstwizqueue = CurrentDb.OpenRecordset("select * from wizardqueuestats where filenumber=" & FileNumber & " and current = true", dbOpenDynaset, dbSeeChanges)
'Set rstFCdetailsCurrent = CurrentDb.OpenRecordset("select * from fcdetails where filenumber=" & FileNumber & " and current = true", dbOpenDynaset, dbSeeChanges)
'Set rstFCdetailsPrior = CurrentDb.OpenRecordset("select * from fcdetails where filenumber=" & FileNumber & " and current = false", dbOpenDynaset, dbSeeChanges)
'rstFCdetailsPrior.MoveLast
rstwizqueue.Edit
'If btn1 = True Then
''MsgBox ("File to Restart Waiting , But see your manager to complete New Demand Buttom")
'
''If Not IsNull(rstWizQueue!DemandComplete) Then rstWizQueue!DemandComplete = Null
''If Not IsNull(rstWizQueue!DemandWaiting) Then rstWizQueue!DemandWaiting = Null
''rstWizQueue!DemandComplete = Null
'
'End If
'
'If btn2 = True Then
''    If JrlTxt = "" Then
''    JrlTxt = "NOI"
''    Else: JrlTxt = JrlTxt & ", NOI"
''    End If
'    rstWizQueue!NOIcomplete = Null
'    rstWizQueue!DateInQueueNOI = Null
'End If

rstwizqueue!RestartWaitingUser = GetStaffID
rstwizqueue!RestartWaiting = Date

rstwizqueue.Update
rstwizqueue.Close




'rstFCdetailsCurrent.Edit
'If btn2 = True Then
'If Not IsNull(Forms!ForeclosureDetails!NOI) Then Forms!ForeclosureDetails!NOI = Null
'If Not IsNull(Forms!ForeclosureDetails!ClientSentNOI) Then Forms!ForeclosureDetails!ClientSentNOI = Null
'
'End If

'If btn1 = True Then
'If Not IsNull(Forms!ForeclosureDetails!AccelerationLetter) Then Forms!ForeclosureDetails!AccelerationLetter = Null 'changed by Diane request 08/6/14 verbale
'If Not IsNull(Forms!ForeclosureDetails!ClientSentAcceleration) Then Forms!ForeclosureDetails!ClientSentAcceleration = Null
'If Not IsNull(Forms!ForeclosureDetails!AccelerationIssued) Then Forms!ForeclosureDetails!ClientSentAcceleration = Null


'End If
'rstFCdetailsCurrent.Update
'rstFCdetailsCurrent.Close
''''''''''

'Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
'If IsNull(rstqueue!RestartWaiting) Then
With rstqueue
.Edit
!RestartWaiting = Now
!RestartUser = GetStaffID
If rstqueue!RestartDocsRecdFlag Then rstqueue!RestartDocsRecdFlag = False
If Not IsNull(rstqueue!RestartComplete) Then rstqueue!RestartComplete = Null

.Update
End With
'End If
Set rstqueue = Nothing

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM WizardSupportTwo where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
If Not IsNull(rstqueue!AttyMilestoneRestart) Then rstqueue!AttyMilestoneRestart = Null
'If Not IsNull(rstqueue!DateSentAttyRestart) Then rstqueue!DateSentAttyRestart = Null stopped on 7/23/2015 to avoied missing file from Att to Manager  Sarab

.Update
End With
Set rstqueue = Nothing

'2/11/14
    DoCmd.SetWarnings False
    strinfo = "This file was added to the Restart waiting queue.  Items missing are:  " & JrlTxt
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color,warning) Values(Forms!EnterRestartReason!FileNumber,Now,GetFullName(),'" & strinfo & "',1,100 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

  'lisa
'  Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  lrs.AddNew
'  lrs![FileNumber] = FileNumber
'  lrs![JournalDate] = Now
'  lrs![Who] = GetFullName()
'  lrs![Info] = "This file was added to the Restart waiting queue.  Items missing are:  " & JrlTxt & vbCrLf
'  lrs![Color] = 1
'  lrs![Warning] = 100
'  lrs.Update
  'lisa
'  imgWarning.Picture = dbLocation & "papertray.emf"
' imgWarning.Visible = True

'lrs.Close


Dim rstvalumeintake As Recordset
Set rstvalumeintake = CurrentDb.OpenRecordset("Select * from ValumeRestart", dbOpenDynaset, dbSeeChanges)
With rstvalumeintake
.AddNew
!CaseFile = FileNumber
!Client = DLookup("ShortClientName", "ClientList", "ClientID = " & Forms![Case List]!ClientID)
!RestartWaiting = Now
!RestartWaitingC = 1
!Name = GetFullName()
.Update
End With
Set rstvalumeintake = Nothing

MsgBox "File sent to Restart Waiting Queue", vbInformation
Call ReleaseFile(FileNumber)
Call RestartWaitingCompletionUpdate(FileNumber)
DoCmd.Close acForm, Me.Name
End Sub
