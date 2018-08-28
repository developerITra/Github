VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterIntakeDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub btn16_AfterUpdate()
If btn16 = True Then
Other.Enabled = True
Else
Other.Enabled = False
End If

End Sub

Private Sub cmdComplete_Click()
Dim ctr As Integer, rstwizqueue As Recordset, rstdocs As Recordset
Dim i As Integer, JrlTxt As String, JrlTxt1 As String, JrlTxt2 As String

If btn1 = True Then
ctr = ctr + 1
End If
If btn2 = True Then
ctr = ctr + 1
End If
If btn3 = True Then
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
If btn8 = True Then
ctr = ctr + 1
End If
If btn9 = True Then
ctr = ctr + 1
End If
If btn10 = True Then
ctr = ctr + 1
End If
If btn11 = True Then
ctr = ctr + 1
End If
If btn12 = True Then
ctr = ctr + 1
End If
If btn13 = True Then
ctr = ctr + 1
End If
If btn14 = True Then
ctr = ctr + 1
End If
If btn15 = True Then
ctr = ctr + 1
End If
If btn16 = True Then
ctr = ctr + 1
End If
If btn17 = True Then
ctr = ctr + 1
End If
If btn18 = True Then
ctr = ctr + 1
End If
If btn19 = True Then
ctr = ctr + 1
End If
If btn20 = True Then
ctr = ctr + 1
End If
If btn21 = True Then
ctr = ctr + 1
End If
If ctr = 0 Then
MsgBox "Please make at least one selection", vbCritical
Exit Sub
End If

If btn1 = True Then
JrlTxt1 = "Affidavits not yet sent: SOT"
'Set rstDocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
'    With rstDocs
'    .AddNew
'    !FileNumber = FileNumber
'    !Timestamp = Now
'    !DocName = "SOT"
'    !DocNeededby = GetStaffID
'    .Update
'    End With
End If
If btn2 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Affidavits not yet sent: MA"
    Else: JrlTxt1 = JrlTxt1 & ", MA"
    End If
'    Set rstDocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
'    With rstDocs
'    .AddNew
'    !FileNumber = FileNumber
'    !Timestamp = Now
'    !DocName = "MA"
'    !DocNeededby = GetStaffID
'    .Update
'    End With
End If
If btn3 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Affidavits not yet sent: Va LNA"
    Else: JrlTxt1 = JrlTxt1 & ", Va LNA"
    End If
'    Set rstDocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
'    With rstDocs
'    .AddNew
'    !FileNumber = FileNumber
'    !Timestamp = Now
'    !DocName = "VA LNA"
'    !DocNeededby = GetStaffID
'    .Update
'    End With
End If
If btn4 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Affidavits not yet sent: SOD"
    Else: JrlTxt1 = JrlTxt1 & ", SOD"
    End If
'    Set rstDocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
'    With rstDocs
'    .AddNew
'    !FileNumber = FileNumber
'    !Timestamp = Now
'    !DocName = "SOD"
'    !DocNeededby = GetStaffID
'    .Update
'    End With
End If
If btn5 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Affidavits not yet sent: ANO"
    Else: JrlTxt1 = JrlTxt1 & ", ANO"
    End If
'    Set rstDocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
'    With rstDocs
'    .AddNew
'    !FileNumber = FileNumber
'    !Timestamp = Now
'    !DocName = "ANO"
'    !DocNeededby = GetStaffID
'    .Update
'    End With
End If
If btn6 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Affidavits not yet sent: NOI Aff"
    Else: JrlTxt1 = JrlTxt1 & ", NOI Aff"
    End If
'    Set rstDocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
'    With rstDocs
'    .AddNew
'    !FileNumber = FileNumber
'    !Timestamp = Now
'    !DocName = "NOI Aff"
'    !DocNeededby = GetStaffID
'    .Update
'    End With
End If
If btn7 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Affidavits not yet sent: PLMA"
    Else: JrlTxt1 = JrlTxt1 & ", PLMA"
    End If
'    Set rstDocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
'    With rstDocs
'    .AddNew
'    !FileNumber = FileNumber
'    !Timestamp = Now
'    !DocName = "PLMA"
'    !DocNeededby = GetStaffID
'    .Update
'    End With
End If
If btn8 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Affidavits not yet sent: FLMA"
    Else: JrlTxt1 = JrlTxt1 & ", FLMA"
    End If
'    Set rstDocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
'    With rstDocs
'    .AddNew
'    !FileNumber = FileNumber
'    !Timestamp = Now
'    !DocName = "FLMA"
'    !DocNeededby = GetStaffID
'    .Update
'    End With
End If
If btn9 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Note"
    Else: JrlTxt2 = JrlTxt2 & ", Note"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "Note"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn10 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on RDOT"
    Else: JrlTxt2 = JrlTxt2 & ", RDOT"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "RDOT"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn11 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on LoanMod"
    Else: JrlTxt2 = JrlTxt2 & ", LoanMod"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "Loan Mod"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn12 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on SSN"
    Else: JrlTxt2 = JrlTxt2 & ", SSN"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "SSN"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn13 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting for Jfigs"
    Else: JrlTxt2 = JrlTxt2 & ", waiting for Jfigs"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "Jfigs"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn14 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Title"
    Else: JrlTxt2 = JrlTxt2 & ", Title"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "Title"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn15 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on NOI"
    Else: JrlTxt2 = JrlTxt2 & ", NOI"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "NOI"
    !DocNeededby = GetStaffID
    .Update
    End With
End If


If btn17 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on AITNO info"
    Else: JrlTxt2 = JrlTxt2 & ", AITNO info"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "AITNO info"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn18 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Executed SOT"
    Else: JrlTxt2 = JrlTxt2 & ", Executed SOT"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "Executed SOT"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn19 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Assignment"
    Else: JrlTxt2 = JrlTxt2 & ", Assignment"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "Assignment"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn20 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Title Review"
    Else: JrlTxt2 = JrlTxt2 & ", Title"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "Title Review"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn16 = True Then
If JrlTxt2 = "" Then
    JrlTxt2 = "Other item:  " & Other
    Else: JrlTxt2 = JrlTxt2 & ", Other item:  " & Other
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !Timestamp = Now
    !DocName = "Other"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
'Not used
'If btn21 = True Then
'    If JrlTxt2 = "" Then
'    JrlTxt2 = "Waiting on " & Other
'    Else: JrlTxt2 = JrlTxt2 & ", " & Other
'    End If
'    Set rstDocs = CurrentDb.OpenRecordset("select * from intakedocsneeded", dbOpenDynaset, dbSeeChanges)
'    With rstDocs
'    .AddNew
'    !FileNumber = FileNumber
'    !Timestamp = Now
'    !DocName = "Other"
'    !DocNeededby = GetStaffID
'    .Update
'    End With
'End If
rstdocs.Close

JrlTxt = JrlTxt1 & "  " & JrlTxt2

Dim lrs As Recordset
'Lisa not working
    DoCmd.SetWarnings False
    strinfo = JrlTxt
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color,warning) Values(Forms!EnterIntakeDocs!FileNumber,Now,GetFullName(),'" & strinfo & "',1,100 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True


'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Warning] = 100
'  ![Info] = JrlTxt & vbCrLf
'  ![Color] = 1
'  .Update
'  End With
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
'!IntakeWaiting = Now
'!IntakeWaitingby = GetStaffID
!IntakeDocsRecdFlag = False
'If Not IsNull(rstqueue!DateSentAttyIntake) Then rstqueue!DateSentAttyIntake = Null
.Update
End With
Set rstqueue = Nothing

Call ReleaseFile(FileNumber)
Call IntakeWaitingCompletionUpdate(FileNumber)
DoCmd.Close
End Sub
