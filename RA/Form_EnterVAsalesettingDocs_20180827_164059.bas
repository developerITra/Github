VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterVAsalesettingDocs"
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
If btn16 = True Then
ctr = ctr + 1
End If

If ctr = 0 Then
MsgBox "Please make at least one selection", vbCritical
Exit Sub
End If

If btn16 = True And IsNull(Other) Then
MsgBox "Please key in a reason for Other", vbCritical
Exit Sub
End If

If btn1 = True Then
JrlTxt1 = "Info missing: Title Not Clear for FC"
Set rstdocs = CurrentDb.OpenRecordset("select * from VAsalesettingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Title Clear"
    !DocNeededby = GetStaffID
    .Update
    End With
End If

If btn9 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Original Note"
    Else: JrlTxt2 = JrlTxt2 & ", Waiting on Original Note"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from VAsalesettingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Note"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn10 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on SOT"
    Else: JrlTxt2 = JrlTxt2 & ", SOT"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from VAsalesettingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "SOT"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn11 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Fair Debt"
    Else: JrlTxt2 = JrlTxt2 & ", Fair Debt"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from VAsalesettingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Fair Debt"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn12 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Demand Letter"
    Else: JrlTxt2 = JrlTxt2 & ", Demand Letter"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from VAsalesettingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Demand"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn13 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting for LNA/Notices Sent"
    Else: JrlTxt2 = JrlTxt2 & ", waiting for LNA/Notices Sent"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from VAsalesettingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "LNA/Notice"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn14 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Assignment"
    Else: JrlTxt2 = JrlTxt2 & ", Assignment"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from VAsalesettingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Assignment"
    !DocNeededby = GetStaffID
    .Update
    End With
End If


If btn16 = True Then
If JrlTxt2 = "" Then
    JrlTxt2 = "Other item:  " & Other
    Else: JrlTxt2 = JrlTxt2 & ", Other item:  " & Other
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from VAsalesettingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Other"
    !DocNeededby = GetStaffID
    .Update
    End With
End If

rstdocs.Close

JrlTxt = JrlTxt1 & "  " & JrlTxt2

If CurrentProject.AllForms("queattymilestone3").IsLoaded = False Then
'-1

    DoCmd.SetWarnings False
    strinfo = "Sent to VA salesetting waiting.  " & JrlTxt
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color,Warning) Values(Forms!EnterVAsalesettingDocs!FileNumber,Now,GetFullName(),'" & strinfo & "',1 ,100)"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

'lisa
'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Warning] = 100
'  ![Info] = "Sent to VA salesetting waiting.  " & JrlTxt & vbCrLf
'  ![Color] = 1
'  .Update
'  End With

Call VAsalesettingWaitingCompletionUpdate(FileNumber)
End If
If IsLoadedF("Case List") = True Then
Call ReleaseFile(FileNumber)
End If
DoCmd.Close
End Sub
