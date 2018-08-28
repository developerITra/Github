VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterDocketingDocs"
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

If ctr = 0 Then
MsgBox "Please make at least one selection", vbCritical
Exit Sub
End If

If btn16 = True And IsNull(Other) Then
MsgBox "Please key in a reason for Other", vbCritical
Exit Sub
End If

If btn1 = True Then
JrlTxt1 = "Info missing: Note"
Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Note"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn2 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Info missing: DOT"
    Else: JrlTxt1 = JrlTxt1 & ", DOT"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "DOT"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn3 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Info missing: Loan Mod"
    Else: JrlTxt1 = JrlTxt1 & ", Loan Mod"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    
    !DocName = "Loan Mod"
    !DocNeededby = GetStaffID
    .Update
    End With
End If

If btn9 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on SOD"
    Else: JrlTxt2 = JrlTxt2 & ", SOD"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "SOD"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn10 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on SOT"
    Else: JrlTxt2 = JrlTxt2 & ", SOT"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
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
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
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
    JrlTxt2 = "Waiting on Acceleration"
    Else: JrlTxt2 = JrlTxt2 & ", Acceleration"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Acceleration"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn13 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting for ANO"
    Else: JrlTxt2 = JrlTxt2 & ", waiting for ANO"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "ANO"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn14 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on NOI Aff"
    Else: JrlTxt2 = JrlTxt2 & ", NOI Aff"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "NOI Aff"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn15 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on NOI"
    Else: JrlTxt2 = JrlTxt2 & ", NOI"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "NOI"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn17 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on LMA info"
    Else: JrlTxt2 = JrlTxt2 & ", LMA info"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "LMA info"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn16 = True Then
If JrlTxt2 = "" Then
    JrlTxt2 = "Other item:  " & Other
    Else: JrlTxt2 = JrlTxt2 & ", Other item:  " & Other
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from Docketingdocsneeded", dbOpenDynaset, dbSeeChanges)
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

'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Warning] = 100
'  ![Info] = "Sent to docketing waiting.  " & JrlTxt & vbCrLf
'  ![Color] = 1
'  .Update
'  End With
  
  DoCmd.SetWarnings False
strinfo = "Sent to docketing waiting.  " & JrlTxt & vbCrLf
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Warning,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),100,'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

Call ReleaseFile(FileNumber)
Call DocketingWaitingCompletionUpdate(FileNumber)
DoCmd.Close
End Sub
