VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterDemandReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdCancel_Click()
Me.Undo
DoCmd.Close
End Sub

Private Sub cmdComplete_Click()
Dim ctr As Integer, rstwizqueue As Recordset, rstdocs As Recordset, JrlTxt1 As String

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

If ctr = 0 Then
MsgBox "Please select a reason", vbCritical
Exit Sub
End If

Set rstdocs = CurrentDb.OpenRecordset("select * from demanddocsneeded where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)

If btn1 = True Then
JrlTxt1 = "Info missing: Need Fee Approval"
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Need Fee Approval"
    !DocNeededby = GetStaffID
    .Update
    End With
End If

If btn2 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Info missing: Figures"
    Else: JrlTxt1 = JrlTxt1 & ", Figures"
    End If
     With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Missing Figures"
    !DocNeededby = GetStaffID
    .Update
    End With
End If

If btn3 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Waiting for Fair Debt"
    Else: JrlTxt1 = JrlTxt1 & ",  Waiting for Fair Debt"
    End If
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Waiting for Fair Debt"
    !DocNeededby = GetStaffID
    .Update
    End With
End If

If btn4 = True Then
    If JrlTxt1 = "" Then
    JrlTxt1 = "Waiting for client demand"
    Else: JrlTxt1 = JrlTxt1 & ",  Waiting for client demand"
    End If
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Waiting for client demand"
    !DocNeededby = GetStaffID
    .Update
    End With
End If

Set rstwizqueue = CurrentDb.OpenRecordset("select * from wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
'If IsNull(rstWizQueue!DemandWaiting) Then
With rstwizqueue
.Edit
!DemandWaiting = Now
!DemandUser = StaffID
If Not IsNull(rstwizqueue!DemandComplete) Then rstwizqueue!DemandComplete = Null
!DemandDocsRecdFlag = False
.Update
End With
'End If
Set rstwizqueue = Nothing
    
    DoCmd.SetWarnings False
    strinfo = "This file was added to the demand waiting queue for the following reasons:  " & JrlTxt1
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EnterDemandReason!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
   ' DoCmd.SetWarnings True
    
    'DoCmd.SetWarnings False
    Dim rstsql As String
    rstsql = "Insert InTo ValumeDemand (CaseFile, Client, Name, DemandWaiting, DemandWaitingC,Demandcompleted, DemandcompletedC,SentByClient ) Values ( Forms!EnterDemandReason!FileNumber, ClientShortName(forms!wizdemand!ClientID),Getfullname(),Now(),1, Null,0, '')"
    DoCmd.RunSQL rstsql
    DoCmd.SetWarnings True



'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  lrs.AddNew
'
'  lrs![FileNumber] = FileNumber
'  lrs![JournalDate] = Now
'  lrs![Who] = GetFullName()
'
'  lrs![Info] = "This file was added to the demand waiting queue for the following reasons:  " & JrlTxt1 & vbCrLf
'  lrs![Color] = 1
'  lrs.Update
'
'lrs.Close



MsgBox "File sent to Demand Waiting Queue", vbInformation
Call ReleaseFile(FileNumber)
Me.Requery
DoCmd.Close acForm, Me.Name
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "wizDemand"
End Sub
