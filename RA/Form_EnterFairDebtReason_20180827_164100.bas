VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterFairDebtReason"
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

Private Sub cmdCancel_Click()
Me.Undo
DoCmd.Close
End Sub

Private Sub cmdComplete_Click()
Dim ctr As Integer, rstwizqueue As Recordset, Other As String, rstdocs As Recordset, JrlTxt As String
Dim rstsql As String


If btn1 = True Then
ctr = ctr + 1
End If
If btn2 = True Then
ctr = ctr + 1
End If
If btn3 = True Then
ctr = ctr + 1
End If

'If btn4 = True Then
'ctr = ctr + 1
'End If

If ctr = 0 Or IsNull(ctr) Then

MsgBox "Please select a reason", vbCritical
Exit Sub
End If

If btn1 = True Then
JrlTxt = "Items missing are:  Note"
    Set rstdocs = CurrentDb.OpenRecordset("select * from FairDebtdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Note"
    !DocNeededby = GetStaffID
    .Update
    End With
End If
If btn2 = True Then
    If JrlTxt = "" Then
    JrlTxt = "Items missing are:  Figures"
    Else: JrlTxt = JrlTxt & ", Figures"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from FairDebtdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Figures"
    !DocNeededby = GetStaffID
    .Update
    End With
End If

If btn3 = True Then
    If JrlTxt = "" Then
    JrlTxt = "Items missing are:  Military Confirmation"
    Else: JrlTxt = JrlTxt & ", Military Confirmation"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from FairDebtdocsneeded", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNumber = FileNumber
    !DocName = "Military"
    !DocNeededby = GetStaffID
    .Update
    End With
End If

If btn4 = True Then
    If JrlTxt = "" Or JrlTxt = Null Then
        JrlTxt = "Items missing are:  " & Me!Other
    Else: JrlTxt = JrlTxt & ", " & Me!Other
    End If
    '08/07/14 - Linda
    'Set rstDocs = CurrentDb.OpenRecordset("select * from FairDebtdocsneeded", dbOpenDynaset, dbSeeChanges)
    'With rstDocs
    '.AddNew
    '!FileNumber = FileNumber
    '!DocName = Me!Other
    '!DocNeededby = GetStaffID
    '.Update
    'End With
End If


Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
'If IsNull(rstqueue!FairDebtWaiting) Then
With rstqueue
.Edit
!FairDebtWaiting = Now
!FairDebtUser = StaffID
!FairDebtDocsRecdFlag = False
.Update
End With
'End If
Set rstqueue = Nothing
'2/11/14
'lisa

    DoCmd.SetWarnings False
    'strInfo = "This file was added to the fair debt waiting queue.  Items missing are:  " & JrlTxt
    strinfo = "This file was added to the fair debt waiting queue.  " & JrlTxt

    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EnterFairDebtReason!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  lrs.AddNew
'  lrs![FileNumber] = FileNumber
'  lrs![JournalDate] = Now
'  lrs![Who] = GetFullName()
'  lrs![Info] = "This file was added to the fair debt waiting queue.  Items missing are:  " & JrlTxt & vbCrLf
'  lrs![Color] = 1
'  lrs.Update
'
'lrs.Close

DoCmd.SetWarnings False
rstsql = "Insert into ValumeFD (CaseFile, Client, Name, FDWaiting, FDIWaitingC,state ) values (Forms!wizFairDebt!FileNumber, ClientShortName(forms!wizFairDebt!ClientID),Getfullname(),Now(),1, Forms!wizFairDebt!State) "
DoCmd.RunSQL rstsql
DoCmd.SetWarnings True


MsgBox "File sent to Fair Debt Waiting Queue", vbInformation
Call ReleaseFile(FileNumber)
DoCmd.Close acForm, Me.Name
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "wizfairdebt"
End Sub
