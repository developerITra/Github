VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterNOIDocs"
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
Dim ctr As Integer, rstwizqueue As Recordset, rstdocs As Recordset
Dim i As Integer, JrlTxt As String, JrlTxt1 As String, JrlTxt2 As String
Dim strC As String
Dim rstsql As String

If btn9 = True Then
ctr = ctr + 1
End If
If btn10 = True Then
ctr = ctr + 1
End If
If btn11 = True Then
ctr = ctr + 1
End If




If ctr = 0 Or IsNull(ctr) Then
MsgBox "Please make at least one selection", vbCritical
Exit Sub
End If

If btn9 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Rfigs"
    Else: JrlTxt2 = JrlTxt2 & ", Rfigs"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from documentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNbr = FileNumber
    !DocName = "Rfigs"
    !DocsPndgby = GetStaffID
    
    .Update
    End With
End If
If btn10 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Payment Dates"
    Else: JrlTxt2 = JrlTxt2 & ", Payment Dates"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from documentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNbr = FileNumber
    !DocName = "Payment Dates"
    !DocsPndgby = GetStaffID
   
    .Update
    End With
End If
If btn11 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Waiting on Copy of Client Sent NOI"
    Else: JrlTxt2 = JrlTxt2 & ", Copy of Client Sent NOI"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from documentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNbr = FileNumber
    !DocName = "Client Sent NOI Copy"
    !DocsPndgby = GetStaffID
    .Update
    End With
End If

If btn13 = True Then
    If JrlTxt2 = "" Then
    JrlTxt2 = "Default dates incorrect"
    Else: JrlTxt2 = JrlTxt2 & ", Default dates incorrect"
    End If
    Set rstdocs = CurrentDb.OpenRecordset("select * from documentmissing", dbOpenDynaset, dbSeeChanges)
    With rstdocs
    .AddNew
    !FileNbr = FileNumber
    !DocName = "Default Dates"
    !DocsPndgby = GetStaffID
   
    .Update
    End With
End If



rstdocs.Close

JrlTxt = JrlTxt2

 
DoCmd.SetWarnings False
strinfo = "Sent to NOI waiting.  " & JrlTxt & vbCrLf
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Warning,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),100,'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

Set rstwizqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
''If IsNull(rstWizQueue!NOICompleteDocsMsng) Then
With rstwizqueue
.Edit
!NOICompleteDocsMsng = Now
If !DocsRecdFlag Then !DocsRecdFlag = False
!DateInWaiitingQueueNOI = Now
!AttyMilestone1_5 = Null
If !AttyMilestone1_5Reject Then !AttyMilestone1_5Reject = False
'!NOIuser = StaffID
.Update
End With
'End If
Set rstwizqueue = Nothing

DoCmd.Close

DoCmd.SetWarnings False
rstsql = "Insert into ValumeNOI (CaseFile, Client, Name, NOISentBy, NOIWaiting, NOIWaitingC ) values (Forms!wizNOI!FileNumber, ClientShortName(forms!wizNOI!ClientID),Getfullname(),'" & Forms!wizNOI!ClientSentNOI & "',Now(),1) "
DoCmd.RunSQL rstsql
DoCmd.SetWarnings True



MsgBox "NOI Wizard - Docs Missing and File sent to Waiting Queue.", vbInformation
'Call ReleaseFile(FileNumber)
'Me.Requery
DoCmd.Close acForm, "EnterNOIDocs"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "wizNOI"

End Sub
