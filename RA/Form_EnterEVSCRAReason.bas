VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterEVSCRAReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdComplete_Click()
Dim ctr As Integer, rstwizqueue As Recordset

If btn1 = True Then
ctr = ctr + 1
End If
If btn2 = True Then
ctr = ctr + 1
End If

If ctr > 1 Then
MsgBox "Please select only 1 reason", vbCritical
Exit Sub
End If

If ctr = 0 Then
MsgBox "Please select a reason", vbCritical
Exit Sub
End If


Set rstwizqueue = CurrentDb.OpenRecordset("select * from wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
rstwizqueue.Edit
If btn1 = True Then
rstwizqueue!EVreason = 1
End If
If btn2 = True Then
rstwizqueue!EVreason = 2
End If

rstwizqueue.Update
rstwizqueue.Close
'2/11/14

    DoCmd.SetWarnings False
    strinfo = "This file was added to the EV SCRA/PACER waiting queue"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!EnterEVSCRAReason!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  lrs.AddNew
'
'  lrs![FileNumber] = FileNumber
'  lrs![JournalDate] = Now
'  lrs![Who] = GetFullName()
'
'  lrs![Info] = "This file was added to the EV SCRA/PACER waiting queue" & vbCrLf
'  lrs![Color] = 1
'  lrs.Update
'
'lrs.Close

MsgBox "File sent to EV SCRA/PACER Waiting Queue", vbInformation
Call ReleaseFile(FileNumber)
DoCmd.Close acForm, Me.Name
Forms!quescra9.Refresh
End Sub
