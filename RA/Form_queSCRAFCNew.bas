VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_queSCRAFCNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click

DoCmd.Close acForm, "SCRAArchive"

DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdRefresh_Click()
lstFiles.Requery
lstfiles2.Requery
Dim rstqueue As Integer
rstqueue = DCount("filenumber", "scraqueuefiles", "completed= 0")
QueueCount = rstqueue

End Sub

Private Sub ComArchive_Click()
DoCmd.OpenForm "SCRAArchive"


End Sub



Private Sub ComRefreshAll_Click()
DoCmd.SetWarnings False


DoCmd.RunSQL "delete * from SCRAQueueFiles"
DoCmd.OpenQuery "SCRAALLNEW"


DoCmd.SetWarnings True
lstFiles.Requery
lstfiles2.Requery
Dim rstqueue As Integer
rstqueue = DCount("filenumber", "scraqueuefiles", "Completed = 0")

QueueCount = rstqueue

MsgBox "SCRA Queue Refresh Complete"
End Sub

Private Sub ComSCRACancel_Click()
Dim strSQL As String
Dim SCRAID As String
Dim Status As String

SCRAID = DLookup("StageID", "SCRA_ALL_Q", "file= " & lstFiles.Column(0) & " And StageID= " & lstFiles.Column(7))


Select Case SCRAID
Case 10, 20, 30
SCRAID = Left(SCRAID, 1)

Case 11, 12
SCRAID = 111
Case 31
SCRAID = 31
Case 32
SCRAID = 32
Case 33
SCRAID = 33
Case 35
SCRAID = 35
Case 36
SCRAID = 36
Case 37
SCRAID = 37
Case 38
SCRAID = 38
Case 39
SCRAID = 39
Case 41
SCRAID = 4
Case 42
SCRAID = 42
Case 43
SCRAID = 43
Case 44
SCRAID = 44
Case 45
SCRAID = 45
Case 8
SCRAID = 8


Case 50, 60, 70
SCRAID = Left(SCRAID, 1) + 1
Case 51 ' JPM Rat
SCRAID = 61
Case 55 ' JPM Nisi
SCRAID = 65
Case 90
SCRAID = 90
Case 34
SCRAID = 34
Case 125
SCRAID = 125
Case 126
SCRAID = 126
Case 127
SCRAID = 127
Case 91
SCRAID = 91
Case 128
SCRAID = 128
Case 93
SCRAID = 93

Case 94
SCRAID = 94

Case 95
SCRAID = 95


Case 96
SCRAID = 96

Case 97
SCRAID = 97

Case 98
SCRAID = 98

Case 99
SCRAID = 99



End Select


Select Case SCRAID

Case 1
Status = "SCRA Check- First Legal, Cleared"
Case 111
Status = "SCRA Check- First Legal JPM/Wells, Cleared"
Case 2
Status = "SCRA Check- Docketing, Cleared"


Case 3

Status = "SCRA Check- Sale Date , Cleared"



Case 31, 45
Status = "SCRA Check- Sale Date 7 day, Cleared"
Case 32
Status = "SCRA Check- Sale Date JPM 3 day, Cleared"
Case 33
Status = "SCRA Check- Sale Date, Cleared"
Case 4
Status = "SCRA Check- Post Sale, Cleared"
Case 42
Status = "SCRA Check- Post Sale, Cleared"
Case 43, 44
Status = "SCRA Check- Post Sale , Cleared"

Case 5
Status = "SCRA Check- Post Sale, Cleared"
Case 6
Status = "SCRA Check- Ratification, Cleared"
Case 7
Status = "SCRA Check- Deeds Sent, Cleared"
Case 8
Status = "SCRA Check- Day of Sale , Cleared"
Case 9
Status = "SCRA Check- DIL Disposition, Cleared"
Case 61
Status = "SCRA Check- Post Sale, Cleared"
Case 34
Status = "SCRA Check- 2 day Sale, Cleared"
Case 65
Status = "SCRA Check- Ratification, Cleared"

Case 35
Status = "SCRA Check- Sale Date BOA 7 day, Cleared"

Case 36
Status = "SCRA Check- Sale Date PHH 1 day, Cleared"

Case 37
Status = "SCRA Check- Sale Date , Cleared"

Case 38
Status = "SCRA Check- Day of Sale, Cleared"

Case 39
Status = "SCRA Check- 1 Day Before Sale, Cleared"

Case 125
Status = "SCRA Check- New Referral, Cleared"
Case 126
Status = "SCRA Check- New Referral, Cleared"


Case 127
Status = "SCRA Check- Borrower Served, Cleared"

Case 128
Status = "SCRA Check- Title Received, Cleared"

Case 91
Status = "SCRA Check - Sale 40 days, Cleared"

Case 93
Status = "SCRA Check - Sale 14 Days, Cleared"

Case 94
Status = "SCRA Check - Sale 10 Days, Cleared"

Case 95
Status = "SCRA Check - Sale 22 Days, Cleared"

Case 96
Status = "SCRA Check - Sent Complaint To Court, Cleared"

Case 97
Status = "SCRA Check - Judgment Entered, Cleared"

Case 98
Status = "SCRA Check - Title Received, Cleared"

Case 99
Status = "SCRA Check - 1 day prior to sale, Cleared"

End Select


Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & lstFiles.Column(0) & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue

.Edit
Select Case SCRAID
Case 1
!SCRAComplete1 = Now
!SCRAUser1 = GetStaffID
.Update


Case 111
!SCRAComplete1a = Now
!SCRAUser1a = GetStaffID
.Update


Case 2
!SCRAComplete2 = Now
!SCRAUser2 = GetStaffID
.Update



Case 3
!SCRAComplete3 = Now
!SCRAUser3 = GetStaffID
.Update


Case 31
!SCRAComplete3a = Now
!SCRAUser3a = GetStaffID
.Update

Case 32
!SCRAComplete3b = Now
!SCRAUser3b = GetStaffID
.Update


Case 33, 45
!SCRAComplete3c = Now
!SCRAUser3c = GetStaffID
.Update

Case 34
!SCRAComplete3d = Now
!SCRAuser3d = GetStaffID
.Update


Case 39
!SCRAComplete3e = Now
!SCRAUser3e = GetStaffID
.Update



Case 4
If lstFiles.Column(4) = "VA" Then
!SCRAComplete4a = Now
!SCRAUser4a = GetStaffID
Else
!SCRAComplete4b = Now
!SCRAUser4b = GetStaffID
End If
.Update

Case 42
!SCRAComplete4b = Now
!SCRAUser4b = GetStaffID
.Update

Case 43
!SCRAComplete4c = Now
!SCRAUser4c = GetStaffID
.Update

Case 44
!SCRAComplete3f = Now
!SCRAUser3f = GetStaffID
.Update

Case 5
!SCRAComplete4b = Now
!SCRAUser4b = GetStaffID
.Update

Case 61
!SCRAComplete5a = Now
!SCRAUser5a = GetStaffID
.Update

Case 65
!SCRAComplete5_5 = Now
!SCRAUser5_5 = GetStaffID
.Update

 Case 125
!SCRAComplete125 = Now
!SCRAUser125 = GetStaffID
.Update


 Case 35, 36, 37, 38
!SCRAComplete3a = Now
!SCRAUser3a = GetStaffID
.Update


  Case 126
!SCRAComplete126 = Now
!SCRAUser126 = GetStaffID
.Update


 Case 127
!SCRAComplete127Borroer = Now
!SCRACompelte127 = GetStaffID
.Update

Case 128
!SCRAComplete3c = Now
!SCRAUser3c = GetStaffID
.Update


Case 93
!SCRAComplete3b = Now
!SCRAUser3b = GetStaffID
.Update


Case "94"
!SCRAComplete3c = Now
!SCRAUser3c = GetStaffID
.Update

Case "95"
!SCRAComplete3b = Now
!SCRAUser3b = GetStaffID
.Update

Case 6
!SCRAComplete5 = Now
!SCRAUser5 = GetStaffID
.Update

Case 7
!SCRAComplete6 = Now
!SCRAUser6 = GetStaffID

.Update

Case 8
!SCRAComplete7 = Now
!SCRAUser7 = GetStaffID
.Update

 Case 9
!SCRAComplete8 = Now
!SCRAUser8 = GetStaffID
.Update


Case 90
!SCRAComplete9 = Now
!SCRAUser9 = GetStaffID
.Update



Case 91
!SCRAComplete3 = Now
!SCRAUser3 = GetStaffID
.Update


Case 96
!SCRASentComplaintToCortCompleted = Now
.Update


Case 97
!SCRAJudgmentEnteredCompleted = Now
.Update



Case 98

DoCmd.SetWarnings False

strSQL = "UPDATE TitleReceivedArchive SET " & " SCRAsearch = #" & Now() & "# , SCRASearchBy = '" & GetFullName() & _
    "' WHERE FileNumber = " & lstFiles.Column(0) & " AND TitleRecieved = (#" & lstFiles.Column(5) & "#)"
    DoCmd.RunSQL strSQL
    strSQL = ""
DoCmd.SetWarnings True

Case 99
!SCRAComplete3c = Now
!SCRAUser3c = GetStaffID
.Update

End Select




End With
Set rstqueue = Nothing

Dim rstLocalQueue2 As Recordset
Set rstLocalQueue2 = CurrentDb.OpenRecordset("Select * FROM SCRA_ALL_Q where File=" & lstFiles.Column(0), dbOpenDynaset)
    With rstLocalQueue2
    .Edit
    !Completed = True
    .Update
    End With
    
    Dim rstSCRAUpdate As Recordset
    Set rstSCRAUpdate = CurrentDb.OpenRecordset("SCRA_All_update", dbOpenDynaset, dbSeeChanges)
    With rstSCRAUpdate

Set rstSCRAUpdate = CurrentDb.OpenRecordset("SCRA_All_update", dbOpenDynaset, dbSeeChanges)
     With rstSCRAUpdate
    .AddNew
    !File = lstFiles.Column(0)
    !Client = rstLocalQueue2!Client
    !Stage = rstLocalQueue2!Stage
    !State = rstLocalQueue2!State
    !RefDate = rstLocalQueue2!RefDate
    !DueDate = rstLocalQueue2!DueDate
    !StageID = rstLocalQueue2!StageID
    !Who = GetFullName
    !DateCompleted = Now
    !Canceled = True
    .Update
     End With
End With

Set rstSCRAUpdate = Nothing
Set rstLocalQueue2 = Nothing

Dim FileJournal As Long
FileJournal = lstFiles.Column(0)
DoCmd.SetWarnings False
strinfo = Replace(Status, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & lstFiles.Column(0) & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

lstFiles.Requery
lstfiles2.Requery
Dim rstqueueQ As Integer
rstqueueQ = DCount("filenumber", "scraqueuefiles", "completed= 0")
QueueCount = rstqueueQ

DoCmd.OpenForm "Journal New Entry", , , , , , FileJournal
Forms![Journal New Entry]!Info = "Reason of cleared: "


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
        Me.lstFiles.Selected(A) = True
    Exit For
    End If
Next J

If Not blnFound Then MsgBox ("File not in the queue.")
lstFiles.SetFocus
End If
End Sub

Private Sub Fnumber_DblClick(Cancel As Integer)
Fnumber.Value = Null

End Sub

Private Sub Form_Close()
DoCmd.Close acForm, "SCRAArchive"
End Sub

Private Sub Form_Current()

Dim rstqueue As Integer
rstqueue = DCount("file", "SCRA_ALL_Q", "completed= 0")

QueueCount = rstqueue



End Sub



Private Sub Form_Open(Cancel As Integer)

If PrivSCRACancelSearch Then ComSCRACancel.Enabled = True

End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)


'Dim rstfiles As Recordset, f As Form, FormClosed As Boolean
'
'Do
'    FormClosed = False
'    For Each f In Forms
'        Select Case f.Name
'            Case "Main", "queRSII", "Wizards" ' leave these forms open"
'            Case Else
'                If UCase$(Left$(f.Name, 8)) <> "WORKFLOW" Then
'                    FormClosed = True
'                    DoCmd.Close acForm, f.Name
'                    DoEvents
'                End If
'        End Select
'    Next
'Loop Until Not FormClosed
'





Dim SCRAID As String
AddToList (lstFiles)

DoCmd.OpenForm "SCRA Search Info"

Forms![SCRA Search Info]!FileNumber = lstFiles

'added 2/9/15
            
strStage = Trim(Me.lstFiles.Column(3))

Call SCRAnames(lstFiles)

SCRAID = DLookup("StageID", "SCRA_ALL_Q", "file=" & lstFiles & " And StageID=" & lstFiles.Column(7))
OpenCase lstFiles
Select Case SCRAID
Case 10, 20, 30
SCRAID = Left(SCRAID, 1)
Case 11, 12
SCRAID = 111
Case 31
SCRAID = 31
Case 32
SCRAID = 32
Case 33
SCRAID = 33
Case 35
SCRAID = 35
Case 36
SCRAID = 36
Case 37
SCRAID = 37
Case 38
SCRAID = 38
Case 39
SCRAID = 39

Case 41
SCRAID = 4
Case 42
SCRAID = 42
Case 43
SCRAID = 43
Case 44
SCRAID = 44
Case 45
SCRAID = 45
Case 8
SCRAID = 8


Case 50, 60, 70, 80
SCRAID = Left(SCRAID, 1) + 1
Case 51 ' JPM Rat
SCRAID = 61
Case 55 ' JPM Nisi
SCRAID = 65
Case 90
SCRAID = 90
Case 34
SCRAID = 34
Case 125
SCRAID = 125
Case 126
SCRAID = 126
Case 127
SCRAID = 127
Case 91
SCRAID = 91
Case 128
SCRAID = 128
Case 93
SCRAID = 93

Case 94
SCRAID = 94

Case 95
SCRAID = 95

Case 96
SCRAID = 96
Case 97
SCRAID = 97
Case 98
SCRAID = 98
Case 99
SCRAID = 99





End Select

Forms![Case List]!SCRAID = SCRAID
Forms![Case List]!Indicatorbox = 1
Forms![Case List]!Page97.SetFocus
Forms![SCRA Search Info].SetFocus

End Sub

Private Sub lstfiles2_DblClick(Cancel As Integer)
Dim SCRAID As String


AddToList (lstfiles2)

DoCmd.OpenForm "SCRA Search Info"
Forms![SCRA Search Info]!FileNumber = lstfiles2

'added 2/9/15
            
strStage = Trim(Me.lstfiles2.Column(3))
Call SCRAnames(lstfiles2)

SCRAID = DLookup("StageID", "SCRA_ALL_Q", "File =" & lstfiles2 & " And StageID=" & lstfiles2.Column(7))
OpenCase lstfiles2
Select Case SCRAID
Case 10, 20, 30
SCRAID = Left(SCRAID, 1)
Case 11
SCRAID = 111
Case 31
SCRAID = 31
Case 32
SCRAID = 32
Case 33
SCRAID = 33
Case 35
SCRAID = 35

Case 36
SCRAID = 36
Case 37
SCRAID = 37
Case 38
SCRAID = 38
Case 39
SCRAID = 39


Case 41
SCRAID = 4
Case 42
SCRAID = 42
Case 43
SCRAID = 43
Case 44
SCRAID = 44
Case 45
SCRAID = 45

Case 8
SCRAID = 8


Case 50, 60, 70, 80
SCRAID = Left(SCRAID, 1) + 1
Case 51 ' JPM Rat
SCRAID = 61
Case 55 ' JPM Nisi
SCRAID = 65
Case 90
SCRAID = 90
Case 34
SCRAID = 34
Case 125
SCRAID = 125
Case 126
SCRAID = 126

Case 127
SCRAID = 127

Case 91
SCRAID = 91

Case 128
SCRAID = 128

Case 93
SCRAID = 93

Case 94
SCRAID = 94

Case 95
SCRAID = 95

Case 96
SCRAID = 96
Case 97
SCRAID = 97

Case 98
SCRAID = 98

Case 99
SCRAID = 99



End Select

Forms![Case List]!SCRAID = SCRAID
Forms![Case List]!Indicatorbox = 1
Forms![Case List]!Page97.SetFocus
Forms![SCRA Search Info].SetFocus
End Sub
