Attribute VB_Name = "Wizards"
Option Compare Database
Option Explicit
Sub ClearSCRAManual()
Dim rstqueue As Recordset, rstUpdate As Recordset, File As Long

Set rstUpdate = CurrentDb.OpenRecordset("Select * FROM tmpscraqueuefiles where completed=yes", dbOpenDynaset, dbSeeChanges)

Do
File = rstUpdate!FileNumber
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & File & " and current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
Select Case rstUpdate!SCRAstageID
Case 10
!SCRAComplete1 = Now
!SCRAUser1 = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 11, 12
!SCRAComplete1a = Now
!SCRAUser1a = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 20
!SCRAComplete2 = Now
!SCRAUser2 = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 30
!SCRAComplete3 = Now
!SCRAUser3 = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 31
!SCRAComplete3a = Now
!SCRAUser3a = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 32
!SCRAComplete3b = Now
!SCRAUser3b = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
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

'Forms!queSCRAFC.Refresh
Case 41
!SCRAComplete4a = Now
!SCRAUser4a = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 43
!SCRAComplete4c = Now
!SCRAUser4c = GetStaffID
.Update

Case 44
!SCRAComplete3f = Now
!SCRAUser3f = GetStaffID
.Update

'Forms!queSCRAFC.Refresh
Case 42
!SCRAComplete4b = Now
!SCRAUser4b = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 51
!SCRAComplete5a = Now
!SCRAUser5a = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 55
!SCRAComplete5_5 = Now
!SCRAUser5_5 = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 50
!SCRAComplete5 = Now
!SCRAUser5 = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 60
!SCRAComplete6 = Now
!SCRAUser6 = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 70
!SCRAComplete7 = Now
!SCRAUser7 = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 80
!SCRAComplete8 = Now
!SCRAUser8 = GetStaffID
.Update
'Forms!queSCRAFC.Refresh
Case 90
!SCRAComplete9 = Now
!SCRAUser9 = GetStaffID
.Update
End Select
.Close
End With
rstUpdate.MoveNext
Loop While Not rstUpdate.EOF



End Sub
'Sub pdftest()
'
'Dim PDFapp As cacroapp
'Dim CurrentPDF As CAcroPDDoc
'Dim MergedPDF As CAcroPDDoc
'Dim DocFolder As String
'Dim DocName As String
'DocFolder = "c:\andrew\"
'DocName = "titleorder.pdf"
'Set PDFapp = CreateObject("acroexch.App")
'PDFapp.Show
'Set CurrentPDF = CreateObject("acroexch.pddoc")
'CurrentPDF.Open (DocFolder & DocName)
'CurrentPDF.Save (pdsavecopy, DocFolder & "test.pdf")
'End Sub

Sub RSIICallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, F As Form, FormClosed As Boolean

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queRSII", "Wizards" ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If

FileLocks = True
    If LockFile(FileNumber) Then
    DoCmd.OpenForm "wizreferralII"
    
   'Added on 9/9/2014
    
    If DCount("*", "CIVDetails", "FCFileNumber= " & FileNumber) > 0 Then
              MsgBox "CAUTION! Litigation in progress, see an attorney!", vbExclamation
              Forms!wizreferralII.Detail.BackColor = vbYellow
    End If

        Forms!wizreferralII.RecordSource = CurrentDb.QueryDefs("qryWizRefSpecII").sql
        Forms!wizreferralII.txtFileNumber = FileNumber
        Forms!wizreferralII.Filter = "FileNumber=" & FileNumber
        Forms!wizreferralII.FilterOn = True
        Call RSIILoanType_AfterUpdate
        Call RSIIConfirmationVisible(True)
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
    rstfiles.Close



On Error GoTo Err_cmdYes_Click
Forms!wizreferralII.lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], DocTitleID FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='' AND Filespec IS NOT NULL and DeleteDate is null"
Forms!wizreferralII.lstDocs.Requery
Call RSIIFieldsVisible(True)

Exit_cmdYes_Click:
    Exit Sub

Err_cmdYes_Click:
    MsgBox Err.Description
    Resume Exit_cmdYes_Click
    
End Sub
Sub NOICallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, F As Form, FormClosed As Boolean

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queNOInew", "queNOIdocs", "Wizards" ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    DoCmd.OpenForm "wizNOI"
    
            If DCount("*", "CIVDetails", "FCFileNumber= " & FileNumber) > 0 Then
              MsgBox "CAUTION! Litigation in progress, see an attorney!", vbExclamation
              Forms!wizNOI.Detail.BackColor = vbYellow
            End If

    
    With Forms!wizNOI
        .RecordSource = CurrentDb.QueryDefs("qryqry45days").sql
        .txtFileNumber = FileNumber
        .Filter = "FileNumber=" & FileNumber
        .FilterOn = True
    End With
        Call NOIConfirmationVisible(True)
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If

rstfiles.Close

On Error GoTo Err_cmdYes_Click

Forms!wizNOI!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name],DocTitleID FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='' AND Filespec IS NOT NULL and DeleteDate is null"
Forms!wizNOI!lstDocs.Requery
Call NOIFieldsVisible(True)

Exit_cmdYes_Click:
    Exit Sub

Err_cmdYes_Click:
    MsgBox Err.Description
    Resume Exit_cmdYes_Click
    
End Sub
Sub Restart1CallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queRestart", "Wizards", "queRSIReview" ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT FileNumber FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "wizRestartFCdetails1"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm stDocName, , , stLinkCriteria
Forms!wizRestartFCdetails1!cmdReturn.Caption = "Remove from Queue"

    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms!wizRestartFCdetails1.SetFocus
Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub

Sub Restart1CallFromReviewMgrQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queRestart", "Wizards", "queRSIReview" ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed


Set rstfiles = CurrentDb.OpenRecordset("SELECT FileNumber FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "wizRestartFCdetails1"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"


rstfiles.Close

    Dim WarningLevel As Integer
    DoCmd.OpenForm "wizRestartCaseList1", , , "FileNumber= " & FileNumber
    Forms!wizRestartCaseList1!WizMag = "Mgr"
    'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & lstConflicts
    DoCmd.OpenForm "Journal", , , "FileNumber = " & FileNumber
    Forms!wizRestartCaseList1.SetFocus
    WarningLevel = Nz(DMax("Warning", "Journal", "FileNumber=" & FileNumber))
    With Forms!wizRestartCaseList1
    Select Case WarningLevel
        Case 50
            .imgWarning.Picture = dbLocation & "dollar.emf"
            .imgWarning.Visible = True
        Case 100
            .imgWarning.Picture = dbLocation & "papertray.emf"
            .imgWarning.Visible = True
        Case 200
            .imgWarning.Picture = dbLocation & "house.emf"
            .imgWarning.Visible = True
        Case 300
            .imgWarning.Picture = dbLocation & "caution.bmp"
            .imgWarning.Visible = True
        Case 400
            .imgWarning.Picture = dbLocation & "stop.emf"
            .imgWarning.Visible = True
        Case Else
            .imgWarning.Visible = False
    End Select
    End With

    Forms!wizRestartCaseList1.AllowEdits = True
    Forms!wizRestartCaseList1.Detail.BackColor = -2147483633
    Forms!wizRestartCaseList1.Page97.Visible = False
    
DoCmd.OpenForm stDocName, , , stLinkCriteria
Forms!wizRestartFCdetails1!cmdReturn.Caption = "Remove from Queue"
Forms!wizRestartFCdetails1!WizMag = "Mgr"
Forms!wizRestartFCdetails1!cmdSetDisposition.Enabled = True


    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms!wizRestartFCdetails1.SetFocus
Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub





Sub RestartCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queRestart", "Wizards", "queRestartWaiting" ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed


Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria

'added on 9/10/14

If DCount("*", "CIVDetails", "FCFileNumber= " & FileNumber) > 0 Then
              'MsgBox "CAUTION! Litigation in progress, see an attorney!", vbExclamation
Forms!foreclosuredetails.Detail.BackColor = vbYellow
End If

Forms!foreclosuredetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!pageAccounting.Visible = True
Forms!foreclosuredetails!WizardSource = "Restart"
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call RestartFieldsVisible(True)
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber

'added on 9/10/14

If DCount("*", "CIVDetails", "FCFileNumber= " & FileNumber) > 0 Then
              'MsgBox "CAUTION! Litigation in progress, see an attorney!", vbExclamation
Forms!DocsWindow.Detail.BackColor = vbYellow

End If
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub

Sub TitleOrderCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queTitelOrder", "Wizards" ', "queRestartWaiting"  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed


Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "Title"
'Forms!ForeclosureDetails!cmdWizComplete.Visible = True
'Forms!ForeclosureDetails!pageAccounting.Visible = True
Forms!foreclosuredetails!WizardSource = "Title"
Forms!foreclosuredetails!TitleThru.Locked = False

    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call TitleFieldsVisible(True)
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub
Sub IntakeCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Intake

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queIntake", "Wizards", "queIntakeWaiting" ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber


Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria


'added on 9/10/14

If DCount("*", "CIVDetails", "FCFileNumber= " & FileNumber) > 0 Then
              'MsgBox "CAUTION! Litigation in progress, see an attorney!", vbExclamation
              Forms!foreclosuredetails.Detail.BackColor = vbYellow
End If


Forms!foreclosuredetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!cmdWaiting.Caption = "Send to Waiting"
Forms!foreclosuredetails!DocstoClient.Locked = True
'Forms!foreclosuredetails![post-sale].AllowEdits = True
'Forms!foreclosuredetails!pageAccounting.Visible = True
Forms!foreclosuredetails!WizardSource = "Intake"
Forms!foreclosuredetails!cmdWaiting1.Visible = True
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call IntakeFieldsVisible(True)
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Exit_Intake:
    Exit Sub

Err_Intake:
    MsgBox Err.Description
    Resume Exit_Intake
    
End Sub
Sub DocketingCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Docketing

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queDocketing", "Wizards", "queDocketingWaiting" ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria
Forms!foreclosuredetails!cmdWaiting1.Visible = True
Forms!foreclosuredetails!cmdWaiting.Caption = "Send to Waiting"
Forms!foreclosuredetails!WizardSource = "Docketing"
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call DocketingFieldsVisible(True)
'DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Exit_Docketing:
    Exit Sub

Err_Docketing:
    MsgBox Err.Description
    Resume Exit_Docketing
    
End Sub
Sub DocketingWaitingCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean, rstwizqueue As Recordset
On Error GoTo Err_Docketing

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queDocketing", "Wizards", "queDocketingWaiting" ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria
Call DocketingFieldsVisible(True)
Set rstwizqueue = CurrentDb.OpenRecordset("SELECT AttyMileStone4 FROM wizardqueuestats WHERE FileNumber=" & FileNumber & " AND current=true", dbOpenSnapshot)
If Not IsNull(rstwizqueue!AttyMilestone4) Then Forms!foreclosuredetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!cmdWaiting.Caption = "Send to Waiting"
Forms!foreclosuredetails!cmdWaiting1.Visible = True
Forms!foreclosuredetails!cmdWaiting1.Caption = "Return to Atty Review"
If Forms!queDocketingWaiting!lstFiles.Column(9) = "2. Rejected" Then
Forms!foreclosuredetails!cmdWizComplete.Visible = False
End If


'Forms!ForeclosureDetails!DocstoClient.Locked = True
'Forms!foreclosuredetails![post-sale].AllowEdits = True
'Forms!foreclosuredetails!pageAccounting.Visible = True
Forms!foreclosuredetails!WizardSource = "Docketing"
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If

DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Exit_Docketing:
    Exit Sub

Err_Docketing:
    MsgBox Err.Description
    Resume Exit_Docketing
    
End Sub
Sub FLMACallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Docketing

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queFLMA", "Wizards" ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call FLMAFieldsVisible(True)
Forms!foreclosuredetails!WizardSource = "FLMA"
'DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Forms![Case List]!SCRAID = "FLMA"
Exit_Docketing:
    Exit Sub

Err_Docketing:
    MsgBox Err.Description
    Resume Exit_Docketing
    
End Sub
Sub ServiceCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Docketing

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queService", "Wizards" ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call ServiceFieldsVisible(True)
Forms!foreclosuredetails!WizardSource = "Service"
'DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Forms![Case List]!SCRAID = "Service"
Exit_Docketing:
    Exit Sub

Err_Docketing:
    MsgBox Err.Description
    Resume Exit_Docketing
    
End Sub
Sub ServiceMailedCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Docketing

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queServiceMailed", "Wizards" ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call ServiceMailedFieldsVisible(True)
Forms!foreclosuredetails!WizardSource = "ServiceMailed"
'DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Forms![Case List]!SCRAID = "ServiceMailed"
Exit_Docketing:
    Exit Sub

Err_Docketing:
    MsgBox Err.Description
    Resume Exit_Docketing
    
End Sub
Sub ServiceMailedFieldsVisible(SetVisible As Boolean)

Forms!foreclosuredetails!WizardSource = "ServiceMailed"
Forms!foreclosuredetails!ServiceMailed.Locked = False
Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWizComplete.Enabled = False
Forms!foreclosuredetails.cmdWaiting.Visible = False
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
Forms![foreclosuredetails]!Page256.Visible = False
Forms![foreclosuredetails]!Trustees.Visible = False
Forms![foreclosuredetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False

End Sub
Public Sub ServiceMailedCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!ServiceMailedComplete) Then
With rstqueue
.Edit
!ServiceMailedComplete = Now
!ServiceMailedUser = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing

If CurrentProject.AllForms("queServiceMailed").IsLoaded = True Then
Forms!queServiceMailed!lstFiles.Requery
Forms!queServiceMailed.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueServiceMailed", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queServiceMailed!QueueCount = cntr
Set rstqueue = Nothing
End If
  
End Sub
Sub SaleSettingCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_SaleSetting

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queSaleSetting", "queSaleSettingwaiting", "Wizards" ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call SaleSettingFieldsVisible(True)
Forms!foreclosuredetails!WizardSource = "SaleSetting"
'DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
'Forms![case list]!SCRAID = "SaleSetting"
Exit_SaleSetting:
    Exit Sub

Err_SaleSetting:
    MsgBox Err.Description
    Resume Exit_SaleSetting
    
End Sub
Sub SaleSettingFieldsVisible(SetVisible As Boolean)
Forms!foreclosuredetails!cmdWaiting.Caption = "Sale Waiting"
Forms!foreclosuredetails!WizardSource = "SaleSetting"
Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWizComplete.Enabled = False
Forms!foreclosuredetails.Sale.Locked = False
Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
'Forms![ForeClosureDetails]!Page256.Visible = False
Forms![foreclosuredetails]!Trustees.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False
End Sub
Public Sub SaleSettingCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!SaleSettingComplete) Then
With rstqueue
.Edit
!SaleSettingComplete = Now
!SaleSettingUser = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing

If CurrentProject.AllForms("queSaleSetting").IsLoaded = True Then
Forms!queSaleSetting!lstFiles.Requery
Forms!queSaleSetting.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueSaleSetting", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queSaleSetting!QueueCount = cntr
Set rstqueue = Nothing
End If
  
If CurrentProject.AllForms("queSaleSettingwaiting").IsLoaded = True Then
Forms!queSaleSettingwaiting!lstFiles.Requery
Forms!queSaleSettingwaiting.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueSaleSettingwaiting", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queSaleSettingwaiting!QueueCount = cntr
Set rstqueue = Nothing
End If

End Sub
Sub VAsalesettingCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_VAsalesetting

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queVAsalesetting", "queVAsalesettingsub", "queVAsalesettingwaiting", "Wizards" ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call VAsalesettingFieldsVisible(True)
If CurrentProject.AllForms("queVAsalesetting").IsLoaded Then Forms!foreclosuredetails.cmdPrint.Enabled = False
Forms!foreclosuredetails!WizardSource = "VAsalesetting"

If CurrentProject.AllForms("queVAsalesettingsub").IsLoaded Then Forms!foreclosuredetails!WizardSource = "VALNNSetting"

Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False

Exit_VAsalesetting:
    Exit Sub

Err_VAsalesetting:
    MsgBox Err.Description
    Resume Exit_VAsalesetting
    
End Sub
Sub VAsalesettingFieldsVisible(SetVisible As Boolean)
Forms!foreclosuredetails!cmdWaiting.Caption = "Sale Waiting"
Forms!foreclosuredetails!PrimaryFirstName.SetFocus
Forms!foreclosuredetails!WizardSource = "VAsalesetting"
Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails!cmdWaiting1.Visible = True
Forms!foreclosuredetails.cmdWizComplete.Enabled = False
Forms!foreclosuredetails.Sale.Locked = False
Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
Forms![foreclosuredetails]!pgNOI.Visible = False
Forms![foreclosuredetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!Page256.Visible = True
'Forms![ForeclosureDetails]!Trustees.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False




End Sub
Public Sub VAsalesettingCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!VASaleSettingComplete) Then
With rstqueue
.Edit
!VASaleSettingComplete = Now
!VASaleSettingUser = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing

If CurrentProject.AllForms("queVAsalesetting").IsLoaded = True Then
Forms!quevasalesetting!lstFiles.Requery
Forms!quevasalesetting.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueVAsalesettinggroupby", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!quevasalesetting!QueueCount = cntr
Set rstqueue = Nothing
End If
  
If CurrentProject.AllForms("queVAsalesettingwaiting").IsLoaded = True Then
Forms!quevasalesettingwaiting!lstFiles.Requery
Forms!quevasalesettingwaiting.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueVAsalesettingwaitinglst", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!quevasalesettingwaiting!QueueCount = cntr
Set rstqueue = Nothing
End If

End Sub

Public Sub VALNNSetting(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
If IsNull(Forms!foreclosuredetails!LostNoteNotice) Then
MsgBox ("The VA LNN wazard is not complete. Add Lost Note Notice date")
Exit Sub
Else
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)

'If IsNull(rstqueue!VALNNComplete) Then
With rstqueue
.Edit
!VALNNComplete = Now
!VALNNUser = GetStaffID
.Update
End With
'End If
Set rstqueue = Nothing

If CurrentProject.AllForms("queVAsalesettingSub").IsLoaded = True Then
Forms!queVAsalesettingSub!lstFiles.Requery
Forms!queVAsalesettingSub.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuevasalesettingred", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queVAsalesettingSub!QueueCount = cntr
Set rstqueue = Nothing
End If
End If
'If CurrentProject.AllForms("queVAsalesettingwaiting").IsLoaded = True Then
'Forms!quevasalesettingwaiting!lstFiles.Requery
'Forms!quevasalesettingwaiting.Requery
'
'Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuevasalesettingred", dbOpenDynaset, dbSeeChanges)
'Do Until rstqueue.EOF
'cntr = cntr + 1
'rstqueue.MoveNext
'Loop
'Forms!quevasalesettingwaiting!QueueCount = cntr
'Set rstqueue = Nothing
'End If
'
End Sub


Public Sub VAsalesettingWaitingCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
!VASaleSettingWaiting = Now
!VASaleSettingWaitingUser = GetStaffID
!VASaleSettingDocsRecdFlag = False
!AttyMilestone3Reject = False
!VASaleSettingReason = 4
.Update
End With

Set rstqueue = Nothing


'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Warning] = 100
'  ![Info] = "Send to VA Sale Setting waiting queue" & vbCrLf
'  ![Color] = 1
'  .Update
'  End With
  
  DoCmd.SetWarnings False
strinfo = "Sent to VA Sale Setting waiting queue" & vbCrLf
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Warning,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),100,'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

MsgBox "File " & FileNumber & " has been added to the VA Sale Setting Waiting queue", , "VA SaleSetting Wizard"
DoCmd.Close acForm, "EntervasalesettingDocs"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

If CurrentProject.AllForms("quevasalesetting").IsLoaded = True Then
Forms!quevasalesetting!lstFiles.Requery
Forms!quevasalesetting.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuevasalesettinggroupby", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!quevasalesetting!QueueCount = cntr
Set rstqueue = Nothing
End If
If CurrentProject.AllForms("quevasalesettingwaiting").IsLoaded = True Then
Forms!quevasalesettingwaiting!lstFiles.Requery
Forms!quevasalesettingwaiting.Requery
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuevasalesettingwaitinglst", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!quevasalesettingwaiting!QueueCount = cntr
Set rstqueue = Nothing
End If
End Sub
Sub BorrowerServedCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_BorrowerServed

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queBorrowerServed", "Wizards" ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call BorrowerServedFieldsVisible(True)
Forms!foreclosuredetails!WizardSource = "BorrowerServed"
'DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Forms![Case List]!SCRAID = "BorrowerServed"
Exit_BorrowerServed:
    Exit Sub

Err_BorrowerServed:
    MsgBox Err.Description
    Resume Exit_BorrowerServed
    
End Sub
Sub BorrowerServedFieldsVisible(SetVisible As Boolean)
Forms!foreclosuredetails!cmdWaiting.Caption = "Waiting"
Forms!foreclosuredetails!WizardSource = "BorrowerServed"
Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWizComplete.Enabled = False
Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms!foreclosuredetails.BorrowerServed.Locked = False
Forms![foreclosuredetails]!Page412.Visible = False
Forms![foreclosuredetails]!Page256.Visible = False
Forms![foreclosuredetails]!Trustees.Visible = False
Forms![foreclosuredetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False

End Sub
Public Sub TitleOrderCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
DoCmd.Hourglass True
Dim strSearchType As String
Dim rstFCdetails As Recordset
Dim rstAcer As Recordset
Dim rstTitleOrders As Recordset
Dim statusMsg As String
Dim t As Recordset
Dim s As Recordset
Dim lrs As Recordset
Dim Jrs As Recordset


Select Case Forms![Print Title Order]!Order
    Case 1
       statusMsg = "Ordered full title search"
       strSearchType = "Full"
    Case 2
       statusMsg = "Ordered title rundown from " & Format$(Forms![Print Title Order]!RundownDate, "m/d/yyyy")
       strSearchType = "Update"
    Case 3
       statusMsg = "Ordered title rundown from present owner"
       strSearchType = "Rundown"
    Case 4
       statusMsg = "Ordered 2 owner search"
       strSearchType = "2 Owner"

End Select


If Forms![Print Title Order]!chForeclosure Then
   ' If MsgBox("Update Title Ordered = " & Format$(Date, "m/d/yyyy") & vbNewLine & "and clear Title Received and Title Through" & vbNewLine & "and add to status?", vbYesNo) = vbYes Then
        Forms![foreclosuredetails]!TitleOrder = Now()
        Forms![foreclosuredetails]!TitleDue = Forms![Print Title Order]!DateRequired
        If Not IsNull(Forms![foreclosuredetails]!TitleBack) Then Forms![foreclosuredetails]!TitleBack = Null
        If Not IsNull(Forms![foreclosuredetails]!TitleThru) Then Forms![foreclosuredetails]!TitleThru = Null
        If Not IsNull(Forms![foreclosuredetails]!TitleReviewToClient) Then Forms![foreclosuredetails]!TitleReviewToClient = Null
        AddStatus FileNumber, Now(), statusMsg



If Forms![Print Title Order]!Abstractor = 89 Then 'Acer
Dim Last As String, First As String, Address As String, City As String, State As String


Set rstFCdetails = CurrentDb.OpenRecordset("SELECT * FROM fcdetails WHERE FileNumber=" & Forms![Print Title Order]!FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
With rstFCdetails
Last = !PrimaryLastName
First = !PrimaryFirstName
Address = !PropertyAddress
City = !City
State = !State

.Close
End With

Set rstAcer = CurrentDb.OpenRecordset("SELECT * FROM RA_Title_Ordered", dbOpenDynaset, dbSeeChanges)

With rstAcer
.AddNew
!RA_Number = Forms![Print Title Order]!FileNumber
!Last_Name = Last
!First_Name = First
!Address = Address
!City = City
!State = State
!DueDate = Forms![Print Title Order]!DateRequired
!Jurisdiction = DLookup("jurisdictionid", "caseList", "filenumber=" & Forms![Print Title Order]!FileNumber)
!Client = DLookup("clientid", "caseList", "filenumber=" & Forms![Print Title Order]!FileNumber)
!RequestDate = Now
.Update
.Close
End With
End If

Set rstTitleOrders = CurrentDb.OpenRecordset("SELECT * FROM TitleOrders", dbOpenDynaset, dbSeeChanges)

With rstTitleOrders
.AddNew
!FileNumber = Forms![Print Title Order]!FileNumber
!Abstractor = Forms![Print Title Order]!Abstractor
!DateOrdered = Now
!OrderedBy = GetStaffID
.Update
.Close
End With

'End If
End If

'NoValid
Forms![foreclosuredetails]!TitleSearchType = strSearchType
'cmdCancel.Caption = "Close"

'DoCmd.Close "Print Title Order"
'DoCmd.Close acForm, "Print Title Order"

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
'If IsNull(rstqueue!TitleCompleted) Then
With rstqueue
.Edit
If Not IsNull(rstqueue!TitleOutDateInQueue) Then rstqueue!TitleOutDateInQueue = Null
If Not IsNull(rstqueue!TitleRevewDateInQueue) Then rstqueue!TitleRevewDateInQueue = Null

!TitleCompleted = Now
!TitleCompUser = GetStaffID
!TitleLastEditDate = Now
!TitleLastEditUser = GetStaffID
!TitelMissingReson = False
.Update
End With
'End If
Set rstqueue = Nothing


If IsLoadedF("queTitelOrder") = True Then
    If FileNumber = Forms!queTitelOrder!lstFiles.Column(0) Then
    
    Set t = CurrentDb.OpenRecordset("TitleOrderHistory", dbOpenDynaset, dbSeeChanges)
    t.AddNew
    t!CaseFile = FileNumber
    t!Client = Forms!queTitelOrder!lstFiles.Column(2)
    t!Name = GetFullName()
    t!TitleOrder = Now
    t!TitleOrderC = 1
    t!DataDisiminated = True
    t!Stage = Forms!queTitelOrder!lstFiles.Column(6)
    t!Days = Forms!queTitelOrder!lstFiles.Column(8)
    t!dateOfStage = Forms!queTitelOrder!lstFiles.Column(7)
    t.Update
    Set t = Nothing
    
    Else
    
    
    Set s = CurrentDb.OpenRecordset("TitleOrderHistory", dbOpenDynaset, dbSeeChanges)
    s.AddNew
    s!CaseFile = FileNumber
    s!Client = ClientShortName(Forms![Case List]!ClientID)
    s!Name = GetFullName()
    s!TitleOrder = Now
    s!TitleOrderC = 1
    s!DataDisiminated = True
    s!Stage = "Hard Order"
    s!Days = 0
    s!UpdateFromDM = 0
    s!dateOfStage = Date
    s.Update
    
    
    Set s = Nothing
    End If
Else
Set s = CurrentDb.OpenRecordset("TitleOrderHistory", dbOpenDynaset, dbSeeChanges)
    s.AddNew
    s!CaseFile = FileNumber
    s!Client = ClientShortName(Forms![Case List]!ClientID)
    s!Name = GetFullName()
    s!TitleOrder = Now
    s!TitleOrderC = 1
    s!DataDisiminated = True
    s!Stage = "Hard Order"
    s!Days = 0
    s!UpdateFromDM = 0
    s!dateOfStage = Date
    s.Update
    
    
    Set s = Nothing
End If


'Dim DM As Recordset
'Set DM = CurrentDb.OpenRecordset("Select * FROM TitleOrderHistory where CaseFile=" & FileNumber, dbOpenDynaset, dbSeeChanges)
'If Not DM.EOF Then
'DM.MoveFirst
'Do While Not DM.EOF
'With DM
'.Edit
'!UpdateFromDM = 0
'.Update
'End With
'DM.MoveNext
'Loop
'End If
'Set DM = Nothing


If CurrentProject.AllForms("queTitelOrder").IsLoaded = True Then
    If FileNumber = Forms!queTitelOrder!lstFiles.Column(0) Then
    
'      Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'      With lrs
'      .AddNew
'      ![FileNumber] = FileNumber
'      ![JournalDate] = Now
'      ![Who] = GetFullName()
'      ![Info] = "Title Ordered " & " - For " & Forms!queTitelOrder!lstFiles.Column(6) & vbCrLf
'      ![Color] = 1
'      .Update
'      .Close
'      End With
      DoCmd.SetWarnings False
strinfo = "Title Ordered " & " - For " & Forms!queTitelOrder!lstFiles.Column(6) & vbCrLf
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
    Else
    
    
'      Set Jrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'      Jrs.AddNew
'      Jrs![FileNumber] = FileNumber
'      Jrs![JournalDate] = Now
'      Jrs![Who] = GetFullName()
'      Jrs![Info] = "Title Ordered " & " - For " & " Hard Order" & vbCrLf
'      Jrs![Color] = 1
'      Jrs.Update
'
'
'    Jrs.Close
    DoCmd.SetWarnings False
strinfo = "Title Ordered " & " - For " & " Hard Order" & vbCrLf
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True
    End If
Else
'Set Jrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'Jrs.AddNew
'Jrs![FileNumber] = FileNumber
'Jrs![JournalDate] = Now
'Jrs![Who] = GetFullName()
'Jrs![Info] = "Title Ordered " & " - For " & " Hard Order" & vbCrLf
'Jrs![Color] = 1
'Jrs.Update
'Jrs.Close

DoCmd.SetWarnings False
strinfo = "Title Ordered " & " - For " & " Hard Order" & vbCrLf
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

End If


AddStatus FileNumber, Date, "Title Ordered"

DoCmd.SetWarnings False
'Dim DeletsSQl As Recordset
''heresarabd
DoCmd.RunSQL "DELETE * FROM TitleOrderFinal WHERE File=" & FileNumber

'Set DeletsSQl = CurrentDb.OpenRecordset("Select * FROM  qryTitle_AllT  Where File =" & FileNumber, dbOpenDynaset, dbSeeChanges)
'If Not DeletsSQl.EOF Then
'DeletsSQl.Delete
'End If
'Set DeletsSQl = Nothing
DoCmd.SetWarnings True


If CurrentProject.AllForms("queTitelOrder").IsLoaded = True Then

'Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryTitle_All", dbOpenDynaset, dbSeeChanges)
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryTitle_All", dbOpenDynaset, dbSeeChanges)
If rstqueue.EOF Then
    cntr = 0
    Else
    rstqueue.MoveLast
    cntr = rstqueue.RecordCount
End If
Set rstqueue = Nothing




'Do Until rstqueue.EOF
'cntr = cntr + 1
'rstqueue.MoveNext
'Loop
'Forms!queTitelOrder!QueueCount = cntr
'Set rstqueue = Nothing
End If


DoCmd.Hourglass False

'MsgBox "   Completed    "
'DoCmd.Hourglass True
If CurrentProject.AllForms("queTitelOrder").IsLoaded = True Then
Forms!queTitelOrder!lstFiles.Requery
Forms!queTitelOrder.Requery
'Forms!queTitelOrder!lstFiles = Null
DoCmd.Hourglass False
End If
DoCmd.Close acForm, "Print Title Order"


  
End Sub

Public Sub BorrowerServedCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!BorrowerServedComplete) Then
With rstqueue
.Edit
!BorrowerServedComplete = Now
!BorrowerServedUser = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing

If CurrentProject.AllForms("queBorrowerServed").IsLoaded = True Then
Forms!queBorrowerServed!lstFiles.Requery
Forms!queBorrowerServed.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueBorrowerServed", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queBorrowerServed!QueueCount = cntr
Set rstqueue = Nothing
End If
  
End Sub



Sub HUDOccCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "Wizards" ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
'Forms![case list].Visible = False
DoCmd.OpenForm stDocName, , , stLinkCriteria
Forms!foreclosuredetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!cmdWizComplete.Enabled = False
Forms!foreclosuredetails!pgNOI.Visible = False
Forms!foreclosuredetails!pgMediation.Visible = False
Forms!foreclosuredetails!pgRealPropTaxes.Visible = False
Forms!foreclosuredetails!pageAccounting.Visible = False
Forms!foreclosuredetails![Post-Sale].Visible = False
Forms!foreclosuredetails.cmdcloserestart.Visible = True
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms!foreclosuredetails!WizardSource = "HUDocc"
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber

Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub
Sub VAappraisalCallFromQueue(FileNumber As Long)

''Cannot add formclosed check to this process, requires case list and FC details to be open

Forms![Case List].SCRAID = 13

Forms!foreclosuredetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!pgNOI.Visible = False
Forms!foreclosuredetails!pgMediation.Visible = False
Forms!foreclosuredetails!pgRealPropTaxes.Visible = False
Forms!foreclosuredetails!pageAccounting.Visible = False
Forms!foreclosuredetails![Post-Sale].Visible = False
Forms!foreclosuredetails.cmdcloserestart.Visible = True
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms!foreclosuredetails!WizardSource = "VAappraisal"


Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub
Sub RestartWaitingCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "Wizards" ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed
stDocName = "wizRestartFCdetails1"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"

Set rstfiles = CurrentDb.OpenRecordset("SELECT FileNumber FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
FileLocks = True
    If LockFile(FileNumber) Then
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If

rstfiles.Close

Call Restart1FieldsVisible(True)

DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber

Exit_Restart:
    Exit Sub
Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub
Sub RSIILoanType_AfterUpdate()
Dim lt As Integer

If IsNull(Forms!wizreferralII.LoanType) Then
    lt = 0
Else
    lt = Forms!wizreferralII.LoanType
End If
Forms!wizreferralII.FHALoanNumber.Enabled = (lt = 2 Or lt = 3)    ' enable for VA or HUD
Forms!wizreferralII.FNMALoanNumber.Enabled = (lt = 4)
Forms!wizreferralII.FHLMCLoanNumber.Enabled = (lt = 5)

End Sub
Sub RSIIConfirmationVisible(SetVisible As Boolean)

Forms!wizreferralII.PrimaryDefName.Visible = SetVisible
Forms!wizreferralII.PropertyAddress.Visible = SetVisible
Forms!wizreferralII.Apt.Visible = SetVisible

'Forms!wizreferralII.txtJurisdiction.Visible = SetVisible
'Forms!wizreferralII.LongClientName.Visible = SetVisible
'Forms!wizreferralII.cmdYes.Enabled = SetVisible
'Forms!wizreferralII.cmdNo.Enabled = SetVisible

End Sub
Sub FLMAFieldsVisible(SetVisible As Boolean)
Forms!foreclosuredetails!cmdWaiting.Caption = "FLMA Incomplete"
Forms!foreclosuredetails!cmdWaiting.FontBold = True
Forms!foreclosuredetails!cmdWaiting.ForeColor = vbBlack

Forms!foreclosuredetails!WizardSource = "FLMA"
Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWizComplete.Enabled = False
Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
Forms![foreclosuredetails]!Page256.Visible = False
Forms![foreclosuredetails]!Trustees.Visible = False
Forms![foreclosuredetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False

End Sub
Sub ServiceFieldsVisible(SetVisible As Boolean)
Forms!foreclosuredetails!cmdWaiting.Caption = "Service Incomplete"
Forms!foreclosuredetails!cmdWaiting.FontBold = True
Forms!foreclosuredetails!cmdWaiting.ForeColor = vbBlack

Forms!foreclosuredetails!WizardSource = "Service"
Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWizComplete.Enabled = False
Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
Forms![foreclosuredetails]!Page256.Visible = False
Forms![foreclosuredetails]!Trustees.Visible = False
Forms![foreclosuredetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False
End Sub
Sub NOIConfirmationVisible(SetVisible As Boolean)

Forms!wizNOI.PrimaryDefName.Visible = SetVisible
Forms!wizNOI.PropertyAddress.Visible = SetVisible
Forms!wizNOI.txtJurisdiction.Visible = SetVisible
Forms!wizNOI.LongClientName.Visible = SetVisible


End Sub
Sub FairDebtConfirmationVisible(SetVisible As Boolean)

Forms!wizfairdebt.PrimaryDefName.Visible = SetVisible
Forms!wizfairdebt.PropertyAddress.Visible = SetVisible
Forms!wizfairdebt.txtJurisdiction.Visible = SetVisible
Forms!wizfairdebt.LongClientName.Visible = SetVisible

End Sub
Sub TitleFieldsVisible(SetVisible As Boolean)

'Forms!ForeclosureDetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
'Forms!ForeclosureDetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
Forms![foreclosuredetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False
Forms![foreclosuredetails]!Trustees.Visible = False
Forms![foreclosuredetails]!pgNOI.Visible = False
Forms![foreclosuredetails]![Post-Sale].Visible = False
Forms![foreclosuredetails]!pageAccounting.Visible = False
Forms!foreclosuredetails!cmdPrint.Visible = False
Forms![Case List]!Page97.Visible = False



End Sub


Sub RestartFieldsVisible(SetVisible As Boolean)

Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
'Forms![ForeclosureDetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False
Forms!foreclosuredetails.cmdWaiting1.Visible = False


End Sub
Sub IntakeFieldsVisible(SetVisible As Boolean)

Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
Forms![foreclosuredetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False

End Sub
Sub DocketingFieldsVisible(SetVisible As Boolean)

Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
Forms![foreclosuredetails]!Page256.Visible = False
Forms![foreclosuredetails]!Trustees.Visible = False
Forms![foreclosuredetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False
Forms!foreclosuredetails!cmdWizComplete.Visible = False

Forms!foreclosuredetails!SentToDocket.Enabled = False
End Sub
Sub Restart1FieldsVisible(SetVisible As Boolean)

Forms!wizRestartFCdetails1.cmdWizComplete.Visible = SetVisible
Forms!wizRestartFCdetails1.cmdClose.Visible = False
Forms![wizRestartFCdetails1]![Post-Sale].Visible = False
Forms![wizRestartFCdetails1]!pgNOI.Visible = False
Forms![wizRestartFCdetails1]!Trustees.Visible = False
Forms![wizRestartFCdetails1]![Pre-Sale].Visible = False



End Sub

Sub RSIIFieldsVisible(SetVisible As Boolean)

Forms!wizreferralII.tabWiz.Visible = SetVisible
Forms!wizreferralII.cmdOK.Enabled = SetVisible
Forms!wizreferralII!AssessedValue.Enabled = (UCase$(Nz(Forms!wizreferralII.State)) = "VA")

If Not IsNull(Forms!wizreferralII.txtDisposition) Then
Forms!wizreferralII.lblDisposition.Visible = True
Forms!wizreferralII.lblDisposition.Caption = "The file has a disposition of:  " & DLookup("Disposition", "FCDisposition", "ID=" & Forms!wizreferralII.txtDisposition)
End If

Select Case Forms!wizreferralII.LoanType.Column(0)

Case "1"

Forms!wizreferralII.lblFHAVA.Visible = False
Forms!wizreferralII.FHALoanNumber.Visible = False
Forms!wizreferralII.lblFNMA.Visible = False
Forms!wizreferralII.FNMALoanNumber.Visible = False
Forms!wizreferralII.lblFHLMC.Visible = False
Forms!wizreferralII.FHLMCLoanNumber.Visible = False
Case "2"

Forms!wizreferralII.lblFHAVA.Visible = True
Forms!wizreferralII.FHALoanNumber.Enabled = True
Forms!wizreferralII.FHALoanNumber.Visible = True
Forms!wizreferralII.lblFNMA.Visible = False
Forms!wizreferralII.FNMALoanNumber.Visible = False
Forms!wizreferralII.lblFHLMC.Visible = False
Forms!wizreferralII.FHLMCLoanNumber.Visible = False
Case "3"
Forms!wizreferralII.lblFHAVA.Visible = True
Forms!wizreferralII.FHALoanNumber.Enabled = True
Forms!wizreferralII.FHALoanNumber.Visible = True
Forms!wizreferralII.lblFNMA.Visible = False
Forms!wizreferralII.FNMALoanNumber.Visible = False
Forms!wizreferralII.lblFHLMC.Visible = False
Forms!wizreferralII.FHLMCLoanNumber.Visible = False
Case "4"
Forms!wizreferralII.lblFNMA.Visible = True
Forms!wizreferralII.FNMALoanNumber.Enabled = True
Forms!wizreferralII.FNMALoanNumber.Visible = True
Forms!wizreferralII.lblFHLMC.Visible = False
Forms!wizreferralII.FHLMCLoanNumber.Visible = False
Forms!wizreferralII.lblFHAVA.Visible = False
Forms!wizreferralII.FHALoanNumber.Visible = False
Case "5"
Forms!wizreferralII.lblFHLMC.Visible = True
Forms!wizreferralII.FHLMCLoanNumber.Enabled = True
Forms!wizreferralII.FHLMCLoanNumber.Visible = True
Forms!wizreferralII.lblFHAVA.Visible = False
Forms!wizreferralII.FHALoanNumber.Visible = False
Forms!wizreferralII.lblFNMA.Visible = False
Forms!wizreferralII.FNMALoanNumber.Visible = False
End Select



If SetVisible Then
    DoCmd.OpenForm "Journal", , , "FileNumber=" & Forms!wizreferralII.FileNumber
Else
    DoCmd.Close acForm, "Journal"
End If

Forms!wizreferralII.SetFocus
End Sub

Sub NOIFieldsVisible(SetVisible As Boolean)

Forms!wizNOI.tabWiz.Visible = SetVisible
Forms!wizNOI.cmdOK.Enabled = SetVisible

If Not IsNull(Forms!wizNOI.txtDisposition) Then
Forms!wizNOI.lblDisposition.Visible = True
Forms!wizNOI.lblDisposition.Caption = "The file has a disposition of:  " & DLookup("Disposition", "FCDisposition", "ID=" & Forms!wizNOI.txtDisposition)
End If

If SetVisible Then
    DoCmd.OpenForm "Journal", , , "FileNumber=" & Forms!wizNOI.txtFileNumber
Else
    DoCmd.Close acForm, "Journal"
End If

Forms!wizNOI.SetFocus
End Sub
Sub FairDebtFieldsVisible(SetVisible As Boolean)

Forms!wizfairdebt.tabWiz.Visible = SetVisible
Forms!wizfairdebt.cmdOK.Enabled = SetVisible

If Not IsNull(Forms!wizfairdebt.txtDisposition) Then
Forms!wizfairdebt.lblDisposition.Visible = True
Forms!wizfairdebt.lblDisposition.Caption = "The file has a disposition of:  " & DLookup("Disposition", "FCDisposition", "ID=" & Forms!wizfairdebt.txtDisposition)
End If

If SetVisible Then
    DoCmd.OpenForm "Journal", , , "FileNumber=" & Forms!wizfairdebt.txtFileNumber
Else
    DoCmd.Close acForm, "Journal"
End If
Forms!wizfairdebt.SetFocus
End Sub
Public Sub RestartCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!RestartComplete) Then
With rstqueue
.Edit
!RestartComplete = Now
!RestartUser = StaffID
.Update
End With
End If
Set rstqueue = Nothing
If CurrentProject.AllForms("querestart").IsLoaded = True Then
Forms!querestart!lstFiles.Requery
Forms!querestart.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuerestart", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!querestart!QueueCount = cntr
Set rstqueue = Nothing
End If

If CurrentProject.AllForms("querestartwaiting").IsLoaded = True Then
Forms!querestartWaiting!lstFiles.Requery
Forms!querestartWaiting.Requery
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuerestartwaiting", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!querestartWaiting!QueueCount = cntr
Set rstqueue = Nothing
End If

End Sub
Public Sub FLMACompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!FLMAComplete) Then
With rstqueue
.Edit
!FLMAComplete = Now
!FLMAUser = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing

Dim lrs As Recordset

DoCmd.SetWarnings False
strinfo = "FLMA sent to court.  Tracking number " & InputBox("Please enter the overnight tracking #")
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

'2/11/14
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Info] = "FLMA sent to court.  Tracking number " & InputBox("Please enter the overnight tracking #") & vbCrLf
'  ![Color] = 1
'  .Update
'  .Close
'  End With
AddStatus FileNumber, Date, "FLMA sent to court"
If CurrentProject.AllForms("queFLMA").IsLoaded = True Then
Forms!queFLMA!lstFiles.Requery
Forms!queFLMA.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueFLMA", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queFLMA!QueueCount = cntr
Set rstqueue = Nothing
End If
  
End Sub
Public Sub ServiceCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!ServiceComplete) Then
With rstqueue
.Edit
!ServiceComplete = Now
!ServiceUser = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing

'2/11/14
DoCmd.SetWarnings False
strinfo = "Proof of Service sent"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True



'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Info] = "Proof of Service sent" & vbCrLf
'  ![Color] = 1
'  .Update
'  .Close
'  End With

AddStatus FileNumber, Date, "Proof of Service sent"
If CurrentProject.AllForms("queService").IsLoaded = True Then
Forms!queService!lstFiles.Requery
Forms!queService.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueService", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queService!QueueCount = cntr
Set rstqueue = Nothing
End If
  
End Sub
Public Sub FLMAWaitingCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
'Set rstQueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
'With rstQueue
'.Edit
'!DocketingWaiting = Now
'!DocketingWaitingby = GetStaffID
'!DocketingDocsRecdFlag = False
'!AttyMilestone4Reject = False
'.Update
'End With

Set rstqueue = Nothing

MsgBox "File " & FileNumber & " has been marked as incomplete", , "FLMA Wizard"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

If CurrentProject.AllForms("queFLMA").IsLoaded = True Then
Forms!queFLMA!lstFiles.Requery
Forms!queFLMA.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueFLMA", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queFLMA!QueueCount = cntr
Set rstqueue = Nothing
End If
End Sub
Public Sub IntakeCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Dim rstvalumeintake As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
'If IsNull(rstQueue!IntakeComplete) Then
    With rstqueue
    .Edit
    !IntakeComplete = Now
    !IntakeCompleteby = GetStaffID
    If IsNull(rstqueue!DateSentAttyIntake) Then rstqueue!DateSentAttyIntake = Null
    .Update
    End With
'End If
Set rstqueue = Nothing
If CurrentProject.AllForms("queintake").IsLoaded = True Then
Forms!queintake!lstFiles.Requery
Forms!queintake.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueintake", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queintake!QueueCount = cntr
Set rstqueue = Nothing
End If
If CurrentProject.AllForms("queintakewaiting").IsLoaded = True Then
Forms!queIntakeWaiting!lstFiles.Requery
Forms!queIntakeWaiting.Requery
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueintakewaitinglst", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queIntakeWaiting!QueueCount = cntr
Set rstqueue = Nothing
End If

Set rstvalumeintake = CurrentDb.OpenRecordset("Select * from ValumeIntake", dbOpenDynaset, dbSeeChanges)
With rstvalumeintake
.AddNew
!CaseFile = FileNumber
!Client = DLookup("ShortClientName", "ClientList", "ClientID = " & Forms![Case List]!ClientID)
!IntakeComplete = Now
!IntakeCompleteC = 1
!Name = GetFullName()
.Update
End With
Set rstvalumeintake = Nothing

'2/11/14
DoCmd.SetWarnings False
strinfo = "Intake Wizard Complete"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True



Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Info] = "Intake Wizard Complete" & vbCrLf
'  ![Color] = 1
'
'  .Update
'  .Close
'  End With
'  Set lrs = Nothing
  
  Set lrs = CurrentDb.OpenRecordset("select * from journal where filenumber=" & FileNumber & " AND warning=100", dbOpenDynaset, dbSeeChanges)
  With lrs
  Do Until .EOF
  .Edit
  ![Warning] = Null
  .Update
  .MoveNext
  Loop
  End With

End Sub
Public Sub DocketingCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!DocketingComplete) Then
With rstqueue
.Edit
!DocketingComplete = Now
!DocketingCompleteby = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing


'2/11/14

DoCmd.SetWarnings False
strinfo = "Docketing wizard complete.  Sent to atty review queue"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True


Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Info] = "Docketing wizard complete.  Sent to atty review queue" & vbCrLf
'  ![Color] = 1
'  .Update
'  .Close
'  End With
Set lrs = CurrentDb.OpenRecordset("select * from journal where filenumber=" & FileNumber & " AND warning=100", dbOpenDynaset, dbSeeChanges)
  With lrs
  Do Until .EOF
  .Edit
  ![Warning] = Null
  .Update
  .MoveNext
  Loop
  End With
  
  If CurrentProject.AllForms("queDocketing").IsLoaded = True Then
Forms!quedocketing!lstFiles.Requery
Forms!quedocketing.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueDocketinggroupby", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!quedocketing!QueueCount = cntr
Set rstqueue = Nothing
End If
If CurrentProject.AllForms("queDocketingwaiting").IsLoaded = True Then
Forms!queDocketingWaiting!lstFiles.Requery
Forms!queDocketingWaiting.Requery
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuedocketingwaitinglst", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queDocketingWaiting!QueueCount = cntr
Set rstqueue = Nothing
End If
  
End Sub
Public Sub DocketingAttyCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!DocketingAttyReview) Then
With rstqueue
.Edit
!DocketingAttyReview = Now
!AttyMilestone4 = Null
!DocketingWaiting = Now
!DocketingWaitingby = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing

'2/11/14

DoCmd.SetWarnings False
strinfo = "Docketing wizard:  Sent to atty review queue"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True


'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Info] = "Docketing wizard:  Sent to atty review queue" & vbCrLf
'  ![Color] = 1
'  .Update
'  .Close
'  End With

If CurrentProject.AllForms("queDocketing").IsLoaded = True Then
Forms!quedocketing!lstFiles.Requery
Forms!quedocketing.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueDocketinggroupby", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!quedocketing!QueueCount = cntr
Set rstqueue = Nothing
End If
If CurrentProject.AllForms("queDocketingwaiting").IsLoaded = True Then
Forms!queDocketingWaiting!lstFiles.Requery
Forms!queDocketingWaiting.Requery
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuedocketingwaitinglst", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queDocketingWaiting!QueueCount = cntr
Set rstqueue = Nothing
End If

End Sub
Public Sub VAsalesettingAttyCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!VASaleSettingAttyReview) Then
With rstqueue
.Edit
!VASaleSettingAttyReview = Now
!AttyMilestone3 = Null
!VASaleSettingWaiting = Now
!VASaleSettingWaitingUser = GetStaffID
!VASaleSettingReason = 3
.Update
End With
End If
Set rstqueue = Nothing


DoCmd.SetWarnings False
strinfo = "VA Sale Setting wizard:  Sent to atty review queue"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True



Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Info] = "VA Sale Setting wizard:  Sent to atty review queue" & vbCrLf
'  ![Color] = 1
'  .Update
'  .Close
'  End With

If CurrentProject.AllForms("quevasalesetting").IsLoaded = True Then
Forms!quevasalesetting!lstFiles.Requery
Forms!quevasalesetting.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuevasalesettinggroupby", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!quevasalesetting!QueueCount = cntr
Set rstqueue = Nothing
End If
If CurrentProject.AllForms("quevasalesettingwaiting").IsLoaded = True Then
Forms!quevasalesettingwaiting!lstFiles.Requery
Forms!quevasalesettingwaiting.Requery
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuevasalesettingwaitinglst", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!quevasalesettingwaiting!QueueCount = cntr
Set rstqueue = Nothing
End If

End Sub
Public Sub SAICompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!SAIcomplete) Then
With rstqueue
.Edit
!SAIcomplete = Now
!SAIcompleteby = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing

End Sub
Public Sub Restart1CompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, rstwiz As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!Current = False
.Update

End With
Set rstqueue = Nothing


Set rstwiz = CurrentDb.OpenRecordset("WizardQueueStats", dbOpenDynaset, dbSeeChanges)
With rstwiz
    .AddNew
    !FileNumber = FileNumber
    !Current = True
    .Update
    .Close
End With


Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
!RestartRSIComplete = Now
!RestartRSIUser = StaffID
!RestartQueue = Now
.Update
End With
Set rstqueue = Nothing

'---- second Wizard update
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM WizardSupportTwo where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!Current = False
.Update
End With
Set rstqueue = Nothing


Set rstwiz = CurrentDb.OpenRecordset("WizardSupportTwo", dbOpenDynaset, dbSeeChanges)
With rstwiz
    .AddNew
    !FileNumber = FileNumber
    !Current = True
    .Update
    .Close
End With








End Sub
Public Sub RestartRSICompletionUpdate(FileNumber As Long, Reason As Long)
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
'If IsNull(rstQueue!RestartRSIComplete) Then
With rstqueue
.Edit
!RestartRSIreviewDateIn = Now
!RestartRSIreviewUser = StaffID
!RestartRSIreviewReason = Reason
.Update
End With
'End If
Set rstqueue = Nothing
'DoCmd.OpenForm "wizIntakeRestart"
'Forms!wizIntakeRestart!txtFileNumber = FileNumber
End Sub

Public Sub RestartRSICompletionUpdateToPutInReviewQ(FileNumber As Long, Reason As Long) ' for reject bottom 05/15 ticket 836
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
'If IsNull(rstQueue!RestartRSIComplete) Then
With rstqueue
.Edit
!RestartRSIreviewDateIn = Now
!RestartRSIreviewUser = StaffID
!RestartRSIreviewReason = Reason
.Update
End With
'End If
Set rstqueue = Nothing
'DoCmd.OpenForm "wizIntakeRestart"
'Forms!wizIntakeRestart!txtFileNumber = FileNumber
End Sub
Public Sub RestartWaitingCompletionUpdate(FileNumber As Long)


'MsgBox "File " & FileNumber & " has been added to the Restarts Waiting queue", , "Restart Wizard"
'DoCmd.Close acForm, "EnterRestartReason"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Journal"
'DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "Case List"



End Sub
Public Sub IntakeWaitingCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Dim rstvalumeintake As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
!IntakeWaiting = Now
!IntakeWaitingby = GetStaffID
'!IntakeDocsRecdFlag = False
If Not IsNull(rstqueue!DateSentAttyIntake) Then rstqueue!DateSentAttyIntake = Null
.Update
End With
Set rstqueue = Nothing

Set rstvalumeintake = CurrentDb.OpenRecordset("Select * from ValumeIntake", dbOpenDynaset, dbSeeChanges)
With rstvalumeintake
.AddNew
!CaseFile = FileNumber
!Client = DLookup("ShortClientName", "ClientList", "ClientID = " & Forms![Case List]!ClientID)
!IntakeWaiting = Now
!IntakeWaitingC = 1
!Name = GetFullName()
.Update
End With
Set rstvalumeintake = Nothing


MsgBox "File " & FileNumber & " has been added to the Intake Waiting queue", , "Intake Wizard"
DoCmd.Close acForm, "EnterIntakeReason"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"



If CurrentProject.AllForms("queintake").IsLoaded = True Then
Forms!queintake!lstFiles.Requery
Forms!queintake.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueintake", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queintake!QueueCount = cntr
Set rstqueue = Nothing
End If
If CurrentProject.AllForms("queintakewaiting").IsLoaded = True Then
Forms!queIntakeWaiting!lstFiles.Requery
Forms!queIntakeWaiting.Requery
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueintakewaitinglst", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queIntakeWaiting!QueueCount = cntr
Set rstqueue = Nothing
End If

End Sub
Public Sub DocketingWaitingCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset, cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
!DocketingWaiting = Now
!DocketingWaitingby = GetStaffID
!DocketingDocsRecdFlag = False
!AttyMilestone4Reject = False
.Update
End With

Set rstqueue = Nothing

DoCmd.SetWarnings False
strinfo = "Send to docket waiting queue"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True


Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Warning] = 100
'  ![Info] = "Send to docket waiting queue" & vbCrLf
'  ![Color] = 1
'  .Update
'  End With

MsgBox "File " & FileNumber & " has been added to the Docketing Waiting queue", , "Docketing Wizard"
DoCmd.Close acForm, "EnterDocketingReason"
DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "DocsWindow"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"

If CurrentProject.AllForms("queDocketing").IsLoaded = True Then
Forms!quedocketing!lstFiles.Requery
Forms!quedocketing.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueDocketinggroupby", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!quedocketing!QueueCount = cntr
Set rstqueue = Nothing
End If
If CurrentProject.AllForms("queDocketingwaiting").IsLoaded = True Then
Forms!queDocketingWaiting!lstFiles.Requery
Forms!queDocketingWaiting.Requery
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuedocketingwaitinglst", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queDocketingWaiting!QueueCount = cntr
Set rstqueue = Nothing
End If
End Sub

Public Sub RSIICompletionUpdate(FileNumber As Long, Exceptions As Long)
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!RSIIcomplete) Then
With rstqueue
.Edit
!RSIIcomplete = Now
!RSIIuser = GetStaffID
!RSIIexceptions = Exceptions
.Update
End With
End If
Set rstqueue = Nothing
If CurrentProject.AllForms("queRSII").IsLoaded = True Then
Forms!queRSII!lstFiles.Requery
Forms!queRSII.Requery
Dim cntr As Integer
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueRSII", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queRSII!QueueCount = cntr
Set rstqueue = Nothing
End If


End Sub
Public Sub NOICompletionUpdate(WizardType As String, FileNumber As Long)
Dim rstqueue As Recordset
Dim rstdocs As Recordset
Dim cntr As Integer

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!NOIcomplete) Then
With rstqueue
.Edit
!NOIcomplete = Now
!NOIuser = GetStaffID
!DateSentAttyNOI = Null
!AttyMilestone1_5 = Null
!DateInQueueNOI = Null
!AttyMilestone1_5Reject = False
!AttyMilestoneMgr1_5 = Null
!DateInWaiitingQueueNOI = Null
!Add45 = ""
.Update
End With
End If
Set rstqueue = Nothing

'Check if needed for docketing
Set rstdocs = CurrentDb.OpenRecordset("Select * FROM docketingdocsneeded where filenumber=" & FileNumber & " AND DocName=""NOI""", dbOpenDynaset, dbSeeChanges)
If Not rstdocs.EOF Then
With rstdocs
.Edit
!DocReceived = Date
!docreceivedby = GetStaffID
.Update
End With
Set rstdocs = Nothing

'Remove from docket queue
Set rstdocs = CurrentDb.OpenRecordset("Select * FROM docketingdocsneeded where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
If rstdocs.EOF Then
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
!DocketingWaiting = Null
!DocketingQueue = Null
!DocketingDocsRecdFlag = False

.Update
End With
Set rstdocs = Nothing
Set rstqueue = Nothing
End If
End If

'Check if needed for intake
Set rstdocs = CurrentDb.OpenRecordset("Select * FROM intakedocsneeded where filenumber=" & FileNumber & " AND DocName=""NOI""", dbOpenDynaset, dbSeeChanges)
If Not rstdocs.EOF Then
With rstdocs
.Edit
!DocReceived = Date
!docreceivedby = GetStaffID
.Update
End With
Set rstdocs = Nothing

Set rstdocs = CurrentDb.OpenRecordset("Select * FROM intakedocsneeded where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
If rstdocs.EOF Then
Set rstqueue = CurrentDb.OpenRecordset("Select intakedocsrecdflag FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
!IntakeDocsRecdFlag = True
.Update
End With
Set rstdocs = Nothing
Set rstqueue = Nothing
End If
End If



'----update missing restart doucment
DoCmd.SetWarnings False

Set rstdocs = CurrentDb.OpenRecordset("Select * FROM RestartDocumentMissing where filenumber=" & FileNumber & " AND docreceived is null AND DocName= ""NOI""", dbOpenDynaset, dbSeeChanges)
If Not rstdocs.EOF Then
    rstdocs.Close
    Set rstdocs = Nothing
    
    
    
            strSQL = "UPDATE RestartDocumentMissing SET DocReceived = #" & Now() & "# , docreceivedby = " & GetStaffID & _
                    " WHERE FileNumber= " & FileNumber & " AND DocName= '" & "NOI" & "'" & " And IsNull(DocReceived)"
                    DoCmd.RunSQL strSQL
                    strSQL = ""
        
                                
                    strinfo = " NOI was removed from the restart waiting list of outstanding items "
                    strinfo = Replace(strinfo, "'", "''")
                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & " ,Now,GetFullName(),'" & strinfo & "',1 )"
                    DoCmd.RunSQL strSQLJournal
                    

        
        
        Set rstdocs = CurrentDb.OpenRecordset("Select * FROM RestartDocumentMissing  where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
            If rstdocs.EOF Then
            
            strSQL = "UPDATE wizardqueuestats SET RestartDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                    DoCmd.RunSQL strSQL
                    strSQL = ""
            
            Set rstqueue = Nothing
            End If
            DoCmd.SetWarnings True


Else

rstdocs.Close
Set rstdocs = Nothing
End If
'----

If CurrentProject.AllForms("queNOInew").IsLoaded = True Then
Forms!queNOInew!lstFiles.Requery
Forms!queNOInew.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueNOI", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queNOInew!QueueCount = cntr
Set rstqueue = Nothing
End If
If CurrentProject.AllForms("queNOIdocs").IsLoaded = True Then
Forms!queNOIdocs!lstFiles.Requery
Forms!queNOIdocs.Requery
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueueNOIdocs", dbOpenDynaset, dbSeeChanges)
Do Until rstqueue.EOF
cntr = cntr + 1
rstqueue.MoveNext
Loop
Forms!queNOIdocs!QueueCount = cntr
Set rstqueue = Nothing
End If
End Sub
Public Sub FairDebtCompletionUpdate(WizardType As String, FileNumber As Long)
Dim rstqueue As Recordset, rstdocs As Recordset, cntr As Integer
Dim strSQL As String
Dim rstsql As String

DoCmd.SetWarnings False

If Forms!wizfairdebt!cmdOK.Caption = "Sent to Atty" Then

    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
 '   If IsNull(rstqueue!FairDebtAttyReview) Then
    With rstqueue
    .Edit
 '   !FairDebtComplete = Now
    !FairDebtUser = GetStaffID
    !FairDebtComplete = Null
    !FairDebtReason = Null
    !AttyMilestone1 = Null
    !AttyMilestone1Reject = False
    !FairDebtAttyReview = Now()
    !FairDebtWaiting = Now()
    !AddFair = ""
    .Update
    End With
    
    
                    'DoCmd.SetWarnings False
                    rstsql = "Insert into ValumeFD (CaseFile, Client, Name, FDAttyReview, FDAttyReviewC, state ) values (Forms!wizFairDebt!FileNumber, ClientShortName(forms!wizFairDebt!ClientID),Getfullname(),Now(),1, Forms!wizFairDebt!State) "
                    DoCmd.RunSQL rstsql
                   ' DoCmd.SetWarnings True
    
    
    
    
 '   End If
Else

 Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
 '   If IsNull(rstqueue!FairDebtAttyReview) Then
    With rstqueue
    .Edit
    !FairDebtComplete = Now
    !FairDebtUser = GetStaffID
  '  !FairDebtWaiting = Null
    !FairDebtReason = Null
    !AttyMilestone1 = Null
    !AttyMilestone1Reject = False
    !FairDebtAttyReview = Null
    !AddFair = ""
    .Update
    End With
 '   End If

End If

Set rstqueue = Nothing

Set rstdocs = CurrentDb.OpenRecordset("Select * FROM docketingdocsneeded where filenumber=" & FileNumber & " AND DocName=""Fair Debt""", dbOpenDynaset, dbSeeChanges)
Do Until rstdocs.EOF
With rstdocs
.Edit
!DocReceived = Date
!docreceivedby = GetStaffID
.Update
End With
rstdocs.MoveNext
Loop
Set rstdocs = Nothing




strSQL = "UPDATE DemandDocsNeeded SET " & " DocReceived = #" & Now() & "# , docreceivedby = " & GetStaffID & _
    " WHERE FileNumber = " & FileNumber & " AND DocName = ('" & "Waiting for Fair Debt" & "')" & " And Isnull(DocReceived)"
    DoCmd.RunSQL strSQL
    strSQL = ""


Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DemandDocsNeeded where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
    If rstdocs.EOF Then
    
    
    strSQL = "UPDATE wizardqueuestats SET DemandDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
            DoCmd.RunSQL strSQL
            strSQL = ""
    End If
Set rstqueue = Nothing
    




Set rstdocs = CurrentDb.OpenRecordset("Select * FROM docketingdocsneeded where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
If rstdocs.EOF Then
Set rstqueue = CurrentDb.OpenRecordset("Select docketingdocsrecdflag FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
!DocketingDocsRecdFlag = True
.Update
End With
Set rstdocs = Nothing
Set rstqueue = Nothing
End If



'--- Adding missing restart document

 '---s1
    Set rstdocs = CurrentDb.OpenRecordset("Select * FROM RestartDocumentMissing where filenumber=" & FileNumber & " AND docreceived is null AND DocName= ""FD""", dbOpenDynaset, dbSeeChanges)
If Not rstdocs.EOF Then
    rstdocs.Close
    Set rstdocs = Nothing
    
    
    
            strSQL = "UPDATE RestartDocumentMissing SET DocReceived = #" & Now() & "# , docreceivedby = " & GetStaffID & _
                    " WHERE FileNumber= " & FileNumber & " AND DocName= '" & "FD" & "'" & " And IsNull(DocReceived)"
                    DoCmd.RunSQL strSQL
                    strSQL = ""
        
                                
                    strinfo = " Fair Dabt was removed from the restart waiting list of outstanding items "
                    strinfo = Replace(strinfo, "'", "''")
                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & " ,Now,GetFullName(),'" & strinfo & "',1 )"
                    DoCmd.RunSQL strSQLJournal
                    

        
        
        Set rstdocs = CurrentDb.OpenRecordset("Select * FROM RestartDocumentMissing  where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
            If rstdocs.EOF Then
            
            strSQL = "UPDATE wizardqueuestats SET RestartDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                    DoCmd.RunSQL strSQL
                    strSQL = ""
            
            Set rstqueue = Nothing
            End If
            DoCmd.SetWarnings True


Else

rstdocs.Close
Set rstdocs = Nothing
End If
    '--------------------



'
'
'
'strsql = "UPDATE RestartDocumentMissing SET DocReceived = #" & Now() & "# , docreceivedby = " & GetStaffID & _
'        " WHERE FileNumber= " & FileNumber & " AND DocName= '" & "FD" & "'" & " And IsNull(DocReceived)"
'        DoCmd.RunSQL strsql
'        strsql = ""
'
'
'Set rstdocs = CurrentDb.OpenRecordset("Select * FROM RestartDocumentMissing  where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
'    If rstdocs.EOF Then
'
'    strsql = "UPDATE wizardqueuestats SET RestartDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
'            DoCmd.RunSQL strsql
'            strsql = ""
'
'    Set rstqueue = Nothing
'    End If






If CurrentProject.AllForms("queFairDebt").IsLoaded = True Then
    Forms!queFairDebt!lstFiles.Requery
    Forms!queFairDebt.Requery
    
    
    
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuefairdebt", dbOpenDynaset, dbSeeChanges)
    Do Until rstqueue.EOF
    cntr = cntr + 1
    rstqueue.MoveNext
    Loop
    Forms!queFairDebt!QueueCount = cntr
    Set rstqueue = Nothing
End If
If CurrentProject.AllForms("queFairDebtwaiting").IsLoaded = True Then
    Forms!queFairDebtWaiting!lstFiles.Requery
    Forms!queFairDebtWaiting.Requery
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuefairdebtwaiting", dbOpenDynaset, dbSeeChanges)
    Do Until rstqueue.EOF
    cntr = cntr + 1
    rstqueue.MoveNext
    Loop
    Forms!queFairDebtWaiting!QueueCount = cntr
    Set rstqueue = Nothing
End If

DoCmd.SetWarnings True
End Sub
Public Sub DemandCompletionUpdate(WizardType As String, FileNumber As Long)
Dim rstqueue As Recordset, rstdocs As Recordset, cntr As Integer
Dim rstDemandNeedDoc As Recordset
Dim strSQL As String
DoCmd.SetWarnings False


', AddDemand = """
strSQL = "UPDATE wizardqueuestats SET DemandComplete = #" & Now() & "# , DemandUser = " & GetStaffID & ", DemandReason = null, AddDemand = '" & "" & "'" & " WHERE filenumber= " & FileNumber & " AND current= true"
DoCmd.RunSQL strSQL
strSQL = ""


strSQL = "UPDATE DemandDocsNeeded SET " & " DocReceived = #" & Now() & "# , docreceivedby = " & GetStaffID & _
    " WHERE FileNumber = " & FileNumber & " AND DocName = ('" & "Waiting for client demand" & "')" & " And Isnull(DocReceived)"
    DoCmd.RunSQL strSQL
    strSQL = ""

    
    
    
    
    Set rstdocs = CurrentDb.OpenRecordset("Select * FROM RestartDocumentMissing where filenumber=" & FileNumber & " AND docreceived is null AND DocName= ""Demand Letter""", dbOpenDynaset, dbSeeChanges)
If Not rstdocs.EOF Then
    rstdocs.Close
    Set rstdocs = Nothing
    
    
    
            strSQL = "UPDATE RestartDocumentMissing SET DocReceived = #" & Now() & "# , docreceivedby = " & GetStaffID & _
                    " WHERE FileNumber= " & FileNumber & " AND DocName= '" & "Demand Letter" & "'" & " And IsNull(DocReceived)"
                    DoCmd.RunSQL strSQL
                    strSQL = ""
        
                                
                    strinfo = " Demand was removed from the restart waiting list of outstanding items "
                    strinfo = Replace(strinfo, "'", "''")
                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & " ,Now,GetFullName(),'" & strinfo & "',1 )"
                    DoCmd.RunSQL strSQLJournal
                    

        
        
        Set rstdocs = CurrentDb.OpenRecordset("Select * FROM RestartDocumentMissing  where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
            If rstdocs.EOF Then
            
            strSQL = "UPDATE wizardqueuestats SET RestartDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                    DoCmd.RunSQL strSQL
                    strSQL = ""
            
            Set rstqueue = Nothing
            End If
            DoCmd.SetWarnings True


Else

rstdocs.Close
Set rstdocs = Nothing
End If
    
    
    
    
    
    
    
    




'Set rstDocs = CurrentDb.OpenRecordset("Select * FROM docketingdocsneeded where filenumber=" & FileNumber & " AND DocName=""Acceleration""", dbOpenDynaset, dbSeeChanges)
'If Not rstDocs.EOF Then
'With rstDocs
'.Edit
'!DocReceived = Date
'!docreceivedby = GetStaffID
'.Update
'End With
'Set rstDocs = Nothing


strSQL = "UPDATE docketingdocsneeded SET " & " DocReceived = #" & Now() & "# , docreceivedby = " & GetStaffID & _
        " WHERE FileNumber = " & FileNumber & " AND DocName = ('" & "Acceleration" & "')" & " And Isnull(DocReceived)"
        DoCmd.RunSQL strSQL
        strSQL = ""





Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DemandDocsNeeded where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
    If rstdocs.EOF Then
    
    
    strSQL = "UPDATE wizardqueuestats SET DemandDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
            DoCmd.RunSQL strSQL
            strSQL = ""
    
    Set rstqueue = Nothing
    End If
'End If
If CurrentProject.AllForms("quedemand").IsLoaded = True Then
Forms!queDemand!lstFiles.Requery
Forms!queDemand.Requery

Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuedemand", dbOpenDynaset, dbSeeChanges)
    If rstqueue.EOF Then
    cntr = 0
    Else
    rstqueue.MoveLast
    cntr = rstqueue.RecordCount
    Forms!queDemand!QueueCount = cntr
    Set rstqueue = Nothing
    End If
End If

If CurrentProject.AllForms("quedemandwaiting").IsLoaded = True Then
Forms!queDemandWaiting!lstFiles.Requery
Forms!queDemandWaiting.Requery
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM qryqueuedemandwaiting", dbOpenDynaset, dbSeeChanges)
    If rstqueue.EOF Then
    cntr = 0
    Else
    rstqueue.MoveLast
    Forms!queDemandWaiting!QueueCount = cntr
    Set rstqueue = Nothing
    End If
End If

DoCmd.SetWarnings True
End Sub
Public Sub HUDOccCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!HUDOCCcomplete) Then
With rstqueue
.Edit
!HUDOCCcomplete = Now
!HUDOCCuser = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing
MsgBox "HUD Occupancy Wizard complete"
End Sub
Public Sub VAappraisalCompletionUpdate(FileNumber As Long)
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!VAappraisalComplete) Then
With rstqueue
.Edit
!VAappraisalComplete = Now
!VAappraisalUser = GetStaffID
.Update
End With
End If
Set rstqueue = Nothing

End Sub
Public Sub SelectDocsTab(FileNum As Long)

FileLocks = True
If LockFile(FileNum) Then
DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNum
Forms![Case List]!Page97.SetFocus
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNum
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
End Sub
Sub FairDebtCallFromQueue(FileNumber As Long)
Dim F As Form, FormClosed As Boolean, rstfiles As Recordset
On Error GoTo Err_FairDebtCallFromQueue_Click

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queFairDebt", "queFairDebtWaiting", "Wizards" ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If

FileLocks = True
    If LockFile(FileNumber) Then
    DoCmd.OpenForm "wizFairDebt"
    
            If DCount("*", "CIVDetails", "FCFileNumber= " & FileNumber) > 0 Then
              MsgBox "CAUTION! Litigation in progress, see an attorney!", vbExclamation
              Forms!wizfairdebt.Detail.BackColor = vbYellow
            End If

    
    
    With Forms!wizfairdebt
        .RecordSource = CurrentDb.QueryDefs("qryqryFairDebt").sql
        .txtFileNumber = FileNumber
        .Filter = "FileNumber=" & FileNumber
        .FilterOn = True
    End With
        Call FairDebtConfirmationVisible(True)
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Forms!wizfairdebt!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], DocTitleID FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='' AND Filespec IS NOT NULL and DeleteDate is null"
Forms!wizfairdebt!lstDocs.Requery

Call FairDebtFieldsVisible(True)

Exit_FairDebtCallFromQueue_Click:
    Exit Sub

Err_FairDebtCallFromQueue_Click:
    MsgBox Err.Description
    Resume Exit_FairDebtCallFromQueue_Click
    
End Sub
Sub DemandCallFromQueue(FileNumber As Long)
Dim F As Form, FormClosed As Boolean, rstfiles As Recordset
On Error GoTo Err_wizDemandCallFromQueue_Click

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queDemand", "queDemandWaiting", "Wizards" ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If

FileLocks = True
    If LockFile(FileNumber) Then
    DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
    DoCmd.OpenForm "wizDemand"
    
           If DCount("*", "CIVDetails", "FCFileNumber= " & FileNumber) > 0 Then
              MsgBox "CAUTION! Litigation in progress, see an attorney!", vbExclamation
              Forms!WizDemand.Detail.BackColor = vbYellow
            End If

    
    With Forms!WizDemand
        .RecordSource = CurrentDb.QueryDefs("qryqryFairDebt").sql
        .txtFileNumber = FileNumber
        .Filter = "FileNumber=" & FileNumber
        .FilterOn = True
    End With
       ' Call DemandConfirmationVisible(True)
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If

Forms!WizDemand!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name],DocTitleID FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='' AND Filespec IS NOT NULL and DeleteDate is null"


If Not IsNull(Forms!WizDemand.txtDisposition) Then
Forms!WizDemand.lblDisposition.Visible = True
Forms!WizDemand.lblDisposition.Caption = "The file has a disposition of:  " & DLookup("Disposition", "FCDisposition", "ID=" & Forms!WizDemand.txtDisposition)
End If

Forms!WizDemand!lstDocs.Requery

'Call wizDemandFieldsVisible(True)

Exit_wizDemandCallFromQueue_Click:
    Exit Sub

Err_wizDemandCallFromQueue_Click:
    MsgBox Err.Description
    Resume Exit_wizDemandCallFromQueue_Click
    
End Sub
Sub SAICallFromQueue(FileNumber As Long)
Dim F As Form, FormClosed As Boolean, rstfiles As Recordset
On Error GoTo Err_wizSAICallFromQueue_Click

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queSAI", "queSAIWaiting", "Wizards" ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    DoCmd.OpenForm "wizSAI"
    With Forms!wizSAI
        .RecordSource = CurrentDb.QueryDefs("qryqryfairdebt").sql
        .txtFileNumber = FileNumber
        .Filter = "FileNumber=" & FileNumber
        .FilterOn = True
    End With
       ' Call SAIConfirmationVisible(True)
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Forms!wizSAI!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name] FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='' AND Filespec IS NOT NULL and DeleteDate is null"
Forms!wizSAI!lstDocs.Requery
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'Call wizSAIFieldsVisible(True)

Exit_wizSAICallFromQueue_Click:
    Exit Sub

Err_wizSAICallFromQueue_Click:
    MsgBox Err.Description
    Resume Exit_wizSAICallFromQueue_Click
    
End Sub
Sub SCRAnames(FileNumber As Long)
Dim rstNames As Recordset, ctr As Integer, FieldName As String, cost As Currency

If DLookup("casetypeid", "caselist", "filenumber=" & FileNumber) = 2 Then
Set rstNames = CurrentDb.OpenRecordset("SELECT Names.FileNumber, Names.First, Names.Last, Names.SSN FROM [Names] GROUP BY Names.FileNumber, Names.First, Names.Last, Names.SSN HAVING (((Names.FileNumber)=" & FileNumber & ") AND ((Names.SSN) Is Not Null)) OR (((Names.FileNumber)=" & FileNumber & ") AND ((Names.SSN)<>""999999999""))", dbOpenDynaset, dbSeeChanges)
Else
Set rstNames = CurrentDb.OpenRecordset("SELECT Names.FileNumber, Names.mortgagor, Names.First, Names.Last, Names.SSN FROM [Names] GROUP BY Names.FileNumber, Names.mortgagor, Names.First, Names.Last, Names.SSN HAVING (((Names.FileNumber)=" & FileNumber & ") AND ((Names.mortgagor)=Yes) AND ((Names.SSN) Is Not Null)) OR (((Names.FileNumber)=" & FileNumber & ") AND ((Names.mortgagor)=Yes) AND ((Names.SSN)<>""999999999""))", dbOpenDynaset, dbSeeChanges)
End If

With rstNames
If .EOF Then
MsgBox "There are no borrowers listed in this file", vbCritical
Exit Sub
End If
.MoveLast
ctr = .RecordCount
.MoveFirst
If ctr > 4 Then MsgBox "There are " & ctr & " borrowers, only the first 4 will be shown.", vbExclamation
If ctr = 0 Then
MsgBox "There are no borrowers listed in this file", vbCritical
Exit Sub
End If

Select Case ctr

Case 1
Forms![SCRA Search Info]!First1 = !First
Forms![SCRA Search Info]!Last1 = !Last
Forms![SCRA Search Info]!SSN1 = !SSN
Case 2
Forms![SCRA Search Info]!First1 = !First
Forms![SCRA Search Info]!Last1 = !Last
Forms![SCRA Search Info]!SSN1 = !SSN
.MoveNext
Forms![SCRA Search Info]!First2 = !First
Forms![SCRA Search Info]!Last2 = !Last
Forms![SCRA Search Info]!SSN2 = !SSN
Case 3
Forms![SCRA Search Info]!First1 = !First
Forms![SCRA Search Info]!Last1 = !Last
Forms![SCRA Search Info]!SSN1 = !SSN
.MoveNext
Forms![SCRA Search Info]!First2 = !First
Forms![SCRA Search Info]!Last2 = !Last
Forms![SCRA Search Info]!SSN2 = !SSN
.MoveNext
Forms![SCRA Search Info]!First3 = !First
Forms![SCRA Search Info]!Last3 = !Last
Forms![SCRA Search Info]!SSN3 = !SSN
Case Else
Forms![SCRA Search Info]!First1 = !First
Forms![SCRA Search Info]!Last1 = !Last
Forms![SCRA Search Info]!SSN1 = !SSN
.MoveNext
Forms![SCRA Search Info]!First2 = !First
Forms![SCRA Search Info]!Last2 = !Last
Forms![SCRA Search Info]!SSN2 = !SSN
.MoveNext
Forms![SCRA Search Info]!First3 = !First
Forms![SCRA Search Info]!Last3 = !Last
Forms![SCRA Search Info]!SSN3 = !SSN
.MoveNext
Forms![SCRA Search Info]!First4 = !First
Forms![SCRA Search Info]!Last4 = !Last
Forms![SCRA Search Info]!SSN4 = !SSN
End Select
.Close
End With

If DLookup("casetypeid", "caselist", "filenumber=" & FileNumber) = 2 Then
    If DLookup("clientid", "caselist", "Filenumber=" & FileNumber) = 97 Then
    cost = ctr * DLookup("ivalue", "db", "ID=" & 32)
'AddInvoiceItem FileNumber, "BK-DOD", "DOD Search", cost, 0, False, True, False, False

'added stage to InvoiceItem  and put under the fee. 2/10/15
    AddInvoiceItem FileNumber, "BK-DOD", "DOD Search - " & strStage & " ", cost, 0, True, True, False, False
    End If
Else
    If IsLoadedF("queSCRAFCNew") = True Then
        cost = ctr * DLookup("ivalue", "db", "ID=" & 32)
        If DLookup("SCRAStageID", "SCRAQueueFiles", "FileNumber=" & FileNumber) < 91 And DLookup("clientid", "caselist", "Filenumber=" & FileNumber) = 97 Then
            'AddInvoiceItem FileNumber, "FC-DOD", "DOD Search", cost, 0, False, True, False, False
            AddInvoiceItem FileNumber, "FC-DOD", "DOD Search - " & strStage & " ", cost, 0, True, True, False, False
        ElseIf DLookup("clientid", "caselist", "Filenumber=" & FileNumber) = 97 Then
            'added stage to InvoiceItem  and put under the fee. 2/10/15
            AddInvoiceItem FileNumber, "FC-DOD", "DOD Search - " & strStage & " ", cost, 0, True, True, False, False
        End If

    End If
End If
strStage = ""

End Sub
Sub VAnames(FileNumber As Long)
Dim rstNames As Recordset, rstClient As Recordset, ctr As Integer
Set rstNames = CurrentDb.OpenRecordset("select * from Names where filenumber=" & FileNumber & " and owner=yes", dbOpenDynaset, dbSeeChanges)

With rstNames
.MoveLast
ctr = .RecordCount
.MoveFirst
If ctr > 4 Then MsgBox "There are " & ctr & " borrowers, only the first 4 will be shown.", vbExclamation
If ctr = 0 Then
MsgBox "There are no borrowers listed in this file", vbCritical
Exit Sub
End If

Select Case ctr

Case 1
Forms![vaappraisal Search Info]!First1 = !First
Forms![vaappraisal Search Info]!Last1 = !Last
Forms![vaappraisal Search Info]!SSN1 = !SSN


Case 2
Forms![vaappraisal Search Info]!First1 = !First
Forms![vaappraisal Search Info]!Last1 = !Last
Forms![vaappraisal Search Info]!SSN1 = !SSN
.MoveNext
Forms![vaappraisal Search Info]!First2 = !First
Forms![vaappraisal Search Info]!Last2 = !Last
Forms![vaappraisal Search Info]!SSN2 = !SSN
Case 3
Forms![vaappraisal Search Info]!First1 = !First
Forms![vaappraisal Search Info]!Last1 = !Last
Forms![vaappraisal Search Info]!SSN1 = !SSN
.MoveNext
Forms![vaappraisal Search Info]!First2 = !First
Forms![vaappraisal Search Info]!Last2 = !Last
Forms![vaappraisal Search Info]!SSN2 = !SSN
.MoveNext
Forms![vaappraisal Search Info]!First3 = !First
Forms![vaappraisal Search Info]!Last3 = !Last
Forms![vaappraisal Search Info]!SSN3 = !SSN
Case Else
Forms![vaappraisal Search Info]!First1 = !First
Forms![vaappraisal Search Info]!Last1 = !Last
Forms![vaappraisal Search Info]!SSN1 = !SSN
.MoveNext
Forms![vaappraisal Search Info]!First2 = !First
Forms![vaappraisal Search Info]!Last2 = !Last
Forms![vaappraisal Search Info]!SSN2 = !SSN
.MoveNext
Forms![vaappraisal Search Info]!First3 = !First
Forms![vaappraisal Search Info]!Last3 = !Last
Forms![vaappraisal Search Info]!SSN3 = !SSN
.MoveNext
Forms![vaappraisal Search Info]!First4 = !First
Forms![vaappraisal Search Info]!Last4 = !Last
Forms![vaappraisal Search Info]!SSN4 = !SSN
End Select
.Close
End With

Set rstClient = CurrentDb.OpenRecordset("select * from clientlist where clientid=" & Forms![Case List]!ClientID, dbOpenDynaset, dbSeeChanges)

With rstClient
Forms![vaappraisal Search Info]!ClientName = !LongClientName
Forms![vaappraisal Search Info]!Address = !StreetAddress
Forms![vaappraisal Search Info]!Address2 = !StreetAddr2
Forms![vaappraisal Search Info]!City = !City
Forms![vaappraisal Search Info]!State = !State
Forms![vaappraisal Search Info]!ZipCode = !ZipCode
Forms![vaappraisal Search Info]!ClientContact = !VAContactLastName & ", " & !VAContactFirstName
Forms![vaappraisal Search Info]!ClientPhone = !VAContactPhoneNum
Forms![vaappraisal Search Info]!ClientEmail = !VAContactEmail
Forms![vaappraisal Search Info]!sponsorid = !VAServicerID
.Close
End With

Forms![vaappraisal Search Info]!LoanNumber = Forms!foreclosuredetails!FHALoanNumber
Forms![vaappraisal Search Info]!ShortLegal = Forms!foreclosuredetails!ShortLegal
Forms![vaappraisal Search Info]!PropAddress = Forms!foreclosuredetails!PropertyAddress
Forms![vaappraisal Search Info]!PropCity = Forms!foreclosuredetails!City
Forms![vaappraisal Search Info]!PropState = Forms!foreclosuredetails!State
Forms![vaappraisal Search Info]!PropZip = Forms!foreclosuredetails!ZipCode
Forms![vaappraisal Search Info]!ClientNumber = Forms![Case List]!ClientNumber

End Sub

Sub TitleOutstandingCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queTitleOut", "Wizards" ', "queRestartWaiting"  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed


Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
Forms![Case List]!cmdSearch.Visible = False
Forms![Case List]!cmdGoToFile.Visible = False
DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "TitleOut"
'Forms!ForeclosureDetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!pageAccounting.Visible = True
Forms!foreclosuredetails!TitleThru.Locked = False
Forms!foreclosuredetails!WizardSource = "TitleOut"
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call TitleOutVisible(True)
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub

Sub TitleOutVisible(SetVisible As Boolean)
Forms![Case List]!Page120.Visible = False
Forms![Case List]!Page91.Visible = False
Forms![Case List]!pageCheckRequest.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pgConflicts.Visible = False
Forms![Case List]!Page97.Visible = True
Forms![Case List]!cmdAddDoc.Visible = False

'Forms!ForeclosureDetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
'Forms!ForeclosureDetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
Forms![foreclosuredetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False
Forms![foreclosuredetails]!Trustees.Visible = True
Forms![foreclosuredetails]!pgNOI.Visible = False
Forms![foreclosuredetails]![Post-Sale].Visible = False
Forms![foreclosuredetails]!pageAccounting.Visible = False
Forms!foreclosuredetails!cmdPrint.Visible = False


End Sub


Public Sub TitleOutComplete(FileNumber As Long)

Dim rstqueueW As Recordset
Dim rstqueueH As Recordset
Dim rstFCdetails As Recordset
Dim lrs As Recordset
Dim DateThrough As Date
Dim FormCaption As String
FormCaption = Forms!foreclosuredetails!cmdWizComplete.Caption
If Forms!foreclosuredetails!cmdWizComplete.Caption <> "Title Cancelled completed" Then
DateThrough = Forms!foreclosuredetails!TitleThru
End If
Dim JourText As String
Set rstqueueW = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueueW
.Edit
!TitleOutComplete = Now
!TitleOutCompleteBy = GetStaffID
.Update
End With
Set rstqueueW = Nothing


Set rstqueueH = CurrentDb.OpenRecordset("ValumeTitleOut", dbOpenDynaset, dbSeeChanges)

rstqueueH.AddNew
rstqueueH!CaseFile = FileNumber
rstqueueH!Client = Forms![Case List]!ClientID
rstqueueH!Name = GetFullName()
rstqueueH!Title = IIf(Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Upload completed", Now, Null)
rstqueueH!TitleCount = IIf(Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Upload completed", 1, 0)
rstqueueH!TitleUpdate = IIf(Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Update Upload completed", Now, Null)
rstqueueH!TitleUpdateCount = IIf(Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Update Upload completed", 1, 0)
rstqueueH!Cancel = IIf(Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Cancelled completed", Now, Null)
rstqueueH!CancelCount = IIf(Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Cancelled completed", 1, 0)
rstqueueH.Update
Set rstqueueH = Nothing

Select Case Forms!foreclosuredetails!cmdWizComplete.Caption

Case "Title Upload completed"
JourText = "Title Uploaded"

Case "Title Update Upload completed"
JourText = "Update Uploaded"

Case "Title Cancelled completed"
JourText = Forms!foreclosuredetails!TitleCancelledReason


Dim rstdocs As Recordset
Set rstdocs = CurrentDb.OpenRecordset("Select Top 1 TitleRecieved From TitleReceivedArchive Where FileNumber =" & FileNumber & " Order by TitleRecieved DESC", dbOpenDynaset, dbSeeChanges)
If Not rstdocs.EOF Then
Forms!foreclosuredetails!TitleBack = rstdocs!TitleRecieved
'Else
'MsgBox (" Please notice that there is no old Ttile received date")
End If
rstdocs.Close

Set rstdocs = CurrentDb.OpenRecordset("Select Top 1 TitleThrough From TitleThroughArchive Where FileNumber =" & FileNumber & " Order by TitleThrough DESC", dbOpenDynaset, dbSeeChanges)
If Not rstdocs.EOF Then
Forms!foreclosuredetails!TitleThru = rstdocs!TitleThrough
'Else
'MsgBox (" Please notice that there is no old Ttile Through date")
End If
rstdocs.Close

Set rstdocs = CurrentDb.OpenRecordset("Select Top 1 TitleReviewToClient From TitleReviewArchive Where FileNumber =" & FileNumber & " Order by TitleReviewToClient DESC", dbOpenDynaset, dbSeeChanges)
If Not rstdocs.EOF Then
Forms!foreclosuredetails!TitleReviewToClient = rstdocs!TitleReviewToClient
'Else
'MsgBox (" Please notice that there is no old Ttile Through date")
End If
rstdocs.Close








End Select

'2/11/14
DoCmd.SetWarnings False
strinfo = JourText
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True



'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Info] = JourText
'  ![Color] = 1
'  .Update
'  .Close
'  End With


If Forms!foreclosuredetails!cmdWizComplete.Caption <> "Title Cancelled completed" Then

DoCmd.Close acForm, "ForeclosureDetails"
Set rstFCdetails = CurrentDb.OpenRecordset("select * from fcdetails where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)

    With rstFCdetails
    .Edit
    !TitleBack = Date
    !TitleReviewToClient = Null
    !TitleThru = DateThrough
    .Update
    .Close
End With

    DoCmd.SetWarnings False
    Dim rstsql As String
    rstsql = "Insert InTo TitleReceivedArchive (FileNumber, TitleRecieved, DateEntered) Values ( " & FileNumber & ", '" & Date & "' , '" & Now() & "')"
    DoCmd.RunSQL rstsql
    DoCmd.SetWarnings True
    

Set rstFCdetails = Nothing
End If

End Sub

Sub TitleOutCallFromQueue(FileNumber As Long)

Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queTitleOut", "Wizards" ', "queRestartWaiting"  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed


Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
Forms![Case List]!cmdSearch.Visible = False
Forms![Case List]!cmdGoToFile.Visible = False
DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "Titleout"
'Forms!ForeclosureDetails!cmdWizComplete.Visible = True
'Forms!ForeclosureDetails!pageAccounting.Visible = True


Forms!foreclosuredetails!WizardSource = "TitleOut"
Forms!foreclosuredetails!TitleThru.Locked = False

    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call TitleOutVisible(True)
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False

End Sub
Sub TitleReviewCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "queTitleReview", "Wizards" ', "queRestartWaiting"  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed


Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "TitleReview"
'Forms!ForeclosureDetails!cmdWizComplete.Visible = True
Forms!foreclosuredetails!pageAccounting.Visible = True
Forms!foreclosuredetails!WizardSource = "TitleReview"
Forms!foreclosuredetails!TitleThru.Locked = False
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call TitleReviewVisible(True)
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
'DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & FileNumber
Forms![Case List].SetFocus
Forms![Case List]!cmdClose.Visible = False
Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub
Sub TitleReviewVisible(SetVisible As Boolean)

Forms![Case List]!Page120.Visible = False
Forms![Case List]!Page91.Visible = False
Forms![Case List]!pageCheckRequest.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pgConflicts.Visible = False
Forms![Case List]!Page97.Visible = True
Forms![Case List]!cmdAddDoc.Visible = False
Forms![Case List]!cmdAddDoc.Visible = True

Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails!cmdWaiting.Caption = "Cancel"
Forms!foreclosuredetails!cmdWaiting.FontBold = True
Forms!foreclosuredetails!cmdWaiting.ForeColor = vbBlack




Forms!foreclosuredetails.cmdWaiting.Visible = SetVisible
Forms!foreclosuredetails.cmdcloserestart.Visible = False
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.Page412.Visible = True
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = True
Forms![foreclosuredetails]!pgMediation.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False
Forms![foreclosuredetails]!Trustees.Visible = True
Forms![foreclosuredetails]!pgNOI.Visible = False
Forms![foreclosuredetails]![Post-Sale].Visible = False
Forms![foreclosuredetails]!pageAccounting.Visible = False
Forms!foreclosuredetails!cmdPrint.Visible = False
Forms![foreclosuredetails]!Page256.SetFocus
Forms![foreclosuredetails]!cmdPrint.Visible = True


End Sub

Public Sub TitleReviewComplete(FileNumber As Long)

Dim rstqueueW As Recordset
Dim rstqueueH As Recordset
Dim rstFCdetails As Recordset
Dim lrs As Recordset




Set rstqueueH = CurrentDb.OpenRecordset("ValumeTitleReview", dbOpenDynaset, dbSeeChanges)

rstqueueH.AddNew
rstqueueH!CaseFile = FileNumber
rstqueueH!Client = Forms![Case List]!ClientID
rstqueueH!Name = GetFullName()
rstqueueH!Title = IIf(Forms!foreclosuredetails!TitleSearchType = "2 Owner", Now, Null)
rstqueueH!TitleCount = IIf(Forms!foreclosuredetails!TitleSearchType = "2 Owner", 1, 0)
rstqueueH!TitleUpdate = IIf(Forms!foreclosuredetails!TitleSearchType = "Update", Now, Null)
rstqueueH!TitleUpdateCount = IIf(Forms!foreclosuredetails!TitleSearchType = "Update", 1, 0)
rstqueueH!Cancel = IIf(IsNull(Forms!foreclosuredetails!TitleSearchType), Now, Null)
rstqueueH!CancelCount = IIf(IsNull(Forms!foreclosuredetails!TitleSearchType), 1, 0)
rstqueueH.Update
Set rstqueueH = Nothing


Dim JourText As String
Set rstqueueW = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
With rstqueueW
.Edit
!TitleReviewComplete = Now
!TitleRevewDateInQueue = Null
.Update
End With
Set rstqueueW = Nothing


DoCmd.SetWarnings False
strinfo = "Title Review complete"
strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal
DoCmd.SetWarnings True

'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  With lrs
'  .AddNew
'  ![FileNumber] = FileNumber
'  ![JournalDate] = Now
'  ![Who] = GetFullName()
'  ![Info] = "Title Review complete"
'  ![Color] = 1
'  .Update
'  .Close
'  End With




DoCmd.Close acForm, "ForeclosureDetails"
Set rstFCdetails = CurrentDb.OpenRecordset("select * from fcdetails where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)

    With rstFCdetails
    .Edit
    !TitleReviewToClient = Date
    .Update
    .Close
    End With
Set rstFCdetails = Nothing

    DoCmd.SetWarnings False
    Dim rstsql As String
    rstsql = "Insert InTo TitleReviewArchive (FileNumber, TitleReviewToClient, DateEntered) Values ( " & FileNumber & ", '" & Date & "' , '" & Now() & "')"
    DoCmd.RunSQL rstsql
    DoCmd.SetWarnings True

AddStatus FileNumber, Date, "Title Review Complete"

End Sub

Public Sub LitigationBillingCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "QueueAccounLitigationBill", "QueueAccountLitigationBillManager" '  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed


FileLocks = True
    If LockFile(FileNumber) Then

stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"


DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber


    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
Forms![Case List]!optDocType = 2
Forms![Case List]!CmdWizLiti.Visible = True

Forms![Case List]!lstDocs.ColumnCount = 6
Forms![Case List]!lstDocs.ColumnWidths = "0 in; 0.4 in; 0.75 in; 3 in; 0 in ;0.3 in "

Dim lstDocs As Recordset

Forms![Case List]!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], [doctitleid] AS DocType , Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='B' AND Filespec IS NOT NULL and DeleteDate is null"
Forms![Case List]!lstDocs.Requery


Dim Fnumber As Long
Dim Value As String
Dim blnFound As Boolean

If IsLoadedF("QueueAccounLitigationBill") = True Then
Forms![Case List]!CmdWizLiti.Visible = True
    If Forms!QueueAccounLitigationBill.lstFiles.Column(6) <> 0 Then
        Fnumber = Forms!QueueAccounLitigationBill.lstFiles.Column(6)
            If Not IsNull(Fnumber) Then
'                Dim Value As String
'                Dim blnFound As Boolean
                blnFound = False
                Dim J As Integer
                Dim A As Integer
                For J = 0 To Forms![Case List]!lstDocs.ListCount - 1
                   Value = Forms![Case List]!lstDocs.Column(0, J)
                   If InStr(Value, Fnumber) Then
                        blnFound = True
                         A = J
                        Forms![Case List].lstDocs.Selected(A) = True
                    Exit For
                    End If
                Next J
                
                If Not blnFound Then MsgBox ("Document not in the document list.")
                Forms![Case List]!lstDocs.SetFocus
                Else
                MsgBox ("Document not in the Document List.")
                Forms![Case List]!lstDocs.SetFocus
            End If
    Else
    
    Forms![Case List].cmdAddDoc.SetFocus
    
    End If
Else

If IsLoadedF("QueueAccountLitigationBillManager") = True Then
 If Forms![Case List]!CmdWizLiti.Visible = True Then Forms![Case List]!CmdWizLiti.Visible = False
    If Forms!QueueAccountLitigationBillManager.lstFiles.Column(9) <> 0 Then
        Fnumber = Forms!QueueAccountLitigationBillManager.lstFiles.Column(9)
            If Not IsNull(Fnumber) Then
                
                blnFound = False
'                Dim j As Integer
'                Dim A As Integer
                For J = 0 To Forms![Case List]!lstDocs.ListCount - 1
                   Value = Forms![Case List]!lstDocs.Column(0, J)
                   If InStr(Value, Fnumber) Then
                        blnFound = True
                         A = J
                        Forms![Case List].lstDocs.Selected(A) = True
                    Exit For
                    End If
                Next J
                
                If Not blnFound Then MsgBox ("Document not in the document list.")
                Forms![Case List]!lstDocs.SetFocus
                Else
                MsgBox ("Document not in the Document List.")
                Forms![Case List]!lstDocs.SetFocus
            End If
    Else
    
    Forms![Case List].cmdAddDoc.SetFocus
    
    End If
Else

Forms![Case List].cmdAddDoc.SetFocus
End If
End If


Forms![Case List]!Page120.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pageCheckRequest.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pgConflicts.Visible = False

Forms![Case List]!SCRAID = "AccLitig"
Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub


Public Sub PSAdvancedCostsCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "QueueAccounPSAdvancedCosts", "QueueAccountPSAdvancedCostManager" '  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

FileLocks = True
If LockFile(FileNumber) Then
    stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
    DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Else
    MsgBox "File is locked", vbCritical
    Exit Sub
End If

DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
Forms![Case List]!optDocType = 2
Forms![Case List]!CmdWizPS.Visible = True

Forms![Case List]!lstDocs.ColumnCount = 6
Forms![Case List]!lstDocs.ColumnWidths = "0 in; 0.4 in; 0.75 in; 3 in; 0 in ;0.3 in "

Dim lstDocs As Recordset
Dim GroupCode As String
Dim selecteddoctype As Integer
selecteddoctype = 0


selecteddoctype = Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(14)



If selecteddoctype = 113 Then
    Forms![Case List]!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], [doctitleid] AS DocType , Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & "  AND Filespec IS NOT NULL and DeleteDate is null"
Else
    Forms![Case List]!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], [doctitleid] AS DocType , Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='B' AND Filespec IS NOT NULL and DeleteDate is null"
    'Forms![Case List]!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], [doctitleid] AS DocType , Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='B' AND Filespec IS NOT NULL and DeleteDate is null or FileNumber=" & FileNumber & " AND isnull(DocGroup)= true and doctitleid = 113 AND Filespec IS NOT NULL and DeleteDate is null"
End If
Forms![Case List]!lstDocs.Requery

Dim Fnumber As Long
Dim Value As String
Dim blnFound As Boolean

If IsLoadedF("QueueAccounPSAdvancedCosts") = True Then
Forms![Case List]!CmdWizPS.Visible = True
    If Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6) <> 0 Then
        Fnumber = Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6)
            If Not IsNull(Fnumber) Then
'                Dim Value As String
'                Dim blnFound As Boolean
                blnFound = False
                Dim J As Integer
                Dim A As Integer
                For J = 0 To Forms![Case List]!lstDocs.ListCount - 1
                   Value = Forms![Case List]!lstDocs.Column(0, J)
                   If InStr(Value, Fnumber) Then
                        blnFound = True
                         A = J
                        Forms![Case List].lstDocs.Selected(A) = True
                    Exit For
                    End If
                Next J
                
                If Not blnFound Then MsgBox ("Document not in the document list.")
                Forms![Case List]!lstDocs.SetFocus
                Else
                MsgBox ("Document not in the Document List.")
                Forms![Case List]!lstDocs.SetFocus
            End If
    Else
    
    Forms![Case List].cmdAddDoc.SetFocus
    
    End If
Else

If IsLoadedF("QueueAccountPSAdvancedCostManager") = True Then
 If Forms![Case List]!CmdWizPS.Visible = True Then Forms![Case List]!CmdWizPS.Visible = False
    If Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(9) <> 0 Then
        Fnumber = Forms!QueueAccountPSAdvancedCostManager.lstFiles.Column(9)
            If Not IsNull(Fnumber) Then
                
                blnFound = False
'                Dim j As Integer
'                Dim A As Integer
                For J = 0 To Forms![Case List]!lstDocs.ListCount - 1
                   Value = Forms![Case List]!lstDocs.Column(0, J)
                   If InStr(Value, Fnumber) Then
                        blnFound = True
                         A = J
                        Forms![Case List].lstDocs.Selected(A) = True
                    Exit For
                    End If
                Next J
                
                If Not blnFound Then MsgBox ("Document not in the document list.")
                Forms![Case List]!lstDocs.SetFocus
                Else
                MsgBox ("Document not in the Document List.")
                Forms![Case List]!lstDocs.SetFocus
            End If
    Else
    
    Forms![Case List].cmdAddDoc.SetFocus
    
    End If
Else

Forms![Case List].cmdAddDoc.SetFocus
End If
End If


Forms![Case List]!Page120.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
'Forms![case list]!pageCheckRequest.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pgConflicts.Visible = False

Forms![Case List]!SCRAID = "AccPSAdvanced"
Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub


Public Sub EscrowCallFromQueue(FileNumber As Long)
Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart
'QueueAccounESC
Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "QueueAccounESC", "QueueESCtManager" '  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed


FileLocks = True
    If LockFile(FileNumber) Then

stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"


DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber




    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
Forms![Case List]!optDocType = 1
Forms![Case List]!ComWizESC.Visible = True

Forms![Case List]!lstDocs.ColumnCount = 6
Forms![Case List]!lstDocs.ColumnWidths = "0 in; 0.4 in; 0.75 in; 3 in; 0 in ;0.3 in "

Dim lstDocs As Recordset

Forms![Case List]!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], [doctitleid] AS DocType , Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND Filespec IS NOT NULL and DeleteDate is null"
Forms![Case List]!lstDocs.Requery


Dim Fnumber As Long
Dim Value As String
Dim blnFound As Boolean

If IsLoadedF("QueueAccounESC") = True Then
Forms![Case List]!ComWizESC.Visible = True
    If Forms!QueueAccounESC.lstFiles.Column(11) <> 0 Then
        Fnumber = Forms!QueueAccounESC.lstFiles.Column(11)
            If Not IsNull(Fnumber) Then
'                Dim Value As String
'                Dim blnFound As Boolean
                blnFound = False
                Dim J As Integer
                Dim A As Integer
                For J = 0 To Forms![Case List]!lstDocs.ListCount - 1
                   Value = Forms![Case List]!lstDocs.Column(0, J)
                   If InStr(Value, Fnumber) Then
                        blnFound = True
                         A = J
                        Forms![Case List].lstDocs.Selected(A) = True
                    Exit For
                    End If
                Next J
                
                If Not blnFound Then MsgBox ("Document not in the document list.")
                Forms![Case List]!lstDocs.SetFocus
                Else
                MsgBox ("Document not in the Document List.")
                Forms![Case List]!lstDocs.SetFocus
            End If
    Else
    
    Forms![Case List].cmdAddDoc.SetFocus
    
    End If
Else

If IsLoadedF("QueueESCtManager") = True Then
 If Forms![Case List]!CmdWizPS.Visible = True Then Forms![Case List]!CmdWizPS.Visible = False
    If Forms!QueueESCtManager.lstFiles.Column(11) <> 0 Then
        Fnumber = Forms!QueueESCtManager.lstFiles.Column(11)
            If Not IsNull(Fnumber) Then
                
                blnFound = False
'                Dim j As Integer
'                Dim A As Integer
                For J = 0 To Forms![Case List]!lstDocs.ListCount - 1
                   Value = Forms![Case List]!lstDocs.Column(0, J)
                   If InStr(Value, Fnumber) Then
                        blnFound = True
                         A = J
                        Forms![Case List].lstDocs.Selected(A) = True
                    Exit For
                    End If
                Next J
                
                If Not blnFound Then MsgBox ("Document not in the document list.")
                Forms![Case List]!lstDocs.SetFocus
                Else
                MsgBox ("Document not in the Document List.")
                Forms![Case List]!lstDocs.SetFocus
            End If
    Else
    
    Forms![Case List].cmdAddDoc.SetFocus
    
    End If
Else

Forms![Case List].ComWizESC.SetFocus
End If
End If


Forms![Case List]!Page120.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pageCheckRequest.Visible = True
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pgConflicts.Visible = False

Forms![Case List]!SCRAID = "AccEsc"
Forms![Case List].ComWizESC.SetFocus
Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
    
End Sub


Public Sub EscrowCallFromQueueR(FileNumber As Long)


Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_Restart
'QueueAccounESC
Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "QueueAccounESC", "QueueESCtManager" '  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed


FileLocks = True
    If LockFile(FileNumber) Then

stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"


DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber




    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
DoCmd.OpenForm "Journal", , , "FileNumber=" & FileNumber
Forms![Case List]!optDocType = 1
Forms![Case List]!ComWizESC.Visible = True

Forms![Case List]!lstDocs.ColumnCount = 6
Forms![Case List]!lstDocs.ColumnWidths = "0 in; 0.4 in; 0.75 in; 3 in; 0 in ;0.3 in "

Dim lstDocs As Recordset

Forms![Case List]!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], [doctitleid] AS DocType , Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND Filespec IS NOT NULL and DeleteDate is null"
Forms![Case List]!lstDocs.Requery


Dim Fnumber As Long
Dim Value As String
Dim blnFound As Boolean

If IsLoadedF("QueueAccounESC") = True Then
Forms![Case List]!ComWizESC.Visible = True
    If Forms!QueueAccounESC.lstFilesR.Column(11) <> 0 Then
        Fnumber = Forms!QueueAccounESC.lstFilesR.Column(11)
            If Not IsNull(Fnumber) Then
'                Dim Value As String
'                Dim blnFound As Boolean
                blnFound = False
                Dim J As Integer
                Dim A As Integer
                For J = 0 To Forms![Case List]!lstDocs.ListCount - 1
                   Value = Forms![Case List]!lstDocs.Column(0, J)
                   If InStr(Value, Fnumber) Then
                        blnFound = True
                         A = J
                        Forms![Case List].lstDocs.Selected(A) = True
                    Exit For
                    End If
                Next J
                
                If Not blnFound Then MsgBox ("Document not in the document list.")
                Forms![Case List]!lstDocs.SetFocus
                Else
                MsgBox ("Document not in the Document List.")
                Forms![Case List]!lstDocs.SetFocus
            End If
    Else
    
    Forms![Case List].cmdAddDoc.SetFocus
    
    End If
Else

If IsLoadedF("QueueESCtManager") = True Then
 If Forms![Case List]!CmdWizPS.Visible = True Then Forms![Case List]!CmdWizPS.Visible = False
    If Forms!QueueESCtManager.lstFilesR.Column(11) <> 0 Then
        Fnumber = Forms!QueueESCtManager.lstFilesR.Column(11)
            If Not IsNull(Fnumber) Then
                
                blnFound = False
'                Dim j As Integer
'                Dim A As Integer
                For J = 0 To Forms![Case List]!lstDocs.ListCount - 1
                   Value = Forms![Case List]!lstDocs.Column(0, J)
                   If InStr(Value, Fnumber) Then
                        blnFound = True
                         A = J
                        Forms![Case List].lstDocs.Selected(A) = True
                    Exit For
                    End If
                Next J
                
                If Not blnFound Then MsgBox ("Document not in the document list.")
                Forms![Case List]!lstDocs.SetFocus
                Else
                MsgBox ("Document not in the Document List.")
                Forms![Case List]!lstDocs.SetFocus
            End If
    Else
    
    Forms![Case List].cmdAddDoc.SetFocus
    
    End If
Else

Forms![Case List].ComWizESC.SetFocus
End If
End If


Forms![Case List]!Page120.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pageCheckRequest.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pgConflicts.Visible = False

Forms![Case List]!SCRAID = "AccEsc"
Forms![Case List].ComWizESC.SetFocus
Exit_Restart:
    Exit Sub

Err_Restart:
    MsgBox Err.Description
    Resume Exit_Restart
End Sub

Sub Limbo_OpenWizard(FileNumber As Long, FormName As String, Col As String)


Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean
On Error GoTo Err_VAsalesetting

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
        
            Case "Main", FormName, "Wizards"   ' leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

Set rstfiles = CurrentDb.OpenRecordset("SELECT casetypeid FROM Caselist WHERE FileNumber=" & FileNumber, dbOpenSnapshot)
If rstfiles!CaseTypeID <> 1 Then
MsgBox "This wizard can only be used with Foreclosure cases", vbCritical
Exit Sub
End If
FileLocks = True
    If LockFile(FileNumber) Then
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"
rstfiles.Close

DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If
Call LimboVisible(True)
If CurrentProject.AllForms("queVAsalesetting").IsLoaded Then Forms!foreclosuredetails.cmdPrint.Enabled = False
Forms!foreclosuredetails!WizardSource = FormName & Col

If CurrentProject.AllForms("queVAsalesettingsub").IsLoaded Then Forms!foreclosuredetails!WizardSource = "VALNNSetting"

Forms![foreclosuredetails].SetFocus
Forms![Case List]!cmdClose.Visible = False

Exit_VAsalesetting:
    Exit Sub

Err_VAsalesetting:
    MsgBox Err.Description
    Resume Exit_VAsalesetting
    
End Sub

Sub LimboVisible(SetVisible As Boolean)
Forms!foreclosuredetails!cmdWaiting.Visible = False
Forms!foreclosuredetails.cmdWizComplete.Visible = SetVisible
Forms!foreclosuredetails.cmdWizComplete.Caption = "LIMBO"
Forms!foreclosuredetails.cmdWizComplete.Enabled = True
Forms!foreclosuredetails.Sale.Locked = False
Forms!foreclosuredetails.cmdcloserestart.Visible = SetVisible
Forms!foreclosuredetails.cmdClose.Visible = False
Forms!foreclosuredetails.cmdAddFC.Visible = False
Forms!foreclosuredetails.cmdSelectFile.Visible = False
Forms!foreclosuredetails.cmdAudit.Visible = False
Forms![foreclosuredetails]!Page412.Visible = False
'Forms![ForeClosureDetails]!Page256.Visible = False
'Forms![foreclosuredetails]!Trustees.Visible = False
Forms![foreclosuredetails]!pgRealPropTaxes.Visible = False
Forms![foreclosuredetails]![Pre-Sale].SetFocus
'Forms!Journal!cmdNewJournalEntry.Visible = False
'Forms!Journal!cmdAttributes.Visible = False

End Sub

Sub Limbo_Prosecc(SourceWizard As String)

DoCmd.OpenForm "Limbo_Prosecc"
Forms!Limbo_Prosecc!Fnumber = Forms!foreclosuredetails!FileNumber
Select Case SourceWizard
Case "Limbo_MDWhite", "Limbo_MDYellow", "Limbo_MDRed" ', "Limbo_VAWhite", "Limbo_VAYellow", "Limbo_VARed", "Limbo_DCWhite", "Limbo_DCYellow", "Limbo_DCRed"
With Forms!Limbo_Prosecc
!DocBackMilAff.Visible = False
!DocBackDOA = Forms!foreclosuredetails!DocBackDOA
!DocBackSOD = Forms!foreclosuredetails!DocBackSOD
!DocBackLossMitPrelim = Forms!foreclosuredetails!DocBackLossMitPrelim
!DocBackLossMitFinal = Forms!foreclosuredetails!DocBackLossMitFinal
!DocBackLostNote.Visible = False
!DocBackOrigNote.Visible = False
!DocBackNoteOwnership = Forms!foreclosuredetails.DocBackNoteOwnership
If Forms!foreclosuredetails.txtClientSentNOI = "C" Then
!DocBackAff7105 = Forms!foreclosuredetails.DocBackAff7105
End If
End With

Case "Limbo_VAWhite", "Limbo_VAYellow", "Limbo_VARed"

With Forms!Limbo_Prosecc
!LMD.Caption = "VA"
!DocBackMilAff.Visible = False
!DocBackDOA = Forms!foreclosuredetails!DocBackDOA
!DocBackSOD.Visible = False '= Forms!foreclosuredetails!DocBackSOD
!DocBackLossMitPrelim.Visible = False ' = Forms!foreclosuredetails!DocBackLossMitPrelim
!DocBackLossMitFinal.Visible = False ' = Forms!foreclosuredetails!DocBackLossMitFinal
!DocBackLostNote = Forms!foreclosuredetails.DocBackLostNote
!DocBackOrigNote = Forms!foreclosuredetails.DocBackOrigNote
!DocBackNoteOwnership.Visible = False ' = Forms!foreclosuredetails.DocBackNoteOwnership
'If Forms!foreclosuredetails.txtClientSentNOI = "C" Then
!DocBackAff7105.Visible = False ' = Forms!foreclosuredetails.DocBackAff7105
'End If
End With


Case "Limbo_DCWhite", "Limbo_DCYellow", "Limbo_DCRed"

With Forms!Limbo_Prosecc
!LMD.Caption = "DC"
!DocBackMilAff.Visible = False
!DocBackDOA = Forms!foreclosuredetails!DocBackDOA
!DocBackSOD.Visible = False '= Forms!foreclosuredetails!DocBackSOD
!DocBackLossMitPrelim.Visible = False ' = Forms!foreclosuredetails!DocBackLossMitPrelim
!DocBackLossMitFinal.Visible = False ' = Forms!foreclosuredetails!DocBackLossMitFinal
!DocBackLostNote.Visible = False ' = Forms!foreclosuredetails.DocBackLostNote
!DocBackOrigNote.Visible = False ' = Forms!foreclosuredetails.DocBackOrigNote
!DocBackNoteOwnership.Visible = False ' = Forms!foreclosuredetails.DocBackNoteOwnership
'If Forms!foreclosuredetails.txtClientSentNOI = "C" Then
!DocBackAff7105.Visible = False ' = Forms!foreclosuredetails.DocBackAff7105
'End If
End With


End Select




End Sub
