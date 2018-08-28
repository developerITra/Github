VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Atty_Restart"
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

Private Sub cmdOK_Click()

'sarab
'************************************************** SA Cods 3/22/15 PS Advanced project
Dim newfilename As String
Dim selecteddoctype As Long
Dim fileextension As String
Dim DocDate As Date
Dim strSQL As String
Dim strSQLValues As String
Dim DocIDNo As Long
Dim clientShor As String
Dim rstBillReasons As Recordset
Dim NInvoiceID As Long
Dim rstqueue As Recordset
'NInvoiceID = Me.InvoiceID


strSQL = ""
strSQLValues = ""

DocDate = Now
selecteddoctype = 1585
Me.CaseFile.SetFocus
'Me.Frame153.Visible = False
'Me.JournalNote.Visible = False
Me.cmdOK.Visible = False
Me.cmdCancel.Visible = False

DoCmd.SetWarnings False

newfilename = "Atty Restart Cover Letter " & " " & Format$(Now(), "yyyymmdd hhnnss")

If Dir$(DocLocation & DocBucket(CaseFile) & "\" & CaseFile & "\" & newfilename) <> "" Then
    MsgBox CaseFile & " already exists.", vbCritical
    Exit Sub
End If

DoCmd.OutputTo acOutputForm, "Atty_Restart", acFormatPDF, DocLocation & DocBucket(CaseFile) & "\" & CaseFile & "\" & newfilename & ".pdf", False, "", 0

newfilename = newfilename & ".pdf"

strSQLValues = CaseFile & "," & selecteddoctype & ",'" & "" & "'," & GetStaffID() & ",'" & DocDate & "','" & Replace(newfilename, "'", "''") & "','" & Replace(newfilename, "'", "''") & "'"
strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
DoCmd.RunSQL (strSQL)





'Me.Frame153.Visible = True
'Me.JournalNote.Visible = True
Me.cmdOK.Visible = True
Me.cmdCancel.Visible = True


'If IsNull(Me.JournalNote) Then
strinfo = "Sent to Attorney Restart Review."
'Else
'strinfo = Me.JournalNote
'End If

strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & CaseFile & ",Now,GetFullName(),'" & strinfo & "',1 )"
DoCmd.RunSQL strSQLJournal


Dim rstvalumeintake As Recordset
Set rstvalumeintake = CurrentDb.OpenRecordset("Select * from ValumeRestart", dbOpenDynaset, dbSeeChanges)
With rstvalumeintake
.AddNew
!CaseFile = FileNumber
!Client = DLookup("ShortClientName", "ClientList", "ClientID = " & Forms![Case List]!ClientID)
!AttySent = Now
!AttySentC = 1
!Name = GetFullName()
.Update
End With
Set rstvalumeintake = Nothing


Dim rstwiz As Recordset
Set rstwiz = CurrentDb.OpenRecordset("select * from wizardqueuestats where current=true and filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
With rstwiz
.Edit

If IsNull(rstwiz!RestartWaiting) Then rstwiz!RestartWaiting = Now()
If Not IsNull(rstwiz!RestartComplete) Then rstwiz!RestartComplete = Null
.Update
End With
Set rstwiz = Nothing





Set rstqueue = CurrentDb.OpenRecordset("Select * FROM WizardSupportTwo where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!DateSentAttyrestart = Now()
!AttyMilestoneMgrRestart = Null
!AttyMilestoneRestart = Null
!AttyMilestonerestartReject = False
!RestartAttySentBy = StaffID
!AttyRestartRemark = "R"
.Update
End With
Set rstqueue = Nothing



DoCmd.SetWarnings True


'********************************************************* SA





Dim F As Form, FormClosed As Boolean

'QueueAccounESC
Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "querestartWaiting", "queRestart" '  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed


FileLocks = False
'***********************************

If IsLoadedF("QueueAccounESC") = True Then

    Forms!QueueAccounESC!lstFiles.Requery
    Forms!QueueAccounESC.Requery
    Dim cntr As Integer
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM QueueESC", dbOpenDynaset, dbSeeChanges)
    Do Until rstqueue.EOF
    cntr = cntr + 1
    rstqueue.MoveNext
    Loop
    Forms!QueueAccounESC!QueueCount = cntr
    Set rstqueue = Nothing
    '***************************************************
    
    '
    
       'If Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(17) <> 0 Then
          '  Fnumber = Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6)
                If Not IsNull(NInvoiceID) Then
                    Dim Value As String
                    Dim blnFound As Boolean
                    blnFound = False
                    Dim J As Integer
                    Dim A As Integer
                    For J = 0 To Forms!QueueAccounESC!lstFiles.ListCount - 1
                       Value = Forms!QueueAccounESC!lstFiles.Column(17, J)
                       If InStr(Value, NInvoiceID) Then
                            blnFound = True
                             A = J
                            Forms!QueueAccounESC.lstFiles.Selected(A) = True
                        Exit For
                        End If
                    Next J
                    
                    If Not blnFound Then MsgBox ("Document not in the document list.")
                    Forms!QueueAccounESC!lstFiles.SetFocus
                    Else
                    MsgBox ("Document not in the Document List.")
                    Forms!QueueAccounESC!lstFiles.SetFocus
                End If
       ' Else
        
      '  Forms![Case list].cmdAddDoc.SetFocus
        
     '   End If

End If

If WizESC Then WizESC = False


'DoCmd.Close acForm, "Audit_ESC"
End Sub

