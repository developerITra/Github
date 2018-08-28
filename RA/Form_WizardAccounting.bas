VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_WizardAccounting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub CmdAccouLitigBill_Click()
Dim FileNumber As Long
       
FileNumber = InputBox("Enter the File Number", "Litigation Billing")
If IsNull(FileNumber) Then
Exit Sub
End If

AddToList (FileNumber)
LitigationBillingCallFromQueue FileNumber

End Sub

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdSCRA1_Click()

    Dim stDocName As String
  
    stDocName = "queSCRA1"
    DoCmd.OpenForm stDocName

End Sub



Private Sub cmdSCRA2_Click()
Dim stDocName As String

    stDocName = "queSCRA2"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA3_Click()
Dim stDocName As String

    stDocName = "queSCRA3"
    DoCmd.OpenForm stDocName
End Sub


Private Sub cmdSCRA4a_Click()
Dim stDocName As String

    stDocName = "queSCRA4a"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA4b_Click()
Dim stDocName As String

    stDocName = "queSCRA4b"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA5_Click()
Dim stDocName As String

    stDocName = "queSCRA5"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA6_Click()
Dim stDocName As String

    stDocName = "queSCRA6"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA7_Click()
Dim stDocName As String

    stDocName = "queSCRA7"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA8_Click()
Dim stDocName As String

    stDocName = "queSCRA8"
    DoCmd.OpenForm stDocName
End Sub

Private Sub cmdSCRA9_Click()
Dim stDocName As String

    stDocName = "queSCRA9"
    DoCmd.OpenForm stDocName
End Sub

Private Sub cmdSCRA9Waiting_Click()
Dim stDocName As String

    stDocName = "queSCRA9Waiting"
    DoCmd.OpenForm stDocName
End Sub


Private Sub cmdSCRAUnionNew_Click()
 stDocName = "queSCRAFCNew"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Command75_Click()
stDocName = "queSCRABK"
    DoCmd.OpenForm stDocName
End Sub

Private Sub ComESc_Click()
WizESC = True
Dim FileNumber As Long
       
FileNumber = InputBox("Enter the File Number", "Escrow Audit")
If IsNull(FileNumber) Then
Exit Sub
End If


Dim rstfiles As Recordset, stDocName As String, stLinkCriteria As String, F As Form, FormClosed As Boolean

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

AddToList (FileNumber)
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

'Forms![Case list]!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], [doctitleid] AS DocType , Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND Filespec IS NOT NULL and DeleteDate is null"
Forms![Case List]!lstDocs.Requery





Forms![Case List]!ComWizESC.Visible = True
    


Forms![Case List].ComWizESC.SetFocus


Forms![Case List]!Page120.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pageCheckRequest.Visible = False
Forms![Case List]!pgDocRequest.Visible = False
Forms![Case List]!pgConflicts.Visible = False

'Forms![Case list]!SCRAID = "AccEsc"
Forms![Case List].ComWizESC.SetFocus













End Sub

Private Sub ComPSAd_Click()
Dim FileNumber As Long
       
FileNumber = InputBox("Enter the File Number", "Post Sale Advanced Sale")
If IsNull(FileNumber) Then
Exit Sub
End If

AddToList (FileNumber)
PSAdvancedCostsCallFromQueue FileNumber
End Sub
