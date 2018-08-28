VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_New Case"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
Dim C As Recordset, rstwiz As Recordset, rstDIL As Recordset, rstInternet As Recordset, rstFCtitle As Recordset
Dim FileNum As Long

On Error GoTo Err_cmdOK_Click

If IsNull(PrimaryDefName) Or IsNull(ReferralDate) Or IsNull(CaseType) Or IsNull(Client) Then
    MsgBox "Referral Date, Project Name, Client, and Type of Case are required.", vbCritical
    Exit Sub
End If

Select Case optCaseNumber
    Case 1
        FileNum = ReserveNextCaseNumber()
    Case 2
        If IsNull(txtCaseNumber) Then
            MsgBox "Enter the new case number.", vbCritical
            Exit Sub
        End If
        FileNum = txtCaseNumber
        Set C = CurrentDb.OpenRecordset("SELECT * FROM CaseList WHERE FileNumber = " & FileNum, dbOpenSnapshot)
        If Not C.EOF Then
            MsgBox "Case number " & FileNum & " is in use by project name " & C("PrimaryDefName"), vbCritical
            C.Close
            Exit Sub
        End If
        C.Close
End Select
'WizardOueueStats
If StaffID = 0 Then Call GetLoginName

Set C = CurrentDb.OpenRecordset("CaseList", dbOpenDynaset, dbSeeChanges)
C.AddNew
C("FileNumber") = FileNum
C("ReferralDate") = ReferralDate
C("PrimaryDefName") = PrimaryDefName
C("CaseTypeID") = CaseType
C("ClientID") = Client
C("ClientNumber") = ClientNumber
C("Active") = True
C("OnStatusReport") = True
C("OpenDate") = Now()
C("OpenBy") = StaffID
C.Update
C.Close


'added on 5_7_15
'Dim rs As Recordset

'Set rs = CurrentDb.OpenRecordset("select * from CaseList where FileNumber = " & Me.FileNum, dbOpenDynaset, dbSeeChanges)

'rs.Edit
'rs!BillCaseUpdateReasonID = 34
'rs.Update

'rs.Close
'Set rs = Nothing


Set rstwiz = CurrentDb.OpenRecordset("WizardQueueStats", dbOpenDynaset, dbSeeChanges)
With rstwiz
.AddNew
!FileNumber = FileNum
!Count = 1
!RSIcomplete = Date
.Update
End With
rstwiz.Close


Set rstwiz = CurrentDb.OpenRecordset("WizardSupportTwo", dbOpenDynaset, dbSeeChanges)
With rstwiz
.AddNew
!FileNumber = FileNum
!Count = 1
.Update
End With
rstwiz.Close

DoCmd.SetWarnings False
Dim Shortclient As String
Dim cbxLoanTypetext As String





'create record in new FC DIL table
Set rstDIL = CurrentDb.OpenRecordset("FCDIL", dbOpenDynaset, dbSeeChanges)
With rstDIL
.AddNew
    !FileNumber = FileNum
    .Update
    .Close
End With

'Add Record to FCtitle Table
Set rstFCtitle = CurrentDb.OpenRecordset("FCtitle", dbOpenDynaset, dbSeeChanges)
With rstFCtitle
    .AddNew
    !FileNumber = FileNum
    .Update
    .Close
End With

'create record in new Internet Sources table
Set rstInternet = CurrentDb.OpenRecordset("InternetSites", dbOpenDynaset, dbSeeChanges)
With rstInternet
.AddNew
    !FileNumber = FileNum
    .Update
    .Close
End With

AddStatus FileNum, ReferralDate, "Received referral"
DoCmd.Close acForm, "New Case"
DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNum

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
lblNextNumber.Caption = "Next available number (" & ReadNextCaseNumber() & ")"
ReferralDate = Now()
End Sub
