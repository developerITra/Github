VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DocsMissing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private Sub cmdOK_Click()
Dim rstdocs As Recordset, rstFC As Recordset, rstNames As Recordset, rstwiz As Recordset
Dim MissingInfo As String, i As Integer, ctr As Integer, JrlTxt As String, FileNbr As Long

On Error GoTo Err_cmdOK_Click


'If Not IsDate(txtReferralDate) Then MissingInfo = MissingInfo & "Referral Date, "
'If IsNull(cbxFileType) Then MissingInfo = MissingInfo & "File Type, "
Set rstdocs = CurrentDb.OpenRecordset("DocumentMissing", dbOpenDynaset, dbSeeChanges)

FileNbr = Forms!wizNOI!txtFileNumber


If Not IsNull(txtDoc1) Then
With rstdocs
    .AddNew
    !FileNbr = FileNbr
    !DocName = txtDoc1
    !DocsPndgby = StaffID
    !ID = FileNbr & 1
    .Update
End With
JrlTxt = txtDoc1
End If

If Not IsNull(txtDoc2) Then
With rstdocs
    .AddNew
    !FileNbr = FileNbr
    !DocName = txtDoc2
    !DocsPndgby = StaffID
    !ID = FileNbr & 2
    .Update
End With
JrlTxt = JrlTxt & ", " & txtDoc2
End If

If Not IsNull(txtDoc3) Then
With rstdocs
    .AddNew
    !FileNbr = FileNbr
    !DocName = txtDoc3
    !DocsPndgby = StaffID
    !ID = FileNbr & 3
    .Update
End With
JrlTxt = JrlTxt & ", " & txtDoc3
End If

If Not IsNull(txtDoc4) Then
With rstdocs
    .AddNew
    !FileNbr = FileNbr
    !DocName = txtDoc4
    !DocsPndgby = StaffID
    !ID = FileNbr & 4
    .Update
End With
JrlTxt = JrlTxt & ", " & txtDoc4
End If

If Not IsNull(txtDoc5) Then
With rstdocs
    .AddNew
    !FileNbr = FileNbr
    !DocName = txtDoc5
    !DocsPndgby = StaffID
    !ID = FileNbr & 5
    .Update
End With
JrlTxt = JrlTxt & ", " & txtDoc5
End If

If Not IsNull(txtDoc6) Then
With rstdocs
    .AddNew
    !FileNbr = FileNbr
    !DocName = txtDoc6
    !DocsPndgby = StaffID
    !ID = FileNbr & 6
    .Update
End With
JrlTxt = JrlTxt & ", " & txtDoc6
End If
 rstdocs.Close
'2/11/14
'lisa
    DoCmd.SetWarnings False
    strinfo = "The following documents necessary for an NOI are missing:  " & JrlTxt
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!DocsMissing!FileNbr,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True


'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  lrs.AddNew
'
'  lrs![FileNumber] = FileNbr
'  lrs![JournalDate] = Now
'  lrs![Who] = GetFullName()
'  lrs![Warning] = 100
'  lrs![Info] = "The following documents necessary for an NOI are missing:  " & JrlTxt & vbCrLf
'  lrs![Color] = 1
'  lrs.Update
'
'lrs.Close


Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNbr, dbOpenDynaset, dbSeeChanges)
If IsNull(rstqueue!NOICompleteDocsMsng) Then
With rstqueue
.Edit
!NOICompleteDocsMsng = Now
!NOIuser = StaffID
.Update
End With
End If
Set rstqueue = Nothing


MsgBox "The documents that you have indicated as missing have been recorded"

DoCmd.Close acForm, Me.Name
'Call ClearForm




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



