VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditInvestor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdInvClient_Click()
Dim C As Recordset

On Error GoTo Err_cmdInvClient_Click
Set C = CurrentDb.OpenRecordset("SELECT * FROM ClientList WHERE ClientID = " & Forms![Case List]!ClientID, dbOpenSnapshot)

If Not C.EOF Then
    If (C!ClientID = 567 And Forms![Case List]!State = "MD") Then
        Investor = "Nationstar Mortgage LLC d/b/a Champion Mortgage Company of Texas"
        InvestorAddr = C("StreetAddress") & IIf(IsNull(C("StreetAddr2")), "", vbNewLine & C("StreetAddr2")) & _
        vbNewLine & C("City") & ", " & C("State") & " " & C("ZipCode")
    ElseIf (C!ClientID = 567 And Forms![Case List]!State = "VA") Then
        Investor = "Nationstar Mortgage LLC, doing business in the Commonwealth of Virginia as Virginia Nationstar LLC d/b/a Champion Mortgage Company"
        InvestorAddr = C("StreetAddress") & IIf(IsNull(C("StreetAddr2")), "", vbNewLine & C("StreetAddr2")) & _
        vbNewLine & C("City") & ", " & C("State") & " " & C("ZipCode")
    Else
        Investor = C("ClientNameAsInvestor")
        InvestorAddr = C("StreetAddress") & IIf(IsNull(C("StreetAddr2")), "", vbNewLine & C("StreetAddr2")) & _
        vbNewLine & C("City") & ", " & C("State") & " " & C("ZipCode")
    End If
End If


C.Close

Exit_cmdInvClient_Click:
    Exit Sub

Err_cmdInvClient_Click:
    MsgBox Err.Description
    Resume Exit_cmdInvClient_Click
End Sub

Private Sub cmdupdate_Click()
NameJournal = ""
Call makeEditInvestorJournaltext
If makeEditInvestorJournaltext = "" Then
DoCmd.Close acForm, Me.Name
Exit Sub

End If

Dim rstFCdetails As Recordset
Dim strinfo As String
Dim strSQLJournal As String

    DoCmd.SetWarnings False
    strinfo = NameJournal
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms![Case List]!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    NameJournal = vbNullString
    Forms![Case List]!Investor = Forms!EditInvestor!Investor
    Forms![Case List]!InvestorAddr = Forms!EditInvestor!InvestorAddr
    Forms![Case List]!AIF = Forms!EditInvestor!AIF

'
'DoCmd.SetWarnings False
'DoCmd.RunSQL ("UPDATE CaseList set Investor = '" & Forms!EditInvestor!Investor & "' AND InvestorAddress = '" & Forms!EditInvestor!InvestorAddress & "' AND AIF = " & Forms!EditInvestor!AIF & " WHERE FileNumber = " & Forms!EditInvestor!FileNumber)
'
'DoCmd.SetWarnings True
'DoCmd.RunCommand acCmdSaveRecord
DoCmd.Close acForm, Me.Name
Forms!Journal.Requery


'DoCmd.SetWarnings False
'
'Forms![Case list].Requery
'DoCmd.Close acForm, Me.Name
'DoCmd.SetWarnings True
End Sub


Public Function makeEditInvestorJournaltext()
NameJournal = ""

If Nz(Investor) <> Nz(Forms![Case List]!Investor) Then
    If IsNull(Investor) And Not IsNull(Forms![Case List]!Investor) Then NameJournal = NameJournal + "Removed Investor " & Forms![Case List]!Investor & ". "
    If IsNull(Forms![Case List]!Investor) And Not IsNull(Investor) Then NameJournal = NameJournal + "Added Investor:" & Investor & ". "
    If Not IsNull(Forms![Case List]!Investor) And Not IsNull(Investor) Then NameJournal = NameJournal + "Edit Investor from " & Forms![Case List]!Investor & " To " & Investor & ". "
End If

If Nz(InvestorAddr) <> Nz(Forms![Case List]!InvestorAddr) Then
    If IsNull(InvestorAddr) And Not IsNull(Forms![Case List]!InvestorAddr) Then NameJournal = NameJournal + "Removed Investor Address " & Forms![Case List]!InvestorAddr & ". "
    If IsNull(Forms![Case List]!InvestorAddr) And Not IsNull(InvestorAddr) Then NameJournal = NameJournal + "Added Investor Address:" & InvestorAddr & ". "
    If Not IsNull(Forms![Case List]!InvestorAddr) And Not IsNull(InvestorAddr) Then NameJournal = NameJournal + "Edit Investor Address from " & Forms![Case List]!InvestorAddr & " To " & InvestorAddr & ". "
End If

If Nz(AIF) <> Nz(Forms![Case List]!AIF) Then
    If AIF <> -1 And Forms![Case List]!AIF = -1 Then NameJournal = NameJournal + "Unchecked AIF. "
    If (Forms![Case List]!AIF <> -1 Or IsNull(Forms![Case List]!AIF)) And AIF = -1 Then NameJournal = NameJournal + "Checked AIF. "
End If

makeEditInvestorJournaltext = NameJournal
End Function

Private Sub Form_Current()
If Forms![Case List]!CaseType.Value = "Foreclosure" Then
    If Forms![Case List].[ClientID] = 97 Then 'JPM
        Me.AIF.Value = False
        Me.AIF.Enabled = False
    ElseIf Forms![Case List].[ClientID] = 451 Then 'Dove
        Me.AIF.Value = False
        Me.AIF.Enabled = False
'    ElseIf Forms![Case List].[ClientID] = 404 Then  'Bogman
'        Me.AIF.Value = False
'        Me.AIF.Enabled = False
    Else
    If Forms![Case List].[ClientID] = 446 Then 'BOA
    Me.AIF.Value = False
    Me.AIF.Enabled = False
'    Else
'    AIF.Value = True
'    AIF.Enabled = True
    End If
    End If
    End If
End Sub
