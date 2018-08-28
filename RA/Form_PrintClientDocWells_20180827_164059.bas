VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PrintClientDocWells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim args() As String

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdOK_Click()
Dim Name As String, Title As String
Dim EffectiveDate As String
Dim AmendedPB As Integer
Dim LoanMod As Boolean
On Error GoTo Err_cmdOK_Click

If (optSign = 1) Then
  If (IsNull(lstContacts)) Then
    MsgBox "Select client contact.", vbExclamation
    Exit Sub
  End If
End If


If optSign = 2 Then
      Name = "_____________________________"
      Title = "___________________________"
ElseIf optSign = 3 Then
    If (IsNull(Forms!Foreclosureprint!Attorney)) Then
      MsgBox "Select designated attorney.", vbExclamation
      Exit Sub
    End If
    Select Case Forms!foreclosuredetails!State
'        Case "VA"
'            Name = DLookup("Name", "Staff", "[ID] = " & Forms!ForeclosurePrint!Attorney)
'            Title = Nz(DLookup("CommonWealthTitle", "Staff", "[ID] = " & Forms!ForeclosurePrint!Attorney))
'        Case "DC"
'            Name = TrusteeNames(0, 3)
'        '    If Forms!ForeclosureDetails!optSubstituteTrustees Then
'                Title = "Substitute Trustee"
'        '        Title = "Trustee"
'        '     End If
        Case "MD"
            Name = Forms!Foreclosureprint!Attorney.Column(1)
            Title = "Substitute Trustee"
    End Select
    If Left(args(0), 8) = "Military" Then
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)
With rstqueue
.Edit
!SCRAQueue9 = Date
.Update
End With
rstqueue.Close
End If
Else
    Name = lstContacts.Column(1)
    Title = lstContacts.Column(2)
End If

'If (args(1) = "Statement of Debt with Figures" Or args(1) = "Statement of Debt") Then
'If MsgBox("Is There A Loan Mod? ", vbYesNo) = vbYes Then
'EffectiveDate = InputBox("Agreement effective Date")
'AmendedPB = InputBox("Amended pricipal balance")
'ReportArgs = Name & "|" & Title & "|" & optSign & "|" & LoanMod & "|" & EffectiveDate & "|" & AmendedPB
'Else
'ReportArgs = Name & "|" & Title & "|" & optSign & "|" & LoanMod
'End If
'Else
ReportArgs = Name & "|" & Title & "|" & optSign
'End If


DoCmd.Close acForm, Me.Name
Call DoReport(args(1), (args(2)))

DoEvents


Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Current()
args = Split(Me.OpenArgs, "|")
lblWho.Caption = "Who will sign the " & args(0) & "?"
Me.Caption = "Print " & args(0)
End Sub

