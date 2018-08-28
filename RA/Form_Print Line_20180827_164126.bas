VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim args() As String
Dim PrintTo As Integer

Private Sub cmdCancel_Click()
DoCmd.Close
End Sub

Private Sub cmdClear_Click()
Me.txtLineDescription = ""
End Sub

Private Sub cmdPrintLine_Click()


If Len(Me.txtLineDescription & "") = 0 Then
MsgBox "Description cannot be blank", vbOKOnly
Exit Sub
Else

    If MsgBox("Do you want to update Line Staying case field?", vbYesNo) = vbYes Then
    
      Forms!foreclosuredetails!sfrmLineStayingCase!LineStayingcase = Now()
      AddStatus FileNumber, Now(), "Line staying case sent to court"
        
     'Forms!ForeclosureDetails!sfrmStatus.Requery
    
       DoCmd.SetWarnings False
        strinfo = "Line sent to court, staying case"
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
        Forms!Journal.Requery
    
    End If


Dim Name As String, Title As String


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
    
    If CurrentProject.AllForms("ForeclosureDetails").IsLoaded Then
'        Forms!ForeclosurePrint!txtDesignator = optSign
'         Forms!ForeclosurePrint!txtDesignatedAttorney = optSign
    Else
        
    End If
    
    
    If (IsNull(Forms!Foreclosureprint!Attorney)) Then
      MsgBox "Select designated attorney.", vbExclamation
      Exit Sub
    End If
    Select Case Forms!foreclosuredetails!State
        Case "VA"
            Name = DLookup("Name", "Staff", "[ID] = " & Forms!Foreclosureprint!Attorney)
            Title = Nz(DLookup("CommonWealthTitle", "Staff", "[ID] = " & Forms!Foreclosureprint!Attorney))
        Case "DC"
            Name = trusteeNames(0, 3)
        '    If Forms!ForeclosureDetails!optSubstituteTrustees Then
                Title = "Substitute Trustee"
        '        Title = "Trustee"
        '     End If
        Case "MD"
            Name = Forms!Foreclosureprint!Attorney.Column(1)
            Title = "Attorney for Substitute Trustee"
    End Select

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



'Call DoReport(args(1), (args(2)))


 



    If Len(Me.txtFileNumber & "") = 0 Then
       Me.txtFileNumber = Forms![Case List]!FileNumber
       If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
       DoCmd.Close acForm, Me.Name
       Call DoReport("Line", (args(0)))
       
       DoEvents
    Else
        If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
        DoCmd.Close acForm, Me.Name
        Call DoReport("Line", (args(0)))
        DoEvents
    End If
End If

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click

End Sub



Private Sub Form_Current()
'PrintTo = Int(Split(Me.OpenArgs, "|")(0))
args = Split(Me.OpenArgs, "|")
'lblWho.Caption = "Who will sign the " & args(2) & "?"
'Me.Caption = "Print " & args(0)

End Sub
