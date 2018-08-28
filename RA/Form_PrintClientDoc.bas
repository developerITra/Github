VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PrintClientDoc"
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

Forms!Foreclosureprint!txtDesignator = 0
Forms!Foreclosureprint!txtDesignatedAttorney = 0

Dim Name As String, Title As String
Dim EffectiveDate As String
Dim AmendedPB As Integer
Dim LoanMod As Boolean
On Error GoTo Err_cmdOK_Click

If CurrentProject.AllForms("ForeclosureDetails").IsLoaded Then
        Forms!Foreclosureprint!txtDesignator = 0
        Forms!Foreclosureprint!txtDesignatedAttorney = 0
Else
End If

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
        Forms!Foreclosureprint!txtDesignator = optSign
         Forms!Foreclosureprint!txtDesignatedAttorney = optSign
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
            If Forms!Foreclosureprint!chMilitaryAffidavitActive Then
                  Name = trusteeNames(0, 3)
                  Title = "Substitute Trustee"
            ElseIf Forms!Foreclosureprint!chNoteOwnership Then
                  Name = trusteeNames(0, 3)
                  Title = "Substitute Trustee"
            ElseIf Forms!Foreclosureprint!chAffMD7105 Then
                   Name = trusteeNames(0, 3)
                  Title = "Substitute Trustee"
            Else
                Name = Forms!Foreclosureprint!Attorney.Column(1)
                Title = "Substitute Trustee"
            End If
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
If Forms!Foreclosureprint!chSOD2 Then
   ' args(2) = prntTo
End If
'PDFS
If (args(2)) = -2 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit Wells" Then
    Call DoReport("45 Day Notice Affidavit", (args(2)))
ElseIf (args(2)) = -2 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit BOA" Then
     Call DoReport("45 Day Notice Affidavit", (args(2)))
ElseIf (args(2)) = -2 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit Chase" Then
    Call DoReport("45 Day Notice Affidavit", (args(2)))
ElseIf (args(2)) = -2 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit" Then
     Call DoReport("45 Day Notice Affidavit", (args(2)))
'Word Docs
ElseIf (args(2)) = -1 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit Wells" Then
    Call DoReport("45 Day Notice Affidavit", (args(2)))
ElseIf (args(2)) = -1 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit BOA" Then
     Call DoReport("45 Day Notice Affidavit", (args(2)))
ElseIf (args(2)) = -1 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit Chase" Then
    Call DoReport("45 Day Notice Affidavit", (args(2)))
ElseIf (args(2)) = -1 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit" Then
     Call DoReport("45 Day Notice Affidavit", (args(2)))
'Views
ElseIf (args(2)) = 2 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit Wells" Then
    Call DoReport("45 Day Notice Affidavit", (args(2)))
ElseIf (args(2)) = 2 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit BOA" Then
     Call DoReport("45 Day Notice Affidavit", (args(2)))
ElseIf (args(2)) = 2 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit Chase" Then
    Call DoReport("45 Day Notice Affidavit", (args(2)))
ElseIf (args(2)) = 2 And Forms!Foreclosureprint!txtDesignatedAttorney = 3 And args(1) = "45 Day Notice Affidavit" Then
     Call DoReport("45 Day Notice Affidavit", (args(2)))
Else 'Normal
Call DoReport(args(1), (args(2)))
End If

DoEvents


Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Current()

If Len(Me.OpenArgs) > 0 Then
    args = Split(Me.OpenArgs, "|")
    lblWho.Caption = "Who will sign the " & args(0) & "?"
    Me.Caption = "Print " & args(0)
End If
End Sub

