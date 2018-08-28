VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PrintClientDocM"
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

On Error GoTo Err_cmdOK_Click

If CurrentProject.AllForms("ForeclosureDetails").IsLoaded Then
        Forms!Foreclosureprint!txtDesignator = 0  'Controls dates being added to military affidavits or not
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
        Forms!Foreclosureprint!txtDesignator = optSign  'Controls dates being added to military affidavits or not
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
        Case "DC", "MD"
            Name = trusteeNames(0, 3)
        '    If Forms!ForeclosureDetails!optSubstituteTrustees Then
            Title = "Substitute Trustee"
        '        Title = "Trustee"
        '     End If
        'Case "MD"
        '    Name = Forms!Foreclosureprint!Attorney.Column(1)
        '    Title = "Substitute Trustee"
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


ReportArgs = Name & "|" & Title & "|" & optSign
DoCmd.Close acForm, Me.Name


Dim JK As Integer
For JK = 1 To BorrowerMorgagorNamesCountSSN(Forms!foreclosuredetails!FileNumber)
CopyNo = JK

'#1272 - 11/13/2014
If Forms![Case List]!ClientID = 328 And args(1) = "MilitaryAffidavitMD SPLS" Then
    Call DoReport("MilitaryAffidavitMD", (args(2)))
Else
    Call DoReport(args(1), (args(2)))
End If

If args(2) = acPreview Then
Wait (1)
If JK < BorrowerMorgagorNamesCountSSN(Forms!foreclosuredetails!FileNumber) Then DoCmd.Close
End If
Next JK

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
