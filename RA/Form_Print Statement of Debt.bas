VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Statement of Debt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cbxReason_AfterUpdate()
Dim LPIPLus As Date
Dim strLPI As String
Me.txtOtherReason.Visible = False
LPIPLus = Forms!foreclosuredetails![LPIDate] + 1

If Me.cbxReason = "Other" Then
    Me.txtOtherReason.Visible = True
    Me.txtOtherReason = ""
ElseIf Me.cbxReason = "Deceased" Then
    Me.txtOtherReason = "are/is deceased"
ElseIf Me.cbxReason = "Non-occupancy" Then
    Me.txtOtherReason = "no longer occupy the subject property"
ElseIf Me.cbxReason = "Failed to pay taxes and insurance" Then
    
    If Forms![Case List]!ClientID = 567 Then 'Champion
    strLPI = "failed to make tax and/or insurance payments and the loan is called due as of " & Format$(Forms!foreclosuredetails![LPIDate], "mmmm dd, yyyy")
    'strLPI = strLPI + " and the loan was in default as of " & Format$(LPIPLus, "mmmm dd, yyyy") & ", and continuing each month thereafter"
    Me.txtOtherReason = strLPI
    Else
    strLPI = "failed to make tax and/or insurance payments and the loan is called due as of " & Format$(Forms!foreclosuredetails![LPIDate], "mmmm dd, yyyy")
    strLPI = strLPI + " and the loan was in default as of " & Format$(LPIPLus, "mmmm dd, yyyy") & ", and continuing each month thereafter"
    Me.txtOtherReason = strLPI
    End If
Else
    Me.txtOtherReason.Visible = False
End If

End Sub

Private Sub Check86_Click()
 
 If Check86 Then
    sfrmSODADJRates.Visible = True
'    Text92.Visible = True
'      Text94.Visible = True
      DescLabel.Visible = True
      Label11.Visible = True
      Label12.Visible = True
      FromDate.Visible = True
      AmountLabel.Visible = True
      
      Text100.Visible = False
      Text102.Visible = False
      Text104.Visible = False
'        Dim rstNS As Recordset
'            Dim strSQL As String
'                strSQL = "SELECT amount From StatementofDebt WHERE ([Desc] LIKE '%interest%') and filenumber=" & FileNumber
'                Set rstNS = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
'                If Not rstNS.EOF Then
                 '   Text94.SetFocus
'Label99.Caption = "" & CStr(Forms![ForeclosureDetails]!RemainingPBal) & ""
          '         Text94.Text = CStr(Forms![ForeclosureDetails]!RemainingPBal)

'                End If

 Else
          'Text92.Visible = False
      'Text94.Visible = False
    Text100.Visible = True
    Text102.Visible = True
    Text104.Visible = True
    sfrmSODADJRates.Visible = False
'    Desc_Label.Visible = False
'      Label11.Visible = False
'      Label12.Visible = False
'      FromDate.Visible = False
'      [Amount Label].Visible = False
 End If
 
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

Private Sub cmdOK_Click()
Dim statusMsg As String
'Dim PrintTo As Integer
On Error GoTo Err_cmdOK_Click

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

If cbxRateType.Visible = False Then
If Forms![Case List]!ClientID = 446 Then
DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt BOA|Statement of Debt with Figures BOA|" & Me.OpenArgs
Else  'Classy Indenting
   If Forms![Case List]!ClientID = 531 And Me.State = "MD" Then 'MDHC
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt with Figures MDCDMT|Statement of Debt with Figures MDCDMT|" & Me.OpenArgs
   ElseIf Forms![Case List]!ClientID = 456 And Me.State = "MD" Then 'M&T
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt with Figures M&T|Statement of Debt with Figures M&T|" & Me.OpenArgs
   ElseIf Forms![Case List]!ClientID = 361 And Me.State = "MD" Then
      DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt with Figures Ocwen|Statement of Debt with Figures Ocwen|" & Me.OpenArgs
   ElseIf Forms![Case List]!ClientID = 567 And Me.State = "MD" Then 'Champion
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt with Figures Cham|Statement of Debt with Figures Cham|" & Me.OpenArgs
   ElseIf Forms![Case List]!ClientID = 404 And Me.State = "MD" Then 'Bogman
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt with Figures|Statement of Debt with Figures Bogman|" & Me.OpenArgs
   ElseIf Forms![Case List]!ClientID = 328 And Me.State = "MD" Then 'SPLS
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt with Figures|Statement of Debt with Figures SPLS|" & Me.OpenArgs
   ElseIf Forms![Case List]!ClientID = 451 Then 'Dove
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt with Figures Dove|Statement of Debt with Figures Dove|" & Me.OpenArgs
   ElseIf Forms![Case List]!ClientID = 532 Then 'Selene
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt with Figures Selene|Statement of Debt with Figures Selene|" & Me.OpenArgs
   
    ElseIf Forms![Case List]!ClientID = 466 Then 'Select
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Combined Affidavit of Compliance Select|Combined Affidavit of Compliance Select|" & Me.OpenArgs
   
   
   ElseIf Forms![Case List]!ClientID = 523 Then 'GreenTree
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt with Figures GreenTree|Statement of Debt with Figures GreenTree|" & Me.OpenArgs
   ElseIf Forms![Case List]!ClientID = 385 And Forms![Foreclosureprint]!chComplianceAffidavit Then
    Call DoReport("Combined Affidavit of Compliance NationStar", Me.OpenArgs)
    Call DoReport("Nationstar Cover Sheet", Me.OpenArgs)
   Else
  
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt|Statement of Debt with Figures|" & Me.OpenArgs
If Forms![Case List]!ClientID = 444 Then
Call DoReport("PHH Cover Sheet", Me.OpenArgs)
End If

End If
End If
Else

    If Forms![Case List]!ClientID = 97 Then
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt|Statement of Debt with Figures JP|" & Me.OpenArgs
    Else
    
    If cbxRateType = "" Then
    MsgBox "Please select a rate type", vbCritical
    Exit Sub
    End If
If Forms![Case List]!ClientID = 6 Or Forms![Case List]!ClientID = 556 Then
DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt|Statement of Debt with Figures Wells|" & Me.OpenArgs
End If
End If
End If
cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdClear_Click()

On Error GoTo Err_cmdClear_Click

CurrentDb.Execute "DELETE * FROM StatementOfDebt WHERE FileNumber=" & FileNumber, dbSeeChanges
sfrmSOD.Requery

Exit_cmdClear_Click:
    Exit Sub

Err_cmdClear_Click:
    MsgBox Err.Description
    Resume Exit_cmdClear_Click
    
End Sub

Private Sub cmdCalcPerDiem_Click()

On Error GoTo Err_cmdCalcPerDiem_Click
'PerDiem = RemainingPBal * InterestRate / 365
PerDiem = RemainingPBal * InterestRate / 100 / 365

Exit_cmdCalcPerDiem_Click:
    Exit Sub

Err_cmdCalcPerDiem_Click:
    MsgBox Err.Description
    Resume Exit_cmdCalcPerDiem_Click
    
End Sub

Private Sub Form_Current()

   Me.TextIntrest.Caption = "From  " & [Text100] & " To " & [Text102] & "  at " & [Text104] & "%"
    If Forms![Case List]!ClientID = 567 Then Me.cbxReason.Visible = True
    If Forms![Case List]!ClientID = 446 Then Me.chVarRate.Visible = True
    
    If Forms!Foreclosureprint!chComplianceAffidavit Then 'Mei 9/23/15
        If Forms![Case List]!ClientID = 466 Then
          Me.Caption = "Combined Affidavit of Compliance Select"
        End If
    
    If Forms![Case List]!ClientID = 385 And Forms!Foreclosureprint!chComplianceAffidavit Then
               Check86.Visible = True
              Check86 = False
              Text92.Visible = True
              Text94.Visible = True
              Text100.Visible = True
              Text102.Visible = True
              Me.Text104.Visible = True
              Me.DescLabel.Visible = True
              Me.FromDate.Visible = True
              Me.Label11.Visible = True
              Me.Label12.Visible = True
              Me.AmountLabel.Visible = True
        '      sfrmSODADJRates.Visible = False
        '        If Check86 Then     'if there's adj. rate    mei
        '            sfrmSODADJRates.Visible = True
        '            Dim rstNS As Recordset
        '            Dim strSQL As String
        '                strSQL = "SELECT amount From StatementofDebt WHERE ([Desc] LIKE '%interest%') and filenumber=" & FileNumber
        '                Set rstNS = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
        '                If Not rstNS.EOF Then
        '    '                Text88.SetFocus
        '    '                Text88.Text = CStr(rstNS.Fields(0))
        '                    Text94.SetFocus
        '                    Text94 = CStr(rstNS.Fields(0))
        '
        '                End If
        '        Else
        '             sfrmSODADJRates.Visible = False
        '            Text88.Visible = False
        '            txtinterestAmt.Visible = False
                    
        '        End If
                 TxtPriorServicer = strPriorServicer
                Me.Caption = "Combined Affidavit of Compliance nationStar"
 
        
      End If
    End If
    
End Sub



Private Sub Form_Open(Cancel As Integer)

'Text94.Value = Format(Forms![ForeclosureDetails]!RemainingPBal, "Currency")
End Sub

Private Sub RemainingPBal_AfterUpdate()
Me.sfrmSOD.Requery
End Sub

Private Sub Text94_AfterUpdate()
Me.sfrmSOD.Requery
End Sub

Private Sub txtDueDate_DblClick(Cancel As Integer)
txtDueDate = Date
End Sub
