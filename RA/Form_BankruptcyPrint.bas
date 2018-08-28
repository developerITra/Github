VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_BankruptcyPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub cmdAcrobat_Click()
Call PrintDocs(-2)
End Sub

Private Sub cmdClear_Click()
On Error GoTo Err_cmdClear_Click

chODC = 0
ChDeclaration = 0
chMotionRelief = 0
chOrderGrantingRelief = 0
chOrderGrantReliefVA = 0
chWithdrawMotionRelief = 0
chConsentOrderTerminating = 0
chConsentOrderModifying = 0
chConsentOrderModifyingVA = 0
chAffDefault = 0
chWithdrawAffDefault = 0
chFinalOrderTerminating = 0
chAppearance = 0
chNoticeOfStay = 0
chTitleOrder = 0
chLossMitigation = 0
chMotionLoanMod = 0
chLoanModOrder = 0
chLoanModHearing = 0

chCh13Note362_Debtor = 0
chCh13MFR_Debtor = 0
PrePetPayHis = 0
PosPetpayHis = 0
PosPetFeeCos = 0
PosPetTaxInsAdvAdd = 0


Exit_cmdClear_Click:
    Exit Sub

Err_cmdClear_Click:
    MsgBox Err.Description
    Resume Exit_cmdClear_Click
    
End Sub

Private Sub PrintDocs(PrintTo As Integer)
Dim ReportName As String, strDate As String, dt As Date, rstBKdebt As Recordset
Dim nDays As Integer, dt0 As Date

'On Error GoTo Err_PrintDocs

If IsNull(cbxAttorney) Then
    MsgBox "Select an attorney to sign the document(s)", vbCritical
    Exit Sub
End If

If chReturnDocs Then
    DoCmd.OpenForm "Print Return Docs", , , "FileNumber = " & Forms![Case List]!FileNumber, , , PrintTo
End If

If chODC Then
    If IsNull(Trustee) Then
        MsgBox "Trustee is required", vbCritical
        Exit Sub
    End If
    If Nz(State) = "" Then
        MsgBox "Property Address (in Foreclosure screen) is incomplete: State is missing", vbCritical
        Exit Sub
    End If
End If

If chODC Then
    dt = DateChooserDialog(Date, "Date of Notice of 362")
    'strDate = InputBox("Date of Notice of 362:", "Notice of 362 Date", Date)
    If 0 <> dt Then
    'If strDate <> "" Then
        dt0 = dt
        strDate = CStr(dt)
        ' 2012.02.27 ODC date should be date printed.
        Forms!BankruptcyDetails!ODC = CStr(Date)
        AddStatus FileNumber, CStr(Date), "Sent notice of motion"
        On Error Resume Next
        If State = "VA" Then
            'UCase$(Format$(NextWeekDay(DateAdd("d",IIf([Districts.Name]="Eastern District of Virginia",14,21),Date())),"mmmm d"", ""yyyy"))
            If Name = "Eastern District of Virginia" Then nDays = 14 Else nDays = 21
        Else
            If State = "DC" Then nDays = 14 Else nDays = 17
        End If
        
        dt = NextWeekDay(DateAdd("d", nDays, dt0))
        dt = DateChooserDialog(dt, "Response date", "Response date for notice dated " & strDate)
        '& Forms!BankruptcyDetails!ODC)
        If 0 <> dt Then
            Me.txtResponseDate.Value = UCase$(Format$(dt, "mmmm d"", ""yyyy"))
            If State = "VA" _
            Then
                DoReport "ODC VA", PrintTo
            Else
                DoReport "ODC", PrintTo
            End If
        Resume Next
        End If
    End If
End If




If ChDeclaration Then
    If Me.Chapter = 13 Then
        If Forms![Case List]![ClientID] = 97 Or Forms![Case List]![ClientID] = 6 Or Forms![Case List]![ClientID] = 556 Or Forms![Case List]![ClientID] = 385 Then
            DoCmd.OpenForm "Print Declaration", , , "FileNumber=" & Forms!BankruptcyPrint!FileNumber, , acDialog, PrintTo
        End If
    End If
End If


If chMotionRelief Then

    If Forms![Case List]!ClientID = 446 Then
        'Do nothing
    Else
        Set rstBKdebt = CurrentDb.OpenRecordset("select filenumber from bkdebt where filenumber=" & FileNumber & " and prepost like 'pre*'", dbOpenDynaset)
            If rstBKdebt.EOF Then
                MsgBox "You must enter at least one line item in Pre-Petition Arrears before printing", vbCritical
                Exit Sub
            End If
        rstBKdebt.Close
        Set rstBKdebt = CurrentDb.OpenRecordset("select filenumber from bkdebt where filenumber=" & FileNumber & " and prepost like 'post*'", dbOpenDynaset)
            If rstBKdebt.EOF Then
                MsgBox "You must enter at least one line item in Post-Petition Arrears before printing", vbCritical
                Exit Sub
            End If
    End If
    
    If Forms![Case List]!ClientID = 446 And Me.Chapter = 7 Then
         DoReport "Motion for Relief BOA", PrintTo
    ElseIf Forms![Case List]!ClientID = 446 And Me.Chapter = 13 Then
        DoReport "Motion for Relief BOA Ch13", PrintTo
    'Else
'added on 4/30/15
    
    ElseIf Forms![Case List]!ClientID = 328 Then
        'DoReport "Motion for Relief_SLS", PrintTo
        DoCmd.OpenForm "Print Motion for Relief", , , "BankruptcyID=" & BankruptcyID, , , "Motion for Relief SLS|" & PrintTo & "|0"
'*****
    ElseIf Forms![Case List]!ClientID = 97 And Me.Chapter = 13 Then
        DoCmd.OpenForm "Print Motion for Relief", , , "BankruptcyID=" & BankruptcyID, , , "Motion for Relief Chase|" & PrintTo & "|0"
        
    ElseIf Forms![Case List]!ClientID = 385 Then
        DoCmd.OpenForm "Print Motion for Relief", , , "BankruptcyID=" & BankruptcyID, , , "Motion for Relief NationStar|" & PrintTo & "|0"

    Else
        DoCmd.OpenForm "Print Motion for Relief", , , "BankruptcyID=" & BankruptcyID, , , "Motion For Relief|" & PrintTo & "|0"
    End If
End If


If chOrderGrantingRelief Then DoReport "Order Granting Relief", PrintTo
If Me.chOrderGrantReliefVA Then
  If (TestDocument("ConsentGrantVA") = True) Then
    DoReport "Order Granting Relief VANew", PrintTo
    'Call Doc_ConsentGrantRelief(True)
  End If
End If

If chWithdrawMotionRelief Then
    DoReport "Withdraw Motion for Relief", PrintTo
    If MsgBox("Update timeline: Motion Withdrawn = Yes ?", vbYesNo + vbQuestion) = vbYes Then
        Forms!BankruptcyDetails!MotionWithdrawn = True
    End If
End If

If chConsentOrderTerminating Then DoCmd.OpenForm "Print Consent Order Terminating " & Chapter, , , "BankruptcyID=" & BankruptcyID, , , PrintTo

If chConsentOrderModifying Then DoCmd.OpenForm "Print Consent Order Modifying", , , "BankruptcyID=" & BankruptcyID, , , PrintTo

If chConsentOrderModifyingVA Then DoCmd.OpenForm "Print Consent Order Modifying VA", , , "BankruptcyID=" & BankruptcyID, , , PrintTo

If chAffDefault Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    Set rstBKdebt = CurrentDb.OpenRecordset("select filenumber from bkdebt where filenumber=" & FileNumber & " and prepost='pre'", dbOpenDynaset)
    If rstBKdebt.EOF Then
    MsgBox "You must enter at least one line item in Pre-Petition Arrears before printing", vbCritical
    Exit Sub
    End If
    rstBKdebt.Close
    Set rstBKdebt = CurrentDb.OpenRecordset("select filenumber from bkdebt where filenumber=" & FileNumber & " and prepost='post'", dbOpenDynaset)
    If rstBKdebt.EOF Then
    MsgBox "You must enter at least one line item in Post-Petition Arrears before printing", vbCritical
    Exit Sub
    End If
    
    DoCmd.OpenForm "Print Affidavit of Default", , , "BankruptcyID=" & BankruptcyID, , , PrintTo
End If

If chWithdrawAffDefault Then
    DoReport "Withdraw Affidavit of Default", PrintTo
    AddStatus FileNumber, Date, "Affidavit of Default withdrawn"
    If Not IsNull(Forms!BankruptcyDetails![3rdAff]) Then
        Forms!BankruptcyDetails![3rdAff] = Null
    ElseIf Not IsNull(Forms!BankruptcyDetails![2ndAff]) Then
        Forms!BankruptcyDetails![2ndAff] = Null
    ElseIf Not IsNull(Forms!BankruptcyDetails!Affidavit) Then
        Forms!BankruptcyDetails!Affidavit = Null
    End If
End If

If chFinalOrderTerminating Then DoReport "Final Order Terminating", PrintTo

If chAppearance Then DoReport "Line Entering Appearance", PrintTo

If chNoticeOfStay Then DoReport "Notice of Stay", PrintTo

If chPayoff Then DoCmd.OpenForm "Print Payoff", , , "FileNumber=" & FileNumber, , , PrintTo & "|BK"

If chTitleOrder Then DoReport "Title Order BK", PrintTo

If chLossMitigation Then DoReport "EMC LMT Letter", PrintTo

If chCh13Note362_Debtor Then
  'If Me.State = "VA" Then

    strDate = InputBox("Date of Notice of 362:", "Notice of 362 Date", Date)
    If strDate <> "" Then
        Forms!BankruptcyDetails!ODC = strDate
        AddStatus FileNumber, strDate, "Sent Debtor notice of motion"
        DoReport "ODC Debtor VA", PrintTo
    End If
  'End If
End If

If chCh13MFR_Debtor Then
 'If Me.State = "VA" Then
    DoCmd.OpenForm "Print Motion for Relief", , , "BankruptcyID=" & BankruptcyID, , , "Motion For Relief Debtor VA|" & PrintTo & "|0"
 'End If
End If

If chCh13Note362_CoDebtor Then
  'If Me.State = "VA" Then

    strDate = InputBox("Date of Notice of 362:", "Notice of 362 Date", Date)
    If strDate <> "" Then
        Forms!BankruptcyDetails!ODC_CoDebtor = strDate
        AddStatus FileNumber, strDate, "Sent Co-Debtor notice of motion"
        DoReport "ODC CoDebtor VA", PrintTo
    End If
 ' End If

End If

If chCh13MFR_CoDebtor Then
 'If Me.State = "VA" Then
    DoCmd.OpenForm "Print Motion for Relief", , , "BankruptcyID=" & BankruptcyID, , , "Motion For Relief CoDebtor VA|" & PrintTo & "|1"
' End If
End If

If chMotionLoanMod Then
    DoCmd.OpenForm "Print Motion Loan Mod", , , "BankruptcyID=" & BankruptcyID, , , "Motion For Loan Modification|Motion For Loan Modification|" & PrintTo

End If



If chLoanModOrder Then
  If (State = "VA") Then
    DoCmd.OpenForm "Print Motion Loan Mod", , , "BankruptcyID=" & BankruptcyID, , , "Order for Loan Modification|Loan Mod Order VA|" & PrintTo


  Else
    DoReport "Loan Mod Order", PrintTo
  End If
End If

If chCreateLabel Then
    DoCmd.OpenForm "Getlabel"
Else
End If

If chLoanModHearing Then
    DoReport "Loan Mod Hearing VA", PrintTo

End If


If PrePetPayHis Then
    DoReport "Exhibit", PrintTo
    DoReport "Pre-PetitionPaymentHistory", PrintTo
End If

If PosPetpayHis Then
     DoReport "Exhibit", PrintTo
    DoReport "Post-PetitionPaymentHistory", PrintTo
End If

If PosPetFeeCos Then
     DoReport "Exhibit", PrintTo
    DoReport "Post-PetitionFeeBreakdownAddendum", PrintTo
End If

If PosPetTaxInsAdvAdd Then
     DoReport "Exhibit", PrintTo
    DoReport "Post-PetitionTaxes-InsuranceAddendum", PrintTo
End If




Exit Sub

Err_PrintDocs:
    MsgBox Err.Description
    Exit Sub

End Sub

Private Sub Form_Current()
  Me.Caption = "Print Bankruptcy " & [CaseList.FileNumber] & " " & [PrimaryDefName]
  
  Me.pgChapter13.Visible = Nz((Chapter = 13), False)
  chLoanModHearing.Enabled = (State = "VA")
  
 ' If Me.State = "VA" Then
 '   cbxAttorney.RowSource = "Select Staff.Name & ', ' & [CommonWealthTitle] AS CWRep , Staff.ID " & _
 '                      "FROM Staff " & _
 '                      "WHERE (((Staff.CommonwealthTitle) Is Not Null)) " & _
 '                      "ORDER BY Staff.CommonwealthTitle, Staff.Sort;"
 ' ElseIf Me.State = "MD" Then
''''    cbxAttorney.RowSource = "SELECT Staff.Name & ', Esq.', Staff.ID FROM Staff WHERE (((Staff.Attorney)=True)) and staff.active = true ORDER BY Staff.Sort;"
 ' Else
 '   cbxAttorney.RowSource = "SELECT Staff.Name, Staff.ID FROM Staff WHERE (((Staff.Attorney)=True)) ORDER BY Staff.Sort;"
'
 ' End If
'''''    cbxAttorney.Value = cbxAttorney.Column(0, 0)
    
 
 If ([Forms]![Case List]!Active = False) Then
 
 chODC.Enabled = False
 ChDeclaration.Enabled = False
 chMotionRelief.Enabled = False
 chOrderGrantingRelief.Enabled = False
 chOrderGrantReliefVA.Enabled = False
 chWithdrawMotionRelief.Enabled = False
 chConsentOrderTerminating.Enabled = False
 chConsentOrderModifying.Enabled = False
 chConsentOrderModifyingVA.Enabled = False
 chAffDefault.Enabled = False
 chWithdrawAffDefault.Enabled = False
 cbxAttorney.Enabled = False
 chElectronicSignature.Enabled = False
 chFinalOrderTerminating.Enabled = False
 chAppearance.Enabled = False
 chNoticeOfStay.Enabled = False
 chPayoff.Enabled = False
 chTitleOrder.Enabled = False
 chLossMitigation.Enabled = False
 NotaryID.Enabled = False
 chCh13Note362_Debtor.Enabled = False
 chCh13MFR_Debtor.Enabled = False
 chCh13Note362_CoDebtor.Enabled = False
 chCh13MFR_CoDebtor.Enabled = False
 chMotionLoanMod.Enabled = False
 chLoanModOrder.Enabled = False
 chLoanModHearing.Enabled = False
 cbxAttorney.Enabled = False
 chElectronicSignature.Enabled = False
 NotaryID.Enabled = False
 
 
 
 
 Label195.Visible = True
 
 
 End If
 

    
    


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

Private Sub cmdWord_Click()
Call PrintDocs(-1)
End Sub

Private Sub cmdPrint_Click()
Call PrintDocs(acViewNormal)
End Sub

Private Sub cmdView_Click()
Call PrintDocs(acPreview)
End Sub
Private Sub Command204_Click()
On Error GoTo Err_Command204_Click

    Dim stDocName As String
    Dim te As String
    stDocName = "Exhibit"
    Dim ss As String
    ss = "dddddddddddddd"
    
    'DoCmd.OpenReport ReportName:="Exhibit", View:=acViewPreview, _
    'OpenArgs:=ss
    DoCmd.OpenReport stDocName, , , , , OpenArgs:=ss

    

Exit_Command204_Click:
    Exit Sub

Err_Command204_Click:
    MsgBox Err.Description
    Resume Exit_Command204_Click
    
End Sub
