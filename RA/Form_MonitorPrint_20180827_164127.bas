VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MonitorPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub chWithdrawBond_AfterUpdate()
chWithdrawBondBK.Enabled = chWithdrawBond
End Sub

Private Sub cmdClear_Click()
On Error GoTo Err_cmdClear_Click
chLine = 0
Me.chWarranty = 0
Me.chQCDeed = 0
ch45Day = 0
Me.chOrderRelease = 0
chTitleOrder = 0
chDeedOfApp = 0
chDOARecordingCover = 0
chDOAAffidavit = 0
chDocket = 0
chAffMD7105 = 0
chDOTAffidavit = 0
ChNoteAffidavit = 0
ChCollateralFileAffidavit = 0
Ch14207Affidavit = 0
'Me.chOrderGrantingDefault = 0
chNoteOwnership = 0
Me.chBaileeLetter = 0
Me.chIntervene = 0
Me.chOrderIntervene = 0
Me.chMotionRelease = 0
Me.chAffCoverLetter = 0
chBondOrder = 0
chPSCover = 0
chAffidavitOfService = 0
chDebtorLetterLabels = 0
chSOD = 0
chSOD2 = 0
chMilitaryAffidavitActive = 0
chMilitaryAffidavit = 0
chMilitaryAffidavitNoSSN = 0


chLossMitPrelim = 0
chLossMitFinal = 0
chLossMitApp = 0
chForecloseMed = 0

chFairDebt = 0
chFairDebtLabels = 0
chHUDOcc = 0
chIRSNotice = 0
chNotice = 0
chNoticeLabels = 0
chCountyAttyLabel = 0
chNoticeToOccupant = 0
chNoticeToOccupantLabel = 0
chAuctioneer = 0
chNewspaperAd = 0
chReadAtSale = 0


chLostNoteAffidavit = 0
chLostNoteNotice = 0
chLostNoteNoticeLabels = 0
chAssignment = 0
chAssignRecCovLtr = 0
chTitleClaim = 0
chTitleReview = 0
chPayoff = 0
chClaimSurplus = 0

chReportOfSale = 0

chSubsPurch = 0
chTrusteeAffidavit = 0
chWithdrawBond = 0
chWithdrawBondBK = 0
chSettlement = 0
chResell = 0
chWithdrawSale = 0
chDismissCase = 0

chDeedConv = 0
chExemptDeed = 0
chDeedHUD = 0
chDeedVA = 0
chDeedSubsPurchaser = 0
chRecordingCoverLetter = 0

chDeedInLieu = 0
chCertofService = 0
chOAH = 0



Exit_cmdClear_Click:
    Exit Sub

Err_cmdClear_Click:
    MsgBox Err.Description
    Resume Exit_cmdClear_Click
    
End Sub

Private Sub PrintDocs(PrintTo As Integer)
Dim ReportName As String, rstLabelData As Recordset, LabelData As String, sql As String, i As Integer, FeeAmount As Currency, noticecnt As Integer, rstJnl As Recordset

'On Error GoTo Err_PrintDocs
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord        ' might need to save the attorney name

If chBaltCityIntake Then Call DoReport("BaltCity Intake Sheet", -1) 'Always print as word
    


If chNOILabels Then     ' same as Fair Debt labels, but do 4 of each
    If PrintTo = acViewNormal Then
        sql = "SELECT Names.Company, Names.First, Names.Last, Names.AKA, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip, CaseList.FileNumber, ClientList.FairDebt, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (ClientList RIGHT JOIN (CaseList RIGHT JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID) LEFT JOIN FCdetails ON CaseList.FileNumber=FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & Forms![Case List]!FileNumber & ") And ((Names.FairDebt)=True) And ((FCdetails.Current)=True));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            For i = 1 To 4
                Call StartLabel
                Print #6, FormatName(rstLabelData!Company, rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
                Print #6, "|FONTSIZE 8"
                Print #6, "|BOTTOM"
                Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
                Call FinishLabel
            Next i
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
    End If
End If

If Me.chIntervene Then Call DoReport("Motion to Intervene", PrintTo)

If Me.chOrderIntervene Then Call DoReport("Order Granting Intervention", PrintTo)

'If Me.chOrderGrantingDefault Then
'    If (Me.ClientID = 6 Or ClientID = 556) Then
'        Call DoReport("DC Order Granting Default - Wells", PrintTo)
'    Else
'        MsgBox ("No Document created for this client")
'    End If
'End If

If Me.chMotionRelease Then Call DoReport("Motion to Release Funds", PrintTo)

If chNoticeOfLisPendens Then Call DoReport("Notice of Lis Pendens", PrintTo)
    

If chTitleOrder Then
    'If SLS file, title must be ordered differently
    If ClientID = 328 Then MsgBox "CAUTION!! See manager for approval before ordering title!!", vbExclamation
    DoCmd.OpenForm "Print Title Order", , , "Caselist.FileNumber=" & Forms!foreclosuredetails!FileNumber, , , PrintTo
End If

If chLine Then
    DoCmd.OpenForm "Print Line", , , "FileNumber=" & Forms![Case List]!FileNumber, , , PrintTo
End If

If chDeedOfApp Then
    DeedOfAppProc (PrintTo)
End If

If chQCDeed Then
    If Me.State = "MD" Then
        Call DoReport("Quit Claim Deed MD", PrintTo)
    Else
        Call DoReport("Quit Claim Deed VA", PrintTo)
    End If
End If

If Me.chWarranty Then
    If Me.State = "MD" Then
        Call DoReport("Special Warranty Deed MD", PrintTo)
    Else
        Call DoReport("Special Warranty Deed VA", PrintTo)
    End If
End If

If chDOARecordingCover Then Call DoReport("DOA Recording Cover", PrintTo)
If chDOAAffidavit Then Call DoReport("Deed of Appointment Affidavit", PrintTo)

If chCourtNotes Then Call DoReport("Mediation Court Notes Wells", PrintTo)


If chDocket Then
    If Forms!foreclosuredetails!WizardSource <> "Docketing" Then
        'If Not PrivDataManager Then
        MsgBox "You can only print the Order to Docket through the wizard", vbCritical
        Exit Sub
        'End If
    End If
    If Me!State = "MD" Then
        Dim Days As Integer
        Dim retval As Boolean
        If (IsNull(SentToDocket)) Then  ' SentToDocket is missing - allow printing
            retval = True
        Else                            ' otherwise, check to ensure SentToDocket date is at least 40 days after NOI
            Days = DateDiff("d", NOI, SentToDocket)
            If (Days < 45) Then
                If DateAdd("d", 45, [NOI]) < Now() Then
                    retval = True
                Else
                    MsgBox "45 Day Notice has not expired", vbCritical
                    Exit Sub
                End If
            Else
                retval = True
            End If
            If IsNull(Forms!foreclosuredetails!FairDebt) Then
                MsgBox "Fair Debt Letter has not been mailed", vbCritical
                Exit Sub
            Else
                retval = True
            End If
            If Date < Forms!foreclosuredetails!AccelerationLetter Then
                MsgBox "Acceleration Letter has not expired", vbCritical
                Exit Sub
            Else
                retval = True
            End If
      End If
      
      If (retval = True) Then '#1558
            'If Forms!ForeclosureDetails!WizardSource <> "Docketing" Then
            '    If Not PrivDataManager Then
            '        MsgBox "Only managers can print the Order to Docket through the wizard", vbCritical
            '        Exit Sub
            '    End If
            'End If
            DoCmd.OpenForm "Print Order To Docket", , , , , , PrintTo
      End If
    Else
        MsgBox "Order to Docket not needed", vbInformation
    End If
End If

If chDebtorLetterLabels Then
    If PrintTo = acViewNormal Then
        sql = "SELECT ClientList.ShortClientName, CaseList.PrimaryDefName, CaseList.FileNumber FROM (ClientList INNER JOIN CaseList ON ClientList.ClientID = CaseList.ClientID) INNER JOIN FCdetails ON CaseList.FileNumber = FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & [Forms]![Case List]![FileNumber] & ") AND ((FCdetails.Current)=True));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        If Forms!foreclosuredetails!State = "MD" Then
        Do While Not rstLabelData.EOF
            
            
            Select Case Forms!foreclosuredetails!City
            Case "Annapolis"
            Call StartLabel
            Print #6, "160 Duke of Gloucester Street"
            Print #6, "Annapolis, MD 21401-2517"
            Print #6, "|FONTSIZE 8"
            
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName


            Case "Poolesville"
            Call StartLabel
            Print #6, "19721 Beall Street, P.O. Box 158"
            Print #6, "Poolesville, Maryland 20837"
            Print #6, "|FONTSIZE 8"
            
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName

            Case "College Park"
            Call StartLabel
            Print #6, "CITY OF COLLEGE PARK, MARYLAND"
            Print #6, "REGISTRATION OF RESIDENTIAL PROPERTY SUBJECT TO FORECLOS"
            Print #6, "|FONTSIZE 8"
            
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Case "Salisbury"
            Call StartLabel
            Print #6, "501B E. Church Street"
            Print #6, "Salisbury, MD 21804"
            Print #6, "|FONTSIZE 8"
           
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Case "Laurel"
            Call StartLabel
            Print #6, "8103 Sandy Spring Road"
            Print #6, "Laurel, Maryland 20707"
            Print #6, "|FONTSIZE 8"
            
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            End Select
            Call FinishLabel
            rstLabelData.MoveNext
         
            Loop
            rstLabelData.Close
            
            End If
            
            If Forms![Case List]!JurisdictionID = 18 Then
            sql = "SELECT ClientList.ShortClientName, CaseList.PrimaryDefName, CaseList.FileNumber FROM (ClientList INNER JOIN CaseList ON ClientList.ClientID = CaseList.ClientID) INNER JOIN FCdetails ON CaseList.FileNumber = FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & [Forms]![Case List]![FileNumber] & ") AND ((FCdetails.Current)=True));"
            Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
            Do While Not rstLabelData.EOF
            Call StartLabel
            Print #6, "1220 Caraway Court, Suite 1050"
            Print #6, "Largo, Maryland 20774"
            Print #6, "|FONTSIZE 8"
            Print #6, ""
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            rstLabelData.MoveNext
            Loop
            rstLabelData.Close
            End If
       End If
        
End If

If chLandInstruments Then
    Call checkInstruments(PrintTo)
    DoCmd.OpenForm "frmLandInstrumentsDetails"
End If

If chAffMD7105 Then
    Call AFFMD7105(PrintTo)
End If

If Me.chBaileeLetter And Forms![Case List]!ClientID = 523 Then
    DoCmd.OpenForm "PrintBaileeLetter", , , "FileNumber = " & Forms![Case List]!FileNumber, , acDialog, PrintTo
End If


If ChCertOfCompliance Then
    If Forms![Case List]!ClientID = 87 Then 'PNC
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Certification of Compliance PNC|Certification of Compliance PNC|" & PrintTo
      
    Else
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Certification of Compliance|Certification of Compliance|" & PrintTo
     
    End If
End If

 If chDOTAffidavit Then
 If Forms![Case List]!State = "MD" Then
 If IsNull(Forms!foreclosuredetails!Disposition) Then
 Call DoReport("DOT Aff Chase", PrintTo)
 Else
 MsgBox ("The file has disposition")
 Exit Sub
 End If
 End If
 End If
 
If ChNoteAffidavit Then
    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If
    Call DoReport("Note Affidavit", PrintTo)
End If


If ChCollateralFileAffidavit Then
    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If
   
    DoCmd.OpenForm "Print Affidavit Collateral file", , , , , acDialog, PrintTo
End If
 
If Ch14207Affidavit Then
If Me!State = "MD" Then
    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If
Call DoReport("Affidavit14207", PrintTo)
End If
End If
    

If chNoteOwnership Then
  
Select Case Forms![Case List]!ClientID
Case 6
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit Wells|Ownership Affidavit Wells|" & PrintTo

Case 532
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit Selene|Ownership Affidavit Selene|" & PrintTo

Case 523
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit GreenTree|Ownership Affidavit Greentree|" & PrintTo

Case 97
   '#1226 Removed Chase ANO - Anne Arundel  - MC - 10/15/2014
   'If Forms![Case List].JurisdictionID = 3 Then
   '     DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!ForeclosureDetails!FileNumber, , acDialog, "Ownership Affidavit Chase Anne|Ownership Affidavit Chase Anne|" & PrintTo
   ' Else
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit Chase|Ownership Affidavit Chase|" & PrintTo
   ' End If
   '/#1226
Case 328
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit SPLS|Ownership Affidavit SPLS|" & PrintTo
Case 446
DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit BOA|Ownership Affidavit BOA|" & PrintTo
    Call DoReport("BOA Cover Sheet", PrintTo)
Case 444
DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit|Ownership Affidavit|" & PrintTo
Call DoReport("PHH Cover Sheet", PrintTo)
Case 556
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit Wells|Ownership Affidavit Wells|" & PrintTo
Case 567
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit|Ownership Affidavit|" & PrintTo
    If Forms![Case List]!ClientID = 567 And Me.State = "MD" Then Call DoReport("CHAM Cover NoteOwnership", PrintTo)
Case 531 'MDHC
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit MDHC|Ownership Affidavit MDHC|" & PrintTo
Case 385
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit|Ownership Affidavit|" & PrintTo
    Call DoReport("Nationstar Cover Sheet", PrintTo)
Case 404
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit|Ownership Affidavit Bogman|" & PrintTo
Case 361
    Call DoReport("Ownership Affidavit Ocwen", PrintTo)
Case Else

    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Ownership Affidavit|Ownership Affidavit|" & PrintTo
End Select
If Forms!foreclosuredetails!WizardSource = "Intake" Then
If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
'Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
'lisa

'2/11/14
    DoCmd.SetWarnings False
    strinfo = "Affidavit of Note Ownership sent to client via Intake Wizard"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True


'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Affidavit of Note Ownership sent to client via Intake Wizard"
'!Color = 1
'.Update
'End With
Else
MsgBox "Remember to manually enter a journal note when the document is sent"
End If
End If
End If

If chLossMitMailing Then
   If (Me.State = "MD") Then
     Call AffadavitOfService(Forms!foreclosuredetails!FileNumber)
     Call DoReport("Loss Mit Mailing Affidavit", PrintTo)

   End If
End If

If chBondOrder Then
   If (Me.State = "MD") Then
     Call DoReport("Bond Order", PrintTo)
   End If
End If

If chPSCover Then

  Call AffadavitOfService(Forms!foreclosuredetails!FileNumber)
  
  DoCmd.OpenForm "Print Process Server Cover", , , , , , PrintTo
End If
If chAffidavitOfService Then
  Call AffadavitOfService(Forms!foreclosuredetails!FileNumber)
  Call DoReport("Order to Docket Affidavit", PrintTo)
 ' DoCmd.OpenForm "Print Affidavit Of Service", , , , , , PrintTo
Forms!foreclosuredetails!cmdWizComplete.Enabled = True
End If


If chSOD Then Call DoReport("Statement of Debt Monitor", PrintTo)
   ' SodProc (PrintTo)
'    If IsNull(Forms!foreclosureDetails![LPIDate]) Then
'    MsgBox ("Cannot print without LPI date")
'    Exit Sub
'    Else
'    MsgBox "Please add the Interest from/to dates once complete"
'        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosureDetails!FileNumber, , acDialog, "Statement of Debt|Statement of Debt|" & PrintTo
'        If Forms![Case List]!ClientID = 446 Then
'        Call DoReport("BOA Cover Sheet", PrintTo)
'        ElseIf Forms![Case List]!ClientID = 444 Then
'        Call DoReport("PHH Cover Sheet", PrintTo)
'        ElseIf Forms![Case List]!ClientID = 385 Then
'        Call DoReport("Nationstar Cover Sheet", PrintTo)
'        ElseIf Forms![Case List]!ClientID = 567 And Me.State = "MD" Then
'        Call DoReport("CHAM Cover SOD", PrintTo)
'        End If
'    End If
'    If Forms!foreclosureDetails!WizardSource = "Intake" Then
'    If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
'    Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
'    With rstJnl
'    .AddNew
'    !FileNumber = FileNumber
'    !JournalDate = Now
'    !Who = GetFullName
'    !Info = "Statement of Debt sent to client via Intake Wizard"
'    !Color = 1
'    .Update
'    End With
'    Else
'    MsgBox "Remember to manually enter a journal note when the document is sent"
'    End If
'    End If
'End If

If chSOD2 Then
    

If IsNull(Forms!foreclosuredetails![LPIDate]) Then
MsgBox ("Cannot print without LPI date")
Exit Sub
Else
Dim Description1 As String, Description2 As String, Description3 As String, Description4 As String, Description5 As String
Dim Description6 As String, Description7 As String, Description8 As String, Description9 As String, Description10 As String, Description11 As String
Dim Description12 As String, Description13 As String, Description14 As String, Description15 As String
Dim Description16, Description17, Description18, Description19, Description20, Description21, Description22, Description23, Description24, Description25 As String
Dim description26, description27, description28, description29, description30 As String
MsgBox "Please add the Interest from/to dates once complete"
Select Case Forms![Case List]![ClientID]


Case 361
Dim rstOcwen As Recordset
Set rstOcwen = CurrentDb.OpenRecordset("select * from statementofdebt where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
If rstOcwen.EOF Then
    description30 = "Interest"
    description29 = "Late Charges"
    description28 = "Escrow Advances for taxes and insurance"
    description27 = "Suspense"
    description26 = "Fees and Expenses"
    With rstOcwen
        .AddNew
        !FileNumber = FileNumber
        !Desc = description26
        !Amount = 0
        !Timestamp = Now
        !Sort_Desc = 1
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = description27
        !Amount = 0
        !Timestamp = Now
        !Sort_Desc = 2
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = description28
        !Amount = 0
        !Sort_Desc = 3
        !Timestamp = Now
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = description29
        !Sort_Desc = 4
        !Amount = 0
        !Timestamp = Now
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = description30
        !Amount = 0
        !Sort_Desc = 5
        !Timestamp = Now
        .Update
       End With
End If
rstOcwen.Close

Case 6
Dim rstWells As Recordset
Set rstWells = CurrentDb.OpenRecordset("select * from statementofdebt where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
If rstWells.EOF Then
Description1 = "Interest"
Description2 = "Pre-Acceleration Late Charges"
Description3 = "Hazard Insurance Disbursements"
Description4 = "Tax Disbursements"
Description5 = "Property Preservation"
Description6 = "PMI/MIP Insurance"
Description7 = "Bankruptcy Fees/Costs"
Description8 = "Other"
Description9 = "Escrow Balance Credit"
Description10 = "Credit to Borrower"
With rstWells
.AddNew
!FileNumber = FileNumber
!Desc = Description1
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description2
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description3
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description4
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description5
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description6
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description7
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description8
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description9
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description10
!Amount = 0
!Timestamp = Now
.Update
End With
End If
rstWells.Close


Case 556
Dim rstWellsH As Recordset
Set rstWellsH = CurrentDb.OpenRecordset("select * from statementofdebt where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
If rstWellsH.EOF Then
Description16 = "Interest"
Description17 = "Pre-Acceleration Late Charges"
Description18 = "Hazard Insurance Disbursements"
Description19 = "Tax Disbursements"
Description20 = "Property Preservation"
Description21 = "PMI/MIP Insurance"
Description22 = "Bankruptcy Fees/Costs"
Description23 = "Other"
Description24 = "Escrow Balance Credit"
Description25 = "Credit to Borrower"
With rstWellsH
.AddNew
!FileNumber = FileNumber
!Desc = Description16
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description17
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description18
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description19
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description20
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description21
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description22
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description23
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description24
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description25
!Amount = 0
!Timestamp = Now
.Update
End With
End If
rstWellsH.Close

Case 446

    If IsNull(Forms![foreclosuredetails]!LPIDate) Then
        MsgBox ("Please add LPI Date in Loan tab, as the Client in this case is Bank of America. ")
        Exit Sub
    End If

Dim rstBOA As Recordset
Set rstBOA = CurrentDb.OpenRecordset("select * from statementofdebt where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
If rstBOA.EOF Then

Description1 = "Interest Amount"
Description2 = "Interest Due"
Description3 = "Assessed Late Charges"
Description4 = "Tax Disbursements"
Description5 = "MIP\PMI Insurance"
Description6 = "Hazard Insurance Disbursements"
Description7 = "Title Fees"
Description8 = "Bankruptcy Fees/Costs"
Description9 = "Property Inspections/Preservation"
Description10 = "Foreclosure Fees"
Description11 = "Escrow Balance Credit"
Description12 = "Unapplied Funds Credit"
Description13 = "Credits"
Description14 = "Other"
Description15 = "Payment Advance - Principal/Interest/Escrow"

With rstBOA
.AddNew
!FileNumber = FileNumber
!Desc = Description1
!Sort_Desc = 1
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description2
!Sort_Desc = 2
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description3
 !Sort_Desc = 3
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description4
 !Sort_Desc = 4
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description5
 !Sort_Desc = 5
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description6
 !Sort_Desc = 6
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description7
 !Sort_Desc = 7
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description8
 !Sort_Desc = 8
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description9
 !Sort_Desc = 9
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description10
 !Sort_Desc = 10
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description11
 !Sort_Desc = 11
!Amount = 0
!Timestamp = Now
!Credit = True
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description12
 !Sort_Desc = 12
!Amount = 0
!Timestamp = Now
!Credit = True
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description13
 !Sort_Desc = 13
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description14
 !Sort_Desc = 14
!Amount = 0
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description15
 !Sort_Desc = 15
!Amount = 0
!Timestamp = Now
.Update
End With
End If
rstBOA.Close

Case 328
Dim rstSPLS As Recordset
Set rstSPLS = CurrentDb.OpenRecordset("select * from statementofdebt where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
If rstSPLS.EOF Then
    description30 = "Interest"
    description29 = "Escrow Advanced"
    description28 = "Attorney Fees and Costs"
    description27 = "Property Preservation"
    
    With rstSPLS
        .AddNew
        !FileNumber = FileNumber
        !Desc = description30
        !Amount = 0
        !Timestamp = Now
        !Sort_Desc = 1
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = description29
        !Amount = 0
        !Timestamp = Now
        !Sort_Desc = 2
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = description28
        !Amount = 0
        !Sort_Desc = 3
        !Timestamp = Now
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = description27
        !Sort_Desc = 4
        !Amount = 0
        !Timestamp = Now
        .Update
      
    End With
End If
rstSPLS.Close

Case 97

    If IsNull(Forms![foreclosuredetails]!LPIDate) Then

        MsgBox ("Please add LPI Date in Loan tab. ")

        Exit Sub

    End If

Dim rstjp As Recordset
Set rstjp = CurrentDb.OpenRecordset("select * from statementofdebt where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
If rstjp.EOF Then
'Description1 = "Unpaid Principal Balance "
Description2 = "Interest Accrued at " & "_______" & " per annum (from ____________ through ____________ ."
Description3 = "Pre-Acceleration Late Charges"
Description4 = "Escrow "
Description5 = "Escrow Deficiency-Real Estate Taxes"
Description6 = "Hazard Insurance"
Description7 = "Mortgage Insurance Premium/Private Mortgage Insurance"
Description8 = "Credits"
Description9 = "Total Escrow"
Description10 = "Broker's Price Opinion/Appraisals"
Description11 = "Property Preservation"
Description12 = "Previous Bankruptcy Fees/Costs"
Description13 = "Suspense"
Description14 = "Miscellaneous Charges/Credits as Follows:"

With rstjp

.AddNew
!FileNumber = FileNumber
!Desc = Description2
!Amount = 0
!Sort_Desc = 13
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description3
!Amount = 0
!Sort_Desc = 12
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description4
!Amount = 0
!Sort_Desc = 11
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description5
!Amount = 0
!Sort_Desc = 10
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description6
!Amount = 0
!Sort_Desc = 9
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description7
!Amount = 0
!Sort_Desc = 8
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description8
!Amount = 0
!Sort_Desc = 7
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description9
!Amount = 0
!Sort_Desc = 6
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description10
!Amount = 0
!Sort_Desc = 5
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description11
!Amount = 0
!Sort_Desc = 4
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description12
!Amount = 0
!Sort_Desc = 3
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description13
!Amount = 0
!Sort_Desc = 2
!Timestamp = Now
.Update
.AddNew
!FileNumber = FileNumber
!Desc = Description14
!Amount = 0
!Sort_Desc = 1
!Timestamp = Now
.Update
End With
End If
rstjp.Close

Case 385
    Call DoReport("Nationstar Cover Sheet", PrintTo)

Case 567
    If Forms![Case List]!ClientID = 567 And Me.State = "MD" Then Call DoReport("CHAM Cover SOD2", PrintTo)
    
End Select

DoCmd.OpenForm "Print Statement of Debt", , , "ForeclosureID=" & Forms!foreclosuredetails!ForeclosureID, , , PrintTo
If Forms![Case List]![ClientID] = 97 Then
Forms![Print Statement of Debt]!cbxRateType.Visible = True
End If

'DoCmd.OpenForm "Print Statement of Debt", , , "ForeclosureID=" & Forms!foreclosuredetails!ForeclosureID, , , PrintTo
If Forms![Case List]![ClientID] = 6 Then
Forms![Print Statement of Debt]!cbxRateType.Visible = True
Else
If Forms![Case List]![ClientID] = 556 Then
Forms![Print Statement of Debt]!cbxRateType.Visible = True
End If
End If
If Forms!foreclosuredetails!WizardSource = "Intake" Then
If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
'2/11/14
'Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)

    DoCmd.SetWarnings False
    strinfo = "Statement of Debt sent to client via Intake Wizard"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
'lisa

'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Statement of Debt sent to client via Intake Wizard"
'!Color = 1
'.Update
'End With
Else
MsgBox "Remember to manually enter a journal note when the document is sent"
End If
End If
End If
End If


If chComplianceAffidavit Then
    
    
    If IsNull(Forms!foreclosuredetails![LPIDate]) Then
        MsgBox ("Cannot print without LPI date")
        Exit Sub
    Else
        Dim Desc1, desc2, desc3, desc4, desc5, desc6, desc7 As String
        
        Select Case Forms![Case List]![ClientID]

        Case 466 'SELECT
            Dim rstSelect As Recordset
            Set rstSelect = CurrentDb.OpenRecordset("Select * from statementofDebt where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
            
            If rstSelect.EOF Then
                desc7 = "Interest"
                desc6 = "Pro Rata Mortgage Insurance"
                desc5 = "Escrow Advance"
                desc4 = "Late Charges"
                desc3 = "NSF Charges"
                desc2 = "Advances Made on the Defendant's Behalf"
                Desc1 = "Suspense Balance"
                
                With rstSelect
                    .AddNew
                    !FileNumber = FileNumber
                    !Desc = desc7
                    !Amount = 0
                    !Timestamp = Now
                    !Sort_Desc = 7
                    .Update
                    .AddNew
                    !FileNumber = FileNumber
                    !Desc = desc6
                    !Amount = 0
                    !Timestamp = Now
                    !Sort_Desc = 6
                    .Update
                    .AddNew
                    !FileNumber = FileNumber
                    !Desc = desc5
                    !Amount = 0
                    !Sort_Desc = 5
                    !Timestamp = Now
                    .Update
                    .AddNew
                    !FileNumber = FileNumber
                    !Desc = desc4
                    !Sort_Desc = 4
                    !Amount = 0
                    !Timestamp = Now
                    .Update
                    .AddNew
                    !FileNumber = FileNumber
                    !Desc = desc3
                    !Amount = 0
                    !Sort_Desc = 3
                    !Timestamp = Now
                    .Update
                    .AddNew
                    !FileNumber = FileNumber
                    !Desc = desc2
                    !Amount = 0
                    !Sort_Desc = 2
                    !Timestamp = Now
                    .Update
                    .AddNew
                    !FileNumber = FileNumber
                    !Desc = Desc1
                    !Amount = 0
                    !Sort_Desc = 1
                    !Timestamp = Now
                    .Update
                End With
            End If
            rstSelect.Close
        End Select
        DoCmd.OpenForm "Print Statement of Debt", , , "ForeclosureID=" & Forms!foreclosuredetails!ForeclosureID, , , PrintTo
    End If
End If





If chMilitaryAffidavitActive Then
  If Me.ClientID = 404 And Me.State = "MD" Then
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Military Affidavit Active|MilitaryAffidavitActiveMD Bogman|" & PrintTo
  ElseIf Me!State = "MD" Then
    DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Military Affidavit Active|MilitaryAffidavitActiveMD|" & PrintTo
  Else

      DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Military Affidavit Active|MilitaryAffidavitActive|" & PrintTo

  End If

    If Forms![Case List]!ClientID = 446 Then
        Call DoReport("BOA Cover Sheet", PrintTo)
    ElseIf Forms![Case List]!ClientID = 444 Then Call DoReport("PHH Cover Sheet", PrintTo)
    ElseIf Forms![Case List]!ClientID = 567 And Me.State <> "DC" Then Call DoReport("CHAM Cover MilitaryAffidavitActive", PrintTo)
            
        
    End If
If Forms!foreclosuredetails!WizardSource = "Intake" Then
If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
'2/11/14
'Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
'lisa

    DoCmd.SetWarnings False
    strinfo = "Military Affidavit sent to client via Intake Wizard"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True


'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Military Affidavit sent to client via Intake Wizard"
'!Color = 1
'.Update
'End With
Else
MsgBox "Remember to manually enter a journal note when the document is sent"
End If
End If
End If

If chMilitaryAffidavit Then
    If Me!State = "MD" Then
        If Me.ClientID = 404 Then
            ReportName = "MilitaryAffidavitMD Bogman"
        ElseIf Me.ClientID = 328 Then '#1093 10/6/2014 MC
            ReportName = "MilitaryAffidavitMD SPLS"
        Else
            ReportName = "MilitaryAffidavitMD"
        End If
    ElseIf Me!State = "DC" Then
        ReportName = "MilitaryAffidavitDC"
    Else
        ReportName = "Military Affidavit"
    End If

    DoCmd.OpenForm "PrintClientDocM", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Military Affidavit|" & ReportName & "|" & PrintTo

    If Forms![Case List]!ClientID = 446 Then
    Call DoReport("BOA Cover Sheet", PrintTo)
    ElseIf Forms![Case List]!ClientID = 444 Then Call DoReport("PHH Cover Sheet", PrintTo)
    ElseIf Forms![Case List]!ClientID = 567 And Me.State <> "DC" Then Call DoReport("CHAM Cover MilitaryAffidavit", PrintTo)
    
    
   End If
If Forms!foreclosuredetails!WizardSource = "Intake" Then
If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
'2/11/14
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
'lisa
'
    DoCmd.SetWarnings False
    strinfo = "Military Affidavit sent to client via Intake Wizard"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Military Affidavit sent to client via Intake Wizard"
'!Color = 1
'.Update
'End With
Else
MsgBox "Remember to manually enter a journal note when the document is sent"
End If
End If
End If

If chMilitaryAffidavitNoSSN Then
    If Me.ClientID = 404 And Me.State = "MD" Then
        ReportName = "Military Affidavit NoSSN MD Bogman"
    ElseIf Me!State = "MD" Then
        Dim n As Recordset
        Set n = CurrentDb.OpenRecordset("SELECT SSN FROM Names WHERE FileNumber = " & Forms!foreclosuredetails!FileNumber & " AND mortgagor = true AND (isnull(SSN) = true  or SSN = ""999999999"" ) ORDER BY ID, Last, First", dbOpenSnapshot)

        ReportName = "Military Affidavit NoSSN MD"
    Else
        ReportName = "Military Affidavit NoSSN"
    End If
    DoCmd.OpenForm "PrintClientDocDefendant", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Military Affidavit - No SSN|" & ReportName & "|" & PrintTo
           
    If Forms![Case List]!ClientID = 446 Then
    Call DoReport("BOA Cover Sheet", PrintTo)
    ElseIf Forms![Case List]!ClientID = 444 Then Call DoReport("PHH Cover Sheet", PrintTo)
    ElseIf Forms![Case List]!ClientID = 567 And Me.State <> "DC" Then Call DoReport("CHAM Cover MilitaryAffidavitNoSSN", PrintTo)
    End If
    
    
If Forms!foreclosuredetails!WizardSource = "Intake" Then
If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
'2/11/14
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
'lisa
    DoCmd.SetWarnings False
    strinfo = "No-SSN Military Affidavit sent to client via Intake Wizard"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "No-SSN Military Affidavit sent to client via Intake Wizard"
'!Color = 1
'.Update
'End With
Else
MsgBox "Remember to manually enter a journal note when the document is sent"
End If
End If
End If


If chLossMitPrelim Then
Select Case Forms![Case List]!ClientID
Case 466 ' SPS
    Call DoReport("Loss Mitigation Preliminary Select", PrintTo)
Case 446
Call DoReport("Loss Mitigation Preliminary BOA", PrintTo)
Call DoReport("BOA Cover Sheet", PrintTo)
Case 97
Call DoReport("Loss Mitigation Preliminary Chase", PrintTo)
Case 328 ' SPLS
    Call DoReport("Loss Mitigation Preliminary SPLS", PrintTo)
Case 6
DoCmd.OpenForm "PrintClientDocWells", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Loss Mitigation Preliminary Wells|Loss Mitigation Preliminary Wells|" & PrintTo

Case 556
DoCmd.OpenForm "PrintClientDocWells", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Loss Mitigation Preliminary Wells|Loss Mitigation Preliminary Wells|" & PrintTo

Case 567
    Call DoReport("Loss Mitigation Preliminary", PrintTo)
    If Forms![Case List]!ClientID = 567 And Me.State = "MD" Then Call DoReport("CHAM Cover LossMitPrelim", PrintTo)

Case 531 'MDHC
Call DoReport("Loss Mitigation Preliminary MDHCD", PrintTo)

Case 456 ' M&T Bank
Call DoReport("Loss Mitigation Preliminary MT", PrintTo)

Case 444
Call DoReport("Loss Mitigation Preliminary", PrintTo)
Call DoReport("PHH Cover Sheet", PrintTo)

Case 385
If Me.State = "MD" Then
    Call DoReport("Loss Mitigation Preliminary Nation Star", PrintTo)
    Call DoReport("Nationstar Cover Sheet", PrintTo)
Else
    Call DoReport("Loss Mitigation Preliminary", PrintTo)
    Call DoReport("Nationstar Cover Sheet", PrintTo)
End If

Case 87
    Call DoReport("Loss Mitigation Preliminary PNC", PrintTo)
Case Else
Call DoReport("Loss Mitigation Preliminary", PrintTo)
End Select


If Forms!foreclosuredetails!WizardSource = "Intake" Then
If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
'2/11/14
'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
'lisa
    DoCmd.SetWarnings False
    strinfo = "PLMA sent to client via Intake Wizard"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!Foreclosureprint!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True


'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "PLMA sent to client via Intake Wizard"
'!Color = 1
'.Update
'End With
Else
MsgBox "Remember to manually enter a journal note when the document is sent"
End If
End If
End If

'If chLossMitFinal Then
'Select Case Forms![Case List]!ClientID
'Case 466 'SPS
'    Call DoReport("Loss Mitigation Final Select", PrintTo)
'Case 446
'Call DoReport("Loss Mitigation Final BOA", PrintTo)
'Call DoReport("BOA Cover Sheet", PrintTo)
'Case 97
'Call DoReport("Loss Mitigation Final Chase", PrintTo)
'
'Case 6
'DoCmd.OpenForm "PrintClientDocWells", , , "FileNumber=" & Forms!ForeclosureDetails!FileNumber, , acDialog, "Loss Mitigation Final Wells|Loss Mitigation Final Wells|" & PrintTo
'
'Case 556
'DoCmd.OpenForm "PrintClientDocWells", , , "FileNumber=" & Forms!ForeclosureDetails!FileNumber, , acDialog, "Loss Mitigation Final Wells|Loss Mitigation Final Wells|" & PrintTo
'
'Case 567
'     Call DoReport("Loss Mitigation Final", PrintTo)
'     If Forms![Case List]!ClientID = 567 And Me.State = "MD" Then Call DoReport("CHAM Cover LossMitFinal", PrintTo)
'
'Case 531 'MHDC
'    Call DoReport("Loss Mitigation Final MDHCD", PrintTo)
'
'Case 456 'M&T Bank
'    Call DoReport("Loss Mitigation Final MT", PrintTo)
'
'Case 444
'Call DoReport("Loss Mitigation Final", PrintTo)
'Call DoReport("PHH Cover Sheet", PrintTo)
'Case 385
'
'    If Me.State = "MD" Then
'    Call DoReport("Loss Mitigation Final Nation Star", PrintTo)
'    Call DoReport("Nationstar Cover Sheet", PrintTo)
'    Else
'    Call DoReport("Loss Mitigation Final", PrintTo)
'    Call DoReport("Nationstar Cover Sheet", PrintTo)
'    End If
'Case 328
'    Call DoReport("Loss Mitigation Final SPLS", PrintTo)
'Case 87
'    Call DoReport("Loss Mitigation Final PNC", PrintTo)
'Case Else
'Call DoReport("Loss Mitigation Final", PrintTo)
'End Select
'
'If Forms!ForeclosureDetails!WizardSource = "Intake" Then
'If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
''2/11/14
''Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
'
''lisa
'    DoCmd.SetWarnings False
'    strinfo = "FLMA sent to client via Intake Wizard"
'    strinfo = Replace(strinfo, "'", "''")
'    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
'    DoCmd.RunSQL strSQLJournal
'    DoCmd.SetWarnings True
'
'
'Else
'MsgBox "Remember to manually enter a journal note when the document is sent"
'End If
'End If
'End If

If chLossMitApp Then Call StartDoc(TemplatePath & "\LossMitApp.pdf")


If chForecloseMed Then Call DoReport("Foreclosure Mediation", PrintTo)

 If chHUDOccLabels Then

        sql = "SELECT fcdetails.PropertyAddress, fcdetails.City, fcdetails.State, fcdetails.ZipCode, CaseList.FileNumber, ClientList.ShortClientName, CaseList.PrimaryDefName"
        sql = sql + " FROM (fcdetails INNER JOIN CaseList ON fcdetails.FileNumber = CaseList.FileNumber) INNER JOIN ClientList ON CaseList.ClientID = ClientList.ClientID"
        sql = sql + " WHERE CaseList.FileNumber =" & Forms![Case List]!FileNumber & ";"
  
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        
        For i = 1 To 3 'Print 3 sets of labels
     
            Call StartLabel
            Print #6, FormatName("", "", MortgagorOwnerNames(0, 7), "", rstLabelData!PropertyAddress, "", rstLabelData!City, rstLabelData!State, rstLabelData!ZipCode)
            Print #6, "|FONTSIZE 8"
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            Call StartLabel
            Print #6, FormatName("", "", "All Occupants", "", rstLabelData!PropertyAddress, "", rstLabelData!City, rstLabelData!State, rstLabelData!ZipCode)
            Print #6, "|FONTSIZE 8"
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            
        Next i
        rstLabelData.Close
        Set rstLabelData = Nothing
    
End If


If chFairDebtLabels Then    ' same as NOI labels
    If PrintTo = acViewNormal Then
        sql = "SELECT Names.Company, Names.First, Names.Last, Names.AKA, Names.Address, Names.Deceased, Names.Address2, Names.City, Names.State, Names.Zip, CaseList.FileNumber, ClientList.FairDebt, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (ClientList RIGHT JOIN (CaseList RIGHT JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID) LEFT JOIN FCdetails ON CaseList.FileNumber=FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & Forms![Case List]!FileNumber & ") And ((Names.FairDebt)=True) And ((FCdetails.Current)=True));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
            If rstLabelData!Deceased = True Then
            Print #6, FormatName("", "The Estate of " & rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            Else
            Print #6, FormatName(rstLabelData!Company, rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            End If
            Print #6, "|FONTSIZE 8"
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
    End If
End If

If chLossMitSol Then
    If LoanType = 5 And State = "VA" Or LoanType = 4 And State = "VA" Then
    Call DoReport("Loss Mitigation Solicitation Letter Wiz", PrintTo)
    
    If MsgBox("Update Loss Mitigation Solicitation Sent = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
        Forms!foreclosuredetails!LossMitSolicitationDate = Now()
        AddStatus [CaseList.FileNumber], Now(), "Loss Mitigation Solicitation Letter sent"
        
    End If
    Else
    MsgBox "The Loss Mitigation Letter can only be printed for FHLMC files in Virginia", vbCritical
    End If
    Dim lossMitSolCnt As Integer
    lossMitSolCnt = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and FairDebt = true")
    If (lossMitSolCnt > 0) Then

    End If
  
End If

If chLossMitSolLabels Then    ' same as NOI labels
    If PrintTo = acViewNormal Then
        sql = "SELECT Names.Company, Names.First, Names.Last, Names.AKA, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip, CaseList.FileNumber, ClientList.FairDebt, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (ClientList RIGHT JOIN (CaseList RIGHT JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID) LEFT JOIN FCdetails ON CaseList.FileNumber=FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & Forms![Case List]!FileNumber & ") And ((Names.FairDebt)=True) And ((FCdetails.Current)=True));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
            Print #6, FormatName(rstLabelData!Company, rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            Print #6, "|FONTSIZE 8"
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
    End If
End If

If chHUDOcc Then
    Call DoReport("HUD Occupancy Letter New", PrintTo)
    Call DoReport("HUD Occupancy Letter All Occ New", PrintTo)
    Call StartDoc(TemplatePath & "HUDOCC Attachments.pdf")
    If MsgBox("Do you wish to update the billing invoice for HUD Occ Letter Postage ?", vbYesNo) = vbYes Then
        AddInvoiceItem [CaseList.FileNumber], "FC-HUDOCC", "HUD Occ Letter Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 8)), 76, False, False, False, True
        AddInvoiceItem [CaseList.FileNumber], "FC-HUDOCC", "HUD Occ Letter Postage", Nz(DLookup("Value", "StandardCharges", "ID=" & 8)), 76, False, False, False, True
    End If
    If MsgBox("Update HUD Occupancy Letter Sent = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
        Forms!foreclosuredetails!HUDOccLetter = Now()
        AddStatus [CaseList.FileNumber], Now(), "HUD Occupancy Letter sent"
    End If
End If

If chIRSNotice Then DoCmd.OpenForm "Print IRS Notice", , , "Caselist.FileNumber=" & Forms!foreclosuredetails!FileNumber, , , PrintTo

If chNotice Then
      If Not IsNull(Forms!foreclosuredetails!Sale) And Not IsNull(Forms!foreclosuredetails!SaleTime) Then
        'MsgBox (DCount("[ID]", "Names", "Nz([NoticeType]) = 0 AND [FileNumber]=" & [Forms]![Case list]![FileNumber])) just for test chagnes on 10/23 to fix print problem
           'If DCount("[ID]", "Names", "Nz([NoticeType]) = 0 AND [FileNumber]=" & [Forms]![Case list]![FileNumber]) = 0 Then
                Call DoReport("Notice " & Me!State, PrintTo)
                If Me!State = "MD" Then
                    Call DoReport("Notice MD County Attorney", PrintTo)
                    Call DoReport("Notice MD All Occupants", PrintTo)
                End If
                If MsgBox("Update Notices Sent = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
                noticecnt = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and nz(NoticeType,0) > 0")
                    Forms!foreclosuredetails!Notices = Now()
                    AddStatus [CaseList.FileNumber], Now(), "Mailed Notice of Foreclosure Sale"
                    AddInvoiceItem FileNumber, "FC-NOT", "Sale Notices - Certified Postage", (Nz(DLookup("Value", "StandardCharges", "ID=" & 8))) * noticecnt, 76, False, False, False, True
                    AddInvoiceItem FileNumber, "FC-NOT", "Sale Notices - First Class Postage", (Nz(DLookup("Value", "StandardCharges", "ID=" & 1))) * noticecnt, 76, False, False, False, True
        
                End If
           ' Else
              '  Call MsgBox("Please review the Names tab for parties without a 'Send Notice' option selected", vbCritical)
            'End If
       Else
        MsgBox ("Please Check the Date of Sale and the Time of Sale")
        Exit Sub
       End If
    
End If
If chNoticeLabels Then
    If PrintTo = acViewNormal Then
        sql = "SELECT CaseList.FileNumber, CaseList.PrimaryDefName, ClientList.ShortClientName, Names.Company, Names.Deceased, Names.Last, Names.First, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip, [Notice Label Copies].CopyNumber " & _
                "FROM [Notice Label Copies], ClientList INNER JOIN (CaseList INNER JOIN Names ON CaseList.FileNumber = Names.FileNumber) ON ClientList.ClientID = CaseList.ClientID " & _
                "WHERE (((CaseList.FileNumber)=" & [Forms]![Case List]![FileNumber] & ") And ((Names.NoticeType) <> 0)) " & _
                "ORDER BY Names.Last, Names.First, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip " & _
                "UNION SELECT CaseList.FileNumber, CaseList.PrimaryDefName, ClientList.ShortClientName, JurisdictionList.CountyAttnyAddr, " & _
                "Null, Null, nULL, Null, Null, Null, Null, Null, [Notice Label Copies].CopyNumber " & _
                "FROM [Notice Label Copies], JurisdictionList INNER JOIN (ClientList INNER JOIN CaseList ON ClientList.ClientID = CaseList.ClientID) ON (JurisdictionList.JurisdictionID = CaseList.JurisdictionID) AND (JurisdictionList.JurisdictionID = CaseList.JurisdictionID) " & _
                "WHERE (((CaseList.FileNumber)=" & [Forms]![Case List]![FileNumber] & ") AND JurisdictionList.CountyAttnyAddr Is Not Null);"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
            Print #6, FormatName(rstLabelData!Company, IIf(rstLabelData!Deceased = True, "Estate of " & rstLabelData!First, rstLabelData!First), rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            Print #6, "|FONTSIZE 8"
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
    End If
End If

If chCountyAttyLabel Then
    If PrintTo = acViewNormal Then
        LabelData = Nz(DLookup("CountyAttnyAddr", "JurisdictionList", "JurisdictionID=" & JurisdictionID))
        If LabelData = "" Then
            MsgBox "County Attorney Address has not been set for this jurisdiction", vbCritical
        Else
            Call StartLabel
            Print #6, LabelData
            Call FinishLabel
        End If
    End If
End If

If chNoticeToOccupant Then Call DoReport("Notice to Occupant", PrintTo)
If chNoticeToOccupantLabel Then
    If PrintTo = acViewNormal Then
        sql = "SELECT PropertyAddress, City, State, ZipCode FROM FCDetails WHERE ForeclosureID=" & ForeclosureID
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
            Print #6, "All Occupants"
            Print #6, rstLabelData!PropertyAddress
            Print #6, rstLabelData!City & ", " & rstLabelData!State & " " & FormatZip(Nz(rstLabelData!ZipCode))
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
    End If
End If

If chAuctioneer Then Call DoReport("Auctioneer " & Me!State, PrintTo)

If chCoCounselLetter Then
    coCounselLetterProc (PrintTo)

     
End If

If chNewspaperAd Then
    NewsPaperAdProc (PrintTo)
End If

If chReadAtSale Then
    Call DoReport("Read at Sale", PrintTo)
    Call DoReport("Contract for Sale", PrintTo)
    If Me!State = "VA" Then Call DoReport("Bid Instructions VA", PrintTo)
End If

If chLostNoteAffidavit Then
    If (ClientID = 6 Or ClientID = 556 Or ClientID = 151) Then
            If Me.State = "MD" Then
            Call DoReport("Lost Note Affidavit Wells", PrintTo)
            Else
            Call DoReport("Lost Note Affidavit Wells VA", PrintTo)
            End If
    ElseIf ClientID = 532 Then
        Call DoReport("Lost Note Affidavit Selene", PrintTo)
    
    ElseIf ClientID = 523 Then
        Call DoReport("Lost Note Affidavit GreenTree", PrintTo)
    
    Else
      Call DoReport("Lost Note Affidavit", PrintTo)
      If (ClientID = 444) Then
      Call DoReport("PHH Cover Sheet", PrintTo)
      ElseIf (ClientID = 385) Then
      Call DoReport("Nationstar Cover Sheet", PrintTo)
      End If
      
End If
 If MsgBox("Add to status: ""Lost Note Affidavit sent""?", vbYesNo) = vbYes Then
        AddStatus [CaseList.FileNumber], Now(), "Executed Lost Note Affidavit"
        Forms!foreclosuredetails!LostNoteAffSent = Date
        If Forms!foreclosuredetails!WizardSource = "Intake" Then
'Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
'2/11/14
'lisa
    DoCmd.SetWarnings False
    strinfo = "Lost Note Affidavit sent to client via Intake Wizard"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True


'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Lost Note Affidavit sent to client via Intake Wizard"
'!Color = 1
'.Update
'End With
End If
Select Case LoanType
    Case 4
    FeeAmount = Nz(DLookup("FeeLostNoteAffidavit", "ClientList", "ClientID=177"))
    Case 5
    FeeAmount = Nz(DLookup("FeeLostNoteAffidavit", "ClientList", "ClientID=263"))
    Case Else
    Nz (DLookup("FeeLostNoteAffidavit", "ClientList", "ClientID=" & Forms![Case List]!ClientID))
    End Select
        If FeeAmount > 0 Then
            AddInvoiceItem FileNumber, "FC-LNA", "Lost Note Affidavit", FeeAmount, 0, True, False, False, False
        Else
            AddInvoiceItem FileNumber, "FC-LNA", "Lost Note Affidavit", 1, 0, True, False, False, False
        End If
                       
    End If
End If

If chLostNoteNotice Then
    Call DoReport("Lost Note Notice", PrintTo)
    If MsgBox("Add to status: ""Executed Lost Note Letter""?", vbYesNo) = vbYes Then
        Dim fairdebtCnt As Integer, qtypstge As Integer
        fairdebtCnt = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and FairDebt = true")
    If (fairdebtCnt > 0) Then
            FeeAmount = Nz(DLookup("Value", "StandardCharges", "ID=" & 1))
            qtypstge = DCount("[FileNumber]", "[qryFairDebt]", "FileNumber=" & [FileNumber])
            AddInvoiceItem FileNumber, "FC-LNN", "Lost Note Notice Postage", (qtypstge * FeeAmount), 76, False, False, False, True
            FeeAmount = DLookup("IVALUE", "DB", "ID=" & 17) / 100
            AddInvoiceItem FileNumber, "FC-LNN", "Lost Note Notice Postage", (qtypstge * FeeAmount), 76, False, False, False, True
    End If
    
    Dim rstqueue As Recordset
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " and current=true", dbOpenDynaset, dbSeeChanges)

    With rstqueue
    .Edit
    !VALNNUser = GetStaffID
    !VALNNComplete = Now
    .Update
    End With
    
    Forms!foreclosuredetails!LostNoteNotice = Date
    AddStatus [CaseList.FileNumber], Now(), "Lost Note Letter sent to Debtor by regular and certified mail"
    End If
End If

If chLostNoteNoticeLabels Then
    If PrintTo = acViewNormal Then
        sql = "SELECT CaseList.FileNumber, CaseList.PrimaryDefName, ClientList.ShortClientName, Names.Company, Names.Last, Names.First, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip, [Notice Label Copies].CopyNumber, Names.FairDebt FROM [Notice Label Copies], ClientList INNER JOIN (CaseList INNER JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID WHERE (((CaseList.FileNumber)=" & Forms![Case List]!FileNumber & ") And ((Names.FairDebt)=True)) ORDER BY Names.Last, Names.First, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip;"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
            Print #6, FormatName(rstLabelData!Company, rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            Print #6, "|FONTSIZE 8"
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
    End If
End If

If PRlabel Then
    If PrintTo = acViewNormal Then
        sql = "SELECT ClientList.ShortClientName, CaseList.PrimaryDefName, CaseList.FileNumber, JurisdictionList.CountyAssessorsName, JurisdictionList.CountyAssessorsAddr, JurisdictionList.Jurisdiction,[Notice Label Copies].CopyNumber FROM [Notice Label Copies], ((CaseList INNER JOIN FCdetails ON CaseList.FileNumber = FCdetails.FileNumber) INNER JOIN JurisdictionList ON CaseList.JurisdictionID = JurisdictionList.JurisdictionID) INNER JOIN ClientList ON CaseList.ClientID = ClientList.ClientID  WHERE (((CaseList.FileNumber)= " & [Forms]![Case List]![FileNumber] & " ) AND ((FCdetails.Current)=True) AND (([Notice Label Copies].CopyNumber)=1));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
           ' Print #6, "Maryland SDAT"
            Print #6, "Attn: " & rstLabelData!CountyAssessorsName
            Print #6, rstLabelData!CountyAssessorsAddr
            Print #6, "|FONTSIZE 8"
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
    Else
        MsgBox "Please click the Printer Icon to print Property Registration Labels"
    End If
End If

If ChLabel Then
    If PrintTo = acViewNormal Then
        sql = "SELECT DISTINCTROW FCdetails.FileNumber, FCdetails.PrimaryLastName, FCdetails.PrimaryFirstName, FCdetails.SecondaryLastName, FCdetails.SecondaryFirstName, FCdetails.PropertyAddress, FCdetails.City, FCdetails.State, FCdetails.ZipCode, FCdetails.LoanNumber, CaseList.Active, FCdetails.Current FROM (CaseList INNER JOIN FCdetails ON CaseList.FileNumber = FCdetails.FileNumber) LEFT JOIN FCDisposition ON FCdetails.Disposition = FCDisposition.ID WHERE (((FCdetails.FileNumber)=" & Forms![Case List]!FileNumber & ") AND ((CaseList.Active)=1) AND ((FCdetails.Current)=1));"

        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
            'Print #6, "|FONTSIZE 11"
            Print #6, rstLabelData!PrimaryFirstName & "  " & rstLabelData!PrimaryLastName
            Print #6, rstLabelData!SecondaryFirstName & "  " & rstLabelData!SecondaryLastName
            Print #6, rstLabelData!PropertyAddress
            Print #6, rstLabelData!City; ", " & rstLabelData!State & " " & rstLabelData!ZipCode
            Print #6, "Loan # " & rstLabelData!LoanNumber
            Print #6, "|BOTTOM"
            Print #6, "File # " & rstLabelData!FileNumber
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
      
        End If
End If

If chAssignment Then DoCmd.OpenForm "Print Assignment", , , "Caselist.FileNumber=" & Forms!foreclosuredetails!FileNumber, , , PrintTo
    
    
If chAssignRecCovLtr Then
   Call DoReport("Assignment Recording Cover Letter", PrintTo)
End If

If chTitleClaim Then DoCmd.OpenForm "Print Title Claim", , , "ForeclosureID=" & Forms!foreclosuredetails!ForeclosureID, , , PrintTo

If chTitleReview Then
'added 1/27/15
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("Select * FROM FCTitle where FileNumber =" & [Forms]![Case List]![FileNumber], dbOpenDynaset, dbSeeChanges)
        If (IsNull(rs!TitleReviewLiens) Or rs!TitleReviewLiens = "") Or (IsNull(rs!TitleReviewJudgments) Or rs!TitleReviewJudgments = "") Or (IsNull(rs!TitleReviewTaxes) Or rs!TitleReviewTaxes = "") Or (IsNull(rs!TitleReviewStatus) Or rs!TitleReviewStatus = "") Or (IsNull(rs!TitleReviewNameOf) Or rs!TitleReviewNameOf = "") Then
        'Or IsNull(rs!LegalDescription) or IsNull(rs!TitleReviewBlank) or IsNull(rs!TitleReviewJunior) Or IsNull(rs!TitleReview3) Then
                            
            MsgBox ("Can't print the title review sheet witout enter all the necessary fields")
        Exit Sub
                                    
        Else
        Call DoReport("Title Review", PrintTo)
        End If
End If



If chReturnDocs Then
    DoCmd.OpenForm "Print Return Docs", , , "FileNumber = " & Forms![Case List]!FileNumber, , , PrintTo
End If

If chPayoff Then
DoCmd.OpenForm "Print Payoff", , , "ForeclosureID=" & Forms!foreclosuredetails!ForeclosureID, , , PrintTo & "|FC"
Forms![Print Payoff]!Option57.Enabled = True
Forms![Print Payoff]!Option59.Enabled = True
End If

If chClaimSurplus Then
    If Not chSOD And Not chSOD2 Then
        MsgBox "You must select the Statement of Debt or Statement of Debt with Figures in order to print the Claim for Surplus", vbCritical
    Else
        Call DoReport("Claim Surplus", PrintTo)
    End If
End If

If chSettlement Then Call DoReport("Settlement Letter", PrintTo)

If chSubsPurch Then
    Call DoReport("Substitute Purchaser Motion", PrintTo)
    Call DoReport("Substitute Purchaser Order", PrintTo)
End If

If chReportOfSale Then
    chReportofSaleProc (PrintTo)
    
End If

If chTrusteeAffidavit Then
    If IsNull(Attorney) Then
        MsgBox "Select an attorney to sign the document", vbCritical
        Exit Sub
    End If
    
    If (Not IsNull([Sale])) Then
      If [Sale] > Date Then
        MsgBox "Trustee Affidavit cannot be printed before Sale date", vbCritical
        Exit Sub
      End If
      
    End If
    
    If Me.State = "DC" Then
    Call DoReport("DC Trustee Affidavit", PrintTo)
    Else
    Call DoReport("Trustees Affidavit", PrintTo)
    End If

End If

If chDeedConv Then

    
    If (IsNull([Forms]![foreclosuredetails]!Purchaser)) Then
    MsgBox ("Missing  Purchaser .")
    Exit Sub
    End If

    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If

    If (Me!State = "MD" Or Me!State = "DC") Then
    SetSaleConductedTrustee
    End If
    'DoCmd.OpenForm "PrintClientAbstractors"
    If Me.State = "MD" Then Call DoReport("Clerk Cover Letter", PrintTo)
    Call DoReport("Conventional Deed " & Me!State, PrintTo)
End If

If chExemptDeed Then
    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If
    
  If (Me!State = "MD" Or Me!State = "DC") Then SetSaleConductedTrustee
  If Me.State = "MD" Then Call DoReport("Clerk Cover Letter", PrintTo)
  Call DoReport("Exempt Deed " & Me!State, PrintTo)
End If


If chDeedHUD Then

    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If

    Dim rst As Recordset
    Dim Atto As String
    Atto = Forms![Foreclosureprint]![Attorney]
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID = " & Atto & ";", dbOpenSnapshot)
  
  If (Me!State = "MD") Then
        If rst![PracticeMD] = 0 Then
        MsgBox (" This is not a MD attorney")
        Exit Sub
        End If
  End If
    
  If (Me!State = "DC") Then
            If rst![PracticeDC] = 0 Then
            MsgBox (" This is not a DC attorney")
            Exit Sub
            End If
            
   End If
        
   If (Me!State = "VA") Then
            
            If rst![PracticeVA] = -1 Then
            
                    If IsNull(rst![VABar]) Then
                    MsgBox ("This Attoreny is missing Virginia bar number")
                    Exit Sub
                    End If
                    Else
                If rst![PracticeVA] = 0 Then
                MsgBox ("This is not a VA attorney")
                Exit Sub
                End If
            End If
    End If

  If (Me!State = "MD" Or Me!State = "DC") Then
    SetSaleConductedTrustee
  End If
    
  If Me.State = "MD" Then Call DoReport("Clerk Cover Letter", PrintTo)
  Call DoReport("HUD Deed " & Me!State, PrintTo)
  
End If

If chDeedVA Then

    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If
    
    Dim rstVA As Recordset
    Dim AttoVA As String
    AttoVA = Forms![Foreclosureprint]![Attorney]
    Set rstVA = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID = " & AttoVA & ";", dbOpenSnapshot)
  
  If (Me!State = "MD") Then
        If rstVA![PracticeMD] = 0 Then
        MsgBox (" This is not a MD attorney")
        Exit Sub
        End If
  End If
    
  If (Me!State = "DC") Then
            If rstVA![PracticeDC] = 0 Then
            MsgBox (" This is not a DC attorney")
            Exit Sub
            End If
            
   End If
        
   If (Me!State = "VA") Then
         
            If rstVA![PracticeVA] = -1 Then
            
                    If IsNull(rstVA![VABar]) Then
                    MsgBox ("This Attoreny is missing Virginia bar number")
                    Exit Sub
                    End If
                    Else
                If rstVA![PracticeVA] = 0 Then
                MsgBox ("This is not a VA attorney")
                Exit Sub
                End If
            End If
    End If
   
    Select Case Me!State
    Case "MD"
        SetSaleConductedTrustee
        If Me!State = "MD" Then Call DoReport("Clerk Cover Letter", PrintTo)
        Call DoReport("Substitute Purchaser Deed MD", PrintTo)
    Case "VA"
        Call DoReport("VA Deed VA", PrintTo)
    Case "DC"
        SetSaleConductedTrustee
        Call DoReport("Conventional Deed DC", PrintTo)
    End Select
End If

If chDeedSubsPurchaser Then

    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If
    
  If Me!State = "MD" Then Call DoReport("Clerk Cover Letter", PrintTo)
  Call DoReport("Substitute Purchaser Deed " & Me!State, PrintTo)
End If

If chDeedConv Or chExemptDeed Or chDeedHUD Or chDeedVA Or chDeedSubsPurchaser Then
    If MsgBox("Update Deed Sent = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
        Forms!foreclosuredetails!DeedtoRec = Now()
        If MsgBox("Deed sent to record?" & vbNewLine & "(No = Deed sent to Title Company)", vbYesNo + vbQuestion) = vbYes Then
            AddStatus [CaseList.FileNumber], Now(), "Sent deed to record"
            Forms!foreclosuredetails!DeedtoTitleCo = False
        Else
            AddStatus [CaseList.FileNumber], Now(), "Sent deed to Title Company"
            Forms!foreclosuredetails!DeedtoTitleCo = True
        End If
    End If
End If

If chWithdrawBond Then
    If chWithdrawBondBK Then
        Call DoReport("Withdraw Bond BK", PrintTo)
    Else
        Call DoReport("Withdraw Bond", PrintTo)
    End If
End If

If chResell Then
    DoReport "Resale", PrintTo
    DoReport "Resale Order", PrintTo
    DoReport "Resale Show Cause", PrintTo
    If MsgBox("Update Resale Motion = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
        Forms!foreclosuredetails!Resell = 1
        Forms!foreclosuredetails!ResellMotion = Date
        AddStatus [CaseList.FileNumber], Date, "Motion to Resell"
    End If
End If

If chWithdrawSale Then
    Call CheckPrintInfo(Forms!foreclosuredetails!FileNumber)
    DoCmd.OpenForm "Print Withdraw Sale", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , , PrintTo
End If


If chRecordingCoverLetter And Me.State = "VA" Then
    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
        MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If
   
    If chRecordingCoverLetter And Me.State = "VA" Then
        Call CheckPrintInfo(Forms!foreclosuredetails!FileNumber)
        DoCmd.OpenForm "Print Deed Recording Cover Letter", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , , PrintTo
    End If
ElseIf chRecordingCoverLetter And Me.State = "MD" Then
        Call CheckPrintInfoMD(Forms!foreclosuredetails!FileNumber)
        If Me.State = "MD" Then Call DoReport("Clerk Cover Letter", PrintTo)
        DoCmd.OpenForm "Print Deed Recording Cover Letter MD", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , , PrintTo
End If

If chDismissCase Then Call DoReport("Dismiss Case", PrintTo)

If Me.chAffCoverLetter Then Call DoReport("Affidavit Cover Letter", PrintTo)

If chDeedInLieu Then
    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If
    
    If Me.State = "MD" Then Call DoReport("Clerk Cover Letter", PrintTo)
    If chDeedInLieu Then Call DoReport("Deed in Lieu", PrintTo)
End If

If chCertofService Then DoCmd.OpenForm "Print Certificate of Service", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , , PrintTo

If chOAH Then Call DoReport("OAH Background Letter", PrintTo)

If Me.chDIL Then Call DoReport("DIL", PrintTo)
If Me.chDILCertificate Then Call DoReport("DIL Certificate", PrintTo)
If Me.chDILJudgment Then Call DoReport("Dil Judgment Affidavit", PrintTo)
If Me.chDILLetter Then Call DoReport("DIL Letter to Borrower", PrintTo)

If PRletter Then 'My KINGDOM for indented IF statements
    If Me.State = "VA" Then
        If MsgBox("Update Prop Reg. = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
            Forms!foreclosuredetails!PropReg = Now()
            AddStatus [CaseList.FileNumber], Now(), "VA Property Registration sent"
            AddInvoiceItem [CaseList.FileNumber], "VA-PropReg", "Virginia Property Registration", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, False, False, True
            'DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter postage for Property Registration|FC|VA Property Registration mailed"
            Call DoReport("PropertyRegistrationVA", PrintTo)
        Else
            Call DoReport("PropertyRegistrationVA", PrintTo)
        End If
    
    ElseIf Me.State = "MD" And Me.CaseTypeID = 1 And Forms![foreclosuredetails].DispositionDesc = "Buy-In" Then
            If (IsNull(Forms![foreclosuredetails]!SaleRat)) Then
                MsgBox ("Missign Sale Rate ")
                Exit Sub
            ElseIf MsgBox("Update Prop Reg. = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
                Forms!foreclosuredetails!PropReg = Now()
                AddStatus [CaseList.FileNumber], Now(), "Property Registration sent"
                Call DoReport("PropertyRegistrationMD", PrintTo)
            
            'AddStatus [CaseList.FileNumber], Forms!ForeclosureDetails!PropReg, "Property Registration sent"
            Call DoReport("PropertyRegistrationMD", PrintTo)
            End If
    End If
End If
 
'If FCRegForm Then
'    If Me.State = "MD" Then
'    Select Case Forms!foreclosureDetails!City
'    Case "Annapolis"
'    Call DoReport("RegAnnapolis", PrintTo)
'    Case "Poolesville"
'    Call DoReport("RegPoolesville", PrintTo)
'    Case "College Park"
'    Call DoReport("RegCollegPark", PrintTo)
'    Case "Salisbury"
'    Call DoReport("RegSalisbury", PrintTo)
'    Case "Laurel"
'    Call DoReport("RegLaurel", PrintTo)
'    End Select
'    If Forms![Case list]!JurisdictionID = 18 Then Call DoReport("RegPrinceGeorge", PrintTo)
'    End If
'End If
'
'If FCRegLable Then
'    If PrintTo = acViewNormal Then
'        sql = "SELECT CaseList.FileNumber, CaseList.PrimaryDefName, ClientList.ShortClientName, Names.Company, Names.Last, Names.First, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip, [Notice Label Copies].CopyNumber, Names.Noteholder FROM [Notice Label Copies], ClientList INNER JOIN (CaseList INNER JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID WHERE (((CaseList.FileNumber)=" & Forms![Case list]!FileNumber & ") And  (([Notice Label Copies].CopyNumber)=1)AND ((Names.Owner)=True))  ORDER BY Names.Last, Names.First, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip;"
'        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
'        Do While Not rstLabelData.EOF
'            Call StartLabel
'            Print #6, FormatName(rstLabelData!Company, rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
'            Print #6, "|FONTSIZE 8"
'            Print #6, "|BOTTOM"
'            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
'            Call FinishLabel
'            rstLabelData.MoveNext
'        Loop
'        rstLabelData.Close
'    End If
'End If

If chCreateLabel Then DoCmd.OpenForm "Getlabel"

'added monitor label print on 6/11/15
If chSODLabel Then

    If PrintTo = acViewNormal Then
        sql = "SELECT Names.Company, Names.First, Names.Last, Names.AKA, Names.Address, Names.Deceased, Names.Address2, Names.City, Names.State, Names.Zip, CaseList.FileNumber, ClientList.FairDebt, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (ClientList RIGHT JOIN (CaseList RIGHT JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID) LEFT JOIN FCdetails ON CaseList.FileNumber=FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & Forms![Case List]!FileNumber & ") And ((Names.owner)=True) And ((FCdetails.Current)=True)and ((Names.Mortgagor) = true or (Names.Noteholder) = true));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
            If rstLabelData!Deceased = True Then
            Print #6, FormatName("", "The Estate of " & rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            Else
            Print #6, FormatName(rstLabelData!Company, rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            End If
            Print #6, "|FONTSIZE 8"
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
    End If
End If

If chMotionToInterveneLable Then

    If PrintTo = acViewNormal Then
        sql = "SELECT Names.Company, Names.First, Names.Last, Names.AKA, Names.Address, Names.Deceased, Names.Address2, Names.City, Names.State, Names.Zip, CaseList.FileNumber, ClientList.FairDebt, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (ClientList RIGHT JOIN (CaseList RIGHT JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID) LEFT JOIN FCdetails ON CaseList.FileNumber=FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & Forms![Case List]!FileNumber & ") And ((Names.owner)=True) And ((FCdetails.Current)=True)and ((Names.Mortgagor) = true or (Names.Noteholder) = true));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
            If rstLabelData!Deceased = True Then
            Print #6, FormatName("", "The Estate of " & rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            Else
            Print #6, FormatName(rstLabelData!Company, rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            End If
            Print #6, "|FONTSIZE 8"
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
    End If
End If

If chMotiontoReleaseFundsLable Then

    If PrintTo = acViewNormal Then
        sql = "SELECT Names.Company, Names.First, Names.Last, Names.AKA, Names.Address, Names.Deceased, Names.Address2, Names.City, Names.State, Names.Zip, CaseList.FileNumber, ClientList.FairDebt, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (ClientList RIGHT JOIN (CaseList RIGHT JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID) LEFT JOIN FCdetails ON CaseList.FileNumber=FCdetails.FileNumber WHERE (((CaseList.FileNumber)=" & Forms![Case List]!FileNumber & ") And ((Names.owner)=True) And ((FCdetails.Current)=True)and ((Names.Mortgagor) = true or (Names.Noteholder) = true));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
            If rstLabelData!Deceased = True Then
            Print #6, FormatName("", "The Estate of " & rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            Else
            Print #6, FormatName(rstLabelData!Company, rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
            End If
            Print #6, "|FONTSIZE 8"
            Print #6, "|BOTTOM"
            Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
    End If
End If


 
Exit Sub

Err_PrintDocs:
    MsgBox Err.Description
    Exit Sub

End Sub

Private Sub Form_Current()
Me.Caption = "Print Foreclosure " & [CaseList.FileNumber] & " " & [PrimaryDefName]

If OpenArgs = "TitleReview" Then
Dim ctrl As Control
For Each ctrl In Me.Form.Controls

If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is Rectangle Or TypeOf ctrl Is Label Then  ' TypeOf ctrl Is CommandButton Then
'If Not ctrl.Locked) Then
If ctrl.Name = "chTitleReview" Or ctrl.Name = "Label179" Then
ctrl.Visible = True
Else
ctrl.Visible = False
End If
'Else
'ctrl.Locked = True
End If
'End If
: Next
End If



If (Not IsNull(FairDebtDispute) And IsNull(FairDebtVerified)) Then
  Me.chBaileeLetter.Enabled = False
  ch45Day.Enabled = False
  chNOILabels.Enabled = False
  Me.chAffCoverLetter.Enabled = False
  chFairDebt.Enabled = False
  chFairDebtLabels.Enabled = False
  chLossMitSol.Enabled = False
  chLossMitSolLabels.Enabled = False
  chTitleOrder.Enabled = False
  chDeedOfApp.Enabled = False
  chHUDOcc.Enabled = False
  chDOARecordingCover.Enabled = False
  chDOAAffidavit.Enabled = False
  chSOD.Enabled = False
  chSOD2.Enabled = False
  chMilitaryAffidavitActive.Enabled = False
  chMilitaryAffidavitNoSSN.Enabled = False
  chMilitaryAffidavit.Enabled = False
  chLostNoteAffidavit.Enabled = False
  chLostNoteNotice.Enabled = False
  chLostNoteNoticeLabels.Enabled = False
  chDocket.Enabled = False
  chDebtorLetterLabels.Enabled = False
  chNoteOwnership.Enabled = False
  chAffMD7105.Enabled = False
  chDOTAffidavit.Enabled = False
  ChNoteAffidavit.Enabled = False
  ChCollateralFileAffidavit.Enabled = False
  Ch14207Affidavit.Enabled = False
  chBondOrder.Enabled = False
  chPSCover.Enabled = False
  chAffidavitOfService.Enabled = False
  chNoticeToOccupant.Enabled = False
  chNoticeToOccupantLabel.Enabled = False
  chAuctioneer.Enabled = False
  chNewspaperAd.Enabled = False
  chIRSNotice.Enabled = False
  chNotice.Enabled = False
  chNoticeLabels.Enabled = False
  chCountyAttyLabel.Enabled = False
  chReadAtSale.Enabled = False
  chAssignment.Enabled = False
  chAssignRecCovLtr = False
  chPayoff.Enabled = True
  chTitleReview.Enabled = False
  chTitleClaim.Enabled = False
    
  Me.chWarranty.Enabled = False
  Me.chQCDeed.Enabled = False
  chLossMitPrelim.Enabled = False
  chLossMitFinal.Enabled = False
  chLossMitApp.Enabled = False
  chForecloseMed.Enabled = False
  chReportOfSale.Enabled = False
  chSubsPurch.Enabled = False
  chTrusteeAffidavit.Enabled = False
  chWithdrawBond.Enabled = False
  chWithdrawBondBK.Enabled = False
  chClaimSurplus.Enabled = False
  chSettlement.Enabled = False
  chResell.Enabled = False
  chWithdrawSale.Enabled = False
  chDismissCase.Enabled = False
  chDeedConv.Enabled = False
  chExemptDeed.Enabled = False
  chDeedHUD.Enabled = False
  chDeedVA.Enabled = False
 ' chDeedSubsPurchaser.Enabled = False
  chRecordingCoverLetter.Enabled = False
  chDeedInLieu.Enabled = False
  chCertofService.Enabled = False
  PRletter.Enabled = False
  PRlabel.Enabled = False
  ChNoteAffidavit.Enabled = False
  ChCollateralFileAffidavit.Enabled = False
  ChLabel.Enabled = False
  
  
Else
  'chAffidavitOfService.Enabled = Nz(State = "MD")
  
  'chDeedSubsPurchaser.Enabled = Nz(State = "DC")
  'chLostNoteNotice.Enabled = Nz((State = "VA"))
  'chLostNoteNoticeLabels.Enabled = Nz((State = "VA"))
  'Disabled temporarily since I want to use this for Maryland TOO
  'chRecordingCoverLetter.Enabled = Nz((State = "VA"))
  'chClaimSurplus.Enabled = (CaseTypeID = 8)
  'chExemptDeed.Enabled = Nz(LoanType = 4) Or Nz(LoanType = 5)
  'chDeedConv.Enabled = Nz(LoanType <> 4) And Nz(LoanType <> 5)
  'chLossMitSol.Enabled = Nz(LoanType = 4) Or Nz(LoanType = 5)
  'chLossMitSolLabels.Enabled = Nz(LoanType = 4) Or Nz(LoanType = 5)
End If

If Me.State = "DC" Then
    Me.chAffCoverLetter.Visible = True
Else
    Me.chAffCoverLetter.Visible = False
End If

'#1249 Lock printing special GSE deeds if loan type doesn't match   10/29/2014

Me.chDeedConv.Enabled = False
Me.chDeedHUD.Enabled = False
'Me.chDeedSubsPurchaser.Enabled = False
Me.chExemptDeed.Enabled = False
Me.chDeedVA.Enabled = False

Select Case Me.LoanType

    Case 1 ' Conventional
        Me.chDeedConv.Enabled = True
    Case 5 ' FHLMC
        Me.chExemptDeed.Enabled = True
    Case 4 ' FNMA
        Me.chExemptDeed.Enabled = True
    Case 3 ' HUD
        Me.chDeedHUD.Enabled = True
        Me.chDeedConv.Enabled = True
    Case 2 ' VA
        If Me.State = "MD" Then
            Me.chDeedSubsPurchaser.Enabled = True
        ElseIf Me.State = "VA" Then
             Me.chDeedVA.Enabled = True
        Else
        End If
End Select
'Always enable Substitute Purchaser Deed for MD Regardless of Loantype
If Me.State = "MD" Then Me.chDeedSubsPurchaser.Enabled = True





If Me.State = "VA" Then
  Attorney.RowSource = "SELECT Staff.ID, Staff.Name & ', ' & [CommonWealthTitle] AS CWRep " & _
                       "FROM Staff " & _
                       "WHERE  ((staff.active = true) And (Staff.Attorney =True) And(staff.PracticeVA = true )) " & _
                       "ORDER BY Staff.CommonwealthTitle, Staff.Sort;"
                   'It was  "WHERE (((Staff.CommonwealthTitle) Is Not Null)) and staff.active = true " S.A.
'staff.active=true
ElseIf Me.State = "MD" Then
  Attorney.RowSource = "SELECT Staff.ID, Staff.Name & ', Esq.' FROM Staff WHERE ((Staff.active = true ) and (Staff.Attorney = True) and (Staff.PracticeMD = True)) ORDER BY Staff.Sort;"
Else
  Attorney.RowSource = "SELECT Staff.ID, Staff.Name FROM Staff WHERE ((staff.active = true) and (Staff.Attorney = True) and (staff.PracticeDC = true ))ORDER BY Staff.Sort;"

End If

If ([Forms]![Case List]!Active = False) Then
    
  Me.chWarranty.Enabled = False
  Me.chQCDeed.Enabled = False
  ch45Day.Enabled = False
  Me.chAffCoverLetter.Enabled = False
  chNOILabels.Enabled = False
  chFairDebt.Enabled = False
  chFairDebtLabels.Enabled = False
  chLossMitSol.Enabled = False
  chLossMitSolLabels.Enabled = False
  chTitleOrder.Enabled = False
  chDeedOfApp.Enabled = False
  Me.chBaileeLetter.Enabled = False
  chHUDOcc.Enabled = False
  chDOARecordingCover.Enabled = False
  chDOAAffidavit.Enabled = False
  chSOD.Enabled = False
  chSOD2.Enabled = False
  chMilitaryAffidavitActive.Enabled = False
  chMilitaryAffidavitNoSSN.Enabled = False
  chMilitaryAffidavit.Enabled = False
  chLostNoteAffidavit.Enabled = False
  chLostNoteNotice.Enabled = False
  chLostNoteNoticeLabels.Enabled = False
  chDocket.Enabled = False
  chDebtorLetterLabels.Enabled = False
  chNoteOwnership.Enabled = False
  chAffMD7105.Enabled = False
  chDOTAffidavit.Enabled = False
  ChNoteAffidavit.Enabled = False
  ChCollateralFileAffidavit.Enabled = False
  Ch14207Affidavit.Enabled = False
  
  chBondOrder.Enabled = False
  chPSCover.Enabled = False
  chAffidavitOfService.Enabled = False
  chNoticeToOccupant.Enabled = False
  chNoticeToOccupantLabel.Enabled = False
  chAuctioneer.Enabled = False
  chNewspaperAd.Enabled = False
  chIRSNotice.Enabled = False
  chNotice.Enabled = False
  chNoticeLabels.Enabled = False
  chCountyAttyLabel.Enabled = False
  chReadAtSale.Enabled = False
  chAssignment.Enabled = False
  chAssignRecCovLtr = False
  chPayoff.Enabled = True
  chTitleReview.Enabled = False
  chTitleClaim.Enabled = False
  chLandInstruments.Enabled = False
  chCourtNotes.Enabled = False
  chLossMitPrelim.Enabled = False
  chLossMitFinal.Enabled = False
  chLossMitApp.Enabled = False
  chForecloseMed.Enabled = False
  chReportOfSale.Enabled = False
  chSubsPurch.Enabled = False
  chTrusteeAffidavit.Enabled = False
  chWithdrawBond.Enabled = False
  chWithdrawBondBK.Enabled = False
  chClaimSurplus.Enabled = False
  chSettlement.Enabled = False
  chResell.Enabled = False
  chWithdrawSale.Enabled = False
  'chDismissCase.Enabled = False  ' Ticket 888 6/9/14 MC
  chDeedConv.Enabled = False
  chExemptDeed.Enabled = False
  chDeedHUD.Enabled = False
  chDeedVA.Enabled = False
 ' chDeedSubsPurchaser.Enabled = False
  chRecordingCoverLetter.Enabled = False
  chDeedInLieu.Enabled = False
  chCertofService.Enabled = False
  chAffidavitOfService.Enabled = False
  chAssignRecCovLtr.Enabled = False
  chPayoff.Enabled = False
  chOAH.Enabled = False
  chDeedConv.Enabled = False
  'chDeedSubsPurchaser.Enabled = False
  Attorney.Enabled = False
  NotaryID.Enabled = False
  PRletter.Enabled = False
  Label263.Visible = True
  PRlabel.Enabled = False
  ChNoteAffidavit.Enabled = False
  ChCollateralFileAffidavit.Enabled = False
  ChLabel.Enabled = False
  
  End If
  
  'Hide certain option based on wizard
'Select Case WizardSource
Select Case OpenArgs
Case "Restart", "Intake", "Docketing", "FLMA", "SaleSetting", "Service", "ServiceMailed", "Title"
With Forms!Foreclosureprint
!chTitleOrder.Visible = False
!chNoticeLabels.Visible = False
!Box134.Visible = False
!Box126.Visible = False
!Label167.Visible = False
!Label113.Visible = False
!chWithdrawBondBK.Visible = False
!Label251.Visible = False
!chBondOrder.Visible = False
!Label212.Visible = False
!chPSCover.Visible = False
!Label199.Visible = False
!chAffidavitOfService.Visible = False
!Label207.Visible = False
!chNoticeToOccupant.Visible = False
!Label139.Visible = False
!chAuctioneer.Visible = False
!Label131.Visible = False
!chNewspaperAd.Visible = False
!Label151.Visible = False
!chIRSNotice.Visible = False
!chBaileeLetter.Visible = False
!Label111.Visible = False
!chNotice.Visible = False
!Label224.Visible = False
!chCountyAttyLabel.Visible = False
!Label209.Visible = False
!chNoticeToOccupantLabel.Visible = False
!Label133.Visible = False
!chReadAtSale.Visible = False
!Label153.Visible = False
!chAssignment.Visible = False
!Label249.Visible = False
!chAssignRecCovLtr.Visible = False
!Label157.Visible = False
!chQCDeed.Visible = False
!chPayoff.Visible = False
!Label179.Visible = False
!chTitleReview.Visible = False
!Label155.Visible = False
!chTitleClaim.Visible = False
'!Label227.Visible = False
'!chLossMitPrelim.Visible = False
'!Label229.Visible = False
'!chLossMitFinal.Visible = False
!Label240.Visible = False
!chLossMitApp.Visible = False
!Label231.Visible = False
!chForecloseMed.Visible = False
!Label262.Visible = False
!chOAH.Visible = False
!Label135.Visible = False
!Label137.Visible = False
!chReportOfSale.Visible = False
!Label147.Visible = False
!chSubsPurch.Visible = False
!Label163.Visible = False
!chTrusteeAffidavit.Visible = False
!Label165.Visible = False
!chWithdrawBond.Visible = False
!Label176.Visible = False
!chClaimSurplus.Visible = False
!Label183.Visible = False
!chSettlement.Visible = False
!Label171.Visible = False
!chResell.Visible = False
!Label173.Visible = False
!chWithdrawSale.Visible = False
!Label187.Visible = False
!chDismissCase.Visible = False
'!chDeedOfApp.SetFocus

!Label127.Visible = False
!Label119.Visible = False
!chDeedConv.Visible = False
!Label253.Visible = False
!chExemptDeed.Visible = False
!Label121.Visible = False
!chDeedHUD.Visible = False
!Label123.Visible = False
!chDeedVA.Visible = False
!Label149.Visible = False
!chDeedSubsPurchaser.Visible = False
!Label193.Visible = False
!chRecordingCoverLetter.Visible = False
!Label189.Visible = False
!chDeedInLieu.Visible = False
!Label246.Visible = False
!chCertofService.Visible = False
!PRletter.Visible = False
!PRlabel.Visible = False
!ChNoteAffidavit.Visible = False
!ChCollateralFileAffidavit.Visible = False
!ChLabel.Visible = False


If OpenArgs = "docketing" Or OpenArgs = "Title" Then
!Ch14207Affidavit.Visible = True
!Label240.Visible = True
!Label231.Visible = True
!chLossMitApp.Visible = True
!chForecloseMed.Visible = True
!chLossMitSol.Visible = False
!chHUDOcc.Visible = False
!chFairDebt.Visible = False
!ch45Day.Visible = False
!chFairDebtLabels.Visible = False
!chLossMitSolLabels.Visible = False
!chNOILabels.Visible = False
!chLostNoteNotice.Visible = False
!chLostNoteNoticeLabels.Visible = False
!chLostNoteAffidavit.Visible = False
!FCRegForm.Visible = True
!FCRegLable.Visible = True
!chBaileeLetter.Visible = False

End If
If OpenArgs = "FLMA" Or OpenArgs = "ServiceMailed" Or OpenArgs = "Title" Then
!chLossMitFinal.SetFocus
!chFairDebtLabels.Visible = False
!chLossMitSolLabels.Visible = False
!chNOILabels.Visible = False
!chNoteOwnership.Visible = False
!chAffMD7105.Visible = False
!chDOTAffidavit.Visible = False
!chBondOrder.Visible = False
!ch45Day.Visible = False
!chFairDebt.Visible = False
!chLossMitSol.Visible = False
!chTitleOrder.Visible = False
!chHUDOcc.Visible = False
!chDeedOfApp.Visible = False
!chDOARecordingCover.Visible = False
!chDOAAffidavit.Visible = False
!chSOD.Visible = False
!chSOD2.Visible = False
!chMilitaryAffidavitActive.Visible = False
!chMilitaryAffidavit.Visible = False
!chMilitaryAffidavitNoSSN.Visible = False
!Label240.Visible = True
!chLossMitApp.Visible = True
!Label231.Visible = True
!chForecloseMed.Visible = True
End If
If OpenArgs = "SaleSetting" Or OpenArgs = "Service" Or OpenArgs = "ServiceMailed" Or OpenArgs = "Title" Then
!chLossMitFinal.SetFocus
!chFairDebtLabels.Visible = False
!chLossMitSolLabels.Visible = False
!chNOILabels.Visible = False
!chNoteOwnership.Visible = False
!chAffMD7105.Visible = False
!chDOTAffidavit.Visible = False
!chBondOrder.Visible = False
!ch45Day.Visible = False
!chFairDebt.Visible = False
!chLossMitSol.Visible = False
!chTitleOrder.Visible = False
!chHUDOcc.Visible = False
!chDeedOfApp.Visible = False
!chDOARecordingCover.Visible = False
!chDOAAffidavit.Visible = False
!chSOD.Visible = False
!chSOD2.Visible = False
!chMilitaryAffidavitActive.Visible = False
!chMilitaryAffidavit.Visible = False
!chMilitaryAffidavitNoSSN.Visible = False
!chAuctioneer.Visible = True
!chAuctioneer.SetFocus
!Label139.Visible = True
!chLostNoteAffidavit.Visible = False
!chLostNoteNotice.Visible = False
!chDocket.Visible = False
!chLossMitPrelim.Visible = False
!chLossMitFinal.Visible = False
!chLossMitApp.Visible = False
!chLossMitMailing.Visible = False
!ChNoteAffidavit.Visible = False
!chLostNoteNoticeLabels.Visible = False
!chDebtorLetterLabels.Visible = False
End If

If OpenArgs = "Service" Then
!chPSCover.Visible = True
!Label212.Visible = True
!chPSCover.SetFocus
!chAuctioneer.Visible = False
End If

If OpenArgs = "ServiceMailed" Then
!chAffidavitOfService.Visible = True
!chNoticeToOccupant.Visible = True
!chNoticeToOccupantLabel.Visible = True
!chLossMitSolLabels.Visible = False
!chNOILabels.Visible = False
!chLostNoteNotice.Visible = False
!chLostNoteNoticeLabels.Visible = False
!chLostNoteAffidavit.Visible = False
!chAffidavitOfService.SetFocus
!Label199.Visible = True
!Label207.Visible = True
!Label209.Visible = True
!chAuctioneer.Visible = False
!chForecloseMed.Visible = False
End If
End With

Case "HUDOcc"
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm "Print HUD Occ Letter"
Forms![print hud occ letter]!FileNumber = FileNumber
End Select
If OpenArgs = "Title" Then

'chTitleOrder.Visible = True
'chTitleOrder.SetFocus
'chTitleOrder.Enabled = True
chAuctioneer.Visible = False
chCoCounselLetter.Visible = False
chForecloseMed.Visible = False
Ch14207Affidavit.Visible = False
'FCRegForm.Visible = False
'FCRegLable.Visible = False

End If

'Deed in Lieu Section

If OpenArgs = "None" Then
    If Len(Forms!foreclosuredetails!sfrmFCDIL![DILReferralReceived] & "") = 0 Then
    Else
        chDIL.Visible = True
        chDILCertificate.Visible = True
        chDILJudgment.Visible = True
        chDILLetter.Visible = True
        Label299.Visible = True
        Label301.Visible = True
        Label295.Visible = True
        Label297.Visible = True
        Label304.Visible = True
        Box298.Visible = True
    End If
Else
End If

If Forms!foreclosuredetails!Disposition = 2 Then
    chDeedConv.Enabled = True
End If


End Sub

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close
Call refreshFCform
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

Private Sub cmdAcrobat_Click()
Call PrintDocs(-2)
End Sub

Private Sub SetSaleConductedTrustee()

 If IsNull(SaleConductedTrusteeID) Then
    SelectedTrusteeID = 0
    DoCmd.OpenForm "SetTrustee", , , , , acDialog
    If SelectedTrusteeID > 0 Then
        
  ' this did not work...had to update through runsql command
  '     If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
  '
  '    Forms!ForeclosureDetails!SaleConductedTrusteeID = SelectedTrusteeID
      DoCmd.SetWarnings False
      DoCmd.RunSQL ("update FCdetails set SaleConductedTrusteeID = " & SelectedTrusteeID & " where [FileNumber] = " & Me.FileNumber & " and current=true")
      DoCmd.SetWarnings True
        
   End If
 End If
End Sub

Private Sub DesignatedAttorneyPrint(PrintTo As Integer)
 If Forms!Foreclosureprint!txtDesignatedAttorney = 3 Then 'A Designated Attorney was Selected 'MARKER
            If Forms![Case List]!ClientID = 446 Then
                DoCmd.Close acReport, "45 Day Notice Affidavit BOA", acSaveYes
                DoCmd.Close acReport, "BOA Cover Sheet", acSaveYes
            ElseIf Forms![Case List]!ClientID = 6 Then
                DoCmd.Close acReport, "45 Day Notice Affidavit Wells", acSaveYes
            ElseIf Forms![Case List]!ClientID = 556 Then
                DoCmd.Close acReport, "45 Day Notice Affidavit Wells", acSaveYes
            ElseIf Forms![Case List]!ClientID = 97 Then
                DoCmd.Close acReport, "45 Day Notice Affidavit Chase", acSaveYes
            ElseIf Forms![Case List]!ClientID = 444 Then
                 DoCmd.Close acReport, "PHH Cover Sheet", acSaveYes
                 DoCmd.Close acReport, "45 Day Notice Affidavit", acSaveYes
            Else
                DoCmd.Close acReport, "45 Day Notice Affidavit", acSaveYes
            End If
            'f PrintTo = -2 Then
             '"""   DoCmd.OpenReport "45 Day Notice Affidavit"
           ' Else
            'DoCmd.OpenReport "45 Day Notice Affidavit", PrintTo, , , , OpenArgs
    Call DoReport("45 Day Notice Affidavit", PrintTo)
            
            
    End If
End Sub

Private Sub coCounselLetterProc(PrintTo As Integer)
If DLookup("AuctioneerCoCounsel", "jurisdictionlist", "JurisdictionID=" & JurisdictionID) = 196 Then
            MsgBox "Trustee is Commonwealth.  Please schedule the sale now.", vbCritical
            Exit Sub
            Else
            
        Dim Recipient As String, BorrowerName As String, PropertyAddress As String, Jurisdictiontxt As String, MessageText As String, DayRule As Integer
        Select Case Forms!foreclosuredetails!LoanType
        Case 1  'Conv
            DayRule = 19
        Case 2 'VA
            If Forms![Case List]!ClientID <> 97 Then
            DayRule = 50
            Else
            DayRule = 65
            End If
        Case 3 'HUD
            DayRule = 30
        Case 4  'FNMA
            DayRule = 25
        Case 5 'FHLMC
            DoCmd.OpenForm "EnterSaleSettingOption", , , , , acDialog
            If Forms!foreclosuredetails!Autovalue = "Autovalue" Then
            DayRule = 22
            End If
            If Forms!foreclosuredetails!Autovalue = "BPO" Then
                DayRule = 30
            End If
        End Select
          
            Jurisdictiontxt = "Jurisdiction:  " & DLookup("jurisdiction", "jurisdictionlist", "JurisdictionID=" & JurisdictionID)
            MessageText = "Please advise when we can schedule the sale for the below referenced property.   The sale must be set " & DayRule & " days out from today or no earlier than " & (Date + DayRule) & "."
            Recipient = DLookup("Email", "Vendors", "ID=" & DLookup("AuctioneerCoCounsel", "jurisdictionlist", "JurisdictionID=" & JurisdictionID))
            PropertyAddress = "Property Address:  " & Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & Forms!foreclosuredetails!ZipCode
            Dim olApp As Object
            Dim olMail As Object
            Set olApp = CreateObject("Outlook.Application")
            Set olMail = olApp.CreateItem(olMailItem)
            
            With olMail
                If Not IsMissing(Recipient) Then .To = Recipient
                .Subject = "Scheduling Sales"
                .Body = MessageText & vbNewLine & BorrowerName & vbNewLine & PropertyAddress & vbNewLine & "Date/Time:" & vbNewLine & Jurisdictiontxt & vbNewLine & "Our File Number:  " & FileNumber
                .Display
            End With
      End If


End Sub

Private Sub chReportofSaleProc(PrintTo As Integer)
If (Me!State = "MD") Then
      Call DoReport("Report of Sale Cover Letter", PrintTo)
    End If
    
    Call DoReport("Report of Sale", PrintTo)
    Call DoReport("Report of Sale Notice", PrintTo)
    Call DoReport("Report of Sale Notice", PrintTo)
    Select Case JurisdictionID
        Case 3          ' Anne Arundel MD
            Call DoReport("Report of Sale Line To Request Ratification", PrintTo)
            Call DoReport("Final Order of Ratification Short", PrintTo)
            Call DoReport("Final Order of Ratification Short", PrintTo)
            Call DoReport("Final Order of Ratification Long", PrintTo)
            Call DoReport("Final Order of Ratification Long", PrintTo)
        Case 10         ' Charles County MD
            Call DoReport("Final Order of Ratification Charles", PrintTo)
            Call DoReport("Final Order of Ratification Charles", PrintTo)
        Case 20         ' St Mary's County MD
            Call DoReport("Final Order of Ratification Short", PrintTo)
            Call DoReport("Final Order of Ratification Short", PrintTo)
        Case 12         'Frederick
            Call DoReport("Final Order of Ratification Fred", PrintTo)
        Case Else
            Call DoReport("Final Order of Ratification Short", PrintTo)
            Call DoReport("Final Order of Ratification Short", PrintTo)
            Call DoReport("Final Order of Ratification Long", PrintTo)
            Call DoReport("Final Order of Ratification Long", PrintTo)
    End Select
    
  
    Call DoReport("Report of Sale Affidavit", PrintTo)
End Sub

Private Sub checkInstruments(PrintTo As Integer)

Dim strCriteria As String
Dim rsLand As Recordset
Set rsLand = CurrentDb.OpenRecordset("LandInstrumentDetails", dbOpenDynaset, dbSeeChanges)

strCriteria = "Filenumber =" & Forms![foreclosuredetails]!FileNumber & ""

    rsLand.FindFirst (strCriteria)

If rsLand.NoMatch Then
    rsLand.AddNew
    rsLand!FileNumber = Forms!foreclosuredetails!FileNumber
    rsLand.Update
Else
End If

rsLand.Close
Set rsLand = Nothing

End Sub

Private Sub NewsPaperAdProc(PrintTo As Integer)
    Select Case Me!State
        Case "DC"
            Call DoReport("Ad DC", PrintTo)
        Case "MD"
            Call DoReport("Ad MD", PrintTo)
        Case "VA"
        If MsgBox("Are you printing notices and don't need an email generated?", vbYesNo) = vbYes Then
        Call DoReport("Ad VA", PrintTo)
        Else
        DoCmd.OpenForm "Print News Ad", , , "ForeclosureID=" & Forms!foreclosuredetails!ForeclosureID, , , PrintTo
        End If
        Case Else
            MsgBox "Newspaper Ad is not available for properties in " & Me!State, vbInformation
    End Select

End Sub
Private Sub DeedOfAppProc(PrintTo As Integer)



'wtf indent!
   If (Me.ClientID = 334 Or Me.ClientID = 477) Then
   DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Saxon|" & PrintTo
   
   ElseIf Me.ClientID = 466 And Me.State = "MD" Then 'SELECT
         DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Select|" & PrintTo
    
   ElseIf Me.ClientID = 345 Then
   DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Kondaur|" & PrintTo
   ElseIf (Me.ClientID = 328 And Me.State = "MD") Then
               DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment SPLS MD|" & PrintTo
      ElseIf (Me.ClientID = 328 And Me.State = "VA") Then
         DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment SPLS VA|" & PrintTo
   ElseIf Me.ClientID = 87 Then 'PNC
   'DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!ForeclosureDetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment PNC|" & PrintTo
   DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment|" & PrintTo
   
   ElseIf Me.ClientID = 451 And Me.State = "VA" Then 'Dove VA
   DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms![Case List]!FileNumber, , acDialog, "Substition of Trustee|Deed of Appointment Dove VA|" & PrintTo
    
   ElseIf Me.ClientID = 451 Then ' Dove all others
   DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms![Case List]!FileNumber, , acDialog, "Substition of Trustee|Deed of Appointment Dove|" & PrintTo
    
   ElseIf Me.ClientID = 523 And Me.State = "VA" Then 'GreenTree VA
         DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms![Case List]!FileNumber, , acDialog, "Substition of Trustee|Deed of Appointment GreenTreeVA|" & PrintTo
   ElseIf Me.ClientID = 523 Then
         DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms![Case List]!FileNumber, , acDialog, "Substition of Trustee|Deed of Appointment GreenTree|" & PrintTo
   
   ElseIf Me.ClientID = 97 And Me.State = "MD" Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Chase|" & PrintTo
   ElseIf Me.ClientID = 97 And Me.State = "VA" Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment ChaseVA|" & PrintTo
   ElseIf Me.ClientID = 6 And Me.State = "MD" Then 'Wells
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Wells|" & PrintTo
   ElseIf Me.ClientID = 556 And Me.State = "MD" Then 'Wells
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Wells|" & PrintTo
   ElseIf Me.ClientID = 6 And Me.State = "VA" Then 'Wells
       ' If Me.JurisdictionID = 153 Or Me.JurisdictionID = 35 Or Me.JurisdictionID = 79 Or Me.JurisdictionID = 47 Or Me.JurisdictionID = 36 Or Me.JurisdictionID = 42 Or Me.JurisdictionID = 58 Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment WellsVA Select|" & PrintTo

   ElseIf Me.ClientID = 556 And Me.State = "VA" Then 'Wells
       ' If Me.JurisdictionID = 153 Or Me.JurisdictionID = 35 Or Me.JurisdictionID = 79 Or Me.JurisdictionID = 47 Or Me.JurisdictionID = 36 Or Me.JurisdictionID = 42 Or Me.JurisdictionID = 58 Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment WellsVA Select|" & PrintTo
   
   ElseIf Me.ClientID = 361 And Me.State = "MD" Then 'Ocwen
   DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Ocwen|" & PrintTo
   ElseIf Me.ClientID = 124 And Me.State = "MD" Then 'Ocwen
   DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Ocwen|" & PrintTo
   ElseIf Me.ClientID = 531 And Me.State = "MD" Then 'MDHC
        DoCmd.OpenForm "PrintClientDoc", , , "Filenumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment MDHC|" & PrintTo
   ElseIf Me.ClientID = 456 And Me.State = "MD" Then 'M&T
        DoCmd.OpenForm "PrintClientDoc", , , "Filenumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment M&T|" & PrintTo
    ElseIf Me.ClientID = 532 And Me.State = "MD" Then 'Selene MD
        DoCmd.OpenForm "PrintClientDoc", , , "Filenumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Selene|" & PrintTo
   
   ElseIf Me.ClientID = 532 And Me.State = "VA" Then 'Selene VA
        DoCmd.OpenForm "PrintClientDoc", , , "Filenumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment SeleneVA|" & PrintTo
   ElseIf Me.ClientID = 532 And Me.State = "DC" Then 'Selene DC
        DoCmd.OpenForm "PrintClientDoc", , , "Filenumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment SeleneDC|" & PrintTo
   
  
   ElseIf Me.ClientID = 404 Then
        DoCmd.OpenForm "PrintClientDoc", , , "Filenumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Bogman|" & PrintTo
   
   ElseIf Me.JurisdictionID = 153 Then 'ACCOMACK
        DoCmd.OpenForm "PrintClientDoc", , , "Filenumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment Accomack|" & PrintTo
            If Me.ClientID = 567 And Me!State <> "DC" Then
                Call DoReport("CHAM Cover DeedOfApp", PrintTo)
            ElseIf Me.ClientID = 385 And Me!State <> "DC" Then
                Call DoReport("Nationstar Cover Sheet", PrintTo)
            ElseIf Forms![Case List]!ClientID = 446 Then
                Call DoReport("BOA Cover Sheet", PrintTo)
            ElseIf Forms![Case List]!ClientID = 444 Then
                Call DoReport("PHH Cover Sheet", PrintTo)
            End If
   ElseIf Me.State = "VA" Then 'Virginia Deed of Appointment
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment VA|" & PrintTo
            If Me.ClientID = 567 And Me!State <> "DC" Then
                Call DoReport("CHAM Cover DeedOfApp", PrintTo)
            ElseIf Me.ClientID = 385 And Me!State <> "DC" Then
                Call DoReport("Nationstar Cover Sheet", PrintTo)
            ElseIf Forms![Case List]!ClientID = 446 Then
                Call DoReport("BOA Cover Sheet", PrintTo)
            ElseIf Forms![Case List]!ClientID = 444 Then
                Call DoReport("PHH Cover Sheet", PrintTo)
            End If
   
   Else
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Substitution of Trustee|Deed of Appointment|" & PrintTo
    
        If Me.ClientID = 567 And Me!State <> "DC" Then
            Call DoReport("CHAM Cover DeedOfApp", PrintTo)
        ElseIf Me.ClientID = 385 And Me!State <> "DC" Then
            Call DoReport("Nationstar Cover Sheet", PrintTo)
        Else
        End If
        If Forms![Case List]!ClientID = 446 Then
            Call DoReport("BOA Cover Sheet", PrintTo)
        ElseIf Forms![Case List]!ClientID = 444 Then
            Call DoReport("PHH Cover Sheet", PrintTo)
        End If
   End If
   
If Forms!foreclosuredetails!WizardSource = "Intake" Then
If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
'Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
'2/11/14
'lisa

    DoCmd.SetWarnings False
    strinfo = "Deed of Appointment sent to client via Intake Wizard"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
'lisa

'Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'With rstJnl
'.AddNew
'!FileNumber = FileNumber
'!JournalDate = Now
'!Who = GetFullName
'!Info = "Deed of Appointment sent to client via Intake Wizard"
'!Color = 1
'.Update
'End With
Else
MsgBox "Remember to manually enter a journal note when the document is sent"
End If
End If


End Sub
Private Sub SodProc(PrintTo As Integer)
Dim rstJnl As Recordset

If IsNull(Forms!foreclosuredetails![LPIDate]) Then
    MsgBox ("Cannot print without LPI date")
    Exit Sub
    Else
    MsgBox "Please add the Interest from/to dates once complete"
        
        If Forms![Case List]!ClientID = 404 Then 'Bogman
             DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt|Statement of Debt Bogman|" & PrintTo
        ElseIf Forms![Case List]!ClientID = 532 Then 'Selene
             DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt|Statement of Debt Selene|" & PrintTo
        ElseIf Forms![Case List]!ClientID = 523 Then 'Green Tree
             DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt|Statement of Debt GreenTree|" & PrintTo
               
        Else
             DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Statement of Debt|Statement of Debt|" & PrintTo
        End If
        
        If Forms![Case List]!ClientID = 446 Then
        Call DoReport("BOA Cover Sheet", PrintTo)
        ElseIf Forms![Case List]!ClientID = 444 Then
        Call DoReport("PHH Cover Sheet", PrintTo)
        ElseIf Forms![Case List]!ClientID = 385 Then
        Call DoReport("Nationstar Cover Sheet", PrintTo)
        ElseIf Forms![Case List]!ClientID = 567 And Me.State = "MD" Then
        Call DoReport("CHAM Cover SOD", PrintTo)
        End If
    End If
    If Forms!foreclosuredetails!WizardSource = "Intake" Then
    If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
    '2/11/14
'    Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
    'Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
    'lisa
    DoCmd.SetWarnings False
    strinfo = "Statement of Debt sent to client via Intake Wizard"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True
    
'    With rstJnl
'    .AddNew
'    !FileNumber = FileNumber
'    !JournalDate = Now
'    !Who = GetFullName
'    !Info = "Statement of Debt sent to client via Intake Wizard"
'    !Color = 1
'    .Update
'    End With
    Else
    MsgBox "Remember to manually enter a journal note when the document is sent"
    End If
    End If
End Sub

 Private Sub AFFMD7105(PrintTo As Integer)
    
    If Forms![Case List]!ClientID = 446 Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "45 Day Notice Affidavit BOA|45 Day Notice Affidavit BOA|" & PrintTo
        If Not Me.txtDesignatedAttorney = 3 Then Call DoReport("BOA Cover Sheet", PrintTo)
        
    ElseIf Forms![Case List]!ClientID = 6 Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "45 Day Notice Affidavit Wells |45 Day Notice Affidavit Wells|" & PrintTo
         
    ElseIf Forms![Case List]!ClientID = 556 Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "45 Day Notice Affidavit Wells |45 Day Notice Affidavit Wells|" & PrintTo
            
    ElseIf Forms![Case List]!ClientID = 567 Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "45 Day Notice Affidavit|45 Day Notice Affidavit|" & PrintTo
        If Forms![Case List]!ClientID = 567 And Me.State = "MD" Then Call DoReport("CHAM Cover AffMD7105", PrintTo)
 
            
    ElseIf Forms![Case List]!ClientID = 97 Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "45 Day Notice Affidavit Chase |45 Day Notice Affidavit Chase|" & PrintTo
    
    ElseIf Forms![Case List]!ClientID = 444 Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "45 Day Notice Affidavit|45 Day Notice Affidavit|" & PrintTo
        If Not Me.txtDesignatedAttorney = 3 Then Call DoReport("PHH Cover Sheet", PrintTo)
    
    ElseIf Forms![Case List]!ClientID = 531 Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "45 Day Notice Affidavit MDHC|45 Day Notice Affidavit MDHC|" & PrintTo
      
    ElseIf Forms![Case List]!ClientID = 404 Then
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "45 Day Notice Affidavit|45 Day Notice Affidavit Bogman|" & PrintTo
            
    Else
        DoCmd.OpenForm "PrintClientDoc", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "45 Day Notice Affidavit|45 Day Notice Affidavit|" & PrintTo
    End If
   
    If Forms!foreclosuredetails!WizardSource = "Intake" Then
        If MsgBox("Will the document be sent today?", vbYesNo) = vbYes Then
            'Set rstJnl = CurrentDb.OpenRecordset("select * from Journal", dbOpenDynaset, dbSeeChanges)
            '2/11/14
            'lisa
    DoCmd.SetWarnings False
    strinfo = "NOI Affidavit sent to client via Intake Wizard"
    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(Forms!ForeclosureDetails!FileNumber,Now,GetFullName(),'" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal
    DoCmd.SetWarnings True

'            Set rstJnl = CurrentDb.OpenRecordset("Journal", dbOpenDynaset, dbSeeChanges)
'                With rstJnl
'                    .AddNew
'                    !FileNumber = FileNumber
'                    !JournalDate = Now
'                    !Who = GetFullName
'                    !Info = "NOI Affidavit sent to client via Intake Wizard"
'                    !Color = 1
'                    .Update
'                End With
        Else
            MsgBox "Remember to manually enter a journal note when the document is sent"
        End If
    End If


End Sub

