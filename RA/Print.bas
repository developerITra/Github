Attribute VB_Name = "Print"
Option Compare Database
Option Explicit

Public WorkflowStates As String
Dim NotaryID As Long, NotaryMissingWarning As Boolean
Dim LabelSequence As Long, LabelSeries As String

Public Function TestDocument(ReportName As String) As Boolean
'
' Make sure that all the required data is present for a document/report.
'
Dim rsDoc As Recordset, rsFields As Recordset, rsTest As Recordset, ErrMsg As String, sql As String, QuerySQL As String, F As Field
On Error GoTo TestDocumentError

If DLookup("iValue", "DB", "Name='TestDocument'") <> 1 Then
    TestDocument = True
    Exit Function
End If

Set rsDoc = CurrentDb.OpenRecordset("SELECT * FROM DocumentList WHERE DocName='" & ReportName & "'", dbOpenSnapshot)

If IsLoadedF("ForeclosurePrint") Then
    If (Forms!Foreclosureprint!chDeedOfApp = True And Forms![Case List]!ClientID = 87) Then
    TestDocument = True  'Do nothing
    Exit Function
    End If
End If

If rsDoc.EOF Then
        'MsgBox "Cannot validate data for " & ReportName & ", not found in DocumentList", vbCritical
        'TestDocument = False
        TestDocument = True     ' until we do all the documents, don't complain!
Else
        On Error Resume Next
        CurrentDb.QueryDefs.Delete "tmpTestDocument"
        On Error GoTo TestDocumentError
        ' Get the report's recordsource.
        QuerySQL = Nz(rsDoc!RecordSource)
        If QuerySQL = "" Then       ' shouldn't happen, but exit cleanly
            TestDocument = True
            Exit Function
        End If
    
        ' If the recordsource doesn't begin with the word SELECT then its the name of a query.  Get the SQL from the query.
        If UCase$(Left$(QuerySQL, 6)) <> "SELECT" Then
            QuerySQL = CurrentDb.QueryDefs(QuerySQL).sql
        End If
    
        ' The record sources (and queries) for documents all reference the file number in the Case List form.
        ' Replace the reference to the form with the actual file number because Access won't reference the form from VBA.
        ' The exact expression varies some, so try each.
        QuerySQL = Replace(QuerySQL, "[Forms]![Case List]![FileNumber]", Forms![Case List]!FileNumber)
        QuerySQL = Replace(QuerySQL, "[Forms]![Case List]!FileNumber", Forms![Case List]!FileNumber)
        QuerySQL = Replace(QuerySQL, "Forms![Case List]![FileNumber]", Forms![Case List]!FileNumber)
        QuerySQL = Replace(QuerySQL, "Forms![Case List]!FileNumber", Forms![Case List]!FileNumber)
    
        ' Create a new query, just like the one used by the report, but with the actual file # inserted in the WHERE clause.
        CurrentDb.CreateQueryDef "tmpTestDocument", QuerySQL
        
        ' Build a query to return each expression that we want to test.  Use the query above as the data source.
        Set rsFields = CurrentDb.OpenRecordset("SELECT * FROM DocumentFields WHERE DocID=" & rsDoc!DocID, dbOpenSnapshot)
        sql = "SELECT "
        Do While Not rsFields.EOF
            sql = sql & rsFields!Expression & " AS F" & rsFields!FieldID & ", "
            rsFields.MoveNext
        Loop
        rsFields.Close
        sql = Left$(sql, Len(sql) - 2)  ' remove trailing comma
        sql = sql & " FROM tmpTestDocument"
            
        
        'Open the query with the test expressions, and see what we have...
        
    
       
             Set rsTest = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
             If rsTest.EOF Then
                 ErrMsg = "Critical data is missing such as the Client, Jurisdiction or Names."
             Else
             
                 For Each F In rsTest.Fields
            
                     If F.Value Then
                         ErrMsg = ErrMsg & DLookup("ErrorText", "DocumentFields", "FieldID=" & Mid$(F.Name, 2)) & vbNewLine
                     End If
                 Next
             End If
             If ErrMsg = "" Then
                 TestDocument = True
             Else
                 TestDocument = False
                 MsgBox "Document " & ReportName & " cannot be generated because:" & vbNewLine & ErrMsg, vbCritical
             End If
       
        
End If

rsDoc.Close
Exit Function

TestDocumentError:
    MsgBox "Document " & ReportName & " cannot be generated because of an unexpected error: " & Err.Description & vbNewLine & "Please submit a trouble ticket with this information and the file # you are working on.", vbCritical
    'Resume
    TestDocument = False
End Function

Public Function NotaryLocation() As String

If NotaryID = 0 Then NotaryID = GetNotaryID()
NotaryLocation = Nz(DLookup("NotaryLocation", "Staff", "ID=" & NotaryID))
End Function

Public Function NotaryName() As String

If NotaryID = 0 Then NotaryID = GetNotaryID()
NotaryName = Nz(DLookup("NotaryName", "Staff", "ID=" & NotaryID))
End Function

Public Function NotaryExpires() As String
If IsNull([Forms]![Foreclosureprint]![NotaryID]) Then
MsgBox (" You have not selected a Notary.")
Exit Function
End If


'Dim Info As Variant
'
'If NotaryID = 0 Then NotaryID = GetNotaryID()
'Info = DLookup("NotaryExpires", "Staff", "ID=" & NotaryID)
'If IsNull(Info) Then
'    NotaryExpires = ""
'Else
    NotaryExpires = Format$(DLookup("NotaryExpires", "Staff", "ID=" & [Forms]![Foreclosureprint]![NotaryID]), "mmmm d, yyyy")
'End If
End Function

Private Function GetNotaryID()
'Dim n As Variant
'
'If StaffID = 0 Then Call GetLoginName
'n = DLookup("NotaryToUse", "Staff", "ID=" & StaffID)
'If IsNull(n) Then
'    If Not NotaryMissingWarning Then
'        MsgBox "You have not selected a Notary.  Click Preferences on the main screen.", vbCritical
'        NotaryMissingWarning = True
'    End If
'    GetNotaryID = 0
'Else
    GetNotaryID = Forms!Foreclosureprint!NotaryID
'End If
End Function

Public Sub CheckPrintInfo(FileNumber As Long)
Dim p As Recordset

If IsNull(DLookup("FileNumber", "PrintInfo", "FileNumber=" & FileNumber)) Then
    Set p = CurrentDb.OpenRecordset("PrintInfo", dbOpenDynaset, dbSeeChanges)
    p.AddNew
    p("FileNumber") = FileNumber
    p.Update
    p.Close
End If
End Sub

Public Sub CheckPrintInfoMD(FileNumber As Long)
Dim p As Recordset

If IsNull(DLookup("FileNumber", "PrintInfoMD", "FileNumber=" & FileNumber)) Then
    Set p = CurrentDb.OpenRecordset("PrintInfoMD", dbOpenDynaset, dbSeeChanges)
    p.AddNew
    p("FileNumber") = FileNumber
    p.Update
    p.Close
End If
End Sub


Public Function OrderDate(Optional DaysOut As Integer) As String
Dim fixedpart As String

If DaysOut = 0 Then DaysOut = 15
fixedpart = "______ day of ________________, "
If Year(Date) <> Year(DateAdd("d", DaysOut, Date)) Then
    OrderDate = fixedpart & "20___"
Else
    OrderDate = fixedpart & Year(Date)
End If
End Function

Public Function DOTWord(DOT As Variant) As String
If DOT Then
    DOTWord = "Deed of Trust"
Else
    DOTWord = "Mortgage"
End If
End Function

Public Sub DoReport(ReportName As String, PrintTo As Integer, Optional FileName As String, Optional OpenArgs As String, Optional Filter As String)
Dim SaveDefaultPrinter As String, ErrCnt As Integer
Dim DefaultPrinterIndex As Integer
Dim strQueryName As String

' Make sure all required fields are complete
If Not TestDocument(ReportName) Then Exit Sub

    ' This is for Access Snapshot format
    'DoCmd.OutputTo acOutputReport, ReportName, "Snapshot Format", EMailPath & FileName & ".snp", False

Select Case PrintTo

    Case -3     ' Excel
    
      Select Case ReportName
      
        Case "Workflow Assignment not Received from Client"
            strQueryName = "rqryAssignmentNotReceived"
               
        Case "Workflow DOA Out"
            strQueryName = "rqryDOAout"
        Case "Workflow Files in Pending Status"
            strQueryName = "rqryFilespending"
        Case "Workflow Hearing Scheduled FC"
            strQueryName = "wkflHearingStatusUnionExceptionQ"
        
        Case "Workflow DOA not Recorded"
            strQueryName = "rqryDeedAppNotRecorded"
        Case "Workflow Docs Out"
            'strQueryName = "rqryDocsOutNew"
            strQueryName = "rqryDocsOutNew_Excel"

            
        Case "Workflow Deeds Not SEnt"
            strQueryName = "rqryDeedsNotSent"
        Case "Eviction Status"
            strQueryName = "rqryEviction"
        Case "qryEV_LPSDesktop"
            strQueryName = "qryEV_LPSdesktop"
        Case "Workflow BK FHLMC Active Cases"
            strQueryName = "rqryBKFHLMCActiveCases"
        Case "Workflow BK FHLMC Chpt 13"
            strQueryName = "rqryBKFHLMCChpt13"
        Case "Workflow BK FHLMC Def Cured Monitor"
            strQueryName = "rqryBKFHLMCDefCuredMonitoring"
        Case "Workflow HUD First Legal"
            strQueryName = "qrywkflHUDfirstlegal"
        Case "Workflow VA Appraisal"
            strQueryName = "qrywkflVAappraisal"
        Case "Workflow Bankruptcy"
            strQueryName = "rqryBankruptcy"
            
        Case "Workflow Docs Not Sent"
            strQueryName = "rqryDocsNotSent"
        Case "Workflow Docs Out"
            strQueryName = "rqryDocsOut"
        'Case "Workflow Docs Out"
          'strQueryName = "rqryDocsOut"
            
        Case "Workflow Fair Debt Title Order Needed"
          strQueryName = "rqryFairDebtNeedTitleOrdered"
          
        Case "Workflow FHLMC BK To be Filed"
            strQueryName = "rqryBKToBeFiled"
      
        Case "Workflow FNMA BK"
            strQueryName = "rqryFNMABKInventory"
        Case "Workflow FNMA Combined"
            strQueryName = "rqryFNMAFCInventory"
        Case "Workflow FNMA FC"
            strQueryName = "rqryFNMAFCInventory new"
        Case "Workflow FNMA Holds"
            strQueryName = "rqryFNMAHolds"
        Case "Workflow FNMA Missing Docs"
            strQueryName = "rqryFNMAMissingDocs"
        Case "Workflow FHLMC Open Files"
            strQueryName = "rqryFHLMCOpenFiles"
        Case "Workflow FNMA Postponements"
            strQueryName = "rqryFNMAPostponements"
            
        Case "Workflow Need To Invoice BK"
            strQueryName = "rqryNeedToInvoiceBankruptcyExcel"
        Case "Workflow Need To Invoice EV"
            strQueryName = "rqryNeedToInvoiceEvictionExcel"
        Case "Workflow Need To Invoice FC"
            strQueryName = "rqryNeedtoInvoiceFC"
        Case "Workflow Need To Invoice FC New"
            strQueryName = "rqryNeedtoInvoiceFCnewExcel"
        Case "Workflow Need To Invoice Rent"
            strQueryName = "rqryNeedToInvoiceRentCollectedExcel"
        Case "Workflow Need To Invoice TR"
            strQueryName = "rqryNeedtoInvoiceTRExcel"
        Case "Workflow Need To Invoice DIL"
            strQueryName = "rqryNeedToInvoiceDILExcel"
        Case "Workflow Need To Invoice Servicer Released"
            strQueryName = "rqryServicerReleaseExcel"
        Case "Workflow Need To Invoice Title"
            strQueryName = "rqryNeedtoInvoiceTitleExcel"
            
        Case "Workflow Judgment Entered Need Set Sale"
            strQueryName = "rptJudgmentEnteredNeedSetSale_excel"


        Case "Workflow Service Deadline DC"
            strQueryName = "rptwkflServiceDeadline_DC_Excel"

'added on 5/11/15

'Workflow Need To Invoice FCMonitor
Case "Workflow Need To Invoice FCMonitor"
            strQueryName = "qryNeedToInvoiceMonitor"

        Case "Workflow NOI"
            strQueryName = "rqryNOI"
   
        Case "Workflow Receivables"
            strQueryName = "rqryReceivables"
        Case "Workflow Receivables_FC"
            strQueryName = "rqry_Receivables_FC"
        Case "Workflow Receivables_BK"
            strQueryName = "rqry_Receivables_BK"
        Case "Workflow Receivables_EV"
            strQueryName = "rqry_Receivables_EV"
        Case "Workflow Receivables_OTH"
            strQueryName = "rqry_Receivables_OTH"
       
        Case "Workflow Waiting for Bills"
            strQueryName = "rqryAttributeBills"
            
        Case "PropertyRegistrationMD"
            strQueryName = "qryPropertyRegistrationMD"
            
        Case "Workflow 362 to be Filed"
            strQueryName = "rqry362ToBeFiled"
            
        Case "Workflow POC to be Filed"
            strQueryName = "rqryPOCToBeFiled"
        
        Case "qryqueuedocketingwaiting"
         strQueryName = "qryqueuedocketingwaitinglst"
         
        Case "Workflow Title To be Reviewed"
            strQueryName = "rqryTitleToBeReviewed"
        Case "Workflow Title Claims not sent"
            strQueryName = "rqryTitleClaimsNotSent"
        Case "Workflow Title Claims Out"
            strQueryName = "rqryTitleClaimsOut"
        Case "WorkFlow DIL All"
            strQueryName = "rqryDILAll"
            
        Case "Workflow Cancel Service Due to Disposition"
            strQueryName = "qryCancelServiceDuetoDispositionDC_Excel"
        'Volume Reports
        
        Case "WizardRSICompleted", PrintTo
            strQueryName = "rqryWizardRSICompleted"
        Case "WizardRSIICompleted", PrintTo
            strQueryName = "qryWizardRSIICompleted"
        Case "WizardFairDebtCompleted", PrintTo
            strQueryName = "qryWizardFairDebtCompleted"
        Case "WizardDemandCompleted", PrintTo
            strQueryName = "qryWizardDemandCompleted"
        Case "WizardDocketingCompleted", PrintTo
            strQueryName = "qryWizardDocketingCompleted"
        Case "WizardServiceCompleted", PrintTo
            strQueryName = "qryWizardServiceComplete"
        Case "WizardborrowerservedCompleted", PrintTo
            strQueryName = "qryWizardBorrowerServedComplete"
        Case "WizardServiceMailedCompleted", PrintTo
            strQueryName = "qryWizardServiceMailedComplete"
        Case "WizardFLMACompleted", PrintTo
            strQueryName = "qryWizardFLMAComplete"
        Case "WizardSaleSettingCompleted", PrintTo
            strQueryName = "qryWizardSaleSettingCompleted"
        Case "WizardSCRAFCCompleted", PrintTo
            strQueryName = "qryWizardSCRAFCCompleted"
        Case "WizardRestartCompleted", PrintTo
            strQueryName = "qryWizardRestartCompleted"
        Case "WizardIntakeCompleted", PrintTo
            strQueryName = "qryWizardIntakeCompleted"
        Case "WizardSAICompleted", PrintTo
            strQueryName = "qryWizardSAIComplete"
        Case "WizardVASaleSettingCompleted", PrintTo
            strQueryName = "qryWizardVASaleSettingCompleted"
        Case "LexisNexisCompleted", PrintTo
            strQueryName = "qryLexisNexisCompleted"
        Case "WizardLNNCompleted", PrintTo
            strQueryName = "qryWizardLNNComplete"
        Case "WizardTiteOrderCompleted", PrintTo
            strQueryName = "qryWizardTiteOrderCompleted"
        Case "WizardTitleOutCompleted", PrintTo
            strQueryName = "qryWizardTitleOutCompleted"
        Case "WizardTitleReviewCompleted", PrintTo
            strQueryName = "qryWizardTitleReviewCompleted"
        Case "WizardNOICompleted", PrintTo
            strQueryName = "qryWizardNOICompleted"
        Case "WizardNOIUpload", PrintTo
            strQueryName = "qryWizardNOIUpload"
        Case "WizardNOIUploadMissing", PrintTo
            strQueryName = "qryWizardNOIUploadMissing"
        Case "VolumeLitigationBilling", PrintTo
            strQueryName = "VolumeLitigationBillingExel"
        Case "TrackingLitigationBilling", PrintTo
            strQueryName = "TrackingLitigationBillingExel"
        Case "TrackingPayOffSent", PrintTo
            strQueryName = "TrackingPayoffSentExel"
        Case "TrackingFcDispositio", PrintTo
            strQueryName = "TrackingFCDispositionExel"
        Case "TrackingDebtVerified", PrintTo
            strQueryName = "TrackingDebtVerifiedExel"
        Case "TrackingReinstatementSent", PrintTo
             strQueryName = "TrackingReinstatementSentExel"
            
        Case "Workflow Receivables_PSAdvanced"
            strQueryName = "wkflReceivablesPSAdvancedExcel"
         'added on 4_28_15
         DoCmd.SetWarnings False
         DoCmd.OpenQuery ("MK_NoticeofAppearance")
         DoCmd.SetWarnings True
        Case "Workflow Notice of Appearance"
            strQueryName = "Notice of Appearance_Excel"
        '----
        Case "Workflow Cases to be Closed", PrintTo
            strQueryName = "rqryCasesToBeClosed"
            
        'added on 7_6_15
        
        Case "workflowInvoiceAll", PrintTo
            strQueryName = "qrywkflInvoiceAll"
        
        Case "workflowInvoicePaid", PrintTo
            strQueryName = "qrywkflInvoicePaid"
        'added 7/7/15
        
        Case "Workflow Loss Mediation_DC", PrintTo
            strQueryName = "qryDCMediiationExcel"
            
        'added on 9/17/15
        Case "Workflow Hearing Scheduled DC", PrintTo
            strQueryName = "qrywkflDCAllHearing"
        
        Case Else
        
          MsgBox "Report " & ReportName & " must be set up to be exported to Excel.", vbExclamation, "Excel Export"
          Exit Sub
      End Select
    
      Call OutputExcel(ReportName, strQueryName)
    Case -2     ' PDF
        ' New way, skip the email stuff, just go to the regular PDF printer
        DefaultPrinterIndex = SetPrinterName("Adobe PDF")
        DoCmd.OpenReport ReportName, acViewNormal, , Filter, , OpenArgs
        GoTo RestorePrinter
        
        ' Old way, drop the PDF into a folder for automatic insertion into email
        If FileName = "" Then FileName = ReportName
        DefaultPrinterIndex = SetPrinterName("Acrobat Distiller")
        DoCmd.OpenReport ReportName, acViewNormal, , , , OpenArgs
        '
        ' Need to wait for the distiller to finish.  If we can rename the output file, then it must be finshed.
        '
        Wait 2          ' give it a chance
        On Error GoTo WaitErr
        Name EMailPath & ReportName & ".pdf" As EMailPath & FileName & ".pdf"
        On Error GoTo 0
RestorePrinter:
        SetPrinterIndex (DefaultPrinterIndex)


Case -5
    Dim StrDraft As String
    Dim FileNumber As Long
    Dim DocID As Integer
    FileNumber = Forms!foreclosuredetails!FileNumber
    
    
    If Forms!Foreclosureprint!chDeedOfApp Then
    DocID = 1577
    StrDraft = "Draft SOT " & Format$(Now(), "yyyymmdd hhnnss") & ".pdf"
    Call SaveDraft(ReportName, FileNumber, DocID, StrDraft)
    
    End If
    
    If Forms!Foreclosureprint!chSOD Then
    DocID = 1576
    StrDraft = "Draft SOD " & Format$(Now(), "yyyymmdd hhnnss") & ".pdf"
    Call SaveDraft(ReportName, FileNumber, DocID, StrDraft)
    End If
    
    
    If Forms!Foreclosureprint!chSOD2 Then
    DocID = 1576
    StrDraft = "Draft SOD " & Format$(Now(), "yyyymmdd hhnnss") & ".pdf"
    Call SaveDraft(ReportName, FileNumber, DocID, StrDraft)
    End If
    
     
    If Forms!Foreclosureprint!chComplianceAffidavit Then
    DocID = 1588
    StrDraft = "Draft Combo AFF Compliance " & Format$(Now(), "yyyymmdd hhnnss") & ".pdf"
    Call SaveDraft(ReportName, FileNumber, DocID, StrDraft)
    End If
    
    
    
    If Forms!Foreclosureprint!chLostNoteAffidavit Then
    DocID = 1589
    StrDraft = "Draft Lost Note Affidavit " & Format$(Now(), "yyyymmdd hhnnss") & ".pdf"
    Call SaveDraft(ReportName, FileNumber, DocID, StrDraft)
    End If
    
    
    If Forms!Foreclosureprint!chNoteOwnership Then
    DocID = 1580
    StrDraft = "Draft ANO " & Format$(Now(), "yyyymmdd hhnnss") & ".pdf"
    Call SaveDraft(ReportName, FileNumber, DocID, StrDraft)
    End If
    
    
    If Forms!Foreclosureprint!chAffMD7105 Then
    DocID = 1584
    StrDraft = "Draft NOI AFF " & Format$(Now(), "yyyymmdd hhnnss") & ".pdf"
    Call SaveDraft(ReportName, FileNumber, DocID, StrDraft)
    End If


    If Forms!Foreclosureprint!chLossMitPrelim Then
    DocID = 1581
    StrDraft = "Draft LMA(s) " & Format$(Now(), "yyyymmdd hhnnss") & ".pdf"
    Call SaveDraft(ReportName, FileNumber, DocID, StrDraft)
    End If

    
    If Forms!Foreclosureprint!chLossMitFinal Then
    DocID = 1581
    StrDraft = "Draft LMA(s) " & Format$(Now(), "yyyymmdd hhnnss") & ".pdf"
    Call SaveDraft(ReportName, FileNumber, DocID, StrDraft)
    End If
    

    Case -1     ' Word or RTF
    

        Select Case ReportName
        
            Case "AssignmentSOT VA Specialized"
                Call Doc_AssignmentSOTVASpecialized(EMailStatus <> 1)
       
            Case "Assignment Recording Cover Letter"
                Call Doc_AssignmentRecordingCoverLetter(EMailStatus <> 1)
            Case "Affidavit of Lien Instrument NationStar"
                Call DOC_AffidavitofLienInstrumentNationStar(EMailStatus <> 1)
            Case "Motion to Release Funds"
                Call Doc_MotionToReleaseFunds(EMailStatus <> 1)
             
            Case "Statement of Debt Monitor"
                Call Doc_StatementOfDebtMonitor(EMailStatus <> 1)
            
            Case "Order Granting Intervention"
                Call Doc_OrderGrantingIntervention(EMailStatus <> 1)
            
            Case "Motion to Intervene"
                Call Doc_MotionToIntervene(EMailStatus <> 1)
             
            Case "Deed of Appointment MD_Nationstar"                   ' SOT doc, Mei 10-6-15
                Call DOC_DeedofAppointmentNationstarMD(EMailStatus <> 1)
                
            Case "Deed of Appointment Select"
                Call Doc_DeedApp(EMailStatus <> 1)
            
            Case "Combined Affidavit of Compliance Nationstar"        'Mei 10/7/15
                Call doc_CombinedAffidavitofComplianceNationStar(EMailStatus <> 1)
                
            Case "Combined Affidavit of Compliance Select"
                Call Doc_ComplianceAffidavit(EMailStatus <> 1)
             
            Case "Line"
                Call Doc_Line(EMailStatus <> 1)
            Case "Deed of Appointment GreenTreeVA"
                Call Doc_DeedApp(EMailStatus <> 1)
            Case "Deed of Appointment GreenTree"
                Call Doc_DeedApp(EMailStatus <> 1)
                
            Case "Deed of APpointment Selene"
                Call Doc_DeedApp(EMailStatus <> 1)
             Case "Deed of APpointment SeleneVA"
                Call Doc_DeedApp(EMailStatus <> 1)
             Case "Deed of APpointment SeleneDC"
                Call Doc_DeedApp(EMailStatus <> 1)
                
            Case "Special Warranty Deed VA"
                Call Doc_SpecialWarranty(EMailStatus <> 1)
            
            Case "Special Warranty Deed MD"
                Call Doc_SpecialWarranty(EMailStatus <> 1)
            
            Case "Quit Claim Deed VA"
                Call Doc_QuitClaimDeed(EMailStatus <> 1)
            
            Case "Quit Claim Deed MD"
                Call Doc_QuitClaimDeed(EMailStatus <> 1)
            Case "Courtesy Eviction Letter"
                Call Doc_CourtesyEvictionLetter(EMailStatus <> 1)
            
            Case "Clerk Cover Letter"
                Call Doc_ClerkCoverLTR(EMailStatus <> 1)
            
            Case "Affidavit Cover Letter"
                Call Doc_AffCoverSheet(EMailStatus <> 1)
            
            Case "Motion for Relief BOA CH13"
                Call Doc_BKBOAMotionForReliefCh13(EMailStatus <> 1)
            
            Case "Motion for Relief BOA"
                Call Doc_BKBOAMotionForRelief(EMailStatus <> 1)
                
            Case "Motion for Relief Chase"              'Mei 10/11/15
                Call Doc_BKChaseMotionForRelief(EMailStatus <> 1)
                
            Case "MD Lease Termination"
                Call Doc_MDLeaseTermination(EMailStatus <> 1)
            
            Case "MD Expired Lease"
                Call Doc_MDExpiredLease(EMailStatus <> 1)
              
            Case "Eviction Affidavit 14-102b"
                Call Doc_Eviction_MD14102b(EMailStatus <> 1)
            
            Case "Eviction Print VA Tenant NTQ 5/3"
                Call Doc_Eviction_VA_Tenant_NTQ_53(EMailStatus <> 1)
            Case "Eviction Print VA Owner NTQ 5/3"
                Call Doc_Eviction_VA_Owner_NTQ_53(EMailStatus <> 1)
            
            
            Case "Eviction Print MD NTQ 5/3"
                Call Doc_Eviction_MD_NTQ_53(EMailStatus <> 1)
            Case "Eviction Print MD Owner NTQ SPS"
                Call Doc_EvictionOwner_NTQ_SPS(EMailStatus <> 1)
            Case "Eviction Print MD Tenant NTQ SPS"
                Call Doc_EvictionTenant_NTQ_SPS(EMailStatus <> 1)
            
            Case "Eviction Print PTFA NTQ 5/3"
                Call Doc_Eviction_PTFA_NTQ_53(EMailStatus <> 1)
            Case "Eviction Print PTFA Owner NTQ SPS"
                Call Doc_EvictionOwner_NTQ_SPS_PTFA(EMailStatus <> 1)
            Case "Eviction Print PTFA Tenant NTQ SPS"
                Call Doc_EvictionTenant_NTQ_SPS_PTFA(EMailStatus <> 1)
            
            Case "Deed of Appointment SPLS VA"
                Call Doc_DeedOfAppointmentSPLS(EMailStatus <> 1)
            Case "Deed of Appointment SPLS MD"
                Call Doc_DeedOfAppointmentSPLS(EMailStatus <> 1)
            
            Case "Loss Mitigation Preliminary SPLS"
                Call Doc_LossMitigationPrelimSPLS(EMailStatus <> 1)
            Case "Statement of Debt with Figures SPLS"
                Call Doc_StatementOfDebtWithFiguresSPLS(EMailStatus <> 1)
                
            Case "GreenTree Bailee Letter"
                Call Doc_BaileeLTR(EMailStatus <> 1)
            
            Case "DC Order Granting Default - Wells"
                Call Doc_DCOrderGrantingDefaultWells(EMailStatus <> 1)
            Case "BaltCity Intake Sheet"
                Call Doc_BaltCityIntakeSheet(EMailStatus <> 1)
            Case "DC Trustee Affidavit"
                Call Doc_DCTrusteeAffidavit(EMailStatus <> 1)
            Case "DC Accounting"
                Call Doc_DCAccounting(EMailStatus <> 1)
            Case "ConsentTerminating13VAWD"
                Call Doc_ConsentTerminatingWD13(EMailStatus <> 1)
            Case "Notice of Lis Pendens"
                Call Doc_LisPendens(EMailStatus <> 1)
            Case "Return Doc Cover LTR"
                Call Doc_ReturnDocCoverLTR(EMailStatus <> 1)
                
            Case "Line Notifying of Audit Filing"
                Call Doc_AuditFiling(EMailStatus <> 1)
            
            Case "45 Day notice Affidavit Bogman"
                Call Doc_MD7_105Affidavit(EMailStatus <> 1)
                        
            Case "Statement of Debt with Figures Bogman"
                Call Doc_StatementOfDebtFigures(EMailStatus <> 1)
            
            Case "Statement of Debt with Figures Dove"
                Call Doc_StatementOfDebtFigures(EMailStatus <> 1)
                
            Case "Statement of Debt with Figures Selene"
                Call Doc_StatementOfDebtFigures(EMailStatus <> 1)
                    
            Case "Statement of Debt with Figures GreenTree"
                Call Doc_StatementOfDebtFigures(EMailStatus <> 1)
            Case "Statement of Debt Bogman"
                Call Doc_StatementOfDebt(EMailStatus <> 1)
            
            Case "Statement of Debt GreenTree"
                 Call Doc_StatementOfDebt(EMailStatus <> 1)
            
            Case "Statement of Debt Selene"
                 Call Doc_StatementOfDebt(EMailStatus <> 1)
                        
            Case "Ownership Affidavit Bogman"
                Call Doc_NoteOwnershipAffidavit(EMailStatus <> 1)
            
            Case "Ownership Affidavit Selene"
                Call Doc_NoteOwnershipAffidavit(EMailStatus <> 1)
            
            Case "Ownership Affidavit GreenTree"
                Call Doc_NoteOwnershipAffidavit(EMailStatus <> 1)
                        
            Case "MilitaryAffidavitMD SpLS"
                Call Doc_MilitaryAffidavitMD(EMailStatus <> 1, False)
            Case "MilitaryAffidavitMD Bogman"
                Call Doc_MilitaryAffidavitMD(EMailStatus <> 1, False)
            Case "Military Affidavit NoSSN MD Bogman"
                Call Doc_MilitaryAffidavitNoSSNMD(EMailStatus <> 1)
                        
            Case "MilitaryAffidavitActive Bogman"
                Call Doc_MilitaryAffidavitActive(EMailStatus <> 1, "MD")
            Case "Deed of Appointment Dove VA"
                Call Doc_DeedApp(EMailStatus <> 1)
            Case "Deed of Appointment Dove"
                Call Doc_DeedApp(EMailStatus <> 1)
            Case "Deed of Appointment Bogman"
                Call Doc_DeedApp(EMailStatus <> 1)
            Case "Contract for Sale"
                Call Doc_ContractForSale(EMailStatus <> 1)
'            Case "Hud Deed VA"
'                Call Doc_HUDDEEDVA(EMailStatus <> 1)
            
            Case "Read at Sale"
                Call Doc_ReadAtSale(EMailStatus <> 1)
            
            Case "PayoffJPRIVA"
                Call Doc_PayoffJPRI(EMailStatus <> 1)
            Case "PayoffJPRI"
                Call Doc_PayoffJPRI(EMailStatus <> 1)
            Case "MD Land Instruments"
                Call Doc_LandInstruments(EMailStatus <> 1)
            Case "Ad DC"
                Call Doc_Ad_DC(EMailStatus <> 1)
            Case "Ad MD"
                Call Doc_Ad_MD(EMailStatus <> 1)
            Case "Ad VA"
                Call Doc_Ad_VA(EMailStatus <> 1)
' in progress -- can't find the template file
'            Case "ConsentModifying13"
'                Call Doc_ConsentModifying13(EMailStatus <> 1)
            Case "ConsentModifying11VAED"
                Call Doc_ConsentModifyingAll(EMailStatus <> 1)
            Case "ConsentModifying13VAED"
                Call Doc_ConsentModifyingAll(EMailStatus <> 1)
            Case "ConsentModifying7VAED"
                 Call Doc_ConsentModifyingAll(EMailStatus <> 1)
                 
            Case "ConsentModifying11"
                Call Doc_ConsentModifyingAll(EMailStatus <> 1)
            Case "ConsentModifying13"
                Call Doc_ConsentModifyingAll(EMailStatus <> 1)
            Case "ConsentModifying7"
                 Call Doc_ConsentModifyingAll(EMailStatus <> 1)
            Case "Debt"
                 Call Doc_DebtBK(EMailStatus <> 1)
                            
            Case "Deed of Appointment", "Deed of Appointment Kondaur", "Deed of Appointment Saxon", "Deed of Appointment PNC"
                Call Doc_DeedApp(EMailStatus <> 1)
                
            Case "Deed of appointment VA"
                Call Doc_DeedApp(EMailStatus <> 1)
                
            Case "Deed of Appointment Chase"
                Call Doc_DeedOfAppointmentChase(EMailStatus <> 1)
            
            Case "Statement of Debt with Figures Ocwen"
                Call Doc_StatementOfDebtWithFiguresOcwen(EMailStatus <> 1)
                           
            Case "Deed of Appointment ChaseVA"
                Call Doc_DeedOfAppointmentChase(EMailStatus <> 1)
            Case "Deed of Appointment MDHC"
                Call Doc_DeedofAppointmentMDHC(EMailStatus <> 1)
                
            Case "Deed of Appointment M&T"
                Call Doc_DeedofAppointmentMT(EMailStatus <> 1)
                
            Case "Mediation Court Notes Wells"
                Call Doc_MediationCourtNotesWells(EMailStatus <> 1)
                
            Case "CHAM Cover Sheet"
                Call Doc_CHAMCoverSheet(EMailStatus <> 1)
            
            Case "CHAM Cover AffMD7105"
                Call Doc_CHAMCoverSheetAFF(EMailStatus <> 1)
             Case "CHAM Cover DeedofApp"
                Call Doc_CHAMCoverSheetDeed(EMailStatus <> 1)
             Case "CHAM Cover LossMitFinal"
                Call Doc_CHAMCoverSheetLossFinal(EMailStatus <> 1)
             Case "CHAM Cover LossMitPrelim"
                Call Doc_CHAMCoverSheetLossPrelim(EMailStatus <> 1)
             Case "CHAM Cover MilitaryAffidavit"
                Call Doc_CHAMCoverSheetMA(EMailStatus <> 1)
             Case "CHAM Cover MilitaryAffidavitActive"
                Call Doc_CHAMCoverSheetMAActive(EMailStatus <> 1)
             Case "CHAM Cover MilitaryAffidavitNoSSn"
                Call Doc_CHAMCoverSheetMaNoSSN(EMailStatus <> 1)
             Case "CHAM Cover Noteownership"
                Call Doc_CHAMCoverSheetNote(EMailStatus <> 1)
             Case "CHAM Cover SOD"
                Call Doc_CHAMCoverSheetSOD(EMailStatus <> 1)
             Case "CHAM Cover SOD2"
                Call Doc_CHAMCoverSheetSOD2(EMailStatus <> 1)
            Case "Deed of appointment Accomack"
                Call Doc_DeedApp(EMailStatus <> 1)
            Case "Deed of Appointment Wells"
                Call Doc_DeedofAppointmentWells(EMailStatus <> 1)
            
            '#1226 MC 10/15/2014
            'Case "Ownership Affidavit Chase Anne"
            '    Call Doc_OwnershipAffidavitChase(EMailStatus <> 1)
            '/#1226
            Case "Deed of Appointment WellsVA"
                Call Doc_DeedofAppointmentWellsVA(EMailStatus <> 1)
              
            Case "Deed of Appointment WellsVA Select"
                Call Doc_DeedofAppointmentWellsVASelect(EMailStatus <> 1)
                
            Case "Military Affidavit"
                Call Doc_MilitaryAffidavit(EMailStatus <> 1)
                
             Case "Military Affidavit NoSSN"
                Call Doc_MilitaryAffidavitNoSSN(EMailStatus <> 1)
            
            Case "Military Affidavit NoSSN MD"
                Call Doc_MilitaryAffidavitNoSSNMD(EMailStatus <> 1)
            
            Case "MilitaryAffidavitDC"
                Call Doc_MilitaryAffidavitDC(EMailStatus <> 1)
                
            Case "MilitaryAffidavitActive"
                Call Doc_MilitaryAffidavitActive(EMailStatus <> 1, "")
                
            Case "MilitaryAffidavitActiveMD"
                Call Doc_MilitaryAffidavitActive(EMailStatus <> 1, "MD")
            Case "MilitaryAffidavitMD"
                Call Doc_MilitaryAffidavitMD(EMailStatus <> 1, False)
            Case "Order Granting Relief"
                Call Doc_OrderGrantingRelief(EMailStatus <> 1)
                
            Case "Final Order Terminating"
                Call Doc_FinalOrderTerminating(EMailStatus <> 1)
                
            
            Case "Order Granting Relief VANew"
                Call Doc_ConsentGrantRelief(True)
                
            Case "Ownership Affidavit BOA"
                Call Doc_OwnershipAffidavitBOA(EMailStatus <> 1)
                
             Case "Ownership Affidavit MDHC"
                Call Doc_OwnershipAffidavitMDHC(EMailStatus <> 1)
            
            Case "Ownership Affidavit Ocwen"
                Call Doc_OwnershipAffidavitOcwen(EMailStatus <> 1)
            
            Case "45 Day Notice Affidavit BOA"
                If Forms!Foreclosureprint!txtDesignatedAttorney = 3 Then
                    Call Doc_MD7_105Affidavit(EMailStatus <> 1)
                Else
                    Call Doc_45DayNoteOwnershipAffidavitBOA(EMailStatus <> 1)
                End If
            
            Case "45 Day Notice Affidavit Wells"
                If Forms!Foreclosureprint!txtDesignatedAttorney = 3 Then
                    Call Doc_MD7_105Affidavit(EMailStatus <> 1)
                Else
                    Call Doc_NOIAffidavitWells(EMailStatus <> 1)
                End If
                
            Case "Ownership Affidavit Wells"
                Call Doc_OwnershipAffidavitWells(EMailStatus <> 1)
            
            Case "BOA Cover Sheet Assignment"
                Call Doc_BOACoverSheetAssignment(EMailStatus <> 1)
                
            Case "BOA Cover Sheet"
                Call Doc_BOACoverSheet(EMailStatus <> 1)
                
            Case "PHH Cover Sheet"
                Call Doc_BOACoverSheet(EMailStatus <> 1)
                
                
            
            Case "Statement of Debt", "Statement of Debt Cenlar"
                Call Doc_StatementOfDebt(EMailStatus <> 1)
            
            Case "Statement of Debt with Figures", "Statement of Debt with Figures Cenlar"
                Call Doc_StatementOfDebtFigures(EMailStatus <> 1)
                
            Case "Statement of Debt with Figures Wells", "Statement of Debt with Figures Cenlar"
                Call Doc_StatementOfDebtWithFiguresWells(EMailStatus <> 1)
                
           Case "Statement of Debt with figures JP"
                Call Doc_StatementOfDebtWithFiguresJP(EMailStatus <> 1)

                
            Case "Statement of Debt with Figures BOA"
                Call Doc_StatementOfDebtWithFiguresBOA(EMailStatus <> 1)
                
            Case "Statement of Debt with Figures MDCDMT"
                 Call Doc_StatementOfDebtFiguresMDCDMT(EMailStatus <> 1)
            
            Case "Loss Mitigation Final BOA"
                Call Doc_LossMitigationFinal(EMailStatus <> 1)
                 
             Case "Loss Mitigation Final SPLS"
                Call Doc_LossMitigationFinal(EMailStatus <> 1)
                              
            Case "Loss Mitigation Preliminary BOA"
                 Call Doc_LossMitigationPre(EMailStatus <> 1)

            Case "Ownership Affidavit Chase"
                Call Doc_OwnershipAffidavitChase(EMailStatus <> 1)
                               
            Case "Loss Mitigation Preliminary MDHCD"
                Call Doc_LossMitigationPrelimMDHCP(EMailStatus <> 1)
            
            Case "Loss Mitigation Preliminary MT"
                Call Doc_LossMitigationPrelimMDHCP(EMailStatus <> 1)
            
            Case "Loss Mitigation Final MDHCD"
                Call Doc_LossMitigationFinalMDHCP(EMailStatus <> 1)
            Case "Loss Mitigation Preliminary Nation star"
                Call Doc_LossMitigationPrelimNationStar(EMailStatus <> 1)
                
            Case "Loss Mitigation Final MT"
                Call Doc_LossMitigationFinalMDHCP(EMailStatus <> 1)
            Case "Loss Mitigation Preliminary"
                Call Doc_LossMitigationPre(EMailStatus <> 1)
                
            Case "Loss Mitigation Preliminary Wells"
                Call Doc_LossMitigationPre(EMailStatus <> 1)
                
            Case "Loss Mitigation Preliminary Chase"
                Call Doc_LossMitigationPre(EMailStatus <> 1)
            
            Case "Loss Mitigation Preliminary Nation star"
                Call Doc_LossMitigationPreliminaryNationStar(EMailStatus <> 1)
            
            Case "Loss Mitigation Final Nation Star"
                Call Doc_LossMitigationFinalNationStarMD(EMailStatus <> 1)
            
            Case "Loss Mitigation Final"
                Call Doc_LossMitigationFinal(EMailStatus <> 1)
                
            Case "Loss Mitigation Final Wells"
                Call Doc_LossMitigationFinal(EMailStatus <> 1)
            Case "Loss Mitigation Final PNC"
                Call Doc_LossMitigationFinalPNC(EMailStatus <> 1)
            Case "Loss Mitigation Preliminary PNC"
                Call Doc_LossMitigationPrelimPNC(EMailStatus <> 1)
            
            Case "Loss Mitigation Final Select"
                Call Doc_LossMitigationFinalSPS(EMailStatus <> 1)
            Case "Loss Mitigation Preliminary Select"
                Call Doc_LossMitigationPrelimSPS(EMailStatus <> 1)
           
                
            Case "Lost Note Affidavit"
                Call Doc_LostNoteAffidavit(EMailStatus <> 1)
            
            Case "Lost Note Affidavit GreenTree"
                Call Doc_LostNoteAffidavit(EMailStatus <> 1)
                        
            Case "Lost Note Affidavit Selene"
                Call Doc_LostNoteAffidavit(EMailStatus <> 1)
            
            Case "Ownership Affidavit"
                Call Doc_NoteOwnershipAffidavit(EMailStatus <> 1)
            Case "Ownership Affidavit SPLS"
                Call Doc_NoteOwnershipAffidavit(EMailStatus <> 1)
                            
            Case "45 Day Notice Affidavit"
                Call Doc_MD7_105Affidavit(EMailStatus <> 1)
            
            Case "45 Day Notice Affidavit MDHC"
                Call Doc_MD7_105Affidavit(EMailStatus <> 1)
            
            Case "Assignment"
                Call Doc_Assignment(EMailStatus <> 1)
            Case "Assignment VA"
                Call Doc_Assignment(EMailStatus <> 1)
            Case "DIL Judgment Affidavit"
                Call Doc_DILJudgmentAffidavit(EMailStatus <> 1)
            Case "DIL Letter to Borrower"
                Call Doc_DILBorrowerLetter(EMailStatus <> 1)
            Case "DIL"
                Call Doc_DIL(EMailStatus <> 1)
            Case "DIL Certificate"
                Call Doc_DILCertificate(EMailStatus <> 1)
            
            Case "Deed of Appointment Ocwen"
                Call Doc_DeedofAppointmentOcwen(EMailStatus <> 1)
                    
            Case "Nationstar Cover Sheet"
                Call Doc_NationstarCoverSheet(EMailStatus <> 1)
            
            Case "Deed Recording Cover MD"
                Call Doc_RecordingDeedCoverMD(EMailStatus <> 1)
            Case "Eviction Notice Baltimore City"
                Call Doc_EVNoticeBalti(EMailStatus <> 1)
            Case "Statement of Debt with Figures CHAM"
                Call Doc_StatementOfDebtFiguresCham(EMailStatus <> 1)
            Case "Dismiss Case"
                Call Doc_DismissCase(EMailStatus <> 1)
            'Doc_DeedofAppointmentNationStarVA
            Case "Deed of Appointment VA_Nationstar"
                Call Doc_DeedofAppointmentNationStarVA(EMailStatus <> 1)
            Case Else ' no Word document available, convert via RTF
                DoCmd.OutputTo acOutputReport, ReportName, acFormatRTF, , True
        End Select

    Case 55     ' Workflow forms
        DoCmd.OpenForm ReportName

    Case Else
        If EMailStatus = 1 Then
            MsgBox "You must select Adobe Acrobat (PDF) or Word formats for EMail", vbCritical
            Exit Sub
        End If
       DoCmd.OpenReport ReportName, PrintTo, , Filter, , OpenArgs
End Select
Exit Sub

WaitErr:
    ErrCnt = ErrCnt + 1
    If ErrCnt > 60 Then     ' more than 2 minutes?
        MsgBox "Conversion to Adobe Acrobat is taking too long for " & ReportName & vbNewLine & "Please make a note of this error.", vbInformation
        Resume RestorePrinter
    End If
    Wait 2      ' give it some more time
    Resume      ' and try again

End Sub

Public Sub SaveDraft(ReportName As String, FileNumber As Long, DocID As Integer, StrDraft As String)
       DoCmd.OutputTo acOutputReport, ReportName, acFormatPDF, DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & StrDraft
        StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & StrDraft
    
        DoCmd.SetWarnings False
        Dim strSQLValues As String: strSQLValues = ""
        strSQL = ""

        strSQLValues = FileNumber & "," & DocID & ",'" & "" & "'," & GetStaffID() & ",#" & Now() & "#,'" & Replace(StrDraft, "'", "''") & "','" & Replace(StrDraft, "'", "''") & "'"
        'Debug.Print strSQLValues
        strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
        'Debug.Print strSQL
        DoCmd.RunSQL (strSQL)
        DoCmd.SetWarnings True
End Sub


Public Function SetPrinterName(PrinterName As String) As Integer
'
' Set the default printer.  Return the previous default printer index.
'
Dim i As Integer

For i = 0 To Application.Printers.Count - 1
    If Application.Printer.DeviceName = Application.Printers(i).DeviceName Then
        SetPrinterName = i
        Exit For
    End If
Next i
If PrinterName = "Acrobat Distiller" Then
    For i = 0 To Application.Printers.Count - 1
        If Left$(Application.Printers(i).DeviceName, 17) = "Acrobat Distiller" Then
            Application.Printer = Application.Printers(i)
            Exit Function
        End If
    Next i
Else
    For i = 0 To Application.Printers.Count - 1
        If Application.Printers(i).DeviceName = PrinterName Then
            Application.Printer = Application.Printers(i)
            Exit Function
        End If
    Next i
End If
MsgBox "Printer " & PrinterName & " is not available", vbExclamation
End Function

Public Sub SetPrinterIndex(PrinterIndex As Integer)
Application.Printer = Application.Printers(PrinterIndex)
End Sub

Public Function Wait(Seconds As Integer, Optional DispHrglass As Boolean)
Dim DelayEnd As Double

DoCmd.Hourglass DispHrglass
DelayEnd = DateAdd("s", Seconds, Now)
While DateDiff("s", Now, DelayEnd) > 0
    DoEvents
Wend
DoCmd.Hourglass False
End Function

Public Sub FirmMargin(rpt As Report, FileNum As Long, Optional Misc As Integer, Optional ProName As String, Optional PropertyAddress As String, Optional RemoveProj As Boolean, Optional APTNum As String)
Dim y1 As Single, y2 As Single

Dim i As Integer
Dim PropAddress() As String
Dim strLength As String
Dim FullAddress As String

Const BIGFONT = 6
Const SMALLFONT = 5
Const FONTSPACE = 30
'
' Simulate "redlines"
'
rpt.ScaleMode = 5    ' measure in inches
rpt.DrawWidth = 2    ' line will be 2 pixels wide
rpt.Line (1.15, 0)-(1.15, 22), 0
rpt.Line (1.18, 0)-(1.18, 22), 0
rpt.Line (7.9, 0)-(7.9, 22), 0
'
' Add Firm's name and address to left margin
'
y1 = 7.5 * 1440
y2 = y1 + 20
With rpt
    .ScaleMode = 1  ' twips
    .FontName = "Georgia"
End With

Select Case Misc
    Case 1
        'With rpt
            '.CurrentX = 280
            '.FontSize = BIGFONT
            '.CurrentY = y1
            '.Print "D"
            '.FontSize = SMALLFONT
            '.CurrentY = y2
            '.Print "IANE"
            '.FontSize = BIGFONT
            '.CurrentY = y1
            '.Print " R"
            '.FontSize = SMALLFONT
            '.CurrentY = y2
            '.Print "OSENBERG"
        'End With
        'y1 = y1 + BIGFONT * 20 + FONTSPACE
        'y2 = y1 + 20
        'With rpt
            '.CurrentX = 420
            '.FontSize = BIGFONT
            '.CurrentY = y1
            '.Print "V"
            '.FontSize = SMALLFONT
            '.CurrentY = y2
            '.Print "A"
            '.FontSize = BIGFONT
            '.CurrentY = y1
            '.Print " B"
            '.FontSize = SMALLFONT
            '.CurrentY = y2
            '.Print "AR"
            '.FontSize = BIGFONT
            '.CurrentY = y1
            '.Print " 35237"
        'End With
        'y1 = y1 + BIGFONT * 20 + FONTSPACE * 5
        'y2 = y1 + 20
        
        With rpt
            .CurrentX = 400
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "M"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "ARK"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " M"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "EYER"
        
        End With
        y1 = y1 + BIGFONT * 30 + FONTSPACE
        y2 = y1 + 20
        With rpt
            .CurrentX = 350
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "DC B"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "AR"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " 475552"
        End With
        y1 = y1 + BIGFONT * 20 + FONTSPACE
        y2 = y1 + 20
        With rpt
            .CurrentX = 350
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "MD B"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "AR"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " 15070"
        End With
        y1 = y1 + BIGFONT * 20 + FONTSPACE
        y2 = y1 + 20
        With rpt
            .CurrentX = 375
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "VA B"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "AR"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " 74290"
        End With
        y1 = y1 + BIGFONT * 20 + FONTSPACE * 5
        y2 = y1 + 20

End Select

With rpt
    .CurrentX = 410
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "R"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "OSENBERG"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " &"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 320
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "A"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "SSOCIATES"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print ", LLC"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 80
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "7910 W"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "OODMONT"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " A"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "VENUE"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 520
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "S"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "UITE"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " 750"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 0
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "B"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ETHESDA"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print ", M"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ARYLAND"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " 20814"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 360
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "(301) 907-8000"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE * 2
y2 = y1 + 20
With rpt
    .CurrentX = 190
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "F"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ILE"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " N"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "UMBER: "
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " " & FileNum
End With

Dim LResult As Long

LResult = Len([ProName])
 
If RemoveProj = True Then
    'Do nothing
Else
    If LResult > 0 Then
      
    y1 = y1 + BIGFONT * 20 + FONTSPACE * 3
    y2 = y1 + 10
    With rpt
        .CurrentX = 0
        .FontSize = BIGFONT
        .CurrentY = y1
        .Print "P"
        .FontSize = SMALLFONT
        .CurrentY = y2
        .Print "roject "
        .FontSize = BIGFONT
        .CurrentY = y1
        .Print "N"
        .FontSize = SMALLFONT
        .CurrentY = y2
        .Print "ame: "
        y1 = y1 + BIGFONT * 20 + FONTSPACE
        y2 = y1 + 5
        .FontSize = BIGFONT
        .CurrentY = y1
        .CurrentX = 0
        .Print " " & ProName
    End With
    
    y1 = y1 + BIGFONT * 20 + FONTSPACE
    y2 = y1 + 5
    With rpt
        .CurrentX = 0
        .FontSize = BIGFONT
        .CurrentY = y1
        .Print "P"
        .FontSize = SMALLFONT
        .CurrentY = y2
        .Print "roperty"
        .FontSize = BIGFONT
        .CurrentY = y1
        .Print " A"
        .FontSize = SMALLFONT
        .CurrentY = y2
        .Print "ddress: "
        y1 = y1 + BIGFONT * 20 + FONTSPACE
        y2 = y1 + 5
        .FontSize = BIGFONT
        .CurrentY = y1
        .CurrentX = 0
    
    strLength = 0
    PropAddress = Split(PropertyAddress, " ")
        For i = 0 To UBound(PropAddress) ' This is good so far...
            strLength = strLength + Len(PropAddress(i))
                If strLength > 40 Then
                    FullAddress = FullAddress + vbCrLf + PropAddress(i) & " "
                    strLength = Len(PropAddress(i))
                Else
                    FullAddress = FullAddress + PropAddress(i) & " "
                    strLength = strLength + Len(PropAddress(i))
                End If
        Next i
        .Print " " & FullAddress + vbCrLf + APTNum
    End With
    End If
End If
   
End Sub




Public Sub FirmMarginVA(rpt As Report, FileNum As Long, Optional Misc As Integer, Optional ProName As String, Optional PropertyAddress As String, Optional APTNum As String)
Dim y1 As Single, y2 As Single
Const BIGFONT = 6
Const SMALLFONT = 5
Const FONTSPACE = 30

Dim i As Integer
Dim PropAddress() As String
Dim strLength As String
Dim FullAddress As String
'
' Simulate "redlines"
'
rpt.ScaleMode = 5    ' measure in inches
rpt.DrawWidth = 2    ' line will be 2 pixels wide
rpt.Line (1.15, 0)-(1.15, 22), 0
rpt.Line (1.18, 0)-(1.18, 22), 0
rpt.Line (7.9, 0)-(7.9, 22), 0
'
' Add Firm's name and address to left margin
'
y1 = 7.5 * 1440
y2 = y1 + 20
With rpt
    .ScaleMode = 1  ' twips
    .FontName = "Georgia"
End With

Select Case Misc
    Case 1
        'With rpt
            '.CurrentX = 280
            '.FontSize = BIGFONT
            '.CurrentY = y1
            '.Print "D"
            '.FontSize = SMALLFONT
            '.CurrentY = y2
            '.Print "IANE"
            '.FontSize = BIGFONT
            '.CurrentY = y1
            '.Print " R"
            '.FontSize = SMALLFONT
            '.CurrentY = y2
            '.Print "OSENBERG"
        'End With
        'y1 = y1 + BIGFONT * 20 + FONTSPACE
        'y2 = y1 + 20
        'With rpt
            '.CurrentX = 420
            '.FontSize = BIGFONT
            '.CurrentY = y1
            '.Print "V"
            '.FontSize = SMALLFONT
            '.CurrentY = y2
            '.Print "A"
            '.FontSize = BIGFONT
            '.CurrentY = y1
            '.Print " B"
            '.FontSize = SMALLFONT
            '.CurrentY = y2
            '.Print "AR"
            '.FontSize = BIGFONT
            '.CurrentY = y1
            '.Print " 35237"
        'End With
        'y1 = y1 + BIGFONT * 20 + FONTSPACE * 5
        'y2 = y1 + 20
        
        With rpt
            .CurrentX = 400
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "M"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "ARK"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " M"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "EYER"
        
        End With
        y1 = y1 + BIGFONT * 30 + FONTSPACE
        y2 = y1 + 20
        With rpt
            .CurrentX = 350
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "DC B"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "AR"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " 475552"
        End With
        y1 = y1 + BIGFONT * 20 + FONTSPACE
        y2 = y1 + 20
        With rpt
            .CurrentX = 350
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "MD B"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "AR"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " 15070"
        End With
        y1 = y1 + BIGFONT * 20 + FONTSPACE
        y2 = y1 + 20
        With rpt
            .CurrentX = 375
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "VA B"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "AR"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " 74290"
        End With
        y1 = y1 + BIGFONT * 20 + FONTSPACE * 5
        y2 = y1 + 20

End Select

With rpt
    .CurrentX = 260
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "C"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "OMMONWEALTH"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 320
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "T"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "RUSTEES"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print ", LLC"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 280
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "8601 W"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ESTWOOD"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 320
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " C"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ENTER"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " D"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "RIVE,"
    
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 540
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "S"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "UITE"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " 255"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 40
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "V"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "IENNA"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print ", V"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "IRGINIA"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " 22182"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 300
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "(703) 752-8500"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE * 2
y2 = y1 + 20
With rpt
    .CurrentX = 190
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "F"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ILE"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " N"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "UMBER: "
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " " & FileNum
End With

Dim LResult As Long

LResult = Len([ProName])
 
If LResult > 0 Then
 
y1 = y1 + BIGFONT * 20 + FONTSPACE * 3
y2 = y1 + 10
With rpt
    .CurrentX = 0
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "P"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "roject "
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "N"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ame: "
    y1 = y1 + BIGFONT * 20 + FONTSPACE
    y2 = y1 + 5
    .FontSize = BIGFONT
    .CurrentY = y1
    .CurrentX = 0
    .Print " " & ProName
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 5
With rpt
    .CurrentX = 0
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "P"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "roperty"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " A"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ddress: "
    y1 = y1 + BIGFONT * 20 + FONTSPACE
    y2 = y1 + 5
    .FontSize = BIGFONT
    .CurrentY = y1
    .CurrentX = 0
    
strLength = 0
PropAddress = Split(PropertyAddress, " ")
    For i = 0 To UBound(PropAddress) ' This is good so far...
        strLength = strLength + Len(PropAddress(i))
            If strLength > 40 Then
                FullAddress = FullAddress + vbCrLf + PropAddress(i) & " "
                strLength = Len(PropAddress(i))
            Else
                FullAddress = FullAddress + PropAddress(i) & " "
                strLength = strLength + Len(PropAddress(i))
            End If
    Next i
    
    .Print " " & FullAddress + vbCrLf + APTNum
End With
End If



End Sub

Public Function GetWorkflowStates() As String
GetWorkflowStates = WorkflowStates
End Function

Public Function RemoveLF(s As String) As String
'
' Remove line feeds from string
'
Dim i As Integer

For i = 1 To Len(s)
    If Mid$(s, i, 1) <> vbLf Then RemoveLF = RemoveLF & Mid$(s, i, 1)
Next i
End Function

Public Function RemoveLF2(s As String) As String
'
Dim newS As String

 newS = Replace(s, vbLf, " ")

RemoveLF2 = s
End Function

Public Function StartLabel() As Boolean
Dim LabelPrinter As Long, PrinterSelected As Boolean, rstStaff As Recordset

'On Error GoTo StartLabelError

StartLabel = False
' See which label printer the user has saved, and validate it.
LabelPrinter = Nz(DLookup("LabelPrinter", "Staff", "ID=" & StaffID))
Do While LabelPrinter < 1 And LabelPrinter > 4     ' nothing saved, or invalid
    PrinterSelected = True
    LabelPrinter = Val(InputBox$("Which label printer do you want to use? (1, 2, 3 or 4)"))
Loop

If LabelPrinter = 0 Then
    If MsgBox("You have not selected a label printer.  Please click OK to choose a label printer.", vbExclamation + vbOKCancel) = vbCancel Then Exit Function
    If StaffID = 0 Then Call GetLoginName
    DoCmd.OpenForm "Preferences", , , "ID=" & StaffID, , acDialog
    DoEvents
    LabelPrinter = Nz(DLookup("LabelPrinter", "Staff", "ID=" & StaffID))
    If LabelPrinter = 0 Then
        MsgBox "Cannot print label because you did not select a label printer.  Click Preferences on the main Rosenberg & Associates screen to select a label printer.", vbCritical
        Exit Function
    End If
End If

'If PrinterSelected Then     ' save new selection
'    Set rstStaff = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID=" & StaffID, dbOpenDynaset, dbSeeChanges)
'    If Not rstStaff.EOF Then
'        rstStaff.Edit
'        rstStaff!LabelPrinter = LabelPrinter
'        rstStaff.Update
'    End If
'    rstStaff.Close
'End If
    
If LabelSequence = 0 Then
    LabelSeries = Format$(Day(Date), "00") & Format$(Now(), "hhmmss") & Format$(StaffID, "000")
    LabelSequence = 1
Else
    LabelSequence = LabelSequence + 1
End If

StartLabel = True
Open LabelRequestInbox & LabelSeries & Format$(LabelSequence, "000") & ".dat" For Output As #6
Print #6, "|Printer " & LabelPrinter
Print #6, "|User " & FullName
Exit Function

StartLabelError:
    MsgBox "Cannot print labels: " & Err.Description, vbExclamation
End Function

Public Sub FinishLabel()
Close #6
End Sub
Sub refreshFCform()

'Added by AW 7/25.  This refreshes the FC details form (Pre-Sale tab) after Title Order is printed per Diane's request.

DoCmd.SelectObject acForm, "ForeclosureDetails"
DoCmd.Requery

End Sub
Sub PrintLienCerts(FileNumber As Long, Jurisdiction As Long, Update As Boolean)
Dim PrintFlag As Boolean, rstJnl As Recordset, cost As Currency, docpath As String

docpath = "\\fileserver\applications\Database\Templates\LienCerts"

Select Case Jurisdiction

Case 4 'Balto City
If Update = False Then

StartDoc (docpath & "\LienCertBaltCity.pdf")
Else

StartDoc (docpath & "\LienCertBaltCity.pdf")
End If
Case 5 'Balto County

StartDoc (docpath & "\LienCertBaltCounty.pdf")
Case 14 'Harford

StartDoc (docpath & "\LienCertHarford.pdf")
Case 15 ' Howard

StartDoc (docpath & "\LienCertHoward.pdf")
Case 3 'Anne Arundel

StartDoc (docpath & "\LienCertAnneArundel.pdf")
Case 8 'Carroll

StartDoc (docpath & "\LienCertCarroll.pdf")
Case 12 'Frederick

StartDoc (docpath & "\LienCertFrederick.pdf")
Case 10 'Charles

StartDoc (docpath & "\LienCertCharles.pdf")
Case Else
Exit Sub
End Select


AddStatus FileNumber, Now, "Lien Certificate Ordered"
Set rstJnl = CurrentDb.OpenRecordset("select * from journal", dbOpenDynaset, dbSeeChanges)
With rstJnl
.AddNew
!FileNumber = FileNumber
!JournalDate = Now
!Who = GetFullName
!Info = "Lien Certificate Ordered"
!Color = 1
.Update
End With
Set rstJnl = Nothing

End Sub

Public Sub FirmMarginVANoLine(rpt As Report, FileNum As Long, Optional Misc As Integer, Optional ProName As String, Optional PropertyAddress As String)


Dim y1 As Single, y2 As Single
Const BIGFONT = 6
Const SMALLFONT = 5
Const FONTSPACE = 30
'
' Simulate "redlines"
'
rpt.ScaleMode = 5    ' measure in inches
rpt.DrawWidth = 2    ' line will be 2 pixels wide
'rpt.Line (1.15, 0)-(1.15, 22), 0
'rpt.Line (1.18, 0)-(1.18, 22), 0
'rpt.Line (7.9, 0)-(7.9, 22), 0
'
' Add Firm's name and address to left margin
'
y1 = 7.5 * 1440
y2 = y1 + 20
With rpt
    .ScaleMode = 1  ' twips
    .FontName = "Georgia"
End With

Select Case Misc
    Case 1
        
        
        With rpt
            .CurrentX = 400
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "M"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "ARK"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " M"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "EYER"
        
        End With
        y1 = y1 + BIGFONT * 30 + FONTSPACE
        y2 = y1 + 20
        With rpt
            .CurrentX = 350
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "DC B"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "AR"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " 475552"
        End With
        y1 = y1 + BIGFONT * 20 + FONTSPACE
        y2 = y1 + 20
        With rpt
            .CurrentX = 350
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "MD B"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "AR"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " 15070"
        End With
        y1 = y1 + BIGFONT * 20 + FONTSPACE
        y2 = y1 + 20
        With rpt
            .CurrentX = 375
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print "VA B"
            .FontSize = SMALLFONT
            .CurrentY = y2
            .Print "AR"
            .FontSize = BIGFONT
            .CurrentY = y1
            .Print " 74290"
        End With
        y1 = y1 + BIGFONT * 20 + FONTSPACE * 5
        y2 = y1 + 20

End Select

With rpt
    .CurrentX = 260
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "C"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "OMMONWEALTH"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 320
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "T"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "RUSTEES"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print ", LLC"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 280
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "8601 W"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ESTWOOD"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 320
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " C"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ENTER"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " D"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "RIVE,"
    
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 540
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "S"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "UITE"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " 255"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 40
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "V"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "IENNA"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print ", V"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "IRGINIA"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " 22182"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 20
With rpt
    .CurrentX = 300
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "(703) 752-8500"
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE * 2
y2 = y1 + 20
With rpt
    .CurrentX = 190
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "F"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ILE"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " N"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "UMBER: "
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " " & FileNum
End With

Dim LResult As Long

LResult = Len([ProName])
 
If LResult > 0 Then
 
y1 = y1 + BIGFONT * 20 + FONTSPACE * 3
y2 = y1 + 10
With rpt
    .CurrentX = 0
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "P"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "roject "
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "N"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ame: "
    y1 = y1 + BIGFONT * 20 + FONTSPACE
    y2 = y1 + 5
    .FontSize = BIGFONT
    .CurrentY = y1
    .CurrentX = 0
    .Print " " & ProName
End With

y1 = y1 + BIGFONT * 20 + FONTSPACE
y2 = y1 + 5
With rpt
    .CurrentX = 0
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print "P"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "roperty"
    .FontSize = BIGFONT
    .CurrentY = y1
    .Print " A"
    .FontSize = SMALLFONT
    .CurrentY = y2
    .Print "ddress: "
    y1 = y1 + BIGFONT * 20 + FONTSPACE
    y2 = y1 + 5
    .FontSize = BIGFONT
    .CurrentY = y1
    .CurrentX = 0
    .Print " " & PropertyAddress
End With
End If


End Sub
