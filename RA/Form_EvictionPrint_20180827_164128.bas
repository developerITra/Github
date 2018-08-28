VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EvictionPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboAttorney_AfterUpdate()
If IsNull(Me.cboAttorney) = False Then Forms!EvictionDetails!Attorney = Me.cboAttorney

End Sub
Private Sub chMotion_Click()
'opt1.Enabled = chMotion
   'If InPossession = 4 Then
       'chNoticeToTenant = 1
   'End If
    
End Sub

Private Sub cmdAcrobat_Click()
Call PrintDocs(-2)
End Sub

Private Sub cmdClear_Click()
On Error GoTo Err_cmdClear_Click

Me.chTermination = 0
Me.chExpiredLease = 0
chCourtesy = 0
chMotion = 0
chNoticeToQuit = 0
chWrit = 0
chWritCover = 0
chWritCoverGen = 0
chComplaintCover = 0
chLineOfAppearance = 0
chAffidavit14102b = 0
Me.chNoticeToOccupantBalt = 0
chNoticeToOccupant = 0
ChNoticeToTenant = 0
'chNoticeToTenant = 0 '2012.02.02 DaveW Control removed
chFinalNoticeToOccupant = 0
'ch90Tenant = 0
'ch90Unknown = 0
'ch90Rent = 0
'ch90Select = 0
'ch90Wilshire = 0
ch90Freddie = 0
ch90McAlla = 0
ch90_7 = 0
ch90Wells = 0
ch90FreddieMD = 0
ch90McAllaMD = 0
ch90_7MD = 0
chLPS = 0

'chRentWelcome = 0 '2012.02.02 DaveW Control removed
chCashForKeys = 0
chDispPerProp = 0

Exit_cmdClear_Click:
    Exit Sub

Err_cmdClear_Click:
    MsgBox Err.Description
    Resume Exit_cmdClear_Click
    
End Sub

Private Sub PrintDocs(PrintTo As Integer)
Dim ReportName As String

On Error GoTo Err_PrintDocs

'If chMotion And IsNull(opt1) Then
    'MsgBox "Select Affidavit of Service or Show Cause Order", vbCritical
'    MsgBox "Select Affidavit of Service to Debtor or to Occupant", vbCritical
'    Exit Sub
'End If

If chMotion Then
    DoReport "Eviction Motion for Possession " & State, PrintTo
    
    If JurisdictionID = 4 Then      ' Baltimore City
        DoReport "Eviction Order Awarding Possession BaltCity", PrintTo
    Else
        DoReport "Eviction Order Awarding Possession", PrintTo
    End If
    
    DoReport "Eviction Affidavit of Service Debtor", PrintTo

    
'    Select Case opt1
'        Case 1
'            DoReport "Eviction Affidavit of Service Debtor", PrintTo
'        Case 2
'            DoReport "Eviction Affidavit of Service Occupant", PrintTo
'        Case 2
'            If JurisdictionID = 4 Then      ' Baltimore City
'                ReportName = "Eviction Show Cause Order BaltCity"
'            Else
'                ReportName = "Eviction Show Cause Order " & State
'            End If
'            DoReport ReportName, PrintTo
'    End Select

    If MsgBox("Add to status: Motion Filed = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
        Forms!EvictionDetails!MotionFiled = Now()
        AddStatus FileNumber, Now(), "Motion for Judgment of Possession filed"
    End If
End If

If Me.chExpiredLease Then
    DoReport "MD Expired Lease", PrintTo
'    Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
    Call doc_SCRANotice90Day(EMailStatus <> 1)
    Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
End If

If Me.chTermination Then
    DoReport "MD Lease Termination", PrintTo
'    Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
    Call doc_SCRANotice90Day(EMailStatus <> 1)
    Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
End If


If chNoticeToQuit Then
    DoReport "Notice To Quit " & State, PrintTo
    If State = "DC" Then
        DoReport "Notice To Quit DC Tenant", PrintTo
        'Call StartDoc(TemplatePath & "EV-SCRA.pDF")
        DoReport "Ev-SCRA", PrintTo
    End If
End If

If chCourtesy Then
    DoReport "Courtesy Eviction Letter", PrintTo
   ' DoReport "Courtesy Eviction Letter ALL Occ", PrintTo
End If
If chLPS Then
    DoReport "qryEV_LPSdesktop", -3
End If

If chWrit Then
    DoReport "Eviction Writ MD", PrintTo
End If

If chWritCover Then
    ' Added chWritCover:  This had been automatic with chWrit '  DaveW 2012.02.20
    ' Baltimore City & County get a cover letter
    If JurisdictionID = 4 Or JurisdictionID = 5 Then
        DoReport "Eviction Writ Cover MD", PrintTo
    ElseIf State = "VA" Then
        DoReport "Eviction Writ Cover VA", PrintTo
    Else
        MsgBox "No Writ cover letter for this jurisdiction."
    End If
End If

If chWritCoverGen Then
    If DLookup("FillingFeePrepaid", "EVdetails", "Current=True AND FileNumber=" & Me.FileNumber) = False Then
        DoReport "Eviction Writ and General Notice Cover", PrintTo
    Else
        DoReport "Eviction Writ and General Notice Cover_Yes", PrintTo
    End If
End If

If chComplaintCover Then
    ' Added chWritCover:  This had been automatic with chWrit '  DaveW 2012.02.20
    ' Baltimore City & County get a cover letter
    ' DaveW 2012.03.08 Added DC & MD
    Select Case State
    Case "VA"
        DoReport "Eviction Complaint Cover VA", PrintTo
    Case "MD"
        DoReport "Eviction Complaint Cover MD", PrintTo
    Case "DC"
        DoReport "Eviction Complaint Cover DC", PrintTo
    Case Else
        MsgBox "No Complaint cover letter for this jurisdiction."
    End Select
End If

If chFormerOwner Then
    Select Case State
    Case "VA"
        DoReport "Eviction Print VA Owner NTQ 5/3", PrintTo
        Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
    Case "DC"
        DoReport "Eviction Print PTFA NTQ 5/3", PrintTo
       ' Call StartDoc(TemplatePath & "EV-SCRA.pDF")
    Case Else
        DoReport "Eviction Print PTFA NTQ 5/3", PrintTo
        Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
  '     Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
       
    End Select
End If

If chFormerOwnerMD Then
    Select Case State
    Case "VA"
        DoReport "Eviction Print VA Owner NTQ 5/3", PrintTo
        Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
    Case "DC"
        DoReport "Eviction Print MD NTQ 5/3", PrintTo
       ' Call StartDoc(TemplatePath & "EV-SCRA.pDF")
    Case Else
        DoReport "Eviction Print MD NTQ 5/3", PrintTo
       ' Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
'       Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
        Call doc_SCRANotice90Day(EMailStatus <> 1)
        Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
    End Select
End If

If chTenantMD Then
    Select Case State
    Case "VA"
        DoReport "Eviction Print VA Tenant NTQ 5/3", PrintTo
        Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
    Case "DC"
        DoReport "Eviction Print MD NTQ 5/3", PrintTo
       ' Call StartDoc(TemplatePath & "EV-SCRA.pDF")
    Case Else
        DoReport "Eviction Print MD NTQ 5/3", PrintTo
        'Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
'        Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
        Call doc_SCRANotice90Day(EMailStatus <> 1)
        Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
    End Select
End If

If chTenant Then
    Select Case State
    Case "VA"
        DoReport "Eviction Print VA Tenant NTQ 5/3", PrintTo
        Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
    Case "DC"
        DoReport "Eviction Print PTFA NTQ 5/3", PrintTo
       ' Call StartDoc(TemplatePath & "EV-SCRA.pDF")
    Case Else
        DoReport "Eviction Print PTFA NTQ 5/3", PrintTo
        Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
   '     Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
        'Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
    End Select
End If

If chOwnerSPSMD Then
    Select Case State
    Case "DC"
        DoReport "Eviction Print MD Owner NTQ SPS", PrintTo
     '   Call StartDoc(TemplatePath & "EV-SCRA.pDF")
 
    Case Else
        DoReport "Eviction Print MD Owner NTQ SPS", PrintTo
        'Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
'        Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
        Call doc_SCRANotice90Day(EMailStatus <> 1)
        Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
    End Select
End If

If chOwnerSPS Then
    Select Case State
    Case "DC"
        DoReport "Eviction Print PTFA Owner NTQ SPS", PrintTo
     '   Call StartDoc(TemplatePath & "EV-SCRA.pDF")
 
    Case Else
        DoReport "Eviction Print PTFA Owner NTQ SPS", PrintTo
        Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
    '    Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
        'Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
    End Select
End If

If chTenantSPS Then
Select Case State
    Case "DC"
       DoReport "Eviction Print PTFA Tenant NTQ SPS", PrintTo
     '   Call StartDoc(TemplatePath & "EV-SCRA.pDF")
    Case Else
        DoReport "Eviction Print PTFA Tenant NTQ SPS", PrintTo
        Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
     '   Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
        'Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
    End Select
End If

If chTenantSPSMD Then
Select Case State
    Case "DC"
       DoReport "Eviction Print MD Tenant NTQ SPS", PrintTo
     '   Call StartDoc(TemplatePath & "EV-SCRA.pDF")
    Case Else
        DoReport "Eviction Print MD Tenant NTQ SPS", PrintTo
        'Call StartDoc(TemplatePath & "PTFA_Scra.pdf")
'        Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
        Call doc_SCRANotice90Day(EMailStatus <> 1)
        Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
    End Select
End If

If chLineOfAppearance Then DoReport "Eviction Line of Appearance", PrintTo

If chAffidavit14102b Then ' DoReport "Eviction Affidavit 14-102b", PrintTo
  Call DoReport("Eviction Affidavit 14-102b", PrintTo) '#1136 10/1/2014 MC

  If MsgBox("Add to status: " & Format$(Date, "mm/dd/yyyy") & "?", vbYesNo + vbQuestion) = vbYes Then
      AddStatus Me!FileNumber, Date, "Eviction Affidavit Pursuant Sent"
  End If


'  If MsgBox("Add to status: " & Format$(Date, "mm/dd/yyyy") & " " & statusMsg, vbYesNo + vbQuestion) = vbYes Then
'      AddStatus Me!FileNumber, Date, "Eviction Affidavit Pursuant Sent"
'  End If


' DoCmd.OpenForm "Print Eviction Affidavit Pursuant", , , "FileNumber=" & Forms!EvictionDetails!FileNumber, , , PrintTo
End If

If chNoticeToOccupant Then
DoReport "Eviction Notice to Occupant", PrintTo
    If (Me.State = "MD" And Forms!EvictionDetails.InPossession = 2) Or (Me.State = "MD" And Forms!EvictionDetails.InPossession = 3) Then
        Call StartDoc(TemplatePath & "MDRule2-321.pdf")
    End If
End If

If ChNoticeToTenant Then DoReport "Eviction MD Notice 1308", PrintTo

If Me.chNoticeToOccupantBalt Then
    DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage for Final Baltimore City Notice|BK-FNEV|Final Notice of Eviction BaltiCi"
    If MsgBox("Do you want to update the status with the current Final BaltCi Notice?", vbYesNo) = vbYes Then
        AddStatus Me.FileNumber, Date, "Final BaltCi Notice"
    Else
    End If
    DoReport "Eviction Notice Baltimore City", PrintTo
End If

If chFinalNoticeToOccupant Then
    
    DoReport "Eviction Notice to Occupant Final", PrintTo
    DoReport "Eviction Final Notice Affidavit of Service", PrintTo
   
End If

'If chNoticeToTenant Then DoReport "Eviction Notice to Tenant", PrintTo

'If ch90Tenant Then DoReport "General 90 Day Notice tenant", PrintTo
'If ch90Unknown Then DoReport "90 Day Notice Unknown Occupant", PrintTo
'If ch90Rent Then DoReport "90 Day Notice W Rent SPS", PrintTo
'If ch90Select Then DoReport "Select Portfolio servicing 90 Day Notice", PrintTo
'If ch90Wilshire Then DoReport "Wilshire Servicing Corporation 90 Day Notice", PrintTo

If ch90Freddie Then
    DoReport "Eviction Print 90 Day Notice - Freddie", PrintTo
'    Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
    Call doc_SCRANotice90Day(EMailStatus <> 1)
    Call StartDoc(TemplatePath & "PTFA.pdf")
    
End If

If ch90FreddieMD Then
    DoReport "Eviction Print 90 Day Notice - FreddieMD", PrintTo
'    Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
    Call doc_SCRANotice90Day(EMailStatus <> 1)
    Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
End If

If ch90Wells Then ' #1347 This document will be phased out 1/1/2015
     DoReport "Eviction Print 90 Day Notice - Wells", PrintTo
'    Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
    Call doc_SCRANotice90Day(EMailStatus <> 1)
    Call StartDoc(TemplatePath & "PTFA.pdf")
    'Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
    'Call StartDoc(TemplatePath & "EV-SCRA.pDF") '8/13/14 Removed 2nd SCRA Document #1029
    
End If

If ch90McAlla Then
    DoReport "Eviction Print 90 Day Notice - McAlla", PrintTo
'    Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
    Call doc_SCRANotice90Day(EMailStatus <> 1)
    Call StartDoc(TemplatePath & "PTFA.pdf")
End If

If ch90McAllaMD Then
    DoReport "Eviction Print 90 Day Notice - McAllaMD", PrintTo
'    Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
    Call doc_SCRANotice90Day(EMailStatus <> 1)
    Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
End If


If ch90_7 Then
    DoReport "Eviction Print 90 Day Notice - 7", PrintTo
'    Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
    Call doc_SCRANotice90Day(EMailStatus <> 1)
    Call StartDoc(TemplatePath & "PTFA.pdf")
End If

If ch90_7MD Then
    DoReport "Eviction Print 90 Day Notice - 7MD", PrintTo
'    Call StartDoc(TemplatePath & "SCRA Notice 90-Day.docx")
    Call doc_SCRANotice90Day(EMailStatus <> 1)
    Call StartDoc(TemplatePath & "Md Code 7-105.6.pdf")
End If

'If chRentWelcome Then DoReport "Rent Welcome Letter", PrintTo '2012.02.02 DaveW Control removed
If chCashForKeys Then
    If MsgBox("Update cash for keys sent = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
        Forms!EvictionDetails!CashForKeysDate = Now()
        AddStatus FileNumber, Now(), "Cash For Keys Letter Sent"
    End If

    DoReport "Eviction Cash For Keys Letter", PrintTo
    Call StartDoc(TemplatePath & "\fw9.pdf")

End If

If chCreateLabel Then
    DoCmd.OpenForm "Getlabel"
Else
End If


If chDispPerProp Then DoReport "Eviction Disposal of Personal Property", PrintTo

Exit Sub

Err_PrintDocs:
    MsgBox Err.Description
    Exit Sub

End Sub

Private Sub Form_Current()

If Me.State = "VA" Then
  cboAttorney.RowSource = "SELECT Staff.ID, Staff.Name " & _
                       "FROM Staff " & _
                       "WHERE  ((staff.active = true) And (Staff.Attorney =True) And(staff.PracticeVA = true )) " & _
                       "ORDER BY Staff.Sort;"
                   'It was  "WHERE (((Staff.CommonwealthTitle) Is Not Null)) and staff.active = true " S.A.
'staff.active=true
ElseIf Me.State = "MD" Then
  cboAttorney.RowSource = "SELECT Staff.ID, Staff.Name & ', Esq.' FROM Staff WHERE ((Staff.active = true ) and (Staff.Attorney = True) and (Staff.PracticeMD = True)) ORDER BY Staff.Sort;"
Else
  cboAttorney.RowSource = "SELECT Staff.ID, Staff.Name FROM Staff WHERE ((staff.active = true) and (Staff.Attorney = True) and (staff.PracticeDC = true ))ORDER BY Staff.Sort;"

End If


'If IsNull(Attorney) Then
'    MsgBox "You must select an attorney before you can print.", vbCritical
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
'    cmdWord.Enabled = False
'End If

Me.Caption = "Print Eviction " & [CaseList.FileNumber] & " " & [PrimaryDefName]
Me.cboAttorney = Forms!EvictionDetails!Attorney

If ([Forms]![Case List]!Active = False) Then
chMotion.Enabled = False
Option147.Enabled = False
Option149.Enabled = False
chNoticeToQuit.Enabled = False
chWrit.Enabled = False
chCourtesy.Enabled = False
chWritCover.Enabled = False
chWritCoverGen.Enabled = False
chComplaintCover.Enabled = False
chLineOfAppearance.Enabled = False
chAffidavit14102b.Enabled = False
chNoticeToOccupant.Enabled = False
ChNoticeToTenant.Enabled = False
chFinalNoticeToOccupant.Enabled = False
ch90Freddie.Enabled = False
ch90McAlla.Enabled = False
ch90_7.Enabled = False
ch90Wells.Enabled = False
chLPS.Enabled = False
NotaryID.Enabled = False
chCashForKeys.Enabled = False
chDispPerProp.Enabled = False
cboAttorney.Enabled = False
Label209.Visible = True
Me.ch90_7MD.Enabled = False
Me.ch90FreddieMD.Enabled = False
Me.ch90_7MD.Enabled = False
Me.chTermination = False
Me.chExpiredLease = False

End If

If Forms![Case List]!JurisdictionID.Column(0) = 4 Then
    Me.chNoticeToOccupantBalt.Visible = True
Else
    Me.chNoticeToOccupantBalt.Visible = False
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

Private Sub Form_Open(Cancel As Integer)
If Me.State = "VA" Then
' cboAttorney.RowSource = "SELECT Staff.ID, Staff.Name & ', ' & [CommonWealthTitle] AS CWRep " & _

 cboAttorney.RowSource = "SELECT Staff.ID, Staff.Name & ' ' " & _
                       "FROM Staff " & _
                       "WHERE  ((staff.active = True) And (Staff.Attorney =True) And(staff.PracticeVA = True )) " & _
                       "ORDER BY  Staff.Sort;"
' "ORDER BY Staff.CommonwealthTitle, Staff.Sort;"
                   'It was  "WHERE (((Staff.CommonwealthTitle) Is Not Null)) and staff.active = true " S.A.
'staff.active=true
ElseIf Me.State = "MD" Then
  cboAttorney.RowSource = "SELECT Staff.ID, Staff.Name & ', Esq.' FROM Staff WHERE ((Staff.active = true ) and (Staff.Attorney = True) and (Staff.PracticeMD = True)) ORDER BY Staff.Sort;"
Else
 cboAttorney.RowSource = "SELECT Staff.ID, Staff.Name FROM Staff WHERE ((staff.active = true) and (Staff.Attorney = True) and (staff.PracticeDC = true ))ORDER BY Staff.Sort;"

End If

Call cmdClear_Click
End Sub

