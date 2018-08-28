VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim CheckVersionCounter As Integer

Private Sub cmd_VolReports_Click()
DoCmd.OpenForm "ReportsMenu_Volume"
End Sub

Private Sub cmdAuctioneer_Click()
On Error GoTo Err_cmdAuctioneer_Click

  DoCmd.OpenForm "Auctioneers"

Exit_cmdAuctioneer_Click:
  Exit Sub

Err_cmdAuctioneer_Click:
  MsgBox Err.Description
  Resume Exit_cmdAuctioneer_Click
  
End Sub

Private Sub cmdBKAttorney_Click()
DoCmd.OpenForm "BKAttorneys"
End Sub

Private Sub cmdCertofServDocs_Click()
On Error GoTo Err_cmdCertofServDocs_Click
DoCmd.OpenForm "Certificate of Service Documents"

Exit_cmdCertofServDocs_Click:
    Exit Sub

Err_cmdCertofServDocs_Click:
    MsgBox Err.Description
    Resume Exit_cmdCertofServDocs_Click
End Sub

Private Sub cmdCheckRequest_Click()
On Error GoTo Err_cmdCheckRequest_Click

  DoCmd.OpenForm "CheckRequests"
  
Exit_cmdCheckRequest_Click:
  Exit Sub
  
Err_cmdCheckRequest_Click:
  MsgBox Err.Description
  Resume Exit_cmdCheckRequest_Click
  
End Sub

Private Sub cmdConflicts_Click()
On Error GoTo Err_cmdConflicts_Click

  DoCmd.OpenForm "Conflicts"

Exit_cmdConflicts_Click:
  Exit Sub
  
Err_cmdConflicts_Click:
  MsgBox Err.Description
  Resume Exit_cmdConflicts_Click
  
End Sub

Private Sub cmdDistrictsEV_Click()

On Error GoTo Err_cmdDistrictsEV_Click
DoCmd.OpenForm "DistrictsEV"

Exit_cmdDistrictsEV_Click:
    Exit Sub

Err_cmdDistrictsEV_Click:
    MsgBox Err.Description
    Resume Exit_cmdDistrictsEV_Click


End Sub

Private Sub cmdDocRequest_Click()

On Error GoTo Err_cmdDocRequest_Click

  DoCmd.OpenForm "DocumentRequests"
  
Exit_cmdDocRequest_Click:
  Exit Sub
  
Err_cmdDocRequest_Click:
  MsgBox Err.Description
  Resume Exit_cmdDocRequest_Click

End Sub

Private Sub cmdGit_Click()
Dim wbPath As String
Dim vbComp As VBComponent
Dim exportPath As String

  wbPath = "C:\Github\RA"
  
  For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
    exportPath = wbPath & "\" & vbComp.Name
'    If vbComp.Type = vbext_ct_ClassModule Or vbComp.Type = vbext_ct_StdModule Then
         vbComp.Export exportPath
        Select Case vbComp.Type
            Case 1 ' Standard Module
                exportPath = exportPath & ".bas"
            Case 2 ' UserForm
                exportPath = exportPath & ".frm"
            Case 3 ' Class Module
                exportPath = exportPath & ".cls"
            Case Else ' Anything else
                exportPath = exportPath & ".bas"
        End Select
'        ToFileExtension (vbComp.Type)
'    End If
    On Error Resume Next
    vbComp.Export exportPath
    On Error GoTo 0
  
  Next

End Sub

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
Select Case vbeComponentType
Case vbext_ComponentType.vbext_ct_ClassModule
ToFileExtension = ".cls"
Case vbext_ComponentType.vbext_ct_StdModule
ToFileExtension = ".bas"
Case vbext_ComponentType.vbext_ct_MSForm
ToFileExtension = ".frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner
Case vbext_ComponentType.vbext_ct_Document
Case Else
ToFileExtension = vbNullString
End Select
 
End Function


Private Sub cmdNewCase_Click()
'DoCmd.OpenForm "wizIntake1"
DoCmd.OpenForm "New Case"
End Sub

Private Sub cmdQueuesLexis_Click()
DoCmd.OpenForm "QueuesLexis"

End Sub

Private Sub cmdReports_Click()
DoCmd.OpenForm "ReportsMenu"
End Sub


Private Sub cmdSearch_Click()
DoCmd.OpenForm "Search"
End Sub

Private Sub cmdStaff_Click()

On Error GoTo Err_cmdStaff_Click
DoCmd.OpenForm "frmStaff", , , "[ID] = " & StaffID, acFormEdit

Exit_cmdStaff_Click:
    Exit Sub

Err_cmdStaff_Click:
    MsgBox Err.Description
    Resume Exit_cmdStaff_Click
    
End Sub

Private Sub cmdJurisdictions_Click()

On Error GoTo Err_cmdJurisdictions_Click
DoCmd.OpenForm "Jursidictions"

Exit_cmdJurisdictions_Click:
    Exit Sub

Err_cmdJurisdictions_Click:
    MsgBox Err.Description
    Resume Exit_cmdJurisdictions_Click
    
End Sub

Private Sub cmdAbstractors_Click()

On Error GoTo Err_cmdAbstractors_Click
DoCmd.OpenForm "Abstractors"

Exit_cmdAbstractors_Click:
    Exit Sub

Err_cmdAbstractors_Click:
    MsgBox Err.Description
    Resume Exit_cmdAbstractors_Click
    
End Sub

Private Sub cmdClients_Click()

On Error GoTo Err_cmdClients_Click
DoCmd.OpenForm "Client List"

Exit_cmdClients_Click:
    Exit Sub

Err_cmdClients_Click:
    MsgBox Err.Description
    Resume Exit_cmdClients_Click
    
End Sub

Private Sub cmdSelectFile_Click()

On Error GoTo Err_cmdSelectCase_Click
DoCmd.OpenForm "Select File"

Exit_cmdSelectCase_Click:
    Exit Sub

Err_cmdSelectCase_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectCase_Click
    
End Sub

Private Sub Command43_Click()

End Sub




Private Sub CmdWizAccon_Click()
DoCmd.OpenForm "ReportMenuTranking"
End Sub

Private Sub ComEvConatact_Click()

Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "EVContactsByClient"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

End Sub


Private Sub ComFCTrustee_Click()
DoCmd.OpenForm "FCTrustees"

End Sub

Private Sub Detail_Click()

End Sub

Private Sub Form_Open(Cancel As Integer)
Dim d As Recordset
'If CurrentProject.Name = "RA.accdb" Or CurrentProject.Name = "RA.accde" Then
    txtVersion = DBVersion
'Else
  '  Dim StaffID As Integer
   ' StaffID = Nz(DLookup("ID", "Staff", "Username=""" & GetLoginName() & """"), 0)
    '    If Nz(DLookup("PrivDataManager", "Staff", "ID=" & StaffID)) = 0 Then
     '       MsgBox "You are not authorized to use Rosie Test .  Please see Diane if you need assistance.", vbCritical
     '       DoCmd.Quit
     '   End If
   '' txtVersion = DBVersionTest
 '   Me.Detail.BackColor = -2147483615
'End If

'Call CheckVersion
'Call CheckSQL

txtLoginName = GetLoginName()
cmdPermissions.Visible = PrivAdmin
cmdOrigTrustBen.Visible = PrivAdmin
cmdCheckRequest.Visible = PrivCheckRequest
cmdConflicts.Visible = PrivAdmin
cmdDocRequest.Visible = PrivDocMgmt

FileLocks = DLookup("iValue", "DB", "Name='FileLocks'")

Set d = CurrentDb.OpenRecordset("SELECT * FROM DB WHERE Name = 'WebUpdate';", dbOpenSnapshot)
d.MoveFirst
If CVDate(d("sValue")) < Date Then MsgBox "Web update has not run.  You can run it manually from the Special screen.", vbExclamation
d.Close

Dim Username As String
Username = LCase(Nz(DLookup("Username", "Staff", "ID=" & StaffID)))

If (CVDate(Nz(DLookup("sValue", "DB", "Name = 'AuctionDotCom'"))) < Date) And (InStr(Nz(DLookup("sValue", "DB", "Name = 'AuctionNotifyUsers'")), Username) > 0) Then

    MsgBox "Auction.com update has not run.  Please notify programming to review and upload.", vbExclamation

End If

If InStr(1, CalendarFolderName, "Test") <> 0 Then MsgBox "CAUTION! Test Calendar in use, fix constant CalenderFolderName in module Calendar before publishing!", vbExclamation
Call Form_Timer

Call CheckSendingEmail

End Sub

Private Sub cmdExit_Click()
On Error GoTo Err_cmdExit_Click
    
If LockedFileNumber <> 0 Then Call ReleaseFile(LockedFileNumber)
Call Unlockedfiles(GetStaffID())
Call StaffSignOut

DoCmd.Quit

Exit_cmdExit_Click:
    Exit Sub

Err_cmdExit_Click:
    MsgBox Err.Description
    Resume Exit_cmdExit_Click
    
End Sub

Private Sub cmdDistricts_Click()

On Error GoTo Err_cmdDistricts_Click
DoCmd.OpenForm "Districts"

Exit_cmdDistricts_Click:
    Exit Sub

Err_cmdDistricts_Click:
    MsgBox Err.Description
    Resume Exit_cmdDistricts_Click
    
End Sub

Private Sub cmdTrustees_Click()

On Error GoTo Err_cmdTrustees_Click
DoCmd.OpenForm "Trustees"

Exit_cmdTrustees_Click:
    Exit Sub

Err_cmdTrustees_Click:
    MsgBox Err.Description
    Resume Exit_cmdTrustees_Click
    
End Sub

Private Sub cmdWorkflow_Click()

On Error GoTo Err_cmdWorkflow_Click
DoCmd.OpenForm "ReportsWorkflow"

Exit_cmdWorkflow_Click:
    Exit Sub

Err_cmdWorkflow_Click:
    MsgBox Err.Description
    Resume Exit_cmdWorkflow_Click
    
End Sub

Private Sub cmdTitleInsurers_Click()

On Error GoTo Err_cmdTitleInsurers_Click
DoCmd.OpenForm "Title Insurance Addresses"

Exit_cmdTitleInsurers_Click:
    Exit Sub

Err_cmdTitleInsurers_Click:
    MsgBox Err.Description
    Resume Exit_cmdTitleInsurers_Click
    
End Sub

Private Sub cmdAuditors_Click()

On Error GoTo Err_cmdAuditors_Click
DoCmd.OpenForm "Auditors"

Exit_cmdAuditors_Click:
    Exit Sub

Err_cmdAuditors_Click:
    MsgBox Err.Description
    Resume Exit_cmdAuditors_Click
    
End Sub

Private Sub Form_Timer()

txtDate = Now()

' About every 15 minutes, check the version again.  Just issue a warning, don't actually upgrade.
' Changed this to check only every 30 minutes JAE 04-15-2015
CheckVersionCounter = CheckVersionCounter + 1
If CheckVersionCounter >= 60 Then
    CheckVersionCounter = 0
    Call CheckVersion(True)
    Exit Sub
ElseIf CheckVersionCounter >= 31 Then
    Call CheckVersion(False)
End If

End Sub

Private Sub cmdProcessServers_Click()

On Error GoTo Err_cmdProcessServers_Click
DoCmd.OpenForm "ProcessServers"

Exit_cmdProcessServers_Click:
    Exit Sub

Err_cmdProcessServers_Click:
    MsgBox Err.Description
    Resume Exit_cmdProcessServers_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)


If LockedFileNumber <> 0 Then Call ReleaseFile(LockedFileNumber)
Call Unlockedfiles(GetStaffID())
Call StaffSignOut



End Sub

Private Sub txtVersion_Click()
cmdExit.SetFocus
DoCmd.RunCommand acCmdAboutMicrosoftAccess
End Sub

Private Sub cmdPreferences_Click()

On Error GoTo Err_cmdPreferences_Click

If StaffID = 0 Then Call GetLoginName
DoCmd.OpenForm "Preferences", , , "ID=" & StaffID

Exit_cmdPreferences_Click:
    Exit Sub

Err_cmdPreferences_Click:
    MsgBox Err.Description
    Resume Exit_cmdPreferences_Click
    
End Sub

Private Sub cmdFileLocks_Click()

On Error GoTo Err_cmdFileLocks_Click
DoCmd.OpenForm "FileLocks"

Exit_cmdFileLocks_Click:
    Exit Sub

Err_cmdFileLocks_Click:
    MsgBox Err.Description
    Resume Exit_cmdFileLocks_Click
    
End Sub

Private Sub cmdFileLocations_Click()

On Error GoTo Err_cmdFileLocations_Click
DoCmd.OpenForm "Vendors"

Exit_cmdFileLocations_Click:
    Exit Sub

Err_cmdFileLocations_Click:
    MsgBox Err.Description
    Resume Exit_cmdFileLocations_Click
    
End Sub

Private Sub cmdDocumentTypes_Click()

On Error GoTo Err_cmdDocumentTypes_Click
DoCmd.OpenForm "Document Types"

Exit_cmdDocumentTypes_Click:
    Exit Sub

Err_cmdDocumentTypes_Click:
    MsgBox Err.Description
    Resume Exit_cmdDocumentTypes_Click
    
End Sub

Private Sub cmdPermissions_Click()

On Error GoTo Err_cmdPermissions_Click
DoCmd.OpenForm "StaffPermissions"

Exit_cmdPermissions_Click:
    Exit Sub

Err_cmdPermissions_Click:
    MsgBox Err.Description
    Resume Exit_cmdPermissions_Click
    
End Sub
Private Sub cmdOrigTrustBen_Click()
On Error GoTo Err_cmdOrigTrustBen_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "OriginalTrusteesAndBeneficiaries"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOrigTrustBen_Click:
    Exit Sub

Err_cmdOrigTrustBen_Click:
    MsgBox Err.Description
    Resume Exit_cmdOrigTrustBen_Click
    
End Sub
Private Sub cmdBrokers_Click()
On Error GoTo Err_cmdBrokers_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Brokers"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdBrokers_Click:
    Exit Sub

Err_cmdBrokers_Click:
    MsgBox Err.Description
    Resume Exit_cmdBrokers_Click
    
End Sub
Private Sub cmdSpecial_Click()
On Error GoTo Err_cmdSpecial_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Special"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdSpecial_Click:
    Exit Sub

Err_cmdSpecial_Click:
    MsgBox Err.Description
    Resume Exit_cmdSpecial_Click
    
End Sub
Private Sub cmdCivilCourts_Click()
On Error GoTo Err_cmdCivilCourts_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "CIV_Court"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdCivilCourts_Click:
    Exit Sub

Err_cmdCivilCourts_Click:
    MsgBox Err.Description
    Resume Exit_cmdCivilCourts_Click
    
End Sub

Private Sub cmdWizards_Click()

On Error GoTo Err_cmdWizards_Click
DoCmd.OpenForm "Wizards"

Exit_cmdWizards_Click:
    Exit Sub

Err_cmdWizards_Click:
    MsgBox Err.Description
    Resume Exit_cmdWizards_Click
    
End Sub

Private Sub cmdQueues_Click()

On Error GoTo Err_cmdQueues_Click
DoCmd.OpenForm "Queues"

Exit_cmdQueues_Click:
    Exit Sub

Err_cmdQueues_Click:
    MsgBox Err.Description
    Resume Exit_cmdQueues_Click
    
End Sub
