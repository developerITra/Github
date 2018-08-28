VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Wizards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdBorrowerServed_Click()
Dim FileNumber As Long
FileNumber = InputBox("Enter the File Number", "Borrower Served Wizard")
    AddToList (FileNumber)
    BorrowerServedCallFromQueue FileNumber
End Sub

Private Sub CmdDemand_Click()
Dim FileNumber As Long
Dim rstqueue As Recordset
Dim rsfC As Recordset
Dim stDocName As String
Dim stLinkCriteria As String
Dim FileNO As Integer

'AccelerationIssued
'AccelerationLetter
FileNumber = InputBox("Enter the File Number", "Demand Wizard")



Set rsfC = CurrentDb.OpenRecordset("Select * from FCdetails where filenumber = " & FileNumber & " and current=true ", dbOpenDynaset, dbSeeChanges)
        With rsfC
            If Not IsNull(rsfC!AccelerationIssued) Or Not IsNull(rsfC!AccelerationLetter) Then
            MsgBox "The File already has Acceleration date please use New bottomn to add it to the queue.", vbCritical
           
            Exit Sub
            End If
            
            Set rsfC = Nothing
        End With


            If Not IsNull(DLookup("File", "qryQueueDemand_P", "File=" & FileNumber)) Then
            MsgBox ("The file is already in Demand queue")
            Exit Sub
            End If

            If Not IsNull(DLookup("File", "qryQueueDemandWaiting_P", "File=" & FileNumber)) Then
            MsgBox ("The file is already in Demand Waiting queue ")
            Exit Sub
            End If

    AddToList (FileNumber)
    DemandCallFromQueue FileNumber
End Sub

Private Sub cmdDocketing_Click()
Dim FileNumber As Long
FileNumber = InputBox("Enter the File Number", "Docketing Wizard")

    DocketingCallFromQueue FileNumber
End Sub

Private Sub cmdFairDebt_Click()
On Error GoTo Err_cmdFairDebt_Click
Dim FileNumber As Long
Dim rsfC As Recordset

FileNumber = InputBox("Enter the File Number", "Fair Debt Wizard")

Set rsfC = CurrentDb.OpenRecordset("Select * from FCdetails where filenumber = " & FileNumber & " and current=true ", dbOpenDynaset, dbSeeChanges)
    With rsfC
        If Not IsNull(rsfC!FairDebt) Then
        MsgBox "The File already has Fair Debt date please use Roise to Add it to the queue.", vbCritical
        Set rsfC = Nothing
        Exit Sub
        End If
    End With
Set rsfC = Nothing
        

            If Not IsNull(DLookup("File", "qryQueueFairDebt_P", "File=" & FileNumber)) Then
            MsgBox ("The file is already in FairDebt queue in Rosie")
            Exit Sub
            End If
            
            If Not IsNull(DLookup("File", "qryQueueFairDebtWaiting_P", "File=" & FileNumber)) Then
            MsgBox ("The file is already in FairDebt Waiting queue in Rosie")
            Exit Sub
            End If


    AddToList (FileNumber)
    FairDebtCallFromQueue FileNumber
    Forms!wizfairdebt!cmdOK.Caption = "Sent to Atty"
    
Exit_cmdFairDebt_Click:
    Exit Sub

Err_cmdFairDebt_Click:
    MsgBox "No file was selected"
    Resume Exit_cmdFairDebt_Click
End Sub

Private Sub cmdFLMA_Click()
Dim FileNumber As Long
FileNumber = InputBox("Enter the File Number", "FLMA Wizard")
    AddToList (FileNumber)
    FLMACallFromQueue FileNumber
End Sub

Private Sub cmdHUDoccLetter_Click()
On Error GoTo Err_cmdFairDebt_Click
Dim FileNumber As Long
FileNumber = InputBox("Enter the File Number", "HUD Occ Letter Wizard")
If DLookup("loantype", "fcdetails", "current=yes and filenumber=" & FileNumber) <> 3 Then
MsgBox "This wizard can only be used for HUD files", vbCritical
Exit Sub
End If
    AddToList (FileNumber)
    HUDOccCallFromQueue FileNumber
    
Exit_cmdFairDebt_Click:
    Exit Sub

Err_cmdFairDebt_Click:
    MsgBox "No file was selected"
    Resume Exit_cmdFairDebt_Click
End Sub

Private Sub cmdIntake_Click()

On Error GoTo Err_cmdIntake_Click
DoCmd.OpenForm "wizIntake1"

Exit_cmdIntake_Click:
    Exit Sub

Err_cmdIntake_Click:
    MsgBox Err.Description
    Resume Exit_cmdIntake_Click
    
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

Private Sub cmdIntakeWiz_Click()
On Error GoTo Err_cmdRefSpecII_Click
Dim FileNumber As Long
Dim rstqueue As Recordset

    FileNumber = InputBox("Enter the File Number", "Intake Wizard")
    
    
    If Not IsNull(DLookup("File", "qryQueueIntake", "File=" & FileNumber)) Then
            MsgBox ("The file is already in Intake queue")
            Exit Sub
            End If
            
            Dim rs As Recordset
            Set rs = CurrentDb.OpenRecordset("Select * FROM qryqueueIntakeWaitinglst_P", dbOpenDynaset, dbSeeChanges)
            rs.Close
            Set rs = Nothing

            If Not IsNull(DLookup("File", "IntakeWaitingQueue", "File=" & FileNumber)) Then
            MsgBox ("The file is already in Intake Waiting queue ")
            Exit Sub
            End If

    AddToList (FileNumber)
    
    IntakeCallFromQueue FileNumber
    
    Forms!foreclosuredetails!cmdWizComplete.Visible = False
    
    Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current=true", dbOpenDynaset, dbSeeChanges)
    
    With rstqueue
    .Edit
    !IntakeLastEdited = Date
    !IntakeIuser = GetStaffID
    .Update
    End With
    
    Set rstqueue = Nothing
    
    
Exit_cmdRefSpecII_Click:
    Exit Sub

Err_cmdRefSpecII_Click:
    MsgBox "No file was selected"
    Resume Exit_cmdRefSpecII_Click
End Sub

Private Sub cmdRefSpecII_Click()
On Error GoTo Err_cmdRefSpecII_Click
Dim FileNumber As Long
    FileNumber = InputBox("Enter the File Number", "RSII Wizard")
    

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "wizReferralII"
    AddToList (FileNumber)
    RSIICallFromQueue FileNumber
    
Exit_cmdRefSpecII_Click:
    Exit Sub

Err_cmdRefSpecII_Click:
    MsgBox "No file was selected"
    Resume Exit_cmdRefSpecII_Click
    
End Sub
Private Sub cmdNOIwiz_Click()
On Error GoTo Err_cmdNOIwiz_Click
Dim FileNumber As Long
Dim rstqueue As Recordset
Dim rsfC As Recordset
Dim stDocName As String
Dim stLinkCriteria As String
Dim FileNO As Integer
    
    
FileNumber = InputBox("Enter the File Number", "NOI Wizard")
If IsNull(FileNumber) Then
Exit Sub
End If


   


        
        If DLookup("State", "FCdetails", "filenumber=" & FileNumber & " AND Current=yes") <> "MD" Then
        MsgBox "The file you wish to open in wizard is not a Maryland file.  Please double check the property address.", vbCritical
        Exit Sub
        End If


        Set rsfC = CurrentDb.OpenRecordset("Select * from FCdetails where filenumber = " & FileNumber & " and current=true ", dbOpenDynaset, dbSeeChanges)
        With rsfC
            If Not IsNull(rsfC!NOI) Then
            MsgBox "The File already has NOI date please use New bottomn to add it to the queue.", vbCritical
           
            Exit Sub
            End If
            
            If IsNull(rsfC!FairDebt) Then
            MsgBox "The File has no FairDebt date yet.", vbCritical
            Exit Sub
            End If
            Set rsfC = Nothing
        End With


            If Not IsNull(DLookup("File", "qryQueueNOI_P", "File=" & FileNumber)) Then
            MsgBox ("The file is already in NOI queue")
            Exit Sub
            End If

            If Not IsNull(DLookup("File", "qryQueueNOIWaiting_P", "File=" & FileNumber)) Then
            MsgBox ("The file is already in NOI Waiting queue ")
            Exit Sub
            End If









AddToList (FileNumber)
NOICallFromQueue FileNumber
    
   

'8/29/14

FileNO = FileNumber

Exit_cmdNOIwiz_Click:
    Exit Sub

Err_cmdNOIwiz_Click:
MsgBox Err.Description
    MsgBox "No file was selected"
    Resume Exit_cmdNOIwiz_Click
    
    End Sub
  Private Sub cmdNOIupload_Click()
DoCmd.OpenForm "queNOIupload"

  End Sub
    
Private Sub cmdRestart_Click()
On Error GoTo Err_cmdRestartwiz_Click
Dim FileNumber As Long
    Dim stDocName As String
    Dim stLinkCriteria As String
    
FileNumber = InputBox("Enter the File Number", "Restart Wizard")
If IsNull(FileNumber) Then
Exit Sub
End If
    AddToList (FileNumber)
    RestartCallFromQueue FileNumber
    Forms!foreclosuredetails!cmdWizComplete.Visible = False
    Forms!foreclosuredetails!cmdWaiting1.Visible = True


Exit_cmdRestartwiz_Click:
    Exit Sub

Err_cmdRestartwiz_Click:
    MsgBox "No file was selected"
    Resume Exit_cmdRestartwiz_Click
End Sub

Private Sub cmdSAI_Click()
Dim FileNumber As Long
   
    
FileNumber = InputBox("Enter the File Number", "SAI Wizard")
If IsNull(FileNumber) Then
Exit Sub
End If
AddToList (FileNumber)
SAICallFromQueue FileNumber
End Sub

Private Sub cmdSaleSetting_Click()
Dim FileNumber As Long
   
    
FileNumber = InputBox("Enter the File Number", "Sale Setting Wizard")
If IsNull(FileNumber) Then
Exit Sub
End If
AddToList (FileNumber)
SaleSettingCallFromQueue FileNumber
End Sub

Private Sub cmdService_Click()
Dim FileNumber As Long
    
    
FileNumber = InputBox("Enter the File Number", "Service Wizard")
If IsNull(FileNumber) Then
Exit Sub
End If
AddToList (FileNumber)
ServiceCallFromQueue FileNumber
End Sub

Private Sub cmdServiceMailed_Click()
Dim FileNumber As Long
    
    
FileNumber = InputBox("Enter the File Number", "Service Mailed Wizard")
If IsNull(FileNumber) Then
Exit Sub
End If

AddToList (FileNumber)
ServiceMailedCallFromQueue FileNumber
End Sub

Private Sub cmdVASaleSetting_Click()
Dim FileNumber As Long
   
    
FileNumber = InputBox("Enter the File Number", "VA Sale Setting Wizard")
If IsNull(FileNumber) Then
Exit Sub
End If

AddToList (FileNumber)
VAsalesettingCallFromQueue FileNumber
End Sub

Private Sub Command63_Click()
DoCmd.OpenForm "WizardAccounting"
End Sub

Private Sub ComTitleOut_Click()
Dim FileNumber As Long
       
FileNumber = InputBox("Enter the File Number", "Title Orderd Wizard")
If IsNull(FileNumber) Then
Exit Sub
End If
AddToList (FileNumber)
TitleOutCallFromQueue FileNumber


End Sub

Private Sub ComTitleReview_Click()
Dim FileNumber As Long
       
FileNumber = InputBox("Enter the File Number", "Title Review Wizard")
If IsNull(FileNumber) Then
Exit Sub
End If
AddToList (FileNumber)
TitleReviewCallFromQueue FileNumber

End Sub

Private Sub Form_Open(Cancel As Integer)
'Select Case GetLoginName()
'Case "Diane S. Rosenberg", "Chris Schwarz", "Sarab Alani", "Angela Henderson", "Lisa Famulare"
'TitleOrderW.Visible = True
'End Select

End Sub
Private Sub Command59_Click()
On Error GoTo Err_Command59_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command59_Click:
    Exit Sub

Err_Command59_Click:
    MsgBox Err.Description
    Resume Exit_Command59_Click
    
End Sub

Private Sub TitleOrderW_Click()

If Not TitleOrder Then
MsgBox ("You are not Authorized to Access this Queue")
Exit Sub
Else
Dim FileNumber As Long
       
FileNumber = InputBox("Enter the File Number", "Title Orderd Wizard")
If IsNull(FileNumber) Then
Exit Sub
End If
AddToList (FileNumber)
TitleOrderCallFromQueue FileNumber
End If
End Sub
