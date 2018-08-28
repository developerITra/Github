VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Payoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim PrintTo As Integer, ContactType As String


Private Sub chSale_AfterUpdate()
If chSale Then GoodThru = DateAdd("d", -1, Sale)
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
Dim statusMsg As String, rptType As String
Dim bIsWizard As Boolean
Dim strInsert As String
Dim clientShor As String
Dim StrJuirs As String



bIsWizard = IsLoaded("wizdemand")
On Error GoTo Err_cmdOK_Click
If IsNull(Me!optDocType) Then
    MsgBox "Choose document type", vbCritical
    Exit Sub
End If

Select Case Me!optDocType

    Case 1
    
   
    
    
                If [Forms]![Case List]![ClientID] = 97 Then
                    If [Forms]![foreclosuredetails]![State] = "VA" Then
                    If CheckAccruedInterest(FileNumber) = True Then Exit Sub
                    rptType = "PayOffJPVA"
                    Else
                    If CheckAccruedInterest(FileNumber) = True Then Exit Sub
                    rptType = "PayOffJP"
                    End If
                        
                Else
                rptType = "Payoff"
                End If
                
                statusMsg = "Sent Payoff figures"
                
                If IsNull(PayoffRequested) Then
                    MsgBox ("Please Add Payoff Requested date")
                    Exit Sub
                Else
                If MsgBox("Update Payoff Sent Date = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
                     Forms!foreclosuredetails!PayoffSent = Now()
                     StrJuirs = DLookup("Jurisdiction", "JurisdictionList", "JurisdictionID= " & Forms![Case List]!JurisdictionID)
                     StrJuirs = Replace(StrJuirs, "'", "''")
                     clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)
                     clientShor = Replace(clientShor, "'", "''")
                     DoCmd.SetWarnings False
                     strInsert = "Insert Into Tracking_PayoffSent (CaseFile,ProjectName,ClientShortName,Juris,Client,DIT, PayoffRequested ,StaffID,StaffName) Values (" & FileNumber & ",'" & Forms![Case List]!PrimaryDefName & "','" & clientShor & "','" & StrJuirs & "', " & Forms![Case List]!ClientID & ", #" & Now() & "#,#" & Forms![foreclosuredetails]!PayoffRequested & "#, " & GetStaffID & ",'" & GetFullName() & "')"
                     DoCmd.RunSQL strInsert
                     DoCmd.SetWarnings True
                    End If
                End If
                
                
                Call DoReport(rptType, PrintTo, , ContactType)   'changes on 07/20/14 SA
                cmdCancel.Caption = "Close"
    
    
   
    
       
            
        
    
    
    
    
 
    
    Case 2
        statusMsg = "Sent Reinstatement figures"
        If [Forms]![Case List]![ClientID] = 97 Then
            If [Forms]![foreclosuredetails]![State] = "VA" Then
             'If CheckAccruedInterest(FileNumber) = True Then Exit Sub
            rptType = "PayOffJPRIVA"
            Else
            rptType = "PayOffJPRI"
            ' If CheckAccruedInterest(FileNumber) = True Then Exit Sub
            End If
                
        Else

        rptType = "Payoff"
        End If
        
        
        If IsNull(ReinstatementRequested) Then
            MsgBox ("Please Add Reinstatement Requeested date")
            Exit Sub
        Else
                If MsgBox("Update Reinstatment Sent Date = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
                        Forms!foreclosuredetails!ReinstatementSent = Now()
                        StrJuirs = DLookup("Jurisdiction", "JurisdictionList", "JurisdictionID= " & Forms![Case List]!JurisdictionID)
                        StrJuirs = Replace(StrJuirs, "'", "''")
                        clientShor = DLookup("ShortClientName", "ClientList", "ClientID= " & Forms![Case List]!ClientID)
                        clientShor = Replace(clientShor, "'", "''")
                        DoCmd.SetWarnings False
                        strInsert = "Insert Into Tracking_ReinstatementSent (CaseFile,ProjectName,ClientShortName,Juris,Client,ReinstRequested,DIT,StaffID,StaffName) Values (" & FileNumber & ",'" & Forms![Case List]!PrimaryDefName & "','" & clientShor & "','" & StrJuirs & "', " & Forms![Case List]!ClientID & ",#" & Forms![foreclosuredetails]!ReinstatementRequested & "#, #" & Now() & "#," & GetStaffID & ",'" & GetFullName() & "')"
                        DoCmd.RunSQL strInsert
                        DoCmd.SetWarnings True
                             
                End If
        End If
        
        
        Call DoReport(rptType, PrintTo, , ContactType)   'changes on 07/20/14 SA
        cmdCancel.Caption = "Close"
        
'        If MsgBox("Add to status: " & Format$(Date, "mm/dd/yyyy") & " " & statusMsg, vbYesNo + vbQuestion) = vbYes _
'        Then
'            AddStatus Me!FileNumber, Date, statusMsg
'        End If
    Case 3
        
        Dim ptntext As String
        Dim ptntext2 As String
'        ptntext2 = ""
'        ptntext = ""
        Dim ptncounter As Integer
        
        If ptn1 Then
        ptncounter = 1 + ptncounter
        ptntext = "make the monthly payments"
        ptntext2 = "the loan is brought completely current"
'        Else
'        ptntext = ""
'        ptntext2 = ""
        End If
           
        If ptn2 Then
        ptncounter = 1 + ptncounter
        If ptncounter > 1 Then
        ptntext = ptntext + " and " + "pay real property taxes and/or homeowner's insurnce"
        ptntext2 = ptntext2 + " and " + "the real property taxes and/or homeowner's insurance is paid and proof is provided"
        Else
        ptntext = "pay real property taxes and/or homeowner's insurnce"
        ptntext2 = "the real property taxes and/or homeowner's insurance is paid and proof is provided"
        End If
'        Else
'        ptntext = ""
'        ptntext = ""
        End If

        If ptn3 Then
        ptncounter = 1 + ptncounter
        If ptncounter > 1 Then
        ptntext = ptntext & " and " & "maintain the property as your principal residence"
        ptntext2 = ptntext2 & " and " & "you move back into the property and deem it your principal residence"
        Else
        ptntext = "maintain the property as your principal residence"
        ptntext2 = "you move back into the property and deem it your principal residence"
        End If
'        Else
'        ptntext = ""
'        ptntext2 = ""
        End If
'
'
        If ptn4 Then
        ptncounter = 1 + ptncounter
        If ptncounter > 1 Then
        ptntext = ptntext & " and " & Forms![Print Payoff].txtDefault
        ptntext2 = ptntext2 & " and " & Forms![Print Payoff].txtCure
        Else
        ptntext = Forms![Print Payoff].txtDefault
        ptntext2 = Forms![Print Payoff].txtCure
        End If
'        Else
'        ptntext = ""
'        ptntext = ""
        End If


        textdefalut = ptntext
        textcure = ptntext2

              
         DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
        
               
        
        statusMsg = "Acceleration Payoff Due " & Format$(Me!GoodThru, "mm/dd/yyyy")
        rptType = "Acceleration Letter Wiz"
        
        
        Call DoReport(rptType, PrintTo, , ContactType)
                Dim OncePDF As Boolean
        OncePDF = True
CheckStatus:  If (SysCmd(acSysCmdGetObjectState, acReport, rptType) = 0) And OncePDF = True Then   ' SA 07/19/14
                 If MsgBox(" Is this the correct One for PDF format?", vbYesNo) = vbYes Then
                 
            
                   ' DoCmd.Close acReport, "Acceleration Letter Wiz"
                        
                    PrintTo = -2
                                                          
         
        '        If MsgBox("Add to status: " & Format$(Date, "mm/dd/yyyy") & " " & statusMsg, vbYesNo + vbQuestion) = vbYes _ ' AS per Diane request on 06/26/2013 sa
        '        Then
        '            AddStatus Me!FileNumber, Date, statusMsg
        '            Forms!wizdemand!AccelerationLetter = Me!GoodThru
        '            Forms!wizdemand!AccelerationIssued = Date
        '        End If
                
'                Dim ClientID As Integer, qtypstge As Integer, LoanType As Integer
'                LoanType = Forms!wizDemand!LoanType
'                ClientID = DLookup("clientid", "caselist", "filenumber=" & FileNumber)
'                    Select Case LoanType
'                        Case 4
'                        FeeAmount = Nz(DLookup("FeeAcceleration", "ClientList", "ClientID=177"))
'                        Case 5
'                        FeeAmount = Nz(DLookup("FeeAcceleration", "ClientList", "ClientID=263"))
'                        Case Else
'                        FeeAmount = Nz(DLookup("FeeAcceleration", "ClientList", "ClientID=" & ClientID))
'                    End Select
'
'                'Discretionary invoicing for demand letters (ability to override)
'
'                'If MsgBox("Do you want to override the standard fee of $" & FeeAmount & " for this client?", vbYesNo) = vbYes Then ' as per Diane request on 06/26 sA
'                'FeeAmount = InputBox("Please enter fee, then rememeber to note the journal")
'                'MsgBox "Please upload fee approval to documents"
'                'End If
'
'                If FeeAmount > 0 Then
'                    AddInvoiceItem FileNumber, "FC-Acc", "Acceleration Letter", FeeAmount, 0, True, True, False, False
'                Else
'                    AddInvoiceItem FileNumber, "FC-Acc", "Acceleration Letter", 1, 0, True, True, False, False 'set unknown fee as $1, per Diane
'                End If
'
'        '        FeeAmount = Nz(DLookup("AccelerationPostage", "ClientList", "ClientID=" & ClientID))
'        '        If FeeAmount > 0 Then
'        '        qtyPstge = DCount("[FileNumber]", "[qryFairDebt]", "FileNumber=" & [FileNumber])
'        '        AddInvoiceItem FileNumber, "FC-Acceleration", "Acceleration Letter mailed", (qtyPstge * FeeAmount), 76, False, True, False, True
'        '        Else
'        '        AddInvoiceItem FileNumber, "FC-Acceleration", "Acceleration Letter mailed", 1, 76, False, True, False, True
'        '        End If
'                Dim fairdebtCnt As Integer
'                fairdebtCnt = DCount("[ID]", "[Names]", "FileNumber = " & [FileNumber] & " and FairDebt = true")
'                 If (fairdebtCnt > 0) Then
'                    FeeAmount = Nz(DLookup("FairDebtPostage", "ClientList", "ClientID=" & ClientID), 0)
'                    If FeeAmount > 0 Then
'                    qtypstge = DCount("[FileNumber]", "[qryFairDebt]", "FileNumber=" & [FileNumber])
'                    AddInvoiceItem FileNumber, "FC-ACC", "Acceleration Letter Postage", (qtypstge * FeeAmount), 76, False, False, False, True
'                    Else
'                    AddInvoiceItem FileNumber, "FC-ACC", "Acceleration Letter Postage", 1, 76, False, False, False, True
'                    End If
'                 End If
'
                
                OncePDF = False
            
                Call DoReport(rptType, PrintTo, , ContactType)
                If MsgBox("Do you want to print a hard copy?", vbYesNo) = vbYes Then  ' changes on 08/04/2014 just to add 2 print copy (as floyed asked)
                    Call DoReport(rptType, acViewNormal, , ContactType)
                    Call DoReport(rptType, acViewNormal, , ContactType)
                End If
                
             '   Forms!WizDemand.cmdOKd.Visible = True
                cmdCancel.Caption = "Close"
              '  DoCmd.Close acForm, "wizdemand"
                 
                 If IsLoaded("Print Payoff") Then DoCmd.Close acForm, "Print Payoff"
                    Exit Sub
                  End If
                
             
                 
            Else
                    If OncePDF = True Then
                    Wait 2
                    GoTo CheckStatus
                    Else
                    Exit Sub
                    End If
            End If
                    
                            
            
            


End Select



'Call DoReport(rptType, PrintTo, , ContactType)
'
'cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdClear_Click()

On Error GoTo Err_cmdClear_Click

DoCmd.RunSQL "DELETE * FROM Payoff WHERE FileNumber=" & Forms!foreclosuredetails!FileNumber
sfrmPayoff.Requery

Exit_cmdClear_Click:
    Exit Sub

Err_cmdClear_Click:
    MsgBox Err.Description
    Resume Exit_cmdClear_Click
    
End Sub





Private Sub DemDate_DblClick(Cancel As Integer)
DemDate = Date

End Sub

Private Sub DutDate_DblClick(Cancel As Integer)
DutDate = Date
End Sub

Private Sub Form_Current()
PrintTo = Int(Split(Me.OpenArgs, "|")(0))
ContactType = Split(Me.OpenArgs, "|")(1)

'Option64.Enabled = (ContactType = "FC") ' allow acceleration letter for Foreclosure only

If IsLoaded("wizdemand") = True Then
sfrmPayoff.Visible = False
cmdClear.Visible = False
ptn1.Visible = True
ptn1t.Visible = True
ptn2.Visible = True
ptn2t.Visible = True
ptn3.Visible = True
ptn3t.Visible = True
ptn4.Visible = True
ptn4t.Visible = True
labOptionD.Visible = True
ptn1 = 1
chSale.Visible = False
Text77.Visible = False
GoodThru.Visible = False
DutDate.Visible = False
Goodthroughdate.Visible = True
boxptns.Visible = True
[To].Visible = False
Dear.Visible = False


End If

End Sub


Private Sub Goodthroughdate_AfterUpdate()
If Goodthroughdate.Value = 2 Then
DemDate.Visible = True
Else
DemDate.Visible = False
End If

End Sub

Private Sub GoodThru_DblClick(Cancel As Integer)
GoodThru = Date
End Sub

Private Sub optDocType_AfterUpdate()
If Me!optDocType = 3 Then   ' acceleration
    Me!To.Enabled = False
    Me!Dear.Enabled = False
    Me!GoodThru = DateAdd("d", 30, Date)
    Else
    If Me!optDocType = 1 Then
    Me!Text77.Locked = False
    Me!Text77.Enabled = True
    Me!To.Enabled = True
    Me!Dear.Enabled = True
    Me!GoodThru = Date
    Else
    Me!To.Enabled = True
    Me!Dear.Enabled = True
    Me!GoodThru = Date
    End If
    
End If
    
End Sub



Private Sub ptn4_AfterUpdate()
If ptn4 Then
txtDefault.Visible = True
txtCure.Visible = True
Else
txtDefault.Visible = False
txtCure.Visible = False
End If

End Sub

Private Sub ptn4_Click()
If ptn4 = True Then
txtDefault.Visible = True
txtCure.Visible = True
Else
txtDefault.Visible = False
txtCure.Visible = False

End If

End Sub

Private Function CheckAccruedInterest(F As Long) As Boolean

If IsNull(DLookup("amount", "PayOff", "FileNumber = " & F & " AND Desc like 'Accrued Interest'+'%'")) Then
            MsgBox ("Please Add the Accrued Interest amount")
            CheckAccruedInterest = True
            
            End If
End Function
