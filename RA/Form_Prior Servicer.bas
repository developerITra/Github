VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Prior Servicer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub ChPrior_Click()

'Mei 9/24/2015
If ChPrior.Value Then   'if it's checked
    TxtPriorServicer.Enabled = True
Else
    TxtPriorServicer.Enabled = False
End If


End Sub

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click

If Forms!Foreclosureprint!chComplianceAffidavit = True Then

        If TxtPriorServicer.Enabled Then
            TxtPriorServicer.SetFocus
            If Len(TxtPriorServicer) > 0 Then
                strPriorServicer = TxtPriorServicer.Text
            End If
        End If
        
        If Me.ChHolder = True Then
            bHolder = True
        End If
        
        If Me.ChTransferee = True Then
            bReferee = True
        End If
        
        If Me.ChLost = True Then
            bLost = True
        End If
        
        If Me.ChPrior = True Then
            bPrior = True
        End If
         
        
     '   Text13.SetFocus
        Call AddStatmentDebt2([Forms]![Foreclosureprint]![FileNumber], Me.OpenArgs)
        
        If bReferee = False And bLost = False And bHolder = False Then
            MsgBox "Please check  Holder, Transferee in Possession, or Lost Note"
        Else
             DoCmd.OpenForm "Print Statement of Debt", , , "ForeclosureID=" & Forms!foreclosuredetails!ForeclosureID, , , Me.OpenArgs
            DoCmd.Close acForm, "Prior Servicer"
           
        End If
Else
 
 If Forms!Foreclosureprint!ChAffiOfLienInstru = True And Forms!foreclosuredetails!State = "MD" And Forms![Case List]!ClientID = 385 Then
 
        If TxtPriorServicer.Enabled Then
            TxtPriorServicer.SetFocus
            If Len(TxtPriorServicer) > 0 Then
                strPriorServicer = TxtPriorServicer.Text
            End If
        End If
          
'  Call DoReport("Nationstar Cover Sheet", Me.OpenArgs)
 Call DoReport("Affidavit of Lien Instrument NationStar", Me.OpenArgs)
 Call DoReport("Nationstar Cover Sheet", Me.OpenArgs)
 
 If Application.CurrentProject.AllForms("Prior Servicer").IsLoaded Then
    DoCmd.Close acForm, "Prior Servicer"
 End If
  
 End If
 
 If Forms!Foreclosureprint!chLossMitPrelim = True And Forms!foreclosuredetails!State = "MD" And Forms![Case List]!ClientID = 385 Then
 
        If TxtPriorServicer.Enabled Then
            TxtPriorServicer.SetFocus
            If Len(TxtPriorServicer) > 0 Then
                strPriorServicer = TxtPriorServicer.Text
            End If
        End If
    Call DoReport("Loss Mitigation Preliminary Nation Star", Me.OpenArgs)
    Call DoReport("Nationstar Cover Sheet", Me.OpenArgs)
    DoCmd.Close acForm, "Prior Servicer"

 End If
 
 
 If Forms!Foreclosureprint!chLossMitFinal = True And Forms!foreclosuredetails!State = "MD" And Forms![Case List]!ClientID = 385 Then
 
        If TxtPriorServicer.Enabled Then
            TxtPriorServicer.SetFocus
            If Len(TxtPriorServicer) > 0 Then
                strPriorServicer = TxtPriorServicer.Text
            End If
        End If
    Call DoReport("Loss Mitigation Final Nation Star", Me.OpenArgs)
    Call DoReport("Nationstar Cover Sheet", Me.OpenArgs)
    
    DoCmd.Close acForm, "Prior Servicer"
 End If
 
End If

 
 


Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub Command3_Click()
Me.Undo
DoCmd.Close

End Sub

Private Sub Form_Current()
If Forms!Foreclosureprint!chLossMitPrelim = True And Forms!foreclosuredetails!State = "MD" And Forms![Case List]!ClientID = 385 Then
 ChTransferee.Enabled = False
 ChHolder.Enabled = False
 ChLost.Enabled = False
End If

If Forms!Foreclosureprint!chLossMitFinal = True And Forms!foreclosuredetails!State = "MD" And Forms![Case List]!ClientID = 385 Then
 ChTransferee.Enabled = False
 ChHolder.Enabled = False
 ChLost.Enabled = False
End If

 If Forms!Foreclosureprint!ChAffiOfLienInstru = True And Forms!foreclosuredetails!State = "MD" And Forms![Case List]!ClientID = 385 Then
    ChTransferee.Enabled = False
    ChHolder.Enabled = False
    ChLost.Enabled = False
 End If
End Sub

Private Sub Form_Load()

'reset the globla variable
strPriorServicer = ""
bHolder = False
bReferee = False
bLost = False
bPrior = False

End Sub
Private Sub AddStatmentDebt2(FileNumber As Long, PrintTo As Integer)
 Dim rstNS As Recordset
 Dim Desc1, desc2, desc3, desc4, desc5, desc6, desc7, desc8, desc9, desc10, desc11, desc12, desc13 As String
            Set rstNS = CurrentDb.OpenRecordset("Select * from statementofDebt where Filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
    If rstNS.EOF Then
 
'    Desc1 = "Interest Amount     "
    Desc1 = "Total Late Charges"
    desc2 = "Property Tax Advances"
    desc3 = "Hazard Insurance Advances"
    desc4 = "MIP\PMI"
    desc5 = "Total Property Inspection Fees"
    desc6 = "Property Preservation Expenses"
    desc7 = "Prior Foreclosure Fees"
    desc8 = "Escrow Balance Credit"
    desc9 = "Credits"
    desc10 = "Other"
    'desc13 = "Payment Advance - Principal/Interest/Escrow"
    

    With rstNS
        .AddNew
        !FileNumber = FileNumber
        !Desc = Desc1
        !Amount = 0
        !Timestamp = Now
        !Sort_Desc = 1
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = desc2
        !Amount = 0
        !Timestamp = Now
        !Sort_Desc = 2
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
        !Desc = desc4
        !Sort_Desc = 4
        !Amount = 0
        !Timestamp = Now
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = desc5
        !Amount = 0
        !Timestamp = Now
        !Sort_Desc = 5
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
        !Desc = desc7
        !Amount = 0
        !Sort_Desc = 7
        !Timestamp = Now
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = desc8
        !Sort_Desc = 8
        !Amount = 0
        !Timestamp = Now
        .Update
              .AddNew
        !FileNumber = FileNumber
        !Desc = desc9
        !Sort_Desc = 9
        !Amount = 0
        !Timestamp = Now
        .Update
        .AddNew
        !FileNumber = FileNumber
        !Desc = desc10
        !Amount = 0
        !Timestamp = Now
        !Sort_Desc = 10
        .Update
'        .AddNew
'        !FileNumber = FileNumber
'        !Desc = desc11
'        !Amount = 0
'        !Timestamp = Now
'        !Sort_Desc = 11
'        .Update
'        .AddNew
'        !FileNumber = FileNumber
'        !Desc = desc12
'        !Amount = 0
'        !Sort_Desc = 12
'        !Timestamp = Now
'        .Update
'        .AddNew
'        !FileNumber = FileNumber
'        !Desc = desc13
'        !Sort_Desc = 13
'        !Amount = 0
'        !Timestamp = Now
'        .Update
    End With
End If
rstNS.Close
End Sub
