VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterPSAdvancedCostDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btn1_AfterUpdate()
If btn1 = True Then
btn2 = False
btn3 = False
btn5 = False
btn4 = False
End If
End Sub

Private Sub btn2_AfterUpdate()
If btn2 = True Then
btn1 = False
btn3 = False
btn5 = False
btn4 = False
End If
End Sub

Private Sub btn3_AfterUpdate()
If btn3 = True Then
btn1 = False
btn2 = False
btn5 = False
btn4 = False
End If
End Sub

Private Sub btn4_AfterUpdate()
If btn4 = True Then
'Other.Enabled = True
btn1 = False
btn2 = False
btn3 = False
btn5 = False
Else
'Other.Enabled = False
End If
End Sub

Private Sub btn5_AfterUpdate()
If btn5 = True Then
btn1 = False
btn2 = False
btn4 = False
btn3 = False
End If
End Sub





Private Sub cmdCancelLIt_Click()
'Dim JurText As String
'
'
'JurText = "PS Advanced costs Queue Cancelled"
'DoCmd.SetWarnings False
'
'DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & Forms![Case list]!FileNumber & "', #" & Now & "#,'" & GetFullName() & "','" & JurText & "',2 )")
'
'DoCmd.SetWarnings True



Call closeformL

End Sub

Private Sub cmdComplete_Click()
Dim ctr As Integer, rstwizqueue As Recordset, Other As String, rstdocs As Recordset, JrlTxt As String, rstFCdetailsCurrent As Recordset, rstFCdetailsPrior As Recordset
Dim s As Recordset
Dim lrs As Recordset
Dim t As Recordset
Dim Jrs As Recordset
Dim AFileNumber As Long
Dim shortClint As String

'DoCmd.Hourglass True
Dim rstqueue As Recordset
ctr = 0
If btn1 = True Then
ctr = ctr + 1
End If
If btn2 = True Then
ctr = ctr + 1
End If
If btn3 = True Then
ctr = ctr + 1
End If
If btn5 = True Then
ctr = ctr + 1
End If
If btn4 = True Then
ctr = ctr + 1
End If

If ctr = 0 Then
MsgBox "Please select a reason", vbCritical
Exit Sub
End If


If btn2 = True Then
    JrlTxt = InputBox("Please Add the reason for VOID")
    JrlTxt = Replace(JrlTxt, "'", "''")
    
   
    
    shortClint = ClientShortName(Forms![Case List]!ClientID)
    
    DoCmd.SetWarnings False
    If IsLoadedF("QueueAccounPSAdvancedCosts") = True Then
    
        If (Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6)) <> 0 Then
           DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Void = True ,Dismissed = True , MangerQ = True ,MangNotic ='" & JrlTxt & "'  WHERE DocIndexID = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6))
        Else
            DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Void = True ,Dismissed = True , MangerQ = True ,MangNotic ='" & JrlTxt & "' WHERE CaseFile = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(0))
        End If
        
        Forms!QueueAccounPSAdvancedCosts!lstFiles.Requery
        Forms!QueueAccounPSAdvancedCosts.Requery
        'Dim rstqueue As Recordset, cntr As Integer
        
        Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueuePSAdvancedCosts ")
        If Not rstqueue.EOF Then
            rstqueue.MoveLast
            Forms!QueueAccounPSAdvancedCosts!QueueCount = rstqueue!ct
        Else
            Forms!QueueAccounPSAdvancedCosts!QueueCount = 0
        End If
        rstqueue.Close
        Set rstqueue = Nothing
            
        
        
        
    Else
        DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Void = True ,Dismissed = True , MangerQ = True  WHERE DocIndexID=0 And CaseFile = " & Forms![Case List]!FileNumber)
    End If
    
    DoCmd.RunSQL ("Insert into ValumePSAdvanced (CaseFile,ClientName,Name,ToManagerQ,ToManagerQCount) Values('" & Forms![Case List]!FileNumber & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1)")
    DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & Forms![Case List]!FileNumber & "', #" & Now & "#,'" & GetFullName() & "','" & JrlTxt & "',2 )")
    
    DoCmd.SetWarnings True

End If



   

    
    If btn3 = True Then
        
    JrlTxt = InputBox("Please Add the reason for HOLD")
    JrlTxt = Replace(JrlTxt, "'", "''")
    
    
    shortClint = ClientShortName(Forms![Case List]!ClientID)
    
    DoCmd.SetWarnings False
    If IsLoadedF("QueueAccounPSAdvancedCosts") = True Then
    
        If (Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6)) <> 0 Then
        DoCmd.RunSQL ("UPDATE DocIndex set Hold = 'H' WHERE DocID = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6))
           DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set HOLD = 'H' ,Dismissed = True , MangerQ = True, Holddate = Now(), MangNotic ='" & JrlTxt & "' WHERE   DocIndexID = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6))
        Else
            DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set HOLD = 'H' ,Dismissed = True , MangerQ = True, Holddate = Now(),MangNotic ='" & JrlTxt & "' WHERE  CaseFile = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(0))
        End If
        
        Forms!QueueAccounPSAdvancedCosts!lstFiles.Requery
        Forms!QueueAccounPSAdvancedCosts.Requery
        'Dim rstqueue As Recordset, cntr As Integer
        
        Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueuePSAdvancedCosts ")
        If Not rstqueue.EOF Then
            rstqueue.MoveLast
            Forms!QueueAccounPSAdvancedCosts!QueueCount = rstqueue!ct
        Else
            Forms!QueueAccounPSAdvancedCosts!QueueCount = 0
        End If
        rstqueue.Close
        Set rstqueue = Nothing
            
        
        
        
    Else
        DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set HOLD = 'H' ,Dismissed = True , MangerQ = True,MangNotic ='" & JrlTxt & "' WHERE DocIndexID=0 And CaseFile = " & Forms![Case List]!FileNumber)
    End If
    
    DoCmd.RunSQL ("Insert into ValumePSAdvanced (CaseFile,ClientName,Name,ToManagerQ,ToManagerQCount) Values('" & Forms![Case List]!FileNumber & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1)")
    DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & Forms![Case List]!FileNumber & "', #" & Now & "#,'" & GetFullName() & "','" & JrlTxt & "',2 )")
    
    DoCmd.SetWarnings True

   End If
'
If btn5 = True Then

 JrlTxt = InputBox("Please Add the reason for Offset")
    JrlTxt = Replace(JrlTxt, "'", "''")
    
    
    shortClint = ClientShortName(Forms![Case List]!ClientID)
    
    DoCmd.SetWarnings False
    If IsLoadedF("QueueAccounPSAdvancedCosts") = True Then
    
        If (Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6)) <> 0 Then
       
           DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Offset = True  ,Dismissed = True , MangerQ = True,  MangNotic ='" & JrlTxt & "' WHERE   DocIndexID = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6))
        Else
            DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Offset =True ,Dismissed = True , MangerQ = True, MangNotic ='" & JrlTxt & "'WHERE CaseFile = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(0))
        End If
        
        Forms!QueueAccounPSAdvancedCosts!lstFiles.Requery
        Forms!QueueAccounPSAdvancedCosts.Requery
        'Dim rstqueue As Recordset, cntr As Integer
        
        Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueuePSAdvancedCosts ")
        If Not rstqueue.EOF Then
            rstqueue.MoveLast
            Forms!QueueAccounPSAdvancedCosts!QueueCount = rstqueue!ct
        Else
            Forms!QueueAccounPSAdvancedCosts!QueueCount = 0
        End If
        rstqueue.Close
        Set rstqueue = Nothing
            
        
        
        
    Else
        DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Offset = True ,Dismissed = True , MangNotic ='" & JrlTxt & "' WHERE DocIndexID=0 And CaseFile = " & Forms![Case List]!FileNumber)
    End If
    
    DoCmd.RunSQL ("Insert into ValumePSAdvanced (CaseFile,ClientName,Name,ToManagerQ,ToManagerQCount) Values('" & Forms![Case List]!FileNumber & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1)")
    DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & Forms![Case List]!FileNumber & "', #" & Now & "#,'" & GetFullName() & "','" & JrlTxt & "',2 )")
    
    DoCmd.SetWarnings True

End If
If btn4 = True Then
JrlTxt = InputBox("Please Add the reason for Other ")
    JrlTxt = Replace(JrlTxt, "'", "''")
    
    
    shortClint = ClientShortName(Forms![Case List]!ClientID)
    
    DoCmd.SetWarnings False
    If IsLoadedF("QueueAccounPSAdvancedCosts") = True Then
    
        If (Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6)) <> 0 Then
       
           DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Other = True  ,Dismissed = True , MangerQ = True,  MangNotic ='" & JrlTxt & "' WHERE   DocIndexID = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6))
        Else
            DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Other = True  ,Dismissed = True , MangerQ = True, MangNotic ='" & JrlTxt & "' CaseFile = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(0))
        End If
        
        Forms!QueueAccounPSAdvancedCosts!lstFiles.Requery
        Forms!QueueAccounPSAdvancedCosts.Requery
        'Dim rstqueue As Recordset, cntr As Integer
        
        Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueuePSAdvancedCosts ")
        If Not rstqueue.EOF Then
            rstqueue.MoveLast
            Forms!QueueAccounPSAdvancedCosts!QueueCount = rstqueue!ct
        Else
            Forms!QueueAccounPSAdvancedCosts!QueueCount = 0
        End If
        rstqueue.Close
        Set rstqueue = Nothing
            
        
        
        
    Else
        DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Other = True ,Dismissed = True , MangNotic ='" & JrlTxt & "' WHERE DocIndexID=0 And CaseFile = " & Forms![Case List]!FileNumber)
    End If
    
    DoCmd.RunSQL ("Insert into ValumePSAdvanced (CaseFile,ClientName,Name,ToManagerQ,ToManagerQCount) Values('" & Forms![Case List]!FileNumber & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1)")
    DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & Forms![Case List]!FileNumber & "', #" & Now & "#,'" & GetFullName() & "','" & JrlTxt & "',2 )")
    
    DoCmd.SetWarnings True

End If


Call closeformL

End Sub

Private Sub cmdLitiCompOption_Click()

DoCmd.OpenForm "DocsWindowPSAdvanced", , , "FileNumber = " & Forms![Case List]!FileNumber
Me.ComDone.Visible = True

End Sub

Private Sub ComDone_Click()
Dim JurText As String
Dim shortClint As String
Dim FileNu As Long
FileNu = Forms![Case List]!FileNumber

shortClint = ClientShortName(Forms![Case List]!ClientID)


    If Forms![Case List]!ChOffset Then
    JurText = "PS Advanced Costs Package and File SetOff sent to Manager queue"
    Else
    JurText = "PS Advanced Costs Package"
    End If



DoCmd.SetWarnings False

        If IsLoadedF("QueueAccounPSAdvancedCosts") = True Then
               
                    If (Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6)) <> 0 Then
                        DoCmd.RunSQL ("UPDATE DocIndex set Hold = '' WHERE DocID = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6))
                     
                                If Forms![Case List]!ChOffset Then
                                 DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Offset = True, Hold = '',Dismissed = True , MangNotic = 'File Offset', MangerQ = true   WHERE DocIndexID = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6) & " And Dismissed = False ")
                                Else
                                 DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Hold = '',Dismissed = True , MangerQ = False  WHERE DocIndexID = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(6) & " And Dismissed = False")
                                End If
                        
                    Else
                        If Forms![Case List]!ChOffset Then
                        DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Offset = True, Hold = '',Dismissed = True , MangNotic = 'File Offset',MangerQ = True  WHERE CaseFile = " & Forms!QueueAccounLitigationBill.lstFiles.Column(0) & " And Dismissed = false")
                        Else
                        DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Hold = '',Dismissed = True , MangerQ = False  WHERE CaseFile = " & Forms!QueueAccounPSAdvancedCosts.lstFiles.Column(0) & " And Dismissed = False")
                        End If
                        
                    End If
            
                
                Forms!QueueAccounPSAdvancedCosts!lstFiles.Requery
                Forms!QueueAccounPSAdvancedCosts.Requery
                Dim rstqueue As Recordset, cntr As Integer
                
                Set rstqueue = CurrentDb.OpenRecordset("select count(*)as ct from QueuePSAdvancedCosts ")
                        If Not rstqueue.EOF Then
                            rstqueue.MoveLast
                            Forms!QueueAccounPSAdvancedCosts!QueueCount = rstqueue!ct
                        Else
                            Forms!QueueAccounPSAdvancedCosts!QueueCount = 0
                        End If
                rstqueue.Close
                Set rstqueue = Nothing
                
            
            
            
        Else
            If Forms![Case List]!ChOffset Then
            DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Offset = True, Hold = '',Dismissed = True , MangNotic = 'File Offset',MangerQ = True  WHERE Not IsNull(DocIndexID) And CaseFile = " & FileNu & " And Dismissed = false")
            Else
            DoCmd.RunSQL ("UPDATE Accou_PSAdvancedCostsPackageQueue set Hold = '',Dismissed = True , MangerQ = False  WHERE Not IsNull(DocIndexID) And CaseFile = " & FileNu & " And Dismissed = false ")
            End If
            
        
      End If
      


DoCmd.RunSQL ("Insert into ValumePSAdvanced (CaseFile,ClientName,Name,CompleteBill,CompleteBillCount) Values('" & FileNu & "','" & shortClint & "','" & GetFullName() & "', #" & Now() & "# ,1)")
DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & FileNu & "', #" & Now & "#,'" & GetFullName() & "','" & JurText & "',2 )")

DoCmd.SetWarnings True



Call closeformL






End Sub

Private Sub Command22_Click()
'btn1.Visible = True
btn2.Visible = True
btn3.Visible = True
btn4.Visible = True
'btn5.Visible = True
'Other.Visible = True
'Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
cmdComplete.Visible = True

    
End Sub

Private Sub closeformL()
Dim F As Form
Dim FormClosed As Boolean

Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "QueueAccounPSAdvancedCosts", "QueueAccountPSAdvancedCostManager" '  leave these forms open
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed

End Sub

