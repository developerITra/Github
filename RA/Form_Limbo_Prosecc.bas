VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Limbo_Prosecc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdCancel_Click()
Me.Undo
DoCmd.Close
End Sub

Private Sub cmdOK_Click()

If Nz([JournalNote], "") = "" Then
MsgBox ("Please Add Journal Note")
Exit Sub
End If



Dim rstwizqueue As Recordset
Dim strinfo As String
Dim rstsql As String
Dim JrlTxt As String
Dim strSQLJournal As String
 Dim rstqueue As Recordset, cntr As Integer
 cntr = 0

DoCmd.SetWarnings False

    Select Case Forms!foreclosuredetails!WizardSource

    'Case "Limbo_MDWhite", "Limbo_MDYellow", "Limbo_MDRed" ', "Limbo_VAWhite", "Limbo_VAYellow", "Limbo_VARed", "Limbo_DCWhite", "Limbo_DCYellow", "Limbo_DCRed"

    Case "Limbo_MDWhite"
        JrlTxt = ""
        
        If Forms!Limbo_MD!lstFiles.Column(6) = "X" Then
        JrlTxt = "Missing SOT"
        Else
        JrlTxt = ""
        End If
       
        If Forms!Limbo_MD!lstFiles.Column(7) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing SOD"
            Else
            JrlTxt = JrlTxt & ", SOD"
            End If
        End If
        
    
        If Forms!Limbo_MD!lstFiles.Column(8) = "X" And Forms!Limbo_MD!lstFiles.Column(9) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing LMA"
            Else
            JrlTxt = JrlTxt & ", LMA"
            End If
        End If
    
'        If Forms!Limbo_MD!lstFiles.Column(9) = "X" Then  As Diane email on 11/10
'            If JrlTxt = "" Then
'            JrlTxt = "Missing FLMA"
'            Else
'            JrlTxt = JrlTxt & ", FLMA"
'            End If
'         End If
   
          
        If Forms!Limbo_MD!lstFiles.Column(10) = "X" Then
          If JrlTxt = "" Then
          JrlTxt = "Missing ANO"
          Else
          JrlTxt = JrlTxt & ", ANO"
          End If
        End If
    
        If Forms!Limbo_MD!lstFiles.Column(11) = "X" Then
         If JrlTxt = "" Then
         JrlTxt = "Missing NOI"
         Else
         JrlTxt = JrlTxt & ", NOI"
         End If
        End If
        
'        If Forms!Limbo_MD!lstFiles.Column(12) = "X" Then Stopped as Diane email on 11-10
'            If JrlTxt = "" Then
'            JrlTxt = "Missing LNA"
'            Else
'            JrlTxt = JrlTxt & ", LNA"
'            End If
'        End If
        
     rstsql = "Insert InTo ValumeLimbo (CaseFile, Client, Name, WhiteComplete, WhiteCompleteC, state ) Values ( Forms!Limbo_MD!lstFiles " & ",'" & Forms!Limbo_MD.lstFiles.Column(1) & "', '" & GetFullName() & "',Now(),1 " & ",'" & Forms!foreclosuredetails!State & "')"
     DoCmd.RunSQL rstsql
     rstsql = ""
     
     rstsql = "UPDATE WizardSupportTwo SET LimboWhit = Now() WHERE FileNumber = " & Forms!Limbo_MD!lstFiles & " AND current = true "
     DoCmd.RunSQL rstsql
     rstsql = ""
     
        
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM LimboMDQueue", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        Forms!Limbo_MD.QueueCount = cntr
        Set rstqueue = Nothing
        Forms!Limbo_MD.lstFiles.Requery
     
     
    
Case "Limbo_MDYellow"
        JrlTxt = ""
        If Forms!Limbo_MD!lstFilesY.Column(6) = "X" Then
        JrlTxt = "Missing SOT"
        Else
        JrlTxt = ""
        End If
    
    
        If Forms!Limbo_MD!lstFilesY.Column(7) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing SOD"
            Else
            JrlTxt = JrlTxt & ", SOD"
            End If
        End If
        
    
        If Forms!Limbo_MD!lstFilesY.Column(8) = "X" And Forms!Limbo_MD!lstFilesY.Column(9) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing LMA"
            Else
            JrlTxt = JrlTxt & ", LMA"
            End If
        End If
    
'        If Forms!Limbo_MD!lstFilesY.Column(9) = "X" Then
'            If JrlTxt = "" Then
'            JrlTxt = "Missing FLMA"
'            Else
'            JrlTxt = JrlTxt & ", FLMA"
'            End If
'         End If
   
          
        If Forms!Limbo_MD!lstFilesY.Column(10) = "X" Then
          If JrlTxt = "" Then
          JrlTxt = "Missing ANO"
          Else
          JrlTxt = JrlTxt & ", ANO"
          End If
        End If
    
        If Forms!Limbo_MD!lstFilesY.Column(11) = "X" Then
         If JrlTxt = "" Then
         JrlTxt = "Missing NOI"
         Else
         JrlTxt = JrlTxt & ", NOI"
         End If
        End If
        
'        If Forms!Limbo_MD!lstFilesY.Column(12) = "X" Then
'            If JrlTxt = "" Then
'            JrlTxt = "Missing LNA"
'            Else
'            JrlTxt = JrlTxt & ", LNA"
'            End If
'        End If
        
     rstsql = "Insert InTo ValumeLimbo (CaseFile, Client, Name, YellowComplete, YellowCompleteC, state ) Values ( Forms!Limbo_MD!lstFilesY " & ",'" & Forms!Limbo_MD.lstFilesY.Column(1) & "', '" & GetFullName() & "',Now(),1 " & ",'" & Forms!foreclosuredetails!State & "')"
     DoCmd.RunSQL rstsql
     rstsql = ""
     
     rstsql = "UPDATE WizardSupportTwo SET LimboYellow = Now() WHERE FileNumber = " & Forms!Limbo_MD!lstFilesY & " AND current = true "
     DoCmd.RunSQL rstsql
     rstsql = ""
        
        cntr = 0
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM LimboMD_Yellow", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        Forms!Limbo_MD.QueueCountY = cntr
        Set rstqueue = Nothing
        Forms!Limbo_MD.lstFilesY.Requery
     
     
Case "Limbo_MDRed"
        JrlTxt = ""
        If Forms!Limbo_MD!lstFilesR.Column(6) = "X" Then
        JrlTxt = "Missing SOT"
        Else
        JrlTxt = ""
        End If
    
    
        If Forms!Limbo_MD!lstFilesR.Column(7) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing SOD"
            Else
            JrlTxt = JrlTxt & ", SOD"
            End If
        End If
        
    
        If Forms!Limbo_MD!lstFilesR.Column(8) = "X" And Forms!Limbo_MD!lstFilesR.Column(9) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing LMA"
            Else
            JrlTxt = JrlTxt & ", LMA"
            End If
        End If
    
'        If Forms!Limbo_MD!lstFilesR.Column(9) = "X" Then
'            If JrlTxt = "" Then
'            JrlTxt = "Missing FLMA"
'            Else
'            JrlTxt = JrlTxt & ", FLMA"
'            End If
'         End If
   
          
        If Forms!Limbo_MD!lstFilesR.Column(10) = "X" Then
          If JrlTxt = "" Then
          JrlTxt = "Missing ANO"
          Else
          JrlTxt = JrlTxt & ", ANO"
          End If
        End If
    
        If Forms!Limbo_MD!lstFilesR.Column(11) = "X" Then
         If JrlTxt = "" Then
         JrlTxt = "Missing NOI"
         Else
         JrlTxt = JrlTxt & ", NOI"
         End If
        End If
        
'        If Forms!Limbo_MD!lstFilesR.Column(12) = "X" Then
'            If JrlTxt = "" Then
'            JrlTxt = "Missing LNA"
'            Else
'            JrlTxt = JrlTxt & ", LNA"
'            End If
'        End If
        
     rstsql = "Insert InTo ValumeLimbo (CaseFile, Client, Name, RedComplete, RedCompleteC, state ) Values ( Forms!Limbo_MD!lstFilesR " & ",'" & Forms!Limbo_MD.lstFilesR.Column(1) & "', '" & GetFullName() & "',Now(),1 " & ",'" & Forms!foreclosuredetails!State & "')"
     DoCmd.RunSQL rstsql
     rstsql = ""
     
     rstsql = "UPDATE WizardSupportTwo SET LimboRed = Now() WHERE FileNumber = " & Forms!Limbo_MD!lstFilesR & " AND current = true "
     DoCmd.RunSQL rstsql
     rstsql = ""
    
        cntr = 0
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM LimboMD_Red", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        Forms!Limbo_MD.QueueCountR = cntr
        Set rstqueue = Nothing
        Forms!Limbo_MD.lstFilesR.Requery
    
    Case "Limbo_VAWhite"
        JrlTxt = ""
        
        If Forms!Limbo_VA!lstFiles.Column(6) = "X" Then
        JrlTxt = "Missing SOT"
        Else
        JrlTxt = ""
        End If
       
        If Forms!Limbo_VA!lstFiles.Column(7) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing OrNote"
            Else
            JrlTxt = JrlTxt & ", OrNote"
            End If
        End If
        
    
        If Forms!Limbo_VA!lstFiles.Column(8) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing LNA"
            Else
            JrlTxt = JrlTxt & ", LNA"
            End If
        End If
    
               
     rstsql = "Insert InTo ValumeLimbo (CaseFile, Client, Name, WhiteComplete, WhiteCompleteC, state ) Values ( Forms!Limbo_VA!lstFiles " & ",'" & Forms!Limbo_VA.lstFiles.Column(1) & "', '" & GetFullName() & "',Now(),1" & ",'" & Forms!foreclosuredetails!State & "')"
     DoCmd.RunSQL rstsql
     rstsql = ""
     
     rstsql = "UPDATE WizardSupportTwo SET LimboWhit = Now() WHERE FileNumber = " & Forms!Limbo_VA!lstFiles & " AND current = true "
     DoCmd.RunSQL rstsql
     rstsql = ""
    
    
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM LimboVAQueue", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        Forms!Limbo_VA.QueueCount = cntr
        Set rstqueue = Nothing
        Forms!Limbo_VA.lstFiles.Requery
    
    
    Case "Limbo_VAYellow"
        JrlTxt = ""
        
        If Forms!Limbo_VA!lstFilesY.Column(6) = "X" Then
        JrlTxt = "Missing SOT"
        Else
        JrlTxt = ""
        End If
       
        If Forms!Limbo_VA!lstFilesY.Column(7) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing OrNote"
            Else
            JrlTxt = JrlTxt & ", OrNote"
            End If
        End If
        
    
        If Forms!Limbo_VA!lstFilesY.Column(8) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing LNA"
            Else
            JrlTxt = JrlTxt & ", LNA"
            End If
        End If
    
               
     rstsql = "Insert InTo ValumeLimbo (CaseFile, Client, Name, YellowComplete, YellowCompleteC, state ) Values ( Forms!Limbo_VA!lstFilesY " & ",'" & Forms!Limbo_VA.lstFilesY.Column(1) & "', '" & GetFullName() & "',Now(),1" & ",'" & Forms!foreclosuredetails!State & "')"
     DoCmd.RunSQL rstsql
     rstsql = ""
     
     rstsql = "UPDATE WizardSupportTwo SET LimboYellow = Now() WHERE FileNumber = " & Forms!Limbo_VA!lstFilesY & " AND current = true "
     DoCmd.RunSQL rstsql
     rstsql = ""
    
        cntr = 0
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM LimboVA_Yellow", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        Forms!Limbo_VA.QueueCountY = cntr
        Set rstqueue = Nothing
        Forms!Limbo_VA.lstFilesY.Requery
    
    Case "Limbo_VARed"
        JrlTxt = ""
        
        If Forms!Limbo_VA!lstFilesR.Column(6) = "X" Then
        JrlTxt = "Missing SOT"
        Else
        JrlTxt = ""
        End If
       
        If Forms!Limbo_VA!lstFilesR.Column(7) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing OrNote"
            Else
            JrlTxt = JrlTxt & ", OrNote"
            End If
        End If
        
    
        If Forms!Limbo_VA!lstFilesR.Column(8) = "X" Then
            If JrlTxt = "" Then
            JrlTxt = "Missing LNA"
            Else
            JrlTxt = JrlTxt & ", LNA"
            End If
        End If
    
               
     rstsql = "Insert InTo ValumeLimbo (CaseFile, Client, Name, RedComplete, RedCompleteC, state ) Values ( Forms!Limbo_VA!lstFilesY " & ",'" & Forms!Limbo_VA.lstFilesR.Column(1) & "', '" & GetFullName() & "',Now(),1" & ",'" & Forms!foreclosuredetails!State & "')"
     DoCmd.RunSQL rstsql
     rstsql = ""
     
     rstsql = "UPDATE WizardSupportTwo SET LimboRed = Now() WHERE FileNumber = " & Forms!Limbo_VA!lstFilesR & " AND current = true "
     DoCmd.RunSQL rstsql
     rstsql = ""
    
        cntr = 0
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM LimboVA_Red", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        Forms!Limbo_VA.QueueCountR = cntr
        Set rstqueue = Nothing
        Forms!Limbo_VA.lstFilesR.Requery
        
    
    
    Case "Limbo_DCWhite"
        JrlTxt = ""
        
        If Forms!Limbo_DC!lstFiles.Column(6) = "X" Then
        JrlTxt = "Missing SOT"
        Else
        JrlTxt = ""
        End If
       
                       
     rstsql = "Insert InTo ValumeLimbo (CaseFile, Client, Name, WhiteComplete, WhiteCompleteC, state ) Values ( Forms!Limbo_DC!lstFiles " & ",'" & Forms!Limbo_DC.lstFiles.Column(1) & "', '" & GetFullName() & "',Now(),1" & ",'" & Forms!foreclosuredetails!State & "')"
     DoCmd.RunSQL rstsql
     rstsql = ""
     
     rstsql = "UPDATE WizardSupportTwo SET LimboWhit = Now() WHERE FileNumber = " & Forms!Limbo_DC!lstFiles & " AND current = true "
     DoCmd.RunSQL rstsql
     rstsql = ""
    
    
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM LimboDCQueue", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        Forms!Limbo_DC.QueueCount = cntr
        Set rstqueue = Nothing
        Forms!Limbo_DC.lstFiles.Requery
    
    Case "Limbo_DCYellow"
        JrlTxt = ""
        
        If Forms!Limbo_DC!lstFilesY.Column(6) = "X" Then
        JrlTxt = "Missing SOT"
        Else
        JrlTxt = ""
        End If
       
                       
     rstsql = "Insert InTo ValumeLimbo (CaseFile, Client, Name, YellowComplete, YellowCompleteC,state ) Values ( Forms!Limbo_DC!lstFilesY " & ",'" & Forms!Limbo_DC.lstFilesY.Column(1) & "', '" & GetFullName() & "',Now(),1" & ",'" & Forms!foreclosuredetails!State & "')"
     DoCmd.RunSQL rstsql
     rstsql = ""
     
     rstsql = "UPDATE WizardSupportTwo SET LimboYellow = Now() WHERE FileNumber = " & Forms!Limbo_DC!lstFilesY & " AND current = true "
     DoCmd.RunSQL rstsql
     rstsql = ""
     
        cntr = 0
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM LimboDC_Yellow", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        Forms!Limbo_DC.QueueCountY = cntr
        Set rstqueue = Nothing
        Forms!Limbo_DC.lstFilesY.Requery
    
    Case "Limbo_DCRed"
        JrlTxt = ""
        
        If Forms!Limbo_DC!lstFilesY.Column(6) = "X" Then
        JrlTxt = "Missing SOT"
        Else
        JrlTxt = ""
        End If
       
                       
     rstsql = "Insert InTo ValumeLimbo (CaseFile, Client, Name, RedComplete, RedCompleteC, state ) Values ( Forms!Limbo_DC!lstFilesR " & ",'" & Forms!Limbo_DC.lstFilesR.Column(1) & "', '" & GetFullName() & "',Now(),1" & ",'" & Forms!foreclosuredetails!State & "')"
     DoCmd.RunSQL rstsql
     rstsql = ""
     
     rstsql = "UPDATE WizardSupportTwo SET LimboRed = Now() WHERE FileNumber = " & Forms!Limbo_DC!lstFilesR & " AND current = true "
     DoCmd.RunSQL rstsql
     rstsql = ""
    
        
        cntr = 0
        Set rstqueue = CurrentDb.OpenRecordset("Select * FROM LimboDC_Yellow", dbOpenDynaset, dbSeeChanges)
        Do Until rstqueue.EOF
        cntr = cntr + 1
        rstqueue.MoveNext
        Loop
        Forms!Limbo_DC.QueueCountR = cntr
        Set rstqueue = Nothing
     Forms!Limbo_DC.lstFilesR.Requery
     
    
    End Select
    
        
     
      
      
    
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!foreclosuredetails!FileNumber & ",Now,'" & GetFullName() & " ','" & JrlTxt & "',1 )"
    DoCmd.RunSQL strSQLJournal
    
 
    strinfo = Forms!Limbo_Prosecc!JournalNote
    strinfo = Replace(strinfo, "'", "''")
    
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms!foreclosuredetails!FileNumber & ",Now,'" & GetFullName() & " ','" & strinfo & "',1 )"
    DoCmd.RunSQL strSQLJournal


DoCmd.SetWarnings True



DoCmd.Close acForm, "Journal"
DoCmd.Close acForm, "ForeclosureDetails"
DoCmd.Close acForm, "Case List"
DoCmd.Close acForm, "Limbo_Prosecc"


End Sub

