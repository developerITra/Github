Attribute VB_Name = "DocMissing"
Option Compare Database
Option Explicit
Global strSQL As String
Global rstdocs As Recordset
Global CheckWiz As Recordset
Global CheckDoc As Recordset
Global DocName As String
Global TextJournal As String
Global rs As Recordset

'Missing Figures

Sub GeneralMissingDoc(FileNumber As Long, DocTitleNO As Integer, Demand As Boolean, FD As Boolean, Intake As Boolean, NOI As Boolean, Dockting As Boolean, Optional ClientID As Integer, Optional Restart As Boolean)

DoCmd.SetWarnings False

Select Case DocTitleNO

Case 1550 ' (Need free Approval)

    If Demand Then
        DocName = "Need Fee Approval"
        TextJournal = "Demand Fee Approval Dcument Uploaded"
        Call UpdateDemand(FileNumber, DocName)
     End If
    
Case 1549
    If Demand Then
        DocName = "Missing Figures"
        TextJournal = "Demand Reinstatement Figures Document Uploaded"
        Call UpdateDemand(FileNumber, DocName)
    End If

    
    If NOI Then
        DocName = "Rfigs"
        TextJournal = "NOI Figures Document Uploaded"
        Call UpdateNOI(FileNumber, DocName)
    End If
    
Case 124 'Demand ("Waiting for client demand")Demand Letter
    If Demand Then
        DocName = "Waiting for client demand"
        TextJournal = "Demand Client Demand Letter upload to system"
        Call UpdateDemand(FileNumber, DocName)
    End If

Case 4, 1517, 1522  'Fair debt + Intake Collateral Documents (Missing note) Note

      
    If FD Then
        DocName = "Note"
        TextJournal = "FairDebt Note upload to system"
        Call UpdateFairDebt(FileNumber, DocName)
    End If
    
    If Intake Then
        DocName = "Note"
        TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
        Call UpdateIntake(FileNumber, DocName)
    End If
    
    
   If Dockting Then
        DocName = "Note"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
    End If
    
    If Restart Then
     DocName = "Note"
            TextJournal = DocName & " was removed from the Restart waiting list of outstanding items"
            Call Updaterestart(FileNumber, DocName)
    End If
    
    
Case 1450, 1371

    If Intake Then
        DocName = "Note"
        TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
        Call UpdateIntake(FileNumber, DocName)
    End If
    
    If FD Then
        DocName = "Note"
        TextJournal = "FairDebt Note upload to system"
        Call UpdateFairDebt(FileNumber, DocName)
    End If
    
    If Dockting Then
        DocName = "Note"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
    End If

    If MsgBox("Is the Original Note included ?", vbQuestion + vbYesNo) = vbYes Then
     
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCDetails WHERE FileNumber = " & FileNumber & " AND Current = True", dbOpenDynaset, dbSeeChanges)
                If rs!DocBackOrigNote = False Then
    
                    rs.Edit
                    rs!DocBackOrigNote = True
                    rs.Update
                    'Add to Status line
                    AddStatus FileNumber, Date, "Received Original Note"
                    'added to Journal
                  
        
                    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "Original Note is received " & "',1 )"
                    DoCmd.RunSQL strSQLJournal
                    strSQLJournal = ""
                    Forms!Journal.Requery
    
                
            End If
        rs.Close
        Set rs = Nothing
    
    Else
      AddStatus FileNumber, Date, "Removed Original Note"

   End If
 


        
Case 1493, 1523
    If FD And ClientID = 1 Then
    DocName = "Figures"
        TextJournal = "FairDebt Figures upload to system"
        Call UpdateFairDebt(FileNumber, DocName)
     End If
    
    If Intake And ClientID = 1 Then
        DocName = "Jfigs"
        TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
        Call UpdateIntake(FileNumber, DocName)
    End If
    
    If Restart Then
        DocName = "Jfigs"
        TextJournal = DocName & " was removed from the restart waiting list of outstanding items"
        Call Updaterestart(FileNumber, DocName)
    End If

Case 1569, 1571
    
    If FD And ClientID = 385 Then
    DocName = "Figures"
        TextJournal = "FairDebt NationStar Figures upload to system"
        Call UpdateFairDebt(FileNumber, DocName)
    End If

Case 1570, 1572

    If FD And ClientID = 446 Then
    DocName = "Figures"
        TextJournal = "FairDebt BOA Figures upload to system"
        Call UpdateFairDebt(FileNumber, DocName)
     End If
    
    
Case 1554
    If FD Then
    DocName = "Military"
        TextJournal = "FairDebt Military Document upload to system"
        Call UpdateFairDebt(FileNumber, DocName)
    End If
    
Case 988
 If NOI Then
    DocName = "Payment Dates"
        TextJournal = "NOI Payment Dates Document Uploaded"
        Call UpdateNOI(FileNumber, DocName)
        DocName = "Default Dates"
        Call UpdateNOI(FileNumber, DocName)
        
 End If

Case 1553
    If NOI Then
    DocName = "Client Sent NOI Copy"
        TextJournal = "NOI Client Sent NOI Copy Document Uploaded"
        Call UpdateNOI(FileNumber, DocName)
        
End If


Case 962, 1526
 If Intake Then
    DocName = "RDOT"
        TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
        Call UpdateIntake(FileNumber, DocName)
 End If

 If Dockting Then
        DocName = "DOT"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
 End If

   

Case 1105
    If Intake Then
        DocName = "Loan Mod"
            TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
            Call UpdateIntake(FileNumber, DocName)
     End If
     
     If Dockting Then
        DocName = "Loan Mod"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
     End If
     
Case 860 'Loan Modification Agreement
    If Intake Then
            DocName = "Loan Mod"
                TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
                Call UpdateIntake(FileNumber, DocName)
         End If
    DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage or overnight costs for the Loan Modification Agreement|FC-OTH|Loan Mod Agreement Postage/Overnight Cost"

    If Dockting Then
        DocName = "Loan Mod"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
     End If
    
Case 0
    If Intake Then
            DocName = "NOI"
                TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
                Call UpdateIntake(FileNumber, DocName)
         End If

Case 1511, 1361, 1525
    If Intake Then
        DocName = "SSN"
            TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
            Call UpdateIntake(FileNumber, DocName)
    End If
    
    If Restart Then
    DocName = "SSN"
            TextJournal = DocName & " was removed from the Restart waiting list of outstanding items"
            Call Updaterestart(FileNumber, DocName)
   End If

Case 1, 591
    If Intake Then
        DocName = "Title"
            TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
            Call UpdateIntake(FileNumber, DocName)
     End If
     
     If Restart Then
     DocName = "Title"
            TextJournal = DocName & " was removed from the Restart waiting list of outstanding items"
            Call Updaterestart(FileNumber, DocName)
    End If
    
    
Case 288
    If Intake Then
        DocName = "Executed SOT"
            TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
            Call UpdateIntake(FileNumber, DocName)
     End If

     If Dockting Then
        DocName = "SOT"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
     End If
    
    If Restart Then
    DocName = "SOT"
            TextJournal = DocName & " was removed from the Restart waiting list of outstanding items"
    Call Updaterestart(FileNumber, DocName)
    End If

Case 362 'Assignment
    DoCmd.OpenForm "GetPostage", , , , , acDialog, "Enter total postage and/or overnight costs|FC-Assng|Assignment Postage/Overnight Costs"

    If Intake Then
        DocName = "Assignment"
            TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
            Call UpdateIntake(FileNumber, DocName)
    End If
    
    If Restart Then
    DocName = "AOM"
            TextJournal = DocName & " was removed from the Restart waiting list of outstanding items"
            Call Updaterestart(FileNumber, DocName)
    End If
    
    

Case 299
    If Intake Then
        DocName = "Title Review"
            TextJournal = DocName & " was removed from the intake waiting list of outstanding items"
            Call UpdateIntake(FileNumber, DocName)
    End If
    

Case 223
    If Dockting Then
        DocName = "SOD"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
     End If
    
Case 464
    If Dockting Then
        DocName = "ANO"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
     End If
    
Case 592
    If Dockting Then
        DocName = "NOI Aff"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
     End If

Case 1345
    If Dockting Then
        DocName = "LMA info"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
     End If
     
Case 1484
    If Dockting Then
        DocName = "LMA info"
            TextJournal = DocName & " was removed from the Docketing waiting list of outstanding items"
            Call Updatedocket(FileNumber, DocName)
     End If



End Select
DoCmd.SetWarnings False
End Sub


Sub UpdateDemand(FileNumber As Long, DocName As String)

Set CheckDoc = CurrentDb.OpenRecordset("Select * FROM DemandDocsNeeded where filenumber=" & FileNumber & " AND DocName ='" & DocName & "' And IsNull(DocReceived)", dbOpenDynaset, dbSeeChanges)

    If Not CheckDoc.EOF Then


            strSQL = "UPDATE DemandDocsNeeded SET " & " DocReceived = #" & Now() & "# , docreceivedby = " & GetStaffID & _
            " WHERE FileNumber = " & FileNumber & " AND DocName = ('" & DocName & "')" & " And Isnull(DocReceived)"
            DoCmd.RunSQL strSQL
            strSQL = ""
        Call UpdateJournal(FileNumber, TextJournal)
        Call CheckDemand(FileNumber)
    End If
    CheckDoc.Close
Set CheckDoc = Nothing
    
End Sub

Sub UpdateJournal(FileNumber As Long, TextJournal As String)
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & TextJournal & "',1 )"
    DoCmd.RunSQL strSQLJournal
    strSQLJournal = ""
    Forms!Journal.Requery
End Sub

Sub CheckDemand(FileNumber As Long)
 Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DemandDocsNeeded where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
    If rstdocs.EOF Then
    
        Set CheckWiz = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current = true", dbOpenDynaset, dbSeeChanges)
        
        If Not CheckWiz!DemandDocsRecdFlag Then
    
            strSQL = "UPDATE wizardqueuestats SET DemandDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                    DoCmd.RunSQL strSQL
                    strSQL = ""
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "All Missing Documents now received for Demand Queue" & "',1 )"
            DoCmd.RunSQL strSQLJournal
            strSQLJournal = ""
            Forms!Journal.Requery
        End If
        CheckWiz.Close
        Set CheckWiz = Nothing
        
    End If
rstdocs.Close
Set rstdocs = Nothing

    

End Sub

Sub UpdateNOI(FileNumber As Long, DocName As String)
Set CheckDoc = CurrentDb.OpenRecordset("Select * FROM DocumentMissing where FileNbr=" & FileNumber & " AND DocRecd = 0  And  DocName ='" & DocName & "'", dbOpenDynaset, dbSeeChanges)

    If Not CheckDoc.EOF Then

        strSQL = "UPDATE DocumentMissing SET " & " DocRecd = -1 , DocRecdBy = " & GetStaffID & _
            " WHERE FileNbr = " & FileNumber & " AND DocName = ('" & DocName & "')"
            DoCmd.RunSQL strSQL
            strSQL = ""
        Call UpdateJournal(FileNumber, TextJournal)
        Call CheckNOI(FileNumber)
    End If
    CheckDoc.Close
Set CheckDoc = Nothing

        
End Sub

Sub CheckNOI(FileNumber)

Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DocumentMissing where FileNbr=" & FileNumber & " AND Not (DocRecd)", dbOpenDynaset, dbSeeChanges)
        If rstdocs.EOF Then
        
         Set CheckWiz = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current = true", dbOpenDynaset, dbSeeChanges)
        
            If Not CheckWiz!DocsRecdFlag Then
        
               strSQL = "UPDATE wizardqueuestats SET DocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                        DoCmd.RunSQL strSQL
                        strSQL = ""
                        
                        
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "All Missing Documents now received for NOI Queue" & "',1 )"
            DoCmd.RunSQL strSQLJournal
            strSQLJournal = ""
            Forms!Journal.Requery
           
            End If
        
     CheckWiz.Close
     Set CheckWiz = Nothing
   
    End If
    rstdocs.Close
End Sub


Sub UpdateFairDebt(FileNumber As Long, DocName As String)

Set CheckDoc = CurrentDb.OpenRecordset("Select * FROM FairDebtDocsNeeded where filenumber=" & FileNumber & " AND DocName ='" & DocName & "' And IsNull(DocReceived)", dbOpenDynaset, dbSeeChanges)

    If Not CheckDoc.EOF Then
        strSQL = "UPDATE FairDebtDocsNeeded SET " & " DocReceived = #" & Now() & "# , DocReceivedBy = " & GetStaffID & _
        " WHERE FileNumber = " & FileNumber & " AND DocName = ('" & DocName & "')" & " And IsNull(DocReceived)"
        DoCmd.RunSQL strSQL
        strSQL = ""
        Call UpdateJournal(FileNumber, TextJournal)
        Call CheckFairDebt(FileNumber)
    End If
CheckDoc.Close
Set CheckDoc = Nothing
    
End Sub

Sub CheckFairDebt(FileNumber As Long)
Set rstdocs = CurrentDb.OpenRecordset("Select * FROM FairDebtDocsNeeded where filenumber=" & FileNumber & " AND docreceived is null", dbOpenDynaset, dbSeeChanges)
    If rstdocs.EOF Then
     
    Set CheckWiz = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current = true", dbOpenDynaset, dbSeeChanges)
        
        If Not CheckWiz!FairDebtDocsRecdFlag Then
            strSQL = "UPDATE wizardqueuestats SET FairDebtDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
            DoCmd.RunSQL strSQL
            strSQL = ""
   
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "All Missing Documents now received for FairDebt Queue" & "',1 )"
            DoCmd.RunSQL strSQLJournal
            strSQLJournal = ""
            Forms!Journal.Requery
           
            
        End If
     CheckWiz.Close
     Set CheckWiz = Nothing
   
    End If
    rstdocs.Close
Set rstdocs = Nothing

End Sub

Sub UpdateIntake(FileNumber As Long, DocName As String)

Set CheckDoc = CurrentDb.OpenRecordset("Select * FROM IntakeDocsNeeded where Filenumber=" & FileNumber & " AND DocName ='" & DocName & "' And IsNull(DocReceived)", dbOpenDynaset, dbSeeChanges)

    If Not CheckDoc.EOF Then
   
        DoCmd.RunSQL ("UPDATE IntakeDocsNeeded set DocReceived =#" & Now() & "#, docreceivedby = " & GetStaffID & " WHERE FileNumber= " & FileNumber & " AND IsNull(DocReceived) AND DocName= '" & DocName & "'")
        Call UpdateJournal(FileNumber, TextJournal)
        Call CheckIntake(FileNumber)
       
    End If
CheckDoc.Close
Set CheckDoc = Nothing

    
    

End Sub

Sub CheckIntake(FileNumber)
Set rstdocs = CurrentDb.OpenRecordset("Select * FROM IntakeDocsNeeded where FileNumber= " & FileNumber & " AND IsNull(DocReceived) ;", dbOpenDynaset, dbSeeChanges)

       If rstdocs.EOF Then
       
            Set CheckWiz = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current = true", dbOpenDynaset, dbSeeChanges)
            
                If Not CheckWiz!IntakeDocsRecdFlag Then
        
               strSQL = "UPDATE wizardqueuestats SET IntakeDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                    DoCmd.RunSQL strSQL
                    strSQL = ""
                    
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "All Missing Documents now received for Intake Queue" & "',1 )"
                DoCmd.RunSQL strSQLJournal
                strSQLJournal = ""
                Forms!Journal.Requery
                End If
            CheckWiz.Close
        Set CheckWiz = Nothing
        End If
    
      rstdocs.Close
  Set rstdocs = Nothing

End Sub

Sub Updatedocket(FileNumber As Long, DocName As String)

Set CheckDoc = CurrentDb.OpenRecordset("Select * FROM DocketingDocsNeeded where Filenumber=" & FileNumber & " AND DocName ='" & DocName & "' And IsNull(DocReceived)", dbOpenDynaset, dbSeeChanges)

    If Not CheckDoc.EOF Then
   
        DoCmd.RunSQL ("UPDATE DocketingDocsNeeded set DocReceived =#" & Now() & "#, docreceivedby = " & GetStaffID & " WHERE FileNumber= " & FileNumber & " AND IsNull(DocReceived) AND DocName= '" & DocName & "'")
        Call UpdateJournal(FileNumber, TextJournal)
        Call Checkdocket(FileNumber)
       
    End If
CheckDoc.Close
Set CheckDoc = Nothing
End Sub
Sub Checkdocket(FileNumber)
Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DocketingDocsNeeded where FileNumber= " & FileNumber & " AND IsNull(DocReceived) ;", dbOpenDynaset, dbSeeChanges)

       If rstdocs.EOF Then
       
            Set CheckWiz = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current = true", dbOpenDynaset, dbSeeChanges)
            
                If Not CheckWiz!DocketingDocsRecdFlag Then
        
               strSQL = "UPDATE wizardqueuestats SET DocketingDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                    DoCmd.RunSQL strSQL
                    strSQL = ""
                    
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "All Missing Documents now received for Docketing Queue" & "',1 )"
                DoCmd.RunSQL strSQLJournal
                strSQLJournal = ""
                Forms!Journal.Requery
                End If
            CheckWiz.Close
        Set CheckWiz = Nothing
        End If
    
      rstdocs.Close
  Set rstdocs = Nothing

End Sub



Sub RemoveDocMissFairDebt(FileNumber As Long)
DoCmd.SetWarnings False
Set CheckDoc = CurrentDb.OpenRecordset("Select * FROM FairDebtDocsNeeded where filenumber=" & FileNumber & " And IsNull(DocReceived)", dbOpenDynaset, dbSeeChanges)

    If Not CheckDoc.EOF Then
        strSQL = "UPDATE FairDebtDocsNeeded SET " & " DocReceived = #" & Now() & "# , DocReceivedBy = " & GetStaffID & _
        " WHERE FileNumber = " & FileNumber & " And IsNull(DocReceived)"
        DoCmd.RunSQL strSQL
        strSQL = ""
        
        strSQL = "UPDATE wizardqueuestats SET FairDebtDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
        DoCmd.RunSQL strSQL
        strSQL = ""
        
    End If
CheckDoc.Close
Set CheckDoc = Nothing
DoCmd.SetWarnings True
End Sub

Sub Updaterestart(FileNumber As Long, DocName As String)

Set CheckDoc = CurrentDb.OpenRecordset("Select * FROM RestartDocumentMissing where Filenumber=" & FileNumber & " AND DocName ='" & DocName & "' And IsNull(DocReceived)", dbOpenDynaset, dbSeeChanges)

    If Not CheckDoc.EOF Then
   
        DoCmd.RunSQL ("UPDATE RestartDocumentMissing set DocReceived =#" & Now() & "#, docreceivedby = " & GetStaffID & " WHERE FileNumber= " & FileNumber & " AND IsNull(DocReceived) AND DocName= '" & DocName & "'")
       ' TextJournal = "Remove Client Figers as outstanding document in Restart queue"
        Call UpdateJournal(FileNumber, TextJournal)
        Call CheckRstart(FileNumber)
       
    End If
CheckDoc.Close
Set CheckDoc = Nothing

End Sub

Sub CheckRstart(FileNumber)
Set rstdocs = CurrentDb.OpenRecordset("Select * FROM RestartDocumentMissing where FileNumber= " & FileNumber & " AND IsNull(DocReceived) ;", dbOpenDynaset, dbSeeChanges)

       If rstdocs.EOF Then
       
            Set CheckWiz = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNumber & " AND current = true", dbOpenDynaset, dbSeeChanges)
            
                If Not CheckWiz!RestartDocsRecdFlag Then
        
               strSQL = "UPDATE wizardqueuestats SET RestartDocsRecdFlag = true WHERE FileNumber = " & FileNumber & " AND current = true "
                    DoCmd.RunSQL strSQL
                    strSQL = ""
                    
                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & "All Missing Documents now received for restart Queue" & "',1 )"
                DoCmd.RunSQL strSQLJournal
                strSQLJournal = ""
                Forms!Journal.Requery
                End If
            CheckWiz.Close
        Set CheckWiz = Nothing
        End If
    
      rstdocs.Close
  Set rstdocs = Nothing
End Sub
