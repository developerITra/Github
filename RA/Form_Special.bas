VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Special"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdClientEvictionStatus_Click()
Dim rstClientContacts As Recordset, QuerySQL As String

On Error GoTo Err_cmdClientEvictionStatus_Click

If MsgBox("Really email Client Eviction Status report?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

txtStatus = "Preparing Reports for Client Contacts"
If Me.optClientContact Then      ' send to all
    Set rstClientContacts = CurrentDb.OpenRecordset("SELECT DISTINCT ClientContactID,Email,ClientName FROM qryEvictionClientContacts", dbOpenSnapshot)
Else
    If IsNull(Me.cbxClientContactID) Then
        MsgBox "Select a client contact.", vbCritical
        Exit Sub
    End If
    Set rstClientContacts = CurrentDb.OpenRecordset("SELECT DISTINCT ClientContactID,Email,ClientName FROM qryEvictionClientContacts WHERE ClientContactID=" & cbxClientContactID, dbOpenSnapshot)
End If

Do While Not rstClientContacts.EOF

    txtStatus = "Sending to " & rstClientContacts!ClientName
    Call EMailInit
    On Error Resume Next
    If EMailStatus <> 1 Then Exit Sub
    
    CurrentDb.QueryDefs.Delete "tmpEvictionClientContacts"
    On Error GoTo Err_cmdClientEvictionStatus_Click
    QuerySQL = CurrentDb.QueryDefs("qryEvictionClientContacts").sql
    QuerySQL = Left$(QuerySQL, Len(QuerySQL) - 3)   ' remove trailing ; and crlf
    QuerySQL = QuerySQL & " AND EVDetails.ClientContactID=" & rstClientContacts!ClientContactID
    CurrentDb.CreateQueryDef "tmpEvictionClientContacts", QuerySQL
    Call DoReport("Eviction Client Contacts", -2, "Property Report")
    Call SendMail2(rstClientContacts!EMail, "Eviction Property Status", "Current Eviction Status report is attached")
   ' Call SendMail2("mikki@systemadix.com", "Eviction Property Status", "Current Eviction Status report is attached")
    rstClientContacts.MoveNext
Loop
rstClientContacts.Close
EMailStatus = 2
txtStatus = "Completed sending"

Exit_cmdClientEvictionStatus_Click:
    Exit Sub

Err_cmdClientEvictionStatus_Click:
    MsgBox Err.Description
    Resume Exit_cmdClientEvictionStatus_Click
    

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

Private Sub cmdWebUpdate_Click()
Dim dbLocation As String
Dim LocalPath As String
dbLocation = "\\FileServer\Applications\Database\"
'Const dbLocation = "\\FileServer\Applications\"
LocalPath = "c:\database\"

'Dim dbSale As Database, Status As Long, d As Recordset
'Dim ftpSales As ChilkatFTP
'Set ftpSales = New ChilkatFTP
'
'On Error GoTo Err_cmdWebUpdate_Click
'
'If MsgBox("Really update the sale list on the web site?" & vbNewLine & "(This may take a few minutes)", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
'
'If dir(Environ("tmp") & "\Sales.mdb") <> "" Then Kill Environ("tmp") & "\Sales.mdb"
'Set dbSale = CreateDatabase(Environ("tmp") & "\Sales.mdb", dbLangGeneral)
'DoCmd.TransferDatabase acExport, "Microsoft Access", Environ("tmp") & "\Sales.mdb", acTable, "qrySaleList", "Sales"
'dbSale.Close
'
'ftpSales.Passive = True
'ftpSales.Hostname = "036ea16.netsolhost.com"
'ftpSales.Username = "ftp1932496"
'ftpSales.Password = "L@w2013$"
'
'Status = ftpSales.Connect()
'If Status = 0 Then
'    MsgBox "Failed to connect to web server: " & ftpSales.ErrorLogText, vbExclamation
'    Exit Sub
'End If
'
''Status = ftpsales.ChangeRemoteDir("www/rosenDATA")
'Status = ftpSales.ChangeRemoteDir("database")
'If Status = 0 Then
'    MsgBox "Failed to change directory: " & ftpSales.ErrorLogText, vbExclamation
'    ftpSales.Disconnect
'    ftpSales.SaveXmlLog JournalPath & "SalesUploadLog.xml"
'    Exit Sub
'End If
'
'Status = ftpSales.PutFile(Environ("tmp") & "\Sales.mdb", "sales.mdb")
'If Status = 0 Then
'    MsgBox "Failed to upload file to web server: " & ftpSales.ErrorLogText, vbExclamation
'    ftpSales.Disconnect
'    ftpSales.SaveXmlLog JournalPath & "SalesUploadLog.xml"
'End If
'
'ftpSales.Disconnect
'ftpSales.SaveXmlLog JournalPath & "SalesUploadLog.xml"
'
'Open JournalPath & "WebUpdate.log" For Append As #1
'Print #1, Format$(Now(), "mm/dd/yyyy hh:nn am/pm") & "  (" & GetLoginName() & ") Web site has been updated"
'Close #1
'
'Set d = CurrentDb.OpenRecordset("SELECT * FROM DB WHERE Name = 'WebUpdate';", dbOpenDynaset, dbSeeChanges)
'd.MoveFirst
'd.Edit
'd("sValue") = Format$(Date, "mm/dd/yyyy")
'd.Update
'd.Close
'
'MsgBox "Web site has been updated", vbInformation
'
'Exit_cmdWebUpdate_Click:
'    Exit Sub
'
'Err_cmdWebUpdate_Click:
'    MsgBox Err.Description
'    Resume Exit_cmdWebUpdate_Click
    
Dim dbSale As Database, Status As Long, d As Recordset

Dim ftpSales As New ChilkatFTP

'Set ftpSales = New ChilkatFTP
'On Error GoTo Err_WebUpdate

txtStatus = "Setting up ..."
DoEvents
If Dir(dbLocation & "Sales.txt") <> "" Then Kill dbLocation & "Sales.txt"
If Dir(LocalPath & "Sales.txt") <> "" Then Kill LocalPath & "Sales.txt"
'Set dbSale = CreateDatabase(LocalPath & "Sales.txt", dbLangGeneral)
'dbSale.Close
txtStatus = "Extracting Sales data ..."
DoEvents
'DoCmd.TransferDatabase acExport, "Microsoft Access", LocalPath & "Sales.txt", acTable, "qrySaleList", "Sales"
'DoCmd.TransferText acExportDelim, "Microsoft Access", LocalPath & "Sales.txt", acTable, "qrySaleList", "Sales"
DoCmd.RunSavedImportExport "Export-Sales-Text"
FileCopy LocalPath & "Sales.txt", dbLocation & "Sales.txt"
Kill LocalPath & "Sales.txt"

txtStatus = "Connecting to web server ..."
DoEvents
'"209.17.116.2"
ftpSales.Passive = True
ftpSales.Hostname = "036ea16.netsolhost.com"
ftpSales.Username = "ftp1932496"
ftpSales.Password = "L@w2013$"

'ftpSales.Hostname = "036ea16.netsolhost.com"
'ftpSales.Username = "ftp1932496"
'ftpSales.Password = "L@w2013$"


'ftpSales.Hostname = "ftp.rosenberg-assoc.com"
'ftpSales.Username = "rosenberg-assoc.com"
'ftpSales.Password = "kh0st3277"

Status = ftpSales.Connect()
If Status = 0 Then
'    LogMsg "(Automatic Update) Failed to connect to web server: " & ftpSales.ErrorLogText
    Exit Sub
End If

Status = ftpSales.ChangeRemoteDir("www/foreclosure")
'status = ftpSales.ChangeRemoteDir("wwwroot/rosenDATA")
If Status = 0 Then
'    LogMsg "(Automatic Update) Failed to change directory: " & ftpSales.ErrorLogText
    ftpSales.Disconnect
    'ftpSales.SaveXmlLog dbLocation & "SalesUploadLog.xml"
    Exit Sub
End If

txtStatus = "Copying data ..."
DoEvents

Status = ftpSales.PutFile(dbLocation & "Sales.txt", "Sales.txt")
If Status = 0 Then
'    LogMsg "(Automatic Update) Failed to upload file to web server: " & ftpSales.ErrorLogText
    ftpSales.Disconnect
    'ftpSales.SaveXmlLog dbLocation & "SalesUploadLog.xml"
    
End If

txtStatus = "Update complete!"
DoEvents

'LogMsg "(Automatic Update) Web site has been updated"

ftpSales.Disconnect
'ftpSales.SaveXmlLog dbLocation & "SalesUploadLog.xml"

Set d = CurrentDb.OpenRecordset("SELECT * FROM DB WHERE Name = 'WebUpdate';", dbOpenDynaset, dbSeeChanges)
d.MoveFirst
d.Edit
d("sValue") = Format$(Date, "mm/dd/yyyy")
d.Update
d.Close


If Status = 0 Then
        Open dbLocation & "WebUpdate.log" For Append As #1
        Print #1, Format$(Now(), "mm/dd/yyyy hh:nn am/pm") & "  (" & GetLoginName() & ") Web site update has failed."
        Close #1
       MsgBox "Web Update Unsuccessful " & ftpSales.LastErrorText
       
    Else
        Open dbLocation & "WebUpdate.log" For Append As #1
        Print #1, Format$(Now(), "mm/dd/yyyy hh:nn am/pm") & "  (" & GetLoginName() & ") Web site has been updated."
        Close #1
        MsgBox "Web Update Successful"

        DoCmd.Close
    End If
    


    
    
    
End Sub

Private Sub cmdEvictionStatus_Click()
Dim rstBrokers As Recordset, QuerySQL As String

On Error GoTo Err_cmdEvictionStatus_Click

If MsgBox("Really email Eviction Status report?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

txtStatus = "Preparing Reports for Brokers"
If optBrokers Then      ' send to all
    Set rstBrokers = CurrentDb.OpenRecordset("SELECT DISTINCT BrokerID,BrokerEmail,BrokerName FROM qryEvictionBrokers", dbOpenSnapshot)
Else
    If IsNull(cbxBrokerID) Then
        MsgBox "Select a broker", vbCritical
        Exit Sub
    End If
    Set rstBrokers = CurrentDb.OpenRecordset("SELECT DISTINCT BrokerID,BrokerEmail,BrokerName FROM qryEvictionBrokers WHERE BrokerID=" & cbxBrokerID, dbOpenSnapshot)
End If

Do While Not rstBrokers.EOF
    txtStatus = "Sending to " & rstBrokers!BrokerName
    Call EMailInit
    On Error Resume Next
    If EMailStatus <> 1 Then Exit Sub
    CurrentDb.QueryDefs.Delete "tmpEvictionBrokers"
    On Error GoTo Err_cmdEvictionStatus_Click
    QuerySQL = CurrentDb.QueryDefs("qryEvictionBrokers").sql
    QuerySQL = Left$(QuerySQL, Len(QuerySQL) - 3)   ' remove trailing ; and crlf
    QuerySQL = QuerySQL & " AND Brokers.BrokerID=" & rstBrokers!BrokerID
    CurrentDb.CreateQueryDef "tmpEvictionBrokers", QuerySQL
    Call DoReport("Eviction Brokers", -2, "Property Report")
    Call SendMail2(rstBrokers!BrokerEMail, "Eviction Property Status", "Current Eviction Status report is attached")
    'Call SendMail2("ebloom@systemadix.com", "Eviction Property Status", "Current Eviction Status report is attached")
    rstBrokers.MoveNext
Loop
rstBrokers.Close
EMailStatus = 2
txtStatus = "Completed sending"

Exit_cmdEvictionStatus_Click:
    Exit Sub

Err_cmdEvictionStatus_Click:
    MsgBox Err.Description
    Resume Exit_cmdEvictionStatus_Click
    
End Sub

Private Sub ComFannieMae_Click()
Call OpenExcel("Fannemae", "FannieMaeQuery")
End Sub

Private Sub optBrokers_AfterUpdate()
If optBrokers Then
    cmdEvictionStatus.Caption = "Send All Reports"
Else
    cmdEvictionStatus.Caption = "Send 1 Report"
End If
End Sub
