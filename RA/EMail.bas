Attribute VB_Name = "EMail"
Option Compare Database
Option Explicit

Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lparam As String) As Long
    
Public Const EMailPath = "c:\database\email\"
Global EMailStatus As Integer   ' 0 = never; 1 = active; 2 = inactive



Private Sub MapiSendMail(Optional SendTo As String)
Dim objSession As Object
Dim objMessage As Object
Dim objRecipient As Object
Dim objAttachment As Object
Dim Filespec As String

On Error GoTo SendMailErr

' Create the Session Object.
Set objSession = CreateObject("mapi.session")
' Log on using the session object.
' Specify a valid profile name if you want to
' avoid the logon dialog box.
'objSession.Logon profileName:="???"
objSession.Logon ShowDialog:=False, NewSession:=False
' Add a new message object to the OutBox.
Set objMessage = objSession.Outbox.Messages.Add
' Set the properties of the message object.
'objMessage.Subject = Subj
objMessage.Text = vbNewLine & vbNewLine

If SendTo <> "" Then
    ' Add a recipient object to the objMessage.Recipients collection.
    Set objRecipient = objMessage.Recipients.Add
    ' Set the properties of the recipient object.
    objRecipient.Name = SendTo
    objRecipient.Resolve
End If

Filespec = Dir(EMailPath & "*.pdf")     ' find attachments
Do While Filespec <> ""                 ' if any attachments found
    Set objAttachment = objMessage.Attachments.Add
    With objAttachment
        .Name = Left$(Filespec, Len(Filespec) - 4)  ' filename without .pdf
        .Type = 1   ' mapiFileData
        .Source = EMailPath & Filespec
        .ReadFromFile FileName:=EMailPath & Filespec
        .Position = 0
    End With
    objMessage.Update
    Filespec = Dir                      ' get next attachment
Loop
' Send the message. Setting showDialog to False
' sends the message without displaying the message
' or requiring user intervention. A setting of True
' displays the message and the user must choose
' to Send from within the message dialog.
objMessage.Send ShowDialog:=True
' Log off using the session object.
objSession.Logoff
Exit Sub

SendMailErr:
    MsgBox "Error " & Err.Number & ", Cannot send mail: " & Err.Description
    Exit Sub
End Sub

Private Sub SendMail(Optional SendTo As String)
'
' Send email using Outlook
'
Dim olApp As Object
Dim olMail As Object
Dim Filespec As String

On Error GoTo SendMail_Err

Set olApp = CreateObject("Outlook.Application")
Set olMail = olApp.CreateItem(olMailItem)     ' needs reference to Microsoft Outlook Object Library

With olMail
    If Not IsMissing(SendTo) Then .To = SendTo
    '.Subject =
    Filespec = Dir(EMailPath & "*.pdf")     ' find attachments
    Do While Filespec <> ""                 ' if any attachments found
        .Attachments.Add EMailPath & Filespec
        Filespec = Dir                      ' get next attachment
    Loop
    Filespec = Dir(EMailPath & "*.doc")     ' find attachments
    Do While Filespec <> ""                 ' if any attachments found
        .Attachments.Add EMailPath & Filespec
        Filespec = Dir                      ' get next attachment
    Loop
    '.ReadReceiptRequested = True
    '.Send
    .Display
End With

SendMail_Exit:
    Set olApp = Nothing
    Set olMail = Nothing
    Exit Sub

SendMail_Err:
    MsgBox "Cannot send mail: " & Err.Description
    Resume SendMail_Exit

End Sub

Public Sub SendMail2(SendTo As String, Subject As String, Message As String)
'
' Send email using Outlook
'
Dim olApp As Object
Dim olMail As Object
Dim Filespec As String

On Error GoTo SendMail2_Err

Set olApp = CreateObject("Outlook.Application")
Set olMail = olApp.CreateItem(olMailItem)     ' needs reference to Microsoft Outlook Object Library

With olMail
    If Not IsMissing(SendTo) Then .To = SendTo
    .Subject = Subject
    .Body = Message
    Filespec = Dir(EMailPath & "*.pdf")     ' find attachments
    Do While Filespec <> ""                 ' if any attachments found
        .Attachments.Add EMailPath & Filespec
        Filespec = Dir                      ' get next attachment
    Loop
    Filespec = Dir(EMailPath & "*.doc")     ' find attachments
    Do While Filespec <> ""                 ' if any attachments found
        .Attachments.Add EMailPath & Filespec
        Filespec = Dir                      ' get next attachment
    Loop
    '.ReadReceiptRequested = True
    .Send
    '.Display
End With

SendMail2_Exit:
    Set olApp = Nothing
    Set olMail = Nothing
    Exit Sub

SendMail2_Err:
    MsgBox "Cannot send mail: " & Err.Description
    Resume SendMail2_Exit

End Sub

Public Function EMailInit() As Boolean
On Error GoTo EMailInitErr

EMailInit = False
If EMailStatus = 0 Then
    If MsgBox("Outlook must be running in order to send mail." & vbNewLine & _
            "Is Outlook running?", vbYesNoCancel) <> vbYes Then Exit Function
End If

If Dir(EMailPath & "*.pdf") <> "" Then Kill EMailPath & "*.pdf"
If Dir(EMailPath & "*.doc") <> "" Then Kill EMailPath & "*.doc"
EMailStatus = 1
EMailInit = True
Exit Function

EMailInitErr:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Exit Function

End Function

Public Function EMailEnd() As Boolean

EMailEnd = True
If Dir(EMailPath & "*.pdf") = "" And Dir(EMailPath & "*.doc") = "" Then
    MsgBox "No documents ready to send", vbInformation
    EMailStatus = 2
    Exit Function
End If

Select Case MsgBox("Ready to send EMail?", vbQuestion + vbYesNoCancel)
    Case vbYes
        Call SendMail          ' use Outlook
        'Call MapiSendMail       ' generic email client
        EMailStatus = 2
    Case vbNo
        MsgBox "EMail not sent, but still active", vbInformation
        EMailEnd = False
        Exit Function
    Case vbCancel
        EMailStatus = 2
        MsgBox "EMail has been cancelled", vbInformation
        Exit Function
End Select
End Function

Public Sub Miktest()
  GetOutlookItemList (2)
End Sub

Public Sub GetOutlookItemList(listType As String)

Dim olApp As Outlook.Application
Dim olNameSpace As Outlook.NameSpace
Dim olInbox As Outlook.MAPIFolder

Dim lrs As Recordset

Dim objItem As Object

Set olApp = CreateObject("Outlook.Application")
Set olNameSpace = olApp.GetNamespace("MAPI")

Select Case listType
  Case "Inbox"
     Set olInbox = olNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
  Case "Sent Items"
     Set olInbox = olNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail)
End Select

DoCmd.SetWarnings False
DoCmd.RunSQL ("delete from emails")
DoCmd.SetWarnings True


Set lrs = CurrentDb.OpenRecordset("Emails")
Dim strBuffer As String

Dim receivedDate As Date
Dim cnt As Integer

cnt = 0

For Each objItem In olInbox.Items

  receivedDate = objItem.ReceivedTime
  cnt = cnt + 1
  
  If (DateDiff("d", receivedDate, Date) < 15 And cnt < 500) Then
    lrs.AddNew
  
    lrs![ID] = objItem.EntryID
    lrs![Who] = Left$(objItem.SenderName, 255)
    lrs![receivedDate] = objItem.ReceivedTime
    lrs![Subject] = Left$(objItem.Subject, 255)
  
    If Len(objItem.Body) < 501 Then
    strBuffer = CleanEmailString(Nz(objItem.Body))
    '500 character limit
    lrs![Content] = strBuffer
  End If
  
    lrs.Update
  End If
Next



Exit_GetOutlookItemList:
  Set olInbox = Nothing
  Set olNameSpace = Nothing
  Set olApp = Nothing
  
End Sub

Private Function CleanEmailString(strBuffer As String) As String

  Dim iCount As Integer
  Dim iFirst As Integer
  Dim iLen As Integer
  
  CleanEmailString = ""
  
 
  
  iFirst = 0
  For iCount = 1 To Len(strBuffer)

    If (Asc(Mid(strBuffer, iCount, 1)) = 10 Or Asc(Mid(strBuffer, iCount, 1)) = 13) Then
     CleanEmailString = CleanEmailString & Mid(strBuffer, iCount, 1)
     iFirst = 0
    ElseIf (Asc(Mid(strBuffer, iCount, 1)) = 34) Then
      CleanEmailString = CleanEmailString & "'"
      iFirst = 0
    ElseIf Asc(Mid(strBuffer, iCount, 1)) > 32 And Asc(Mid(strBuffer, iCount, 1)) < 127 Then
      CleanEmailString = CleanEmailString & Mid(strBuffer, iCount, 1)
      iFirst = 0
    ElseIf Asc(Mid(strBuffer, iCount, 1)) = 32 Then
      iFirst = iFirst + 1
      If (iFirst < 3) Then
        CleanEmailString = CleanEmailString & Mid(strBuffer, iCount, 1)
      End If
    End If
    
  Next
  

End Function



Public Sub DeleteItem(EntryID As String)
'
' Delete the item.
'


Dim olApp As Outlook.Application
Dim olNameSpace As Outlook.NameSpace
Dim olInbox As Outlook.MAPIFolder

Dim msg As Object
Dim StoreID As String

Set olApp = CreateObject("Outlook.Application")
Set olNameSpace = olApp.GetNamespace("MAPI")

Set msg = olNameSpace.GetItemFromID(EntryID, StoreID)

If msg Is Nothing Then
    MsgBox "Unable to find message.", vbCritical
Else
    msg.Delete
End If
Set msg = Nothing
Set olNameSpace = Nothing
Set olApp = Nothing

End Sub

Public Sub cmd_EmailBT_Click(distriputionList As Integer, FileNumber As Long, Bodytext As String, Subjecttext As String)
On Error GoTo Err_cmd_EmailToBT_Click


    Dim SendTo As String
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim i As Long
    Dim rstqueue As Recordset
    Dim OutlookObj As Object, NewMsg As Object
    Dim AttachmentPath As String, AttachmentPath1 As String, AttachmentPath2 As String, AttachmentPath3 As String
    Dim rstJnl As Recordset

SendTo = Nz(DLookup("sValue", "DB", "ID=" & distriputionList))
'AttachmentPath = DocLocation & DocBucket(FileID) & "\" & FileID & "\Title Order " & Format$(Now(), "yyyymmdd") & " 7.pdf"

Set OutlookObj = CreateObject("Outlook.Application")
Set NewMsg = OutlookObj.CreateItem(0)

 'Create new email


   With NewMsg ' this is this test
      .To = SendTo
      '.CC = "tr@acertitle.com"
      .Body = Bodytext
      .Subject = Subjecttext
      .Send
   End With

' .DeferredDeliveryTime = defertime



  

Exit_cmd_EmailToBT_Click:
    Exit Sub

Err_cmd_EmailToBT_Click:
    MsgBox Err.Description
    Resume Exit_cmd_EmailToBT_Click

End Sub

Public Sub Email_Reminder(distriputionList As Integer, FileNumber As Long, Bodytext As String, Subjecttext As String)
On Error GoTo Err_cmd_EmailToBT_Click
', defertime As String
 '   Const Morningtime As String = "09:30:00"
    Dim SendTo As String
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim i As Long
    Dim rstqueue As Recordset
    Dim OutlookObj As Object, NewMsg As Object
    Dim AttachmentPath As String, AttachmentPath1 As String, AttachmentPath2 As String, AttachmentPath3 As String
    Dim rstJnl As Recordset

SendTo = Nz(DLookup("sValue", "DB", "ID=" & distriputionList))
'AttachmentPath = DocLocation & DocBucket(FileID) & "\" & FileID & "\Title Order " & Format$(Now(), "yyyymmdd") & " 7.pdf"

Set OutlookObj = CreateObject("Outlook.Application")
Set NewMsg = OutlookObj.CreateItem(0)

 'Create new email


   With NewMsg ' this is this test
      .To = SendTo
      '.CC = "tr@acertitle.com"
      .Body = Bodytext
      .Subject = Subjecttext
    '  .DeferredDeliveryTime = defertime & " " & Morningtime
      .Send
      
   End With

' .DeferredDeliveryTime = defertime



  

Exit_cmd_EmailToBT_Click:
    Exit Sub

Err_cmd_EmailToBT_Click:
    MsgBox Err.Description
    Resume Exit_cmd_EmailToBT_Click

End Sub


Public Sub CheckSendingEmail()
Dim i As Integer
Dim J As Integer
Dim str As Recordset
Dim UsrId As Integer

UsrId = GetStaffID()


Set str = CurrentDb.OpenRecordset("Select * From ScheduledEmail Where SentEmail = False", dbOpenDynaset, dbSeeChanges)
    If Not str.EOF Then
    
        Do Until str.EOF
            If Format(Date, "mm/dd/yyyy") = Format(str!SendingDAte, "mm/dd/yyyy") And (UsrId = str!Sender1 Or UsrId = str!Sender2 Or UsrId = str!Sender3) Then
            
            i = i + 1
            
            Call Email_Reminder(48, str!FileNumber, str!Emailboudy, str!EmailSub)
            
            With str
            .Edit
            !SentEmail = True
            .Update
            End With
            End If
          
            str.MoveNext
        
            
        Loop
    
    Else
    Exit Sub
    End If

Set str = Nothing

 

End Sub
