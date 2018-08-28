Attribute VB_Name = "Calendar"
Option Compare Database
Option Explicit
 
Public Const CalendarFolderName = "Public Folders/All Public Folders/Shared Calendar"
' Public Const CalendarFolderName = "Public Folders/All Public Folders/Test Calendar"

Dim StoreID As String
Dim objNS As Outlook.NameSpace

Public Function AddCalendarEvent(EventDateTime As Date, AllDay As Boolean, Subject As String, Location As String, intColor As Long, emailGroup As String) As String
'
' Return the EntryID of the item created.  Or an empty string if it fails.
'
Dim Folder As Outlook.MAPIFolder, RecipientList() As String, i As Integer
Dim Recip() As Outlook.Recipient
Dim Appointment As Outlook.AppointmentItem

Set Folder = GetMAPIFolder(CalendarFolderName)
If Not Folder Is Nothing Then
    Set Appointment = Folder.Items.Add
    If Not Appointment Is Nothing Then
        RecipientList = Split(DLookup("sValue", "DB", "Name='" & emailGroup & "'"), ";")
        'RecipientList = Split(DLookup("sValue", "DB", "Name='SharedCalRecipTest'"), ";")
        ReDim Recip(0 To UBound(RecipientList)) As Outlook.Recipient
        For i = 0 To UBound(RecipientList)
            Set Recip(i) = Appointment.Recipients.Add(RecipientList(i))
            With Recip(i)
                .Type = olRequired
                .Resolve
            End With
        Next i
        With Appointment
            .Subject = Subject
            .Start = EventDateTime
            .End = DateAdd("n", 1, EventDateTime)
            .Location = Location
            .BusyStatus = olBusy
            .Body = ""
            .AllDayEvent = AllDay
            .MeetingStatus = olMeeting
            .Save
            .Send
        End With
        Call SetApptColorLabel(Appointment, intColor)
        AddCalendarEvent = Appointment.EntryID
        DoEvents
    End If
Else
    AddCalendarEvent = ""
    MsgBox "Cannot find folder " & CalendarFolderName
End If
Set Appointment = Nothing
Set Folder = Nothing

End Function

Public Sub UpdateCalendarEvent(EntryID As String, EventDateTime As Date, AllDay As Boolean, Subject As String, Location As String, intColor As Long)
'
' Update the item.
'
Dim Folder As Outlook.MAPIFolder
Dim Appointment As Outlook.AppointmentItem

Set Folder = GetMAPIFolder(CalendarFolderName)
Set Appointment = objNS.GetItemFromID(EntryID, StoreID)

If Appointment Is Nothing Then
    MsgBox "Unable to find calendar entry", vbCritical
Else
    With Appointment
        .Subject = Subject
        .Start = EventDateTime
        .End = DateAdd("n", 1, EventDateTime)
        .Location = Location
        .BusyStatus = olBusy
        .Body = ""
        .AllDayEvent = AllDay
        .Save
        .Send
    End With
    Call SetApptColorLabel(Appointment, intColor)
End If
Set Appointment = Nothing
Set Folder = Nothing

End Sub

Public Sub DeleteCalendarEvent(EntryID As String)
'
' Delete the item.
'
Dim Folder As Outlook.MAPIFolder
Dim Appointment As Outlook.AppointmentItem

If EntryID = "X" Then
    MsgBox "Cannot remove entry from Shared Calendar because the entry was created manually."
    Exit Sub
End If

Set Folder = GetMAPIFolder(CalendarFolderName)
Set Appointment = objNS.GetItemFromID(EntryID, StoreID)

If Appointment Is Nothing Then
    MsgBox "Unable to find calendar entry", vbCritical
Else
    Appointment.MeetingStatus = olMeetingCanceled
    Appointment.Save
    Appointment.Send
    Appointment.Delete
End If
Set Appointment = Nothing
Set Folder = Nothing

End Sub

Private Sub SetApptColorLabel(objAppt As Outlook.AppointmentItem, intColor As Long)  ' Integer)
' This code from http://www.outlookcode.com/codedetail.aspx?id=139

    ' requires reference to CDO 1.21 Library
    ' adapted from sample code by Randy Byrne
    ' intColor corresponds to the ordinal value of the color label
        '1=Important, 2=Business, etc.
        '1 = Red
        '2 = Blue
        '3 = Light Green
        '4 = Gray
        '5 = Orange
        '6 = Light Blue
        '7 = Light Yellow
        '8 = Purple
        '9 = Dark Green
        '10= Yellow
'8421376 = teal
    Const CdoPropSetID1 = "0220060000000000C000000000000046"
    Const CdoAppt_Colors = "0x8214"
    Dim objCDO As MAPI.Session
    Dim objMsg As MAPI.Message
    Dim colFields As MAPI.Fields
    Dim objField As MAPI.Field
    Dim strMsg As String
    Dim intAns As Integer
    On Error Resume Next

    Set objCDO = CreateObject("MAPI.Session")
    objCDO.Logon "", "", False, False
    If Not objAppt.EntryID = "" Then
        Set objMsg = objCDO.GetMessage(objAppt.EntryID, objAppt.Parent.StoreID)
        Set colFields = objMsg.Fields
        Set objField = colFields.Item(CdoAppt_Colors, CdoPropSetID1)
        If objField Is Nothing Then
            Err.Clear
            Set objField = colFields.Add(CdoAppt_Colors, vbLong, intColor, CdoPropSetID1)
        Else
            objField.Value = intColor
        End If
        objMsg.Update True, True
    Else
        strMsg = "You must save the appointment before you add a color label. Do you want to save the appointment now?"
        intAns = MsgBox(strMsg, vbYesNo + vbDefaultButton1, "Set Appointment Color Label")
        If intAns = vbYes Then
            Call SetApptColorLabel(objAppt, intColor)
        End If
    End If

    Set objMsg = Nothing
    Set colFields = Nothing
    Set objField = Nothing
    objCDO.Logoff
    Set objCDO = Nothing
End Sub

Private Function GetMAPIFolder(FolderName As String) As MAPIFolder
Dim objOutlook As Outlook.Application
Dim objFolder As Outlook.MAPIFolder
Dim objFolders As Outlook.Folders
Dim arrName() As String
Dim i As Integer
Dim blnFound As Boolean

blnFound = False  'not required

On Error Resume Next
Set objOutlook = GetObject(, "Outlook.Application")
If objOutlook Is Nothing Then
    Set objOutlook = New Outlook.Application
    If objOutlook Is Nothing Then
        MsgBox "Outlook is not installed on this computer"
        GoTo ExitHere
    End If
End If

Set objNS = objOutlook.GetNamespace("MAPI")

arrName = Split(FolderName, "/")

Set objFolders = objNS.Folders

For i = 0 To UBound(arrName)
    For Each objFolder In objFolders
        If objFolder.Name = arrName(i) Then
            Set objFolders = objFolder.Folders
            blnFound = True
            Exit For
        Else
            blnFound = False
        End If
    Next
    If blnFound = False Then
        Exit For
    End If
Next

If blnFound = True Then
    StoreID = objFolder.StoreID
    Set GetMAPIFolder = objFolder
End If

ExitHere:
   Set objOutlook = Nothing
   'Set objNS = Nothing
   Set objFolder = Nothing
   Set objFolders = Nothing
End Function

Private Sub Demo_OutlookEntryID()

   ' The Outlook object library must be referenced.
   Dim ol As Outlook.Application
   Dim olns As Outlook.NameSpace
   Dim objFolder As Outlook.MAPIFolder
   Dim AllContacts As Outlook.Items
   Dim Item As Outlook.ContactItem
   Dim i As Integer

   ' If there are more than 500 contacts, change the following line:
   Dim MyEntryID(500) As String
   Dim StoreID As String
   Dim strFind As String

   ' Set the application object
   Set ol = New Outlook.Application

   ' Set the namespace object
   Set olns = ol.GetNamespace("MAPI")

   ' Set the default Contacts folder.
   Set objFolder = olns.GetDefaultFolder(olFolderContacts)

   ' Get the StoreID, which is a property of the folder.
   StoreID = objFolder.StoreID

   ' Set objAllContacts equal to the collection of all contacts.
   Set AllContacts = objFolder.Items
   i = 0

   ' Loop to get all of the EntryIDs for the contacts.
   For Each Item In AllContacts
      i = i + 1
      ' The EntryID is a property of the item.
      MyEntryID(i) = Item.EntryID
   Next

   ' Randomly choose the 2nd Contact to retrieve.
   ' In a larger solution, this might be the index from a list box.
   ' Both the StoreID and EntryID must be used to retrieve the item.
   Set Item = olns.GetItemFromID(MyEntryID(2), StoreID)
   Item.Display

End Sub

Public Sub CalendarEventInfo(EntryID As String)

Dim Folder As Outlook.MAPIFolder
Dim Appointment As Outlook.AppointmentItem

Set Folder = GetMAPIFolder(CalendarFolderName)
Set Appointment = objNS.GetItemFromID(EntryID, StoreID)

If Appointment Is Nothing Then
    MsgBox "Unable to find calendar entry", vbCritical
Else
    With Appointment
        If .AllDayEvent Then
            MsgBox "Sale Date: " & Format$(.Start, "m/d/yyyy")
        Else
            MsgBox "Sale Date & Time: " & .Start
        End If
    End With
    Call SetApptColorLabel(Appointment, 1)
End If
Set Appointment = Nothing
Set Folder = Nothing

End Sub

Public Sub ShowCalendarEvent(EntryID As String)
Dim Folder As Outlook.MAPIFolder
Dim Appointment As Outlook.AppointmentItem

Set Folder = GetMAPIFolder(CalendarFolderName)
Set Appointment = objNS.GetItemFromID(EntryID, StoreID)

If Appointment Is Nothing Then
    MsgBox "Unable to find calendar entry", vbCritical
Else
    Appointment.Display
End If
Set Appointment = Nothing
Set Folder = Nothing

End Sub
