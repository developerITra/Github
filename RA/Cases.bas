Attribute VB_Name = "Cases"
Option Compare Database
Option Explicit

Private Const MaxRecent = 10    ' number of recent case numbers to keep
Dim recent As Recordset
Global LockedFileNumber As Long
Global FileReadOnly As Boolean
Global ReadOnlyColor As Long
Global FairDebtColor As Long
Global Rejection As Boolean


Public Sub OpenCase(FileNumber As Long)
Dim F As Form, FormClosed As Boolean
'
' Open a case.  Update the recent list.

Set recent = CurrentDb.OpenRecordset("SELECT * FROM Recent where StaffID=" & GetStaffID() & " AND filenumber=" & FileNumber & " ORDER BY AccessTime DESC", dbOpenDynaset, dbSeeChanges)
'
' Attempt to find the case number in the recent list.
'
If recent.EOF Then 'not in recent list
    recent.AddNew
    recent!FileNumber = FileNumber
    recent!AccessTime = Now()
    recent!StaffID = StaffID
    recent.Update
Else                    ' found in recent list, update access time
    recent.Edit
    recent!AccessTime = Now()
    recent.Update
End If
'
' Close any other forms that may be open.  This is tricky because as you close a form, the Forms collection changes.
' So after closing a form, must check the entire collection again.
'
Do
    FormClosed = False
    For Each F In Forms
        Select Case F.Name
            Case "Main", "Search", "Journal", "Select File", "queRestart", "queService", "queborrowerserved", "queSaleSetting", "queSaleSettingwaiting", "queRSIReview", "queRestartWaiting", "queAttyMilestone1_5", "queAttyMilestone1", "queAttyMilestone2", "queAttyMilestone3", "queAttyMilestone4", "queAttyMilestone5", "queAttyMilestone6", "queAttyMilestone1mgr", "queAttyMilestone1_5mgr", "queAttyMilestone2mgr", "queAttyMilestone3mgr", "queAttyMilestone4mgr", "queAttyMilestone5mgr", "queAttyMilestone6mgr", "SCRA Search Info", "queSCRA1", "queSCRA2", "queSCRA3", "queSCRA4a", "queSCRA4b", "queSCRA5", "queSCRA6", "queSCRA7", "queSCRA8", "queSCRA9", "queSCRABK", "queSCRAFCNew", "queSCRA9waiting", "queVAappraisal", "vaappraisal search info", "queSAI", "queAttyMilestone1_25", "queAttyMilestone1_25mgr" ' leave these forms open"
            Case Else
                If UCase$(Left$(F.Name, 8)) <> "WORKFLOW" Then
                    FormClosed = True
                    DoCmd.Close acForm, F.Name
                    DoEvents
                End If
        End Select
    Next
Loop Until Not FormClosed
'
' Test file lock and open the file
'
Call LockFile(FileNumber)
DoCmd.OpenForm "Case List", , , "[FileNumber]=" & FileNumber

End Sub
Public Sub OpenCaseDONTCloseForms(FileNumber As Long)
Dim F As Form, FormClosed As Boolean
'
' Open a case.  Update the recent list.
'
' Keep the recordset open for better performance.
' If the recordset has not been created, then do so.
'
Set recent = CurrentDb.OpenRecordset("SELECT * FROM Recent where StaffID=" & GetStaffID() & " AND filenumber=" & FileNumber & " ORDER BY AccessTime DESC", dbOpenDynaset, dbSeeChanges)
'
' Attempt to find the case number in the recent list.
'
If recent.EOF Then 'not in recent list
    recent.AddNew
    recent!FileNumber = FileNumber
    recent!AccessTime = Now()
    recent!StaffID = StaffID
    recent.Update
Else                    ' found in recent list, update access time
    recent.Edit
    recent!AccessTime = Now()
    recent.Update
End If
'
' This is different from OpenCase in that other forms are NOT Closed.
' This is where the form closing code was removed.
'
' Test file lock and open the file
'
Call LockFile(FileNumber)
DoCmd.OpenForm "Case List", , , "[FileNumber]=" & FileNumber
DoCmd.OpenForm "ForeclosureDetails", , , "[FileNumber]=" & FileNumber


End Sub


Public Sub CheckConflicts(FileNumber As Long)
Dim qdf As QueryDef, sSQL As String

If StaffID = 0 Then Call GetLoginName

Set qdf = CurrentDb.QueryDefs("qryConflictsCheck")
sSQL = "exec spConflictsCheck"
qdf.sql = sSQL & " " & FileNumber & "," & StaffID
qdf.Execute

End Sub

Public Function CheckStaffConflict(FileNumber As Long) As Boolean
Dim rstCon As Recordset
If StaffID = 0 Then Call GetStaffID

Set rstCon = CurrentDb.OpenRecordset("SELECT * FROM StaffConflict where StaffID=" & StaffID & " AND Filenumber=" & FileNumber & " And ConflictStatus = 'Approved'", dbOpenDynaset, dbSeeChanges)
If Not rstCon.EOF Then
CheckStaffConflict = True
Else
CheckStaffConflict = False
End If

End Function

Public Function LockFile(FileNumber As Long) As Boolean
Dim rstLocks As Recordset, LockUserID As Long, rstLocksArchive As Recordset
Dim str_SQL As String

On Error GoTo LockFileErr

If FileLocks Then
    If LockedFileNumber <> 0 Then Call ReleaseFile(LockedFileNumber)
    Set rstLocks = CurrentDb.OpenRecordset("SELECT * FROM Locks WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
    'Removed by JE 07-14-2014
    'Set rstLocksArchive = CurrentDb.OpenRecordset("SELECT * FROM LocksArchive WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
    If rstLocks.EOF Then    ' no lock record for this file
        rstLocks.AddNew
        rstLocks!FileNumber = FileNumber
        rstLocks!StaffID = StaffID
        rstLocks!Timestamp = Now()
        rstLocks.Update
        LockFile = True
        LockedFileNumber = FileNumber
    Else
        LockUserID = rstLocks!StaffID
        If LockUserID = StaffID Then    ' this user already has this file locked
            LockFile = True
            LockedFileNumber = FileNumber
        ElseIf LockUserID = 0 Then      ' its available
            rstLocks.Edit
            rstLocks!StaffID = StaffID
            rstLocks!Timestamp = Now()
            rstLocks.Update
            LockFile = True
            LockedFileNumber = FileNumber
            'Removed by JE on 07-14-2014
            'change to append to archive lock records
            'rstLocksArchive.AddNew
            'rstLocksArchive!StaffID = StaffID
            'rstLocksArchive!FileNumber = FileNumber
            'rstLocksArchive!Timestamp = Now()
            'rstLocksArchive!Type = "L"
            'rstLocksArchive.Update
            'MsgBox "INSERT INTO LocksArchive " & FileNumber & StaffID & Now() & "L"
            'Added by JE 07-14-2014
            str_SQL = "INSERT INTO LocksArchive(FileNumber,StaffID,[TimeStamp],[Type]) VALUES (" & FileNumber & "," & StaffID & ",'" & Now() & "','L')"
            Debug.Print str_SQL
            RunSQL (str_SQL)
            
        ElseIf DateDiff("d", rstLocks!Timestamp, Date) > 0 Then ' re-use old lock
       
            rstLocks.Edit
            rstLocks!StaffID = StaffID
            rstLocks!Timestamp = Now()
            rstLocks.Update
            LockFile = True
            LockedFileNumber = FileNumber
            'Removed by JE on 07-14-2014
            'rstLocksArchive.AddNew
            'rstLocksArchive!StaffID = StaffID
            'rstLocksArchive!FileNumber = FileNumber
            'rstLocksArchive!Timestamp = Now()
            'rstLocksArchive!Type = "L"
            'rstLocksArchive.Update
            'Added by JE 07-14-2014
            str_SQL = "INSERT INTO LocksArchive(FileNumber,StaffID,[TimeStamp],[Type]) VALUES (" & FileNumber & "," & StaffID & ",'" & Now() & "','L')"
            Debug.Print str_SQL
            RunSQL (str_SQL)
        Else                            ' in use
            LockFile = False
        End If
    End If
    rstLocks.Close
    'rstLocksArchive.Close
Else                        ' locks not enabled, allow access
    LockFile = True
End If
If LockFile Then
    FileReadOnly = False
Else
    If ReadOnlyColor = 0 Then ReadOnlyColor = DLookup("iValue", "DB", "Name='ReadOnlyColor'")
    FileReadOnly = True
    MsgBox "File " & FileNumber & " is in use by " & DLookup("Name", "Staff", "ID=" & LockUserID), vbExclamation
End If
Exit Function

LockFileErr:
    LockFile = Not FileLocks    ' if locks enabled then don't allow, otherwise allow
    Exit Function

End Function

Public Sub ReleaseFile(FileNumber As Long)
Dim rstLocks As Recordset, rstLocksArchive As Recordset
Dim str_SQL As String

On Error GoTo ReleaseFileErr

If LockedFileNumber <> 0 Then
    Set rstLocks = CurrentDb.OpenRecordset("SELECT * FROM Locks WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
    'Removed by JE 07-14-2014
    'Set rstLocksArchive = CurrentDb.OpenRecordset("SELECT * FROM LocksArchive WHERE FileNumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
    If Not rstLocks.EOF Then    ' should exist
        rstLocks.Edit
        rstLocks!StaffID = 0
        rstLocks.Update
        'rstLocksArchive.AddNew
        'rstLocksArchive!StaffID = StaffID
        'rstLocksArchive!FileNumber = FileNumber
        'rstLocksArchive!Timestamp = Now()
        'rstLocksArchive!Type = "U"
        'rstLocksArchive.Update
        str_SQL = "INSERT INTO LocksArchive(FileNumber,StaffID,[TimeStamp],[Type]) VALUES (" & FileNumber & "," & StaffID & ",'" & Now() & "','U')"
        Debug.Print str_SQL
        RunSQL (str_SQL)
    End If
    rstLocks.Close
    'rstLocksArchive.Close
    LockedFileNumber = 0
End If
FileReadOnly = False
Exit Sub

ReleaseFileErr:
    MsgBox "Error releasing file lock: " & Err.Description
    Exit Sub

End Sub

Public Function ReadNextCaseNumber() As Long
Dim d As Recordset

On Error GoTo ReadErr
Set d = CurrentDb.OpenRecordset("SELECT iValue FROM DB WHERE Name = 'NextCaseNumber';", dbOpenSnapshot)
d.MoveFirst
ReadNextCaseNumber = d("iValue")
d.Close
Exit Function

ReadErr:
    MsgBox Err.Description
    ReadNextCaseNumber = 0
    Exit Function
End Function

Public Function ReserveNextCaseNumber() As Long
Dim d As Recordset
Dim NextNumber As Long

On Error GoTo ReserveErr
Set d = CurrentDb.OpenRecordset("SELECT iValue FROM DB WHERE Name = 'NextCaseNumber';", dbOpenDynaset, dbSeeChanges)
d.MoveFirst
NextNumber = d("iValue")
d.Edit
d("iValue") = NextNumber + 1
d.Update
d.Close
ReserveNextCaseNumber = NextNumber
Exit Function

ReserveErr:
    MsgBox Err.Description
    ReserveNextCaseNumber = 0
    Exit Function

End Function

Public Function ReserveNextClosedNumber() As Long
Dim d As Recordset
Dim NextNumber As Long

On Error GoTo ReserveClosedErr
Set d = CurrentDb.OpenRecordset("SELECT iValue FROM DB WHERE Name = 'NextClosedNumber';", dbOpenDynaset, dbSeeChanges)
d.MoveFirst
NextNumber = d("iValue")
d.Edit
d("iValue") = NextNumber + 1
d.Update
d.Close
ReserveNextClosedNumber = NextNumber
Exit Function

ReserveClosedErr:
    MsgBox Err.Description
    ReserveNextClosedNumber = 0
    Exit Function

End Function

Public Function AddDetailRecord(CaseType As Long, FileNumber As Long, ReferralDate As Date)
Dim Details As Recordset, JurisdictionID As Variant

Select Case CaseType

    Case 1  ' Foreclosure or Monitor
        JurisdictionID = DLookup("JurisdictionID", "CaseList", "FileNumber=" & FileNumber)
        Set Details = CurrentDb.OpenRecordset("FCDetails", dbOpenDynaset, dbSeeChanges)
        Details.AddNew
        Details!Referral = ReferralDate
        Details!FileNumber = FileNumber
        Details!Current = True
        Details!DOT = True
        Details!SubstituteTrustees = True
        
        If (CaseType = 1) Then
          If (Not IsNull(JurisdictionID)) Then
            If (JurisdictionID > 0) Then ' set jurisdiction state to foreclosure state
              Details!State = DLookup("[State]", "[JurisdictionList]", "[JurisdictionID] = " & JurisdictionID)
              
            End If
          End If
        End If
        
        Details.Update
        
        
    Case 8    '  Monitor
        JurisdictionID = DLookup("JurisdictionID", "CaseList", "FileNumber=" & FileNumber)
        Set Details = CurrentDb.OpenRecordset("FCDetails", dbOpenDynaset, dbSeeChanges)
        Details.AddNew
        Details!Referral = ReferralDate
        Details!FileNumber = FileNumber
        Details!Current = True
        Details!DOT = True
        Details!SubstituteTrustees = True
        Details.Update

        Set Details = CurrentDb.OpenRecordset("BKDetails", dbOpenDynaset, dbSeeChanges)
        Details.AddNew
        Details!FileNumber = FileNumber
        Details!Current = True
        Details.Update
        
        


    Case 2      ' Bankruptcy
        Set Details = CurrentDb.OpenRecordset("BKDetails", dbOpenDynaset, dbSeeChanges)
        Details.AddNew
        Details!FileNumber = FileNumber
        Details!Current = True
        Details.Update

    Case 4      ' Collection
        Set Details = CurrentDb.OpenRecordset("COLDetails", dbOpenDynaset, dbSeeChanges)
        Details.AddNew
        Details!FileNumber = FileNumber
        Details.Update
    
    Case 5      ' Civil
        Set Details = CurrentDb.OpenRecordset("CIVDetails", dbOpenDynaset, dbSeeChanges)
        Details.AddNew
        Details!FileNumber = FileNumber
        Details.Update
    
    Case 7      ' Eviction
        Set Details = CurrentDb.OpenRecordset("EVDetails", dbOpenDynaset, dbSeeChanges)
        Details.AddNew
        Details!FileNumber = FileNumber
        Details!Current = True
        Details.Update

    Case 9      ' REO
        Set Details = CurrentDb.OpenRecordset("REODetails", dbOpenDynaset, dbSeeChanges)
        Details.AddNew
        Details!FileNumber = FileNumber
        Details.Update

    Case 10      ' Title Resolution
        JurisdictionID = DLookup("JurisdictionID", "CaseList", "FileNumber=" & FileNumber)
        Set Details = CurrentDb.OpenRecordset("TRDetails", dbOpenDynaset, dbSeeChanges)
        Details.AddNew
        Details!FileNumber = FileNumber
        Details!Referral = ReferralDate
        Details!Current = True
        
          If (Not IsNull(JurisdictionID)) Then
            If (JurisdictionID > 0) Then ' set jurisdiction state to foreclosure state
              Details!State = DLookup("[State]", "[JurisdictionList]", "[JurisdictionID] = " & Forms![Case List]!JurisdictionID)
            End If
          End If

        Details.Update

End Select
Details.Close

End Function

Public Function CheckOpenJournalEntry() As Boolean

CheckOpenJournalEntry = True
If (IsLoaded("Journal New Entry")) Then

  If (MsgBox("Do you want to complete the journal entry?", vbYesNo, "Complete Journal Entry") = vbYes) Then
    CheckOpenJournalEntry = False
  Else
    DoCmd.Close acForm, "Journal New Entry"
  End If
End If

End Function

Public Sub AddDefaultTrustees(FileNumber As Long)
Dim rstTrustees As Recordset, rstStaff As Recordset, State As String, ctr As Integer

On Error GoTo Err_AddDefaultTrustees

State = Nz(DLookup("State", "FCdetails", "Current<>0 AND FileNumber=" & FileNumber))
Select Case State
    Case "VA", "DC"
        Set rstTrustees = CurrentDb.OpenRecordset("Select * from Trustees where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
        If rstTrustees.EOF Then
        Set rstStaff = CurrentDb.OpenRecordset("SELECT ID, Name  FROM Staff WHERE AutoAddTrustee" & State & " Is Not Null ORDER BY AutoAddTrustee" & State, dbOpenSnapshot)
        Do While Not rstStaff.EOF
            rstTrustees.AddNew
            rstTrustees!FileNumber = FileNumber
            rstTrustees!Trustee = rstStaff!ID
            rstTrustees!Assigned = Now()
            rstTrustees!Name = rstStaff!Name
            rstTrustees.Update
            rstStaff.MoveNext
        Loop
        rstStaff.Close
        End If
        Case "MD"
        Set rstTrustees = CurrentDb.OpenRecordset("Select * from Trustees where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
        If rstTrustees.EOF Then
        Set rstStaff = CurrentDb.OpenRecordset("SELECT ID, Name FROM Staff WHERE AutoAddTrustee" & State & " Is Not Null ORDER BY AutoAddTrustee" & State, dbOpenSnapshot)
        Do While Not rstStaff.EOF
        'Only add first 4 trustees for Chase MD files only
            'If DLookup("clientid", "caselist", "filenumber=" & FileNumber) = 97 And ctr = 4 Then GoTo Exit_Proc
            rstTrustees.AddNew
            rstTrustees!FileNumber = FileNumber
            rstTrustees!Trustee = rstStaff!ID
            rstTrustees!Assigned = Now()
            rstTrustees!Name = rstStaff!Name
            rstTrustees.Update
            rstStaff.MoveNext
            ctr = ctr + 1
        Loop
        rstStaff.Close
        End If

    Case Else
        MsgBox "Usual trustees can be assigned in MD, DC or VA only.  Make sure you entered the state in the property address.", vbCritical
Exit Sub
End Select

Exit_Proc:


        
        rstTrustees.Close
        TrusteeWordFile = 0         ' invalidate cache
Exit Sub

Err_AddDefaultTrustees:
    MsgBox Err.Description

End Sub

'Public Function AddDocPreIndex(FileNumber As Long, DocTypeID As Long, Optional Notes As String) As String
''
'' Add a record to the DocIndex table for a document which has not yet been scanned.
'' Return the Record ID that will be printed as a barcode on the document.
''
'Dim rstDoc As Recordset
'
''Current Edits
'Set rstDoc = CurrentDb.OpenRecordset("DocIndex", dbOpenDynaset, dbSeeChanges)  ' an append only shouldnt require a bookmark
'With rstDoc
'    .AddNew
'    !FileNumber = FileNumber 'Numeric
'    !DocTitleID = DocTypeID  'Numeric
'    !DocGroup = ""  'Usually null/ Text Field
'    !StaffID = GetStaffID()  'Numeric
'    !DateStamp = Now()  'Date
'    !Filespec = Null
'    !Notes = Notes 'This is the PROBLEM Field  that can have " '  "
'    .Update
'    .Bookmark = .LastModified  ' Replace with a movelast, if using appendOnly
'    AddDocPreIndex = "*RA" & !DocID & "*"
'    .Close
'End With
'
'End Function

Public Function AddDocPreIndex(FileNumber As Long, DocTypeID As Long, Optional Notes As String) As String
'
' Add a record to the DocIndex table for a document which has not yet been scanned.
' Return the Record ID that will be printed as a barcode on the document.
'
Dim rstDoc As Recordset

'Current Edits
Set rstDoc = CurrentDb.OpenRecordset("DocIndex", dbOpenDynaset, dbSeeChanges + dbAppendOnly)  ' an append only shouldnt require a bookmark
With rstDoc
    ' MsgBox .RecordCount for testing
    .AddNew
    !FileNumber = FileNumber 'Numeric
    !DocTitleID = DocTypeID  'Numeric
    !DocGroup = ""  'Usually null/ Text Field
    !StaffID = GetStaffID()  'Numeric
    !DateStamp = Now()  'Date
    !Filespec = Null
    !Notes = Notes 'This is the PROBLEM Field  that can have " '  " when using inserts
    .Update
    .MoveLast
    '.Bookmark = .LastModified  ' Replace with a movelast, if using appendOnly
    AddDocPreIndex = "*RA" & !DocID & "*"
    .Close
End With

End Function


Public Function ParseWord(varPhrase As Variant, ByVal iWordNum As Integer, Optional strDelimiter As String = " ", _
    Optional bRemoveLeadingDelimiters As Boolean, Optional bIgnoreDoubleDelimiters As Boolean) As Variant
On Error GoTo Err_Handler
    'Purpose:   Return the iWordNum-th word from a phrase.
    'Return:    The word, or Null if not found.
    'Arguments: varPhrase = the phrase to search.
    '           iWordNum = 1 for first word, 2 for second, ...
    '               Negative values for words form the right: -1 = last word; -2 = second last word, ...
    '               (Entire phrase returned if iWordNum is zero.)
    '           strDelimiter = the separator between words. Defaults to a space.
    '           bRemoveLeadingDelimiters: If True, leading delimiters are stripped.
    '               Otherwise the first word is returned as null.
    '           bIgnoreDoubleDelimiters: If true, double-spaces are treated as one space.
    '               Otherwise the word between spaces is returned as null.
    'Author:    Allen Browne. http://allenbrowne.com. June 2006.
    Dim varArray As Variant     'The phrase is parsed into a variant array.
    Dim strPhrase As String     'varPhrase converted to a string.
    Dim strResult As String     'The result to be returned.
    Dim lngLen As Long          'Length of the string.
    Dim lngLenDelimiter As Long 'Length of the delimiter.
    Dim bCancel As Boolean      'Flag to cancel this operation.

    '*************************************
    'Validate the arguments
    '*************************************
    'Cancel if the phrase (a variant) is error, null, or a zero-length string.
    If IsError(varPhrase) Then
        bCancel = True
    Else
        strPhrase = Nz(varPhrase, vbNullString)
        If strPhrase = vbNullString Then
            bCancel = True
        End If
    End If
    'If word number is zero, return the whole thing and quit processing.
    If iWordNum = 0 And Not bCancel Then
        strResult = strPhrase
        bCancel = True
    End If
    'Delimiter cannot be zero-length.
    If Not bCancel Then
        lngLenDelimiter = Len(strDelimiter)
        If lngLenDelimiter = 0& Then
            bCancel = True
        End If
    End If

    '*************************************
    'Process the string
    '*************************************
    If Not bCancel Then
        strPhrase = varPhrase
        'Remove leading delimiters?
        If bRemoveLeadingDelimiters Then
            strPhrase = Nz(varPhrase, vbNullString)
            Do While Left$(strPhrase, lngLenDelimiter) = strDelimiter
                strPhrase = Mid(strPhrase, lngLenDelimiter + 1&)
            Loop
        End If
        'Ignore doubled-up delimiters?
        If bIgnoreDoubleDelimiters Then
            Do
                lngLen = Len(strPhrase)
                strPhrase = Replace(strPhrase, strDelimiter & strDelimiter, strDelimiter)
            Loop Until Len(strPhrase) = lngLen
        End If
        'Cancel if there's no phrase left to work with
        If Len(strPhrase) = 0& Then
            bCancel = True
        End If
    End If

    '*************************************
    'Parse the word from the string.
    '*************************************
    If Not bCancel Then
        varArray = Split(strPhrase, strDelimiter)
        If UBound(varArray) >= 0 Then
            If iWordNum > 0 Then        'Positive: count words from the left.
                iWordNum = iWordNum - 1         'Adjust for zero-based array.
                If iWordNum <= UBound(varArray) Then
                    strResult = varArray(iWordNum)
                End If
            Else                        'Negative: count words from the right.
                iWordNum = UBound(varArray) + iWordNum + 1
                If iWordNum >= 0 Then
                    strResult = varArray(iWordNum)
                End If
            End If
        End If
    End If

    '*************************************
    'Return the result, or a null if it is a zero-length string.
    '*************************************
    If strResult <> vbNullString Then
        ParseWord = strResult
    Else
        ParseWord = Null
    End If

Exit_Handler:
    Exit Function
    
Err_Handler:
    MsgBox "Error " & Err.Number & ", Cannot Parse " & Err.Description
    Resume Exit_Handler
    
End Function


Public Sub OpenCaseDONTCloseForms_S2(FileNumber As Long)

Dim F As Form, FormClosed As Boolean
'
' Open a case.  Update the recent list.
'
' Keep the recordset open for better performance.
' If the recordset has not been created, then do so.
'
Set recent = CurrentDb.OpenRecordset("SELECT * FROM Recent where StaffID=" & GetStaffID() & " AND filenumber=" & FileNumber & " ORDER BY AccessTime DESC", dbOpenDynaset, dbSeeChanges)
'
' Attempt to find the case number in the recent list.
'
If recent.EOF Then 'not in recent list
    recent.AddNew
    recent!FileNumber = FileNumber
    recent!AccessTime = Now()
    recent!StaffID = StaffID
    recent.Update
Else                    ' found in recent list, update access time
    recent.Edit
    recent!AccessTime = Now()
    recent.Update
End If
'
' This is different from OpenCase in that other forms are NOT Closed.
' This is where the form closing code was removed.
'
' Test file lock and open the file
'
'Call LockFile(FileNumber)

    If LockFile(FileNumber) Then
    Dim stDocName, stLinkCriteria As String
    stDocName = "ForeclosureDetails"
stLinkCriteria = "[FileNumber]=" & FileNumber & " AND Current = True"


DoCmd.OpenForm "Case List", , , "FileNumber = " & FileNumber
Forms![Case List].Visible = True
DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        MsgBox "File is locked", vbCritical
        Exit Sub
    End If



'DoCmd.OpenForm "Case List", , , "[FileNumber]=" & FileNumber
'DoCmd.OpenForm "ForeclosureDetails", , , "[FileNumber]=" & FileNumber
End Sub

