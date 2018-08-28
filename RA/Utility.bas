Attribute VB_Name = "Utility"
Option Compare Database
Option Explicit
Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)



Public Const dbLocation = "\\FileServer\Applications\Database\"
Public Const ClosedScanLocation = "\\FileServer\FortDoxBackup\ROSENBERG\Foreclosures\"
Public Const DocLocation = "\\FileServer\ForeclosureDocs\"
Public Const ClientDocLocation = "\\FileServer\ClientDocs\"
Public Const LabelRequestInbox = "\\PrintServer\LabelRequests\"
Public Const JournalPath = "\\FileServer\Applications\Journals\RA\"      ' backup journal files here
Public Const TemplatePath = dbLocation & "Templates\"   ' Word Document templates

Global FullName As String, StaffID As Long
Global FileLocks As Boolean    ' enable file locks
Global FileLocked As Boolean    ' current file is locked
Global ReportArgs As String
Global selecteddoctype As Long
Global selectedgrouptype As Long
Global SelectedDispositionID As Long
Global SelectedTrusteeID As Long
Global AddSection  As Boolean
Global CopyNo As Integer
Global SelectedLMTrusteeID As Long 'For Mediations
Global BillTitle As Boolean
Global BillTitleUpdate As Boolean
Global InOrOut As Boolean
Global UpdateName As Boolean
Global strinfo As String
Global strSQLJournal As String
Global NameJournal As String
Global confilictName As Boolean
Global conflictAddress As Boolean
Global EditFormRSI As Boolean
Global ClientSentDemand As Boolean
Global CheckCancelDisposition As Boolean
Global checkCmdEdit As Boolean
Global SortTable As String
Global WizESC As Boolean
Global MonitorChoose As Boolean
Global recent As Recordset
Global SumD As Integer
Global CaseNuUpdate As Boolean
Global EditDispute As Boolean



Global PrivAdmin As Boolean
Global PrivAccounting As Boolean
Global PrivReceivePayments As Boolean
Global PrivJournalFlags As Boolean
Global PrivSetDisposition As Boolean
Global PrivRescindDisposition As Boolean
Global PrivCloseFiles As Boolean
Global PrivSetSale As Boolean
Global PrivDocDate As Boolean
Global PrivAdjustDeposit As Boolean
Global PrivDeleteDocs As Boolean
Global PrivFinalReview As Boolean
Global PrivCheckRequest As Boolean
Global PrivClients As Boolean
Global PrivReportedVacant As Boolean
Global PrivDocMgmt As Boolean
Global PrivClientFeeCost As Boolean
Global PrivNotices As Boolean
Global PrivAttyQueue As Boolean
Global PrivWriteOff As Boolean
Global PrivPND As Boolean
Global PrivSCRA As Boolean
Global IstheNoteLost As Boolean
Global PrivDataManager As Boolean
Global PrivWaitingForBill As Boolean
Global PrivWaitingForDoc As Boolean
Global PrivTitleIssue As Boolean
Global PriveCaution As Boolean
Global PrivStop As Boolean
Global PrivJurisdic As Boolean
Global PrivBillingEdits As Boolean
Global PrivFC As Boolean
Global PrivEV As Boolean
Global PrivBK As Boolean
Global PrivPrintPostSale As Boolean
Global LexisNexis As Boolean
Global PrivSSN As Boolean
Global MERS As Boolean
Global TitleOrder As Boolean
Global SaleSetter As Boolean
Global LastPostSale As Boolean
Global ClientOrder As Boolean
Global PrivNewNOIFDDemaind As Boolean
Global PrivitLimitedView As Boolean
Global PrivEvictionCashForKeys As Boolean
Global PrivSCRACancelSearch As Boolean
Global PrivChrono As Boolean
Global conflicts As Boolean
Global privProject As Boolean
Global BHproject As Boolean
Global DCTabView As Boolean




Global FileNum As Long

Global ZipChange As Boolean
Global NoticeTypechange As Boolean
Global SSNContainer As String
Global SSNChange As Boolean
Global FileNO As Long
Global FCDis As Boolean
Global EscFilter As Boolean



Global PostageAmount As Currency
Global FeeAmount As Currency
Global Vendor As Integer

Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Const SW_SHOWNORMAL = 1
Const ERROR_FILE_NOT_FOUND = 2&
Const ERROR_PATH_NOT_FOUND = 3&
Const ERROR_BAD_FORMAT = 11&
Declare Function GetDesktopWindow& Lib "user32" ()

Global strDateChooserCpt As String
Global strDateChooserTxt As String
Global dtDateChooser As Date
Global strStage As String


'Global strOpenArgs As String ' Because DoReport's OpenArgs does not seem to work correctly!

Public Function DocBucket(FileNumber As Long) As String
DocBucket = FileNumber \ 1000
End Function

Public Sub StartDoc(DocName As String)
Dim Scr_hDC As Long, msg$

Scr_hDC = GetDesktopWindow()
Select Case ShellExecute(Scr_hDC, "Open", DocName, vbNullString, ClosedScanLocation, SW_SHOWNORMAL)
    Case ERROR_FILE_NOT_FOUND
        msg$ = "File not found."
    Case ERROR_PATH_NOT_FOUND
        msg$ = "Path was not found."
    Case ERROR_BAD_FORMAT
        msg$ = "Invalid executable file format."
    Case 31
        msg$ = "No association for the specified file type."
    Case Is > 32    ' success
        Exit Sub
    Case Else
        msg$ = "Unexpected return code."
End Select
MsgBox "Cannot open " & DocName & vbNewLine & msg$, vbOKOnly + vbCritical
End Sub

Public Function ThisDate(d As Date) As String
Select Case Day(d)
    Case 1, 21, 31
        ThisDate = Trim$(str$(Day(d))) & "st"
    Case 2, 22
        ThisDate = Trim$(str$(Day(d))) & "nd"
    Case 3, 23
        ThisDate = Trim$(str$(Day(d))) & "rd"
    Case Else
        ThisDate = Trim$(str$(Day(d))) & "th"
End Select
ThisDate = ThisDate & " day of " & Format$(d, "mmmm, yyyy")
End Function

Public Function CurrencyWords(Amount As Double) As String
'
' Translate currency (under a billion) to words
'
Dim cents As Integer
Dim millions As Integer
Dim thousands As Integer
Dim dollars As Integer
Dim tmp$

If Amount >= 1000000000 Then    ' can't do billions
    CurrencyWords = "(Amount too large)"
    Exit Function
End If

If Amount < 0 Then
    CurrencyWords = "(Amount is negative)"
    Exit Function
End If

millions = Int(Amount / 1000000)
thousands = Int((Amount - CLng(millions) * 1000000) / 1000)
dollars = Int(Amount - CLng(millions) * 1000000 - CLng(thousands) * 1000)
cents = Right(Format(Amount, "Currency"), 2)

If millions > 0 Then
    tmp$ = NumberWords(millions) & " Million, "
Else
    tmp$ = ""
End If

If thousands > 0 Then
    tmp$ = tmp$ & NumberWords(thousands) & " Thousand, "
End If

If dollars > 0 Then
    tmp$ = tmp$ & NumberWords(dollars)
End If

If Amount < 1 Then
    CurrencyWords = Format$(cents, "00") & "/100"
Else
    CurrencyWords = Trim$(tmp$) & " and " & Format$(cents, "00") & "/100"
End If
End Function

Public Function Ordinal(num As Integer) As String
If num >= 11 And num <= 13 Then
    Ordinal = num & "th"
Else
    Select Case Right$(str$(num), 1)
        Case "1"
            Ordinal = num & "st"
        Case "2"
            Ordinal = num & "nd"
        Case "3"
            Ordinal = num & "rd"
        Case Else
            Ordinal = num & "th"
    End Select
End If
End Function

Function NumberWords(Amount As Integer) As String
'
' Return words for an integer amount less than 1000.
'
Dim hundreds As Integer
Dim tens As Integer
Dim ones As Integer
Dim tmp$
Dim dash$

hundreds = Int(Amount / 100)
tens = Int((Amount - hundreds * 100) / 10)
ones = Int(Amount - hundreds * 100 - tens * 10)

If hundreds > 0 Then
    tmp$ = NumberWords(hundreds) & " Hundred "
Else
    tmp$ = ""
End If

If tens > 1 And ones > 0 Then
    dash$ = "-"
Else
    dash$ = " "
End If
Select Case tens
    Case 0
        Select Case ones
            Case 0
            Case 1
                tmp$ = tmp$ & "One"
            Case 2
                tmp$ = tmp$ & "Two"
            Case 3
                tmp$ = tmp$ & "Three"
            Case 4
                tmp$ = tmp$ & "Four"
            Case 5
                tmp$ = tmp$ & "Five"
            Case 6
                tmp$ = tmp$ & "Six"
            Case 7
                tmp$ = tmp$ & "Seven"
            Case 8
                tmp$ = tmp$ & "Eight"
            Case 9
                tmp$ = tmp$ & "Nine"
        End Select
    Case 1
        Select Case ones
            Case 0
                tmp$ = tmp$ & "Ten"
            Case 1
                tmp$ = tmp$ & "Eleven"
            Case 2
                tmp$ = tmp$ & "Twelve"
            Case 3
                tmp$ = tmp$ & "Thirteen"
            Case 4
                tmp$ = tmp$ & "Fourteen"
            Case 5
                tmp$ = tmp$ & "Fifteen"
            Case 6
                tmp$ = tmp$ & "Sixteen"
            Case 7
                tmp$ = tmp$ & "Seventeen"
            Case 8
                tmp$ = tmp$ & "Eighteen"
            Case 9
                tmp$ = tmp$ & "Nineteen"
        End Select
    Case 2
        tmp$ = tmp$ & "Twenty" & dash$ & NumberWords(ones)
    Case 3
        tmp$ = tmp$ & "Thirty" & dash$ & NumberWords(ones)
    Case 4
        tmp$ = tmp$ & "Forty" & dash$ & NumberWords(ones)
    Case 5
        tmp$ = tmp$ & "Fifty" & dash$ & NumberWords(ones)
    Case 6
        tmp$ = tmp$ & "Sixty" & dash$ & NumberWords(ones)
    Case 7
        tmp$ = tmp$ & "Seventy" & dash$ & NumberWords(ones)
    Case 8
        tmp$ = tmp$ & "Eighty" & dash$ & NumberWords(ones)
    Case 9
        tmp$ = tmp$ & "Ninety" & dash$ & NumberWords(ones)
End Select
NumberWords = tmp$
End Function

Public Function GetLoginName() As String
'
' Find the user's login name.  Then find his/her real name.
'
Dim LoginName As String, Initials As String, cnt As Long
Dim s As Recordset

If FullName = "" Then           ' name not known yet
    cnt = 20
    LoginName = String$(cnt, 0)
    Call GetUserName(LoginName, cnt)
    LoginName = Left$(LoginName, cnt - 1)
    
    
    Set s = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE Username=""" & LoginName & """", dbOpenSnapshot)
    If s.EOF Then           ' not in database, a new user!
  '      s.Close
        FullName = ""
        MsgBox "User is not in the database.  Please add the user through the Data Manager.", vbCritical, "User Status"
  '      Set s = CurrentDb.OpenRecordset("Staff", dbOpenDynaset, dbSeeChanges)
  '      Do While FullName = ""
  '          FullName = InputBox$("Please enter your name (First Last):", "New User")
  '      Loop
  '      FullName = Trim$(FullName)
  '      s.AddNew
  '      s!Name = FullName
  '      s!Username = LoginName
  '      s!Sort = Trim$(Mid$(FullName, InStr(1, FullName, " ")))
  '      s!Processor = True
  '      s!PrivAdmin = False
  '      s!PrivAccounting = False
  '      s!PrivReceivePayments = False
  '      s!PrivJournalFlags = False
  '      s!PrivAdjustDeposit = False
  '      s!PrivDeleteDocs = False
  '      s!PrivSetSale = False
  '      s!PrivSetDisposition = False
  '      s!PrivCloseFiles = False
  '      s!PrivDocDate = False
  '      s!PrivRescindDisposition = False
  '      s.Update
  '      s.Bookmark = s.LastModified
  '      StaffID = s!ID
     Else                        ' name is in the database
        FullName = s!Name    ' cache it for quick access next time
        StaffID = s!ID
        PrivAdmin = s!PrivAdmin
        PrivAccounting = s!PrivAccounting
        PrivReceivePayments = s!PrivReceivePayments
        PrivJournalFlags = s!PrivJournalFlags
        PrivAdjustDeposit = s!PrivAdjustDeposit
        PrivDeleteDocs = s!PrivDeleteDocs
        PrivSetSale = s!PrivSetSale
        PrivSetDisposition = s!PrivSetDisposition
        PrivCloseFiles = s!PrivCloseFiles
        PrivDocDate = s!PrivDocDate
        PrivRescindDisposition = s!PrivRescindDisposition
        PrivFinalReview = s!PrivFinalReview
        PrivCheckRequest = s!PrivCheckRequest
        PrivClients = s!PrivClients
        PrivReportedVacant = s!PrivReportedVacant
        PrivDocMgmt = s!PrivDocMgmt
        PrivClientFeeCost = s!PrivClientFeeCost
        PrivNotices = s!PrivNotices
        PrivAttyQueue = s!PrivAttyQueue
        PrivWriteOff = s!PrivWriteOff
        PrivPND = s!PrivPND
        PrivSCRA = s!PrivSCRA
        PrivDataManager = s!PrivDataManager
        
        PrivWaitingForBill = s!PrivWaitingForBill
        PrivWaitingForDoc = s!PrivWaitingForDoc
        PrivTitleIssue = s!PrivTitleIssue
        PriveCaution = s!PriveCaution
        PrivStop = s!PrivStop
        PrivJurisdic = s!PrivJurisdic
        PrivBillingEdits = s!PrivBillingEdits
        PrivFC = s!PrivFC
        PrivBK = s!PrivBK
        PrivEV = s!PrivEV
        PrivPrintPostSale = s!PrivPrintPostSale
        LexisNexis = s!LexisNexis
        TitleOrder = s!TitleOrder
        SaleSetter = s!SaleSetter
        LastPostSale = s!LastPostSale
        PrivSSN = s!SSN
        PrivNewNOIFDDemaind = s!NewNOIDebtDemaind
        PrivitLimitedView = s!PrivitLimitedView
        PrivEvictionCashForKeys = s!PrivEvictionCFK
        PrivSCRACancelSearch = s!PrivSCRACancelSearch
        PrivChrono = s!PrivChrono
        conflicts = s!conflicts
        privProject = s!Project
        DCTabView = s!DCTab
        
        
        
        
        
        
    End If
    'If s!Developer Then DoCmd.OpenForm "Developer"
    s.Close
    
    If (FullName <> "") Then   ' name is in database
      If Nz(DLookup("Initials", "Staff", "ID=" & StaffID)) = "" Then
        Do While Initials = ""
            Initials = InputBox$("Please enter your initals:", "Initials")
        Loop
        Set s = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID=" & StaffID, dbOpenDynaset, dbSeeChanges)
        s.Edit
        s!Initials = UCase$(Initials)
        s.Update
        s.Close
      End If
    End If
End If
GetLoginName = FullName
End Function

Public Function GetFullName() As String
If FullName = "" Then Call GetLoginName
GetFullName = FullName
End Function

Public Function GetStaffID() As Integer
If StaffID = 0 Then Call GetLoginName
GetStaffID = StaffID
End Function

Public Function GetStaffEmail() As String
GetStaffEmail = Nz(DLookup("Email", "Staff", "ID=" & StaffID))
End Function

Public Function GetStaffFullName(Idin As Integer) As String
GetStaffFullName = DLookup("Name", "Staff", "ID=" & Idin)
End Function

Public Function GetStaffInitials(Idin As Integer) As String
GetStaffInitials = DLookup("Initials", "Staff", "ID=" & Idin)
End Function

Public Function PCase(s As String) As String
Dim i As Long, up As Boolean, o As String, C As String

up = True
For i = 1 To Len(s)
    C = Mid$(s, i, 1)
    If up Then
        o = o & UCase$(C)
    Else
        o = o & LCase$(C)
    End If
    up = (UCase$(C) < "A" Or UCase$(C) > "Z")
Next i
PCase = o
End Function

Public Function SCase(s As String) As String
'
' Make sure start of string is upper case
'
SCase = UCase$(Left$(s, 1)) & Mid$(s, 2)
End Function

Public Function LiberFolio(Liber As Variant, Folio As Variant, State As String, Optional JurisdictionID As Integer) As String
Dim lf As String
Dim strWord As Variant
Select Case State
    Case "DC"
        lf = "at Instrument Number "
        If Nz(Liber) = "" Then
            lf = lf & "_______"
        Else
            lf = lf & Liber
        End If
    Case "MD"
        lf = "at "
        
        strWord = DLookup("[LiberWord]", "[JurisdictionList]", "JurisdictionID = " & JurisdictionID)
        If Not IsNull(strWord) Then
          lf = lf & strWord & " "
        Else
          lf = lf & "Liber "
        End If
          
        If Nz(Liber) = "" Then
            lf = lf & "_______"
        Else
            lf = lf & Liber
        End If
        
        strWord = DLookup("[FolioWord]", "[JurisdictionList]", "JurisdictionID = " & JurisdictionID)
        If Not IsNull(strWord) Then
          lf = lf & ", " & strWord & " "
        Else
          lf = lf & " Folio "
        End If
        
        If Nz(Folio) = "" Then
            lf = lf & "_______"
        Else
            lf = lf & Folio
        End If
    Case "VA"
        If Not IsNull(Liber) And IsNull(Folio) Then     ' Instrument #
            lf = "at Instrument Number " & Liber
        Else    ' Book & Page
            lf = "in Deed Book "
            If Nz(Liber) = "" Then
                lf = lf & "_______"
            Else
                lf = lf & Liber
            End If
            lf = lf & ", Page "
            If Nz(Folio) = "" Then
                lf = lf & "_______"
            Else
                lf = lf & Folio
            End If
        End If
    Case Else
        MsgBox "Unrecognized State '" & State & "', Liber/Folio information will not print", vbExclamation
        lf = ""
        LiberFolio = ""
End Select
LiberFolio = lf
End Function

Public Function CheckPropertyAddress(File As Long, varAddress1 As Variant, varCity As Variant, varState As Variant, varZip As Variant)
 
CheckPropertyAddress = 0

Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("select PropertyAddress, City, State, ZipCode from FCDetails where FileNumber = " & File & " and [Current] = 1")
If (Not rs.EOF) Then
  
  CheckPropertyAddress = (rs![PropertyAddress] = varAddress1) And (rs![City] = varCity) And (rs![State] = varState) And (rs!ZipCode = varZip)
  
End If


rs.Close
Set rs = Nothing


 
End Function


Public Function FirmName(Optional State As String = "MD") As String
If State = "VA" Then
    FirmName = "Commonwealth Trustees, LLC"
Else
    FirmName = "Rosenberg & Associates, LLC"
End If
End Function

Public Function FirmAddress(Optional NewLine As String = vbNewLine) As String
FirmAddress = "Rosenberg & Associates, LLC" & NewLine & _
    "7910 Woodmont Avenue, Suite 750" & NewLine & "Bethesda, Maryland 20814"
End Function

Public Function FirmAddressAttn(Attn As String, Optional NewLine As String = vbNewLine) As String
' 2012.02.03 DaveW:  Added to allow Eviction Dept
FirmAddressAttn = "Rosenberg & Associates, LLC" & NewLine & _
    "Attn: " & Attn & NewLine & _
    "7910 Woodmont Avenue, Suite 750" & NewLine & "Bethesda, Maryland 20814"
End Function

Public Function FirmAddressVA(Optional NewLine As String = vbNewLine) As String
FirmAddressVA = "Rosenberg & Associates, LLC" & NewLine & _
    "(Attorney for Commonwealth Trustees, LLC)" & NewLine & _
    "7910 Woodmont Avenue, Suite 750" & NewLine & "Bethesda, Maryland 20814"
End Function
Public Function FirmShortAddressVA(Optional NewLine As String = vbNewLine) As String
FirmShortAddressVA = "8601 Westwood Center Drive, Suite 255" & NewLine & "Vienna, Virginia 22182"
End Function

Public Function FirmPhoneVA() As String
FirmPhoneVA = "(703) 752-8500"
End Function


Public Function FirmShortAddressOneLine(Optional NewLine As String = vbNewLine) As String
FirmShortAddressOneLine = "7910 Woodmont Avenue, Suite 750, Bethesda, Maryland 20814"
End Function

Public Function FirmShortAddress(Optional NewLine As String = vbNewLine) As String
FirmShortAddress = "7910 Woodmont Avenue, Suite 750" & NewLine & "Bethesda, Maryland 20814"
End Function

Public Function FirmAddressPhone(Optional NewLine As String = vbNewLine) As String
FirmAddressPhone = FirmAddress(NewLine) & NewLine & "(301) 907-8000"
End Function

Public Function FirmPhone() As String
FirmPhone = "(301) 907-8000"
End Function

Public Function FirmFax() As String
FirmFax = "(301) 907-8101"
End Function

Public Function FirmAddressPhoneFax(Optional NewLine As String = vbNewLine) As String
FirmAddressPhoneFax = FirmAddress() & NewLine & "(301) 907-8000" & NewLine & "(301) 907-8101"
End Function

Public Function FormatZip(Zip As Variant) As String
If IsNull(Zip) Then
    FormatZip = ""
    Exit Function
End If

If Len(Zip) > 5 Then
    FormatZip = Left$(Zip, 5) & "-" & Right$(Zip, 4)
Else
    FormatZip = Zip
End If
End Function

'Public Function FormatPhone(Phone As Variant) As String  'This only works for PURE numbers currently.
'If IsNull(Phone) Then
'    FormatPhone = ""
'    Exit Function
'Else
'    FormatPhone = Format(Phone, "(000) 000-0000")
'End If
'End Function

Public Sub AddStatus(FileNumber As Long, StatusDate As Variant, StatusDesc As String, Optional Clock As Integer = 0)
Dim s As Recordset
Dim strSQL As String
On Error GoTo AddStatusErr
If IsNull(StatusDate) Then Exit Sub
'MsgBox (FileNumber)
'MsgBox (StatusDate)
'MsgBox (StaffID)


If StaffID = 0 Then Call GetLoginName
'Status speed SA10/30/14

    If Len(StatusDesc) > 255 Then
        StatusDesc = Left$(StatusDesc, 255)
        MsgBox "Status update too long, data will be truncated in Status Report" & vbNewLine & vbNewLine & Left$(StatusDesc, 255), vbExclamation
    
    End If
    
    
        DoCmd.SetWarnings False
        StatusDesc = Replace(StatusDesc, "'", "''")
        DoCmd.RunSQL "Insert Into StatusList (FileNumber,StatusDate,StatusWho,StatusDesc,StatusClock,EntryTime) VALUES ( '" & FileNumber & "' , '" & StatusDate & "' , '" & StaffID & "' , '" & StatusDesc & "','" & Clock & "' , '" & Now() & " ')"
        DoCmd.SetWarnings True

'Set s = CurrentDb.OpenRecordset("StatusList", dbOpenDynaset, dbSeeChanges)
's.AddNew
's("FileNumber") = FileNumber
's("StatusDate") = StatusDate
's("StatusWho") = StaffID
'If Len(StatusDesc) > 255 Then
'    s("StatusDesc") = Left$(StatusDesc, 255)
'    MsgBox "Status update too long, data will be truncated in Status Report" & vbNewLine & vbNewLine & Left$(StatusDesc, 255), vbExclamation
'Else
'    s("StatusDesc") = StatusDesc
'End If
's("StatusClock") = Clock
's("EntryTime") = Now()
's.Update
's.Close
Exit Sub

AddStatusErr:
    MsgBox "Unexpected error: " & Err.Description & vbNewLine & "Status information not updated"
    Exit Sub
End Sub

Public Sub AddInvoiceItem(FileNumber As Long, Process As String, Description As String, Amount As Currency, VendorID As Integer, Fee As Boolean, Actual As Boolean, ApprovalNeeded As Boolean, Advanced As Boolean)

Dim strSQLINvoice As String
Dim strDescription As String
Dim strValues As String
Dim strInsert As String

On Error GoTo AddIIErr
If StaffID = 0 Then Call GetLoginName


strSQLINvoice = "Insert into InvoiceItems "
strValues = " Values "

strSQLINvoice = strSQLINvoice + "(Filenumber, [Timestamp],StaffID,Process"
strValues = strValues + "(" & FileNumber & " , '" & Now & "', " & StaffID & ", '" & Process & "'"


If Len(Description) > 255 Then
    strDescription = Left$(Description, 255)

    MsgBox "Description too long, will be truncated" & vbNewLine & vbNewLine & Left$(Description, 255), vbExclamation
Else
    strDescription = Description
End If


strDescription = Replace(strDescription, "'", "''")

strSQLINvoice = strSQLINvoice + ",Description, Fee"
strValues = strValues + ", '" & strDescription & "', " & Fee

If Actual Then
    strSQLINvoice = strSQLINvoice + ", ActualAmount"
    strValues = strValues + ", " & Amount
Else
    strSQLINvoice = strSQLINvoice + ", EstimatedAmount"
    strValues = strValues + ", " & Amount
End If

strSQLINvoice = strSQLINvoice + ", ApprovalNeeded, AdvancedByRA, VendorID)"
strValues = strValues + ", " & ApprovalNeeded & ", " & Advanced & ", " & VendorID & ")"

'Call MsgBox(strSQLINvoice)
'Call MsgBox(strValues)
'Call MsgBox(strSQLINvoice + strValues)
DoCmd.SetWarnings False
strInsert = strSQLINvoice + strValues
DoCmd.RunSQL strInsert
'MsgBox "Stop here"
DoCmd.SetWarnings True
Exit Sub

AddIIErr:
    MsgBox "Unexpected error: " & Err.Description & vbNewLine & "Invoice item not created"
    Exit Sub

End Sub


'Public Sub AddInvoiceItem(FileNumber As Long, Process As String, Description As String, Amount As Currency, VendorID As Integer, Fee As Boolean, Actual As Boolean, ApprovalNeeded As Boolean, Advanced As Boolean)
'Dim ii As Recordset
'
'On Error GoTo AddIIErr
'If StaffID = 0 Then Call GetLoginName
'
'Set ii = CurrentDb.OpenRecordset("InvoiceItems", dbOpenDynaset, dbSeeChanges)
'ii.AddNew
'ii!FileNumber = FileNumber
'ii!Timestamp = Now()
'ii!StaffID = StaffID
'ii!Process = Process
'
'If Len(Description) > 255 Then
'    ii!Description = Left$(Description, 255)
'    MsgBox "Description too long, will be truncated" & vbNewLine & vbNewLine & Left$(Description, 255), vbExclamation
'Else
'    ii!Description = Description
'End If
'
'ii!Fee = Fee
'If Actual Then
'    ii!ActualAmount = Amount
'Else
'    ii!EstimatedAmount = Amount
'End If
'
'ii!ApprovalNeeded = ApprovalNeeded
'ii!AdvancedByRA = Advanced
'ii!VendorID = VendorID
'ii.Update
'ii.Close
'Exit Sub
'
'Vendor = 0
'
'AddIIErr:
'    MsgBox "Unexpected error: " & Err.Description & vbNewLine & "Invoice item not created"
'    Exit Sub
'End Sub

Public Sub AddFileLocationHistory(FileNumber As Long, FileLocation As String)
Dim rstHist As Recordset

On Error GoTo AddFileLocHistErr

If StaffID = 0 Then Call GetLoginName

Set rstHist = CurrentDb.OpenRecordset("FileLocationHistory", dbOpenDynaset, dbSeeChanges)
rstHist.AddNew
rstHist!FileNumber = FileNumber
rstHist!MoveDate = Now()
rstHist!StaffID = StaffID
rstHist!FileLocation = Left$(FileLocation, 255)
rstHist.Update
rstHist.Close
Exit Sub

AddFileLocHistErr:
    MsgBox "Unexpected error: " & Err.Description & vbNewLine & "File location history not updated"
    Exit Sub
End Sub

Public Function FetchLatestFileLocation(FileNumber As Long)

  Dim lngID As Long
  lngID = Nz(DMax("[ID]", "qryLastFileLocation", "[FileNumber] = " & FileNumber), 0)
  
  If (lngID = 0) Then
    FetchLatestFileLocation = "Not Determined"
  Else
    FetchLatestFileLocation = DLookup("[FileLocation]", "[FileLocationHistory]", "[ID]= " & lngID)
  End If
End Function

Public Sub AddFileResponsibilityHistory(FileNumber As Long, DepartmentID As Integer, StaffID As Long)
Dim rstHist As Recordset

On Error GoTo AddFileRespHistErr

Set rstHist = CurrentDb.OpenRecordset("FileResponsibilityHistory", dbOpenDynaset, dbSeeChanges)
rstHist.AddNew
rstHist!FileNumber = FileNumber
rstHist!RespDate = Now()
rstHist!StaffID = StaffID
rstHist!DepartmentID = DepartmentID
rstHist.Update
rstHist.Close
Exit Sub

AddFileRespHistErr:
    MsgBox "Unexpected error: " & Err.Description & vbNewLine & "File responsibility history not updated"
    Exit Sub
End Sub

Public Sub AddCheckRequest(FileNumber As Long, amt As Currency, Desc As String, PayableTo As String, RequestType As Integer, strFeeType As String, BulkCheck As Boolean, CheckNumber As Integer, PreviouslyBilled As Boolean, Location As String, Optional FCType As String)

'Public Sub AddCheckRequest(FileNumber As Long, amt As Currency, Desc As String, PayableTo As String, RequestType As Integer, strFeeType As String, BulkCheck As Boolean, CheckNumber As Integer, PreviouslyBilled As Boolean)
Dim rstCheckReq As Recordset

On Error GoTo AddCheckRequestErr

If StaffID = 0 Then Call GetLoginName

Set rstCheckReq = CurrentDb.OpenRecordset("CheckRequest", dbOpenDynaset, dbSeeChanges)
rstCheckReq.AddNew
rstCheckReq!FileNumber = FileNumber
rstCheckReq!RequestDate = Now()
rstCheckReq!RequestBy = StaffID
rstCheckReq!Amount = amt
rstCheckReq!PayableTo = PayableTo
rstCheckReq!StatusID = 1
rstCheckReq!Description = Desc
rstCheckReq!RequestTypeID = RequestType
rstCheckReq!FeeType = strFeeType
rstCheckReq!BulkCheck = BulkCheck
rstCheckReq!CheckNumber = CheckNumber
rstCheckReq!PreviousBilled = PreviouslyBilled
rstCheckReq!Location = Location
rstCheckReq!FCType = FCType

rstCheckReq.Update
rstCheckReq.Close
Exit Sub

AddCheckRequestErr:
    MsgBox "Unexpected error: " & Err.Description & vbNewLine & "Check Request not updated"
    Exit Sub
End Sub

Public Sub AddDocumentRequest(FileNumber As Long, DocType As Integer, DocLoc As Long)
Dim rstDocReq As Recordset

On Error GoTo AddDocRequestErr

If StaffID = 0 Then Call GetLoginName

Set rstDocReq = CurrentDb.OpenRecordset("DocumentRequest", dbOpenDynaset, dbSeeChanges)
rstDocReq.AddNew
rstDocReq!FileNumber = FileNumber
rstDocReq!RequestDate = Now()
rstDocReq!RequestBy = StaffID
rstDocReq!DocumentType = DocType
rstDocReq!DocumentLocation = DocLoc

rstDocReq.Update
rstDocReq.Close
Exit Sub

AddDocRequestErr:
    MsgBox "Unexpected error: " & Err.Description & vbNewLine & "Document Request not updated"
    Exit Sub
End Sub




Function NextWeekDay(d As Date) As Date
Select Case Weekday(d)
    Case vbSaturday
        NextWeekDay = DateAdd("d", 2, d)
    Case vbSunday
        NextWeekDay = DateAdd("d", 1, d)
    Case Else
        NextWeekDay = d
End Select
End Function

Function NextWeekDayAfterMonday(d As Date) As Date
Select Case Weekday(d)
    Case vbSaturday
        NextWeekDayAfterMonday = DateAdd("d", 3, d)
    Case vbSunday
        NextWeekDayAfterMonday = DateAdd("d", 2, d)
    Case vbMonday
        NextWeekDayAfterMonday = DateAdd("d", 1, d)
        
    Case Else
        NextWeekDayAfterMonday = d
End Select
End Function


Public Function OneLine(S1 As Variant) As String
Dim S2 As String
Dim i As Integer

If IsNull(S1) Then
    OneLine = ""
    Exit Function
End If

For i = 1 To Len(S1)
    Select Case Mid$(S1, i, 1)
        Case vbNewLine, Chr$(13)
            S2 = S2 & ", "
        Case Chr$(10)
            ' do nothing
        Case Else
            S2 = S2 & Mid$(S1, i, 1)
    End Select
Next i
OneLine = S2
End Function

Public Function FirstLine(S1 As String) As String
Dim S2 As String
Dim i As Integer

For i = 1 To Len(S1)
    Select Case Mid$(S1, i, 1)
        Case vbNewLine, Chr$(13)
            FirstLine = S2
            Exit Function
        Case Chr$(10)
            ' do nothing
        Case Else
            S2 = S2 & Mid$(S1, i, 1)
    End Select
Next i
FirstLine = S2
End Function

Public Function StateFromAddress(Addr As String) As String
'
' Parse an address, return the state
'
Dim i As Integer, StateAndZip As String, State As String
Dim sa As Recordset
'
' Look backwards through the string and find a comma
'
For i = Len(Addr) To 1 Step -1
    If Mid(Addr, i, 1) = "," Then Exit For
Next i

If i = 1 Then       ' didn't find a comma
    StateFromAddress = ""
    Exit Function
End If

StateAndZip = Trim(Mid(Addr, i + 1))

i = InStr(1, StateAndZip, " ")      ' find space between state and zip
If i = 0 Then
    State = StateAndZip             ' assume no zip
Else
    State = Left(StateAndZip, i - 1)
End If

If Len(State) = 2 Then              ' expand abbreviation
    Set sa = CurrentDb.OpenRecordset("SELECT StateName FROM StateAbbreviations WHERE StateAbbreviation='" & State & "';", dbOpenSnapshot)
    If sa.EOF Then                  ' didn't find abbreviation in table, use it anyway
        StateFromAddress = State
    Else
        StateFromAddress = sa("StateName")  ' found it, use full state name
    End If
    sa.Close
Else                                ' not abbreviated, use it
StateFromAddress = State
End If
End Function

Public Sub BackupJournal()
Dim J As Recordset, FN As Integer
Dim FileNumber As Long, cnt As Long

On Error GoTo Err_cmdOK_Click
Set J = CurrentDb.OpenRecordset("SELECT * FROM Journal ORDER BY JournalDate", dbOpenSnapshot)
Do While Not J.EOF

    FileNumber = J("FileNumber")
    cnt = J("ID")
    If cnt / 100 = cnt \ 100 Then Debug.Print cnt
    
    FN = FreeFile(1)
    Open JournalPath & "\" & Format$(FileNumber Mod 100, "00") & "\" & FileNumber & ".txt" For Append As FN
    Print #FN, J("Who")
    Print #FN, J("JournalDate")
    Print #FN, J("Info")
    Print #FN,
    Print #FN,
    Close FN
    DoEvents
    J.MoveNext
Loop
DoCmd.Close

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    If Err.Number = 76 Then     ' path not found
        MkDir JournalPath & "\" & Format$(FileNumber Mod 100, "00") & "\"
        Resume
    End If
    MsgBox "Error encountered attempting to add to journal: " & Err.Description, vbExclamation
    Resume Exit_cmdOK_Click
    
End Sub

Public Function MaximumValue(N1 As Variant, N2 As Variant) As Variant
If N1 > N2 Then
    MaximumValue = N1
Else
    MaximumValue = N2
End If
End Function

Public Function OpenFile(CurrentForm As Form) As String
Dim cdlg As New CommonDialogAPI
Dim lngFormHwnd As Long
Dim lngAppInstance As Long
Dim strInitDir As String
Dim strFileFilter As String
Dim lngResult As Long

lngFormHwnd = CurrentForm.hWnd
lngAppInstance = Application.hWndAccessApp
strInitDir = ""
strFileFilter = "All Files (*.*)" & Chr(0)
lngResult = cdlg.OpenFileDialog(lngFormHwnd, lngAppInstance, strInitDir, strFileFilter)

If cdlg.GetStatus = True Then
    OpenFile = cdlg.GetName
Else
    OpenFile = ""
End If

End Function

Public Function SelectDocumentType() As Long
DoCmd.OpenForm "Select Document Type", acNormal, , , , acDialog
SelectDocumentType = selecteddoctype
End Function

Public Function GetPostageAmount(Prompt As String, Process As String, Desc As String) As Currency

    DoCmd.OpenForm "GetPostage", , , , , acDialog, Prompt

End Function

Public Function GetFeeAmount(Prompt As String) As Currency
FeeAmount = 0
Do
    DoCmd.OpenForm "GetFee", , , , , acDialog, Prompt
   
Loop Until FeeAmount > 0
GetFeeAmount = FeeAmount
End Function
Public Function GetTitleFeeAmount(Prompt As String) As Currency
FeeAmount = 0
Do
    DoCmd.OpenForm "GetFeeTitle", , , , , acDialog, Prompt
Loop Until FeeAmount > 0
GetTitleFeeAmount = FeeAmount
End Function


Public Function GetAmount(Prompt As String) As Currency
FeeAmount = 0
Do
    DoCmd.OpenForm "GetAmount", , , , , acDialog, Prompt
Loop Until FeeAmount >= 0
GetAmount = FeeAmount
End Function

Public Sub AffadavitOfService(ID As Long)

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * from AffadavitOfServiceRecipients"
DoCmd.RunSQL "INSERT into AffadavitOfServiceRecipients (NameID, Company, First, Last, AKA, Address, Address2, City, State, Zip, Deceased) " & _
             "SELECT ID, Company, First, Last, AKA, Address, Address2, City, State, Zip, Deceased " & _
             "FROM Names " & _
             "WHERE FileNumber = " & ID & " and Owner = True"

DoCmd.SetWarnings True

End Sub


Function GetAttorneyForSummary(tid As Variant)

Dim lrs As Recordset
Dim sqlstr As String

GetAttorneyForSummary = ""

If (IsNull(tid)) Then
  Exit Function
End If

Set lrs = CurrentDb.OpenRecordset("select Company, First, Last from NamesCIVAttorneyRep inner join NamesCIV on NamesCIV.ID = NamesCIVAttorneyRep.AttorneyForNameID where AttorneyNameID = " & tid)

With lrs
  Do While Not .EOF
  
    GetAttorneyForSummary = GetAttorneyForSummary & IIf(Not IsNull(lrs![Company]), lrs![Company], lrs![First] & " " & lrs![Last]) & vbNewLine
    .MoveNext
  Loop
  .Close
  
End With

Set lrs = Nothing


End Function


Function GetAttorneyForSummaryTR(tid As Variant)

Dim lrs As Recordset
Dim sqlstr As String

GetAttorneyForSummaryTR = ""

If (IsNull(tid)) Then
  Exit Function
End If

Set lrs = CurrentDb.OpenRecordset("select Company, First, Last from NamesTRAttorneyRep inner join NamesTR on NamesTR.ID = NamesTRAttorneyRep.AttorneyForNameID where AttorneyNameID = " & tid)

With lrs
  Do While Not .EOF
  
    GetAttorneyForSummaryTR = GetAttorneyForSummaryTR & IIf(Not IsNull(lrs![Company]), lrs![Company], lrs![First] & " " & lrs![Last]) & vbNewLine
    .MoveNext
  Loop
  .Close
  
End With

Set lrs = Nothing


End Function



Function GetStaffDepartments(tid As Variant)

Dim lrs As Recordset
Dim sqlstr As String

GetStaffDepartments = ""

If (IsNull(tid)) Then
  Exit Function
End If

Set lrs = CurrentDb.OpenRecordset("select department from StaffDepartments inner join Department on StaffDepartments.DepartmentID = Department.DepartmentID where StaffDepartments.StaffID = " & tid)

With lrs
  Do While Not .EOF
  
    GetStaffDepartments = GetStaffDepartments & lrs![Department] & vbNewLine
    .MoveNext
  Loop
  .Close
  
End With

Set lrs = Nothing


End Function


Public Function GetResponsibility(ID As Long, resptype As String) As String

GetResponsibility = ""

Dim lrs As Recordset

Set lrs = CurrentDb.OpenRecordset("SELECT TOP 1 Staff.Initials " & _
                                  "FROM (FileResponsibilityHistory INNER JOIN Department ON FileResponsibilityHistory.DepartmentID = Department.DepartmentID) INNER JOIN Staff ON FileResponsibilityHistory.StaffID = Staff.ID " & _
                                  "WHERE FileNumber = " & ID & " and Department = '" & resptype & "' " & _
                                  "ORDER BY RespDate DESC")
                                                                    
If Not lrs.EOF Then
  GetResponsibility = lrs![Initials]
End If

lrs.Close
Set lrs = Nothing

End Function

Public Function GetLastDocIDNo(ID As Long, DocNo As Long, FileNO As Long) As Long


GetLastDocIDNo = 0

Dim lrs As Recordset

Set lrs = CurrentDb.OpenRecordset("select Top 1 DocID from DocIndex " & _
                                "where FileNumber = " & FileNO & " and DocTitleID = " & DocNo & " and StaffID = " & ID & " order by DocID desc", dbOpenDynaset, dbSeeChanges)
                                                                    
If Not lrs.EOF Then
GetLastDocIDNo = lrs![DocID]
End If

lrs.Close
Set lrs = Nothing

End Function



Public Sub SetObjectAttributes(obj As Object, tenabled As Boolean)

If (tenabled = True) Then
    obj.Enabled = True
    obj.Locked = False
    obj.BackColor = -2147483643
    obj.BackStyle = 1
Else
    obj.Enabled = False
    obj.Locked = True
    obj.BackColor = 16777215
    obj.BackStyle = 0
End If


End Sub

Public Sub TestPassword()

  MsgBox "password = " & PasswordGenerator(6)
End Sub


Public Function PasswordGenerator(Length As Long) As String
On Error GoTo Err_PasswordGenerator

Dim iChr As Integer
Dim C As Long
Dim strResult As String
Dim iAsc As String

Randomize Timer

For C = 1 To Length

  iAsc = Int(3 * Rnd + 1)
  Select Case iAsc
    Case 1
      iChr = Int((Asc("Z") - Asc("A") + 1) * Rnd + Asc("A"))
    Case 2
      iChr = Int((Asc("z") - Asc("a") + 1) * Rnd + Asc("a"))
    Case 3
      iChr = Int((Asc("9") - Asc("0") + 1) * Rnd + Asc("0"))
    Case Else
      Err.Raise 200000, , "Error generating a password."
  End Select

  strResult = strResult & Chr(iChr)
Next C

PasswordGenerator = strResult

Exit_PasswordGenerator:
  Exit Function

Err_PasswordGenerator:
  MsgBox Err.Description
  PasswordGenerator = vbNullString
  Resume Exit_PasswordGenerator

End Function

Function IsLoaded(ByVal strFormName As String) As Integer
 ' Returns True if the specified form is open in Form view or Datasheet view.
    
    Const conObjStateClosed = 0
    Const conDesignView = 0
    
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
    
End Function

Function FetchBKDaysClosed(dateClosed, dateDismissed, dateDischarged, dateNoticeTerminating) As Long

FetchBKDaysClosed = 0
If Not IsNull(dateClosed) Then
  FetchBKDaysClosed = DateDiff("d", dateClosed, Date)
ElseIf Not IsNull(dateDismissed) Then
  FetchBKDaysClosed = DateDiff("d", dateDismissed, Date)
ElseIf Not IsNull(dateDischarged) Then
  FetchBKDaysClosed = DateDiff("d", dateDischarged, Date)
ElseIf Not IsNull(dateNoticeTerminating) Then
  FetchBKDaysClosed = DateDiff("d", dateNoticeTerminating, Date)

End If
End Function

Function FetchLastAction(dateReferral, datePub, dateSale, strState, action_flg) As Variant

Dim dateTemp As Date
Dim strLastAction As String

If Not IsNull(dateReferral) Then
  dateTemp = dateTemp
  strLastAction = "Referral"
End If

If (Not IsNull(datePub)) Then
  If (datePub > dateTemp) Then
    dateTemp = datePub
    strLastAction = "1st Pub"
  End If
End If

If (Not IsNull(dateSale)) Then
  If (dateSale > dateTemp) Then
    dateTemp = dateSale
    strLastAction = "Sale"
  End If
End If

If (action_flg = True) Then
  FetchLastAction = strLastAction
Else
  FetchLastAction = Format(dateTemp, "mm/dd/yyyy")
End If
  






End Function

Function GetLatestIntakeResponsibility(File)

Dim R As Recordset

GetLatestIntakeResponsibility = ""
Set R = CurrentDb.OpenRecordset("select Top 1 initials from qryFileResponsibilityHistory " & _
                                "where FileNumber = " & File & " and Department = 'Intake' " & _
                                "order by respdate desc")
If (Not R.EOF) Then
  GetLatestIntakeResponsibility = R![Initials]
End If
R.Close
Set R = Nothing



End Function

Function GetMediationLocation(JurisdictionID As Variant)

If IsNull(JurisdictionID) Then
  GetMediationLocation = ""
  Exit Function
End If


GetMediationLocation = Nz(DLookup("OAHAddress", "JurisdictionList", "JurisdictionID = " & JurisdictionID) & ", ", "") & _
                       Nz(DLookup("OAHCity", "JurisdictionList", "JurisdictionID = " & JurisdictionID) & ", ", "") & _
                       Nz(DLookup("OAHState", "JurisdictionList", "JurisdictionID = " & JurisdictionID), "")
                       
End Function

Public Sub FetchZipCodeCityState(strZipCode As String, objCity As Object, objState As Object)


Dim strCity As String

strCity = Nz(DLookup("City", "ZipCodes", "ZipCode = '" & Mid(strZipCode, 1, 5) & "' and Preferred = 'Yes'"))
If (Len(strCity) > 0) Then

  strCity = StrConv(strCity, vbProperCase)
  If (IsNull(objCity.Value)) Then
    objCity.Value = strCity
  ElseIf (objCity.Value <> strCity) Then
      If (MsgBox("The USPS City (" & strCity & ") does not match the city entered (" & objCity.Value & ").  Do you want to update the city?", vbYesNo, "City Update") = vbYes) Then
        objCity.Value = strCity
      End If
      
  End If
End If

' need to work on the jurisdiction/state relationships with the zipcode table

End Sub


Public Function CheckFutureDate(varDate As Variant)

Dim retval As Integer

retval = 0
If Not IsNull(varDate) Then
  If (varDate > Date) Then
    retval = 1
    MsgBox "Date cannot be in the future. ", vbCritical, "Date Error"
  End If
  
End If

CheckFutureDate = retval

End Function

Public Function FetchBarNumber(tid As Integer, tstate As String, Optional sta As String)

Dim barNum As Variant

Select Case tstate
    Case "MD"
    
        barNum = DLookup("[MDBar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchBarNumber = "MD Bar " & barNum
        Else
            If Not IsNull(sta) Then ' Using Md bar name even with Fed number. SA
            barNum = DLookup("[FedBar]", "[Staff]", "[ID] = " & tid)
            FetchBarNumber = "Md Bar " & barNum
            End If
        End If

        
    Case "VA"
        barNum = DLookup("[VABar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchBarNumber = "VA Bar " & barNum
        End If
    Case "DC"
        barNum = DLookup("[DCBar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchBarNumber = "DC Bar " & barNum
        End If

    Case "Fed"
        barNum = DLookup("[FedBar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchBarNumber = "Federal Bar " & barNum
        End If
    
    Case Else
        barNum = Null
        
End Select

If barNum = Null Then
  FetchBarNumber = ""
End If

End Function


Public Function FetchbarNumberPoundSign(tid As Integer, tstate As String, Optional sta As String)

Dim barNum As Variant

Select Case tstate
    Case "MD"
    
        barNum = DLookup("[MDBar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchbarNumberPoundSign = "MD BAR #" & barNum
        Else
            If Not IsNull(sta) Then ' Using Md bar name even with Fed number. SA
            barNum = DLookup("[FedBar]", "[Staff]", "[ID] = " & tid)
            FetchbarNumberPoundSign = "MD BAR #" & barNum
            End If
        End If

        
    Case "VA"
        barNum = DLookup("[VABar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchbarNumberPoundSign = "VA BAR # " & barNum
        End If
    Case "DC"
        barNum = DLookup("[DCBar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchbarNumberPoundSign = "DC BAR # " & barNum
        End If

    Case "Fed"
        barNum = DLookup("[FedBar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchbarNumberPoundSign = "FEDERAL BAR # " & barNum
        End If
    
    Case Else
        barNum = Null
        
End Select

If barNum = Null Then
  FetchbarNumberPoundSign = ""
End If

End Function 'This puts a # sign in before the Barnum
Public Function FetchBarNumberEvictions(tid As Integer, tstate As String, Optional sta As String)

Dim barNum As Variant

Select Case tstate
        
    Case "VA"
        barNum = DLookup("[VABar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchBarNumberEvictions = "VA BAR # " & barNum
        End If
    Case "DC"
        barNum = DLookup("[DCBar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchBarNumberEvictions = "DC BAR # " & barNum
        End If

    Case "Fed"
        barNum = DLookup("[FedBar]", "[Staff]", "[ID] = " & tid)
        If (Not IsNull(barNum)) Then
          FetchBarNumberEvictions = "FEDERAL BAR # " & barNum
        End If
    
    Case Else 'MD  Should be blank
        barNum = Null
        
End Select

If barNum = Null Then
  FetchBarNumberEvictions = ""
End If

End Function 'This puts a # sign in before the Barnum / Never lists a Bar number for MD

Public Function FetchNextMedHearing(ForeclosureID As Long)
Dim rs As Recordset
Dim sqlstr As String


FetchNextMedHearing = "No Mediation Date Scheduled"
sqlstr = "SELECT top 1 hearing FROM LMHearings where foreclosureid=" & ForeclosureID & " and datediff('d',date(),hearing) > 1 order by hearing asc"
Set rs = CurrentDb.OpenRecordset(sqlstr)

If Not rs.EOF Then

  FetchNextMedHearing = Format(rs![Hearing], "mmmm d, yyyy h:nn am/pm")
End If

rs.Close
Set rs = Nothing

End Function

Public Function FetchNoteOwner(LoanType As Integer, strInvestor As String, strState As String)
  FetchNoteOwner = strInvestor
  
  If (strState = "MD") Then
  
    If (LoanType = 5) Then
      FetchNoteOwner = "Federal Home Loan Mortgage Corporation"
    ElseIf (LoanType = 4) Then
      FetchNoteOwner = "Federal National Mortgage Association"
    End If
    
  End If
End Function

Public Function FetchNoteNames(LoanType As Integer, strInvestor As String, strState As String)

FetchNoteNames = GetNames(0, 2, "Noteholder=True") & " and " & strInvestor
  If (strState = "MD") Then
  
    If (LoanType = 5 Or LoanType = 4) Then
      FetchNoteNames = strInvestor
    End If
    
  End If

End Function

Public Sub mikki()
  MsgBox testthis
  
End Sub
Public Function DateChooserDialog(dt As Date, capt As String, Optional msg As String = "") As Date
    Dim fm As Form
    ' Shows the DateChooser dialog, and lets user choose date.
    ' Returns 0 if user cancels
  
    On Error Resume Next
    DoCmd.Close acForm, "DateChooser", acSaveNo
    On Error GoTo 0
    
    Form_DateChooser.dtDateChooserSet (dt)
    Form_DateChooser.strDateChooserCaptionSet (capt)
    If 0 < Len(msg) Then Form_DateChooser.strDateChooserTextSet (msg)
    
    DoCmd.OpenForm "DateChooser", acNormal, , , acFormEdit, acDialog
    DateChooserDialog = Form_DateChooser.dtDateChooserGet()
    
    Form_DateChooser.dtDateChooserSet (0)
    Form_DateChooser.strDateChooserCaptionSet ("")
    Form_DateChooser.strDateChooserTextSet ("")
    
    On Error Resume Next
    DoCmd.Close acForm, "DateChooser", acSaveNo
    Resume Next
    
    
    'DoCmd.DeleteObject acForm, "DateChooser"
    'acFormEdit, acDialog
    'fm.Show Modal
End Function

Public Function testthis()
 testthis = Nz(DLookup("[Location]", "[qryDocumentLocationList]", "[ID] = '1'"))
 
End Function

Public Function LastMonth(LPIDate As Date)

LastMonth = DateAdd("m", -1, [LPIDate])


'If Day(LPIDate) = 1 Then
'If Month(LPIDate) > 1 Then
'LastMonth = DateSerial(Year(LPIDate), Month(LPIDate) - 1, 1)
'Else
'LastMonth = DateSerial(Year(LPIDate) - 1, 12, 1)
'End If
'Else
'LastMonth = InputBox("Please enter the Interest as of date")
'End If


End Function


Public Function RemoveCR(strTestText As String) As String 'SA

Dim i As Integer, strTemp As String

For i = 1 To Len(strTestText)
If Mid$(strTestText, i, 1) <> Chr$(13) And Mid$(strTestText, i, 1) <> Chr$(10) Then
strTemp = strTemp & Mid$(strTestText, i, 1)
End If
Next i

RemoveCR = strTemp

End Function


Public Function HearingCheking(HearingApp As Date, Akind As Integer) ' cheacking hearing date SA.
Dim HCancel, DateHearing, DayHearing, TimeHearing As Integer
HCancel = 0
DateHearing = 0
DayHearing = 0
TimeHearing = 0
If Not IsNull(HearingApp) Then

Select Case Akind

Case 1
    If Now() > HearingApp Then
    MsgBox ("Hearing date must be in the future")
    DateHearing = 1
    End If

Case 2
    If Weekday(HearingApp) = vbSunday Or Weekday(HearingApp) = vbSaturday Then
    MsgBox ("Exceptions Hearing date cannot be Saturday or Sunday"), vbCritical
    DayHearing = 1
    End If

Case 3
    If Hour(HearingApp) < 8 Or Hour(HearingApp) > 18 Then
    MsgBox ("Invalid Exceptions Hearing time: " & Format$(HearingApp, "h:nn am/pm"))
    TimeHearing = 1
    End If

End Select


If DateHearing = 1 Or DayHearing = 1 Or TimeHearing = 1 Then HCancel = 1
HearingCheking = HCancel
End If

End Function


Public Function NumSuffix(MyNum As Variant) As String

Dim n      As Integer
Dim X      As Integer
Dim strSuf As String

    n = Right(MyNum, 2)
    X = n Mod 10
    strSuf = Switch(n <> 11 And X = 1, "st", n <> 12 And X = 2, "nd", _
             n <> 13 And X = 3, "rd", True, "th")
    NumSuffix = LTrim(str(MyNum)) & strSuf

End Function

Public Function Workdays(startDate As Date, endDate As Date, Optional strHolidays As String = "Holidays") As Integer
    ' Returns the number of workdays between startDate
    ' and endDate inclusive.  Workdays excludes weekends and
    ' holidays. Optionally, pass this function the name of a table
    ' or query as the third argument. If you don't the default
    ' is "Holidays".
    On Error GoTo Workdays_Error
    Dim nWeekdays As Integer
    Dim nHolidays As Integer
    Dim strWhere As String
   
    startDate = DateValue(startDate)
    endDate = DateValue(endDate)
    
    nWeekdays = Weekdays(startDate, endDate)
    If nWeekdays = -1 Then
        Workdays = -1
        GoTo Workdays_Exit
    End If
    
    strWhere = "Holiday >= #" & startDate _
        & "# AND Holiday <= #" & endDate & "#"
    
    ' Count the number of holidays.
    nHolidays = DCount(Expr:="Holiday", _
        Domain:=strHolidays, _
        Criteria:=strWhere)
    
    Workdays = nWeekdays - nHolidays
    
    
Workdays_Exit:
    Exit Function
    
Workdays_Error:
    Workdays = -1
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, "Workdays"
    Resume Workdays_Exit
    
End Function

Public Function Weekdays(startDate As Date, endDate As Date) As Integer
    ' Returns the number of weekdays in the period from startDate
    ' to endDate inclusive. Returns -1 if an error occurs.
    ' If your weekend days do not include Saturday and Sunday and
    ' do not total two per week in number, this function will
    ' require modification.
    On Error GoTo Weekdays_Error
    
    ' The number of weekend days per week.
    Const ncNumberOfWeekendDays As Integer = 2
    
    ' The number of days inclusive.
    Dim varDays As Variant
    
    ' The number of weekend days.
    Dim varWeekendDays As Variant
    
    ' Temporary storage for datetime.
    Dim dtmX As Date
    
    ' If the end date is earlier, swap the dates.
'    If endDate < startDate Then
'        dtmX = startDate
'        startDate = endDate
'        endDate = dtmX
'    End If
    
    ' Calculate the number of days inclusive (+ 1 is to add back startDate).
    Select Case endDate
    Case Is < startDate
        varDays = DateDiff(Interval:="d", _
        date1:=startDate - 1, _
        date2:=endDate) - 1
    Case Is > startDate
    varDays = DateDiff(Interval:="d", _
        date1:=startDate + 1, _
        date2:=endDate) + 1
    Case 0
    varDays = 0
    End Select
    
    ' Calculate the number of weekend days.
    varWeekendDays = (DateDiff(Interval:="ww", _
        date1:=startDate, _
        date2:=endDate) _
        * ncNumberOfWeekendDays) _
        + IIf(DatePart(Interval:="w", _
        Date:=startDate) = vbSunday, 1, 0) _
        + IIf(DatePart(Interval:="w", _
        Date:=endDate) = vbSaturday, 1, 0)
    
    ' Calculate the number of weekdays.
    Weekdays = (varDays - varWeekendDays)
    
Weekdays_Exit:
    Exit Function
    
Weekdays_Error:
    Weekdays = -1
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, "Weekdays"
    Resume Weekdays_Exit
End Function

Public Function FunYesNo(QuesText As String) As Boolean
 If MsgBox(QuesText, vbYesNo) = vbYes Then
 FunYesNo = True
 Else
 FunYesNo = False
 End If

 
End Function

Public Function CopyNoR() As Integer ' for military affidiavit reports SA
CopyNoR = CopyNo
End Function

Public Function IsLoadedF(ByVal strFormName As String) As Boolean
 ' Returns True if the specified form is open in Form view or Datasheet view.
    
    Const conObjStateClosed = 0
    Const conDesignView = 0
    
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoadedF = True
            Else
            IsLoadedF = False
        End If
    End If
    
End Function


Public Function ClientShortName(ClientIDF As Integer) As String
Dim Shortclient As String

Shortclient = DLookup("ShortClientName", "ClientList", "ClientID= " & ClientIDF)
ClientShortName = Shortclient

End Function


Public Function CloseAllForms()
'It will close all forms before opening the new form (if required)

Dim obj As Object
Dim strName As String

For Each obj In Application.CurrentProject.AllForms
    DoCmd.Close acForm, obj.Name, acSaveYes
Next obj

End Function

Public Function CloseAllReports()
'It will close all reports before opening the new report (if required)

Dim obj As Object
Dim strName As String

For Each obj In Application.CurrentProject.AllReports
    'DoCmd.Close acForm, obj.Name, acSaveYes
    DoCmd.Close acReport, obj.Name, acSaveYes
Next obj

End Function

Public Function CheckNameEdit()
Dim R1 As String
Dim R2 As String
Dim RC1 As String
Dim RC2 As String
Dim IsComplete As Boolean
Dim IsFormOpen As Boolean
Dim R As Recordset
IsComplete = False

If IsLoadedF("wizReferralII") = True Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
End If

If IsLoadedF("WizDemand") = True Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
End If

If IsLoadedF("wizFairDebt") = True Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
End If

If IsLoadedF("wizNOI") = True Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
End If

If IsLoadedF("wizRestartFCdetails1") = True Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
End If


If IsLoadedF("ForeclosureDetails") = True Then
If (Forms!foreclosuredetails!WizardSource) = "Restart" Then
IsComplete = True
CheckNameEdit = IsComplete
Exit Function
Else

If IsNull(Forms!foreclosuredetails!WizardSource) Then

 Set R = CurrentDb.OpenRecordset("Select * From WizardQueueStats WHERE FileNumber = " & Forms!foreclosuredetails.[FileNumber] & " And Current = True", dbOpenDynaset, dbSeeChanges)
    If Not IsNull(R!RSIcomplete) Or Not IsNull(R!RSIcomplete) Or Not IsNull(R!RestartRSIComplete) Or Not IsNull(R!RestartComplete) Then
    IsComplete = True
    CheckNameEdit = IsComplete
    Exit Function

End If
End If
End If
End If


End Function


Public Function LastDay90(dteAny As Date) As Date

Dim days90 As Date
days90 = DateAdd("d", 90, dteAny)

LastDay90 = DateSerial(Year(days90), Month(days90) + 1, 1) - 1
End Function



Public Sub StaffSignIn()

Dim rstUser As Recordset
Set rstUser = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID=" & StaffID, dbOpenDynaset, dbSeeChanges)
If Not rstUser.EOF Then
  With rstUser
  .Edit
  !InOrOut = True
  .Update
  End With
Set rstUser = Nothing
End If

End Sub

Public Function StaffCheck()

Dim rstUser As Recordset
Set rstUser = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID=" & StaffID, dbOpenDynaset, dbSeeChanges)
If Not rstUser.EOF Then
If rstUser!InOrOut = True Then
StaffCheck = True
Else
StaffCheck = False
End If
End If
Set rstUser = Nothing

End Function

Public Sub StaffSignOut()

Dim rstUser As Recordset
Set rstUser = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID=" & StaffID, dbOpenDynaset, dbSeeChanges)
If Not rstUser.EOF Then
  With rstUser
  .Edit
  !InOrOut = False
  .Update
  End With
Set rstUser = Nothing
End If

End Sub



Public Sub RemoveDates()

    If Not IsNull(Forms!foreclosuredetails!FirstPub.Value) Then Forms!foreclosuredetails!FirstPub.Value = Null
    If Not IsNull(Forms!foreclosuredetails!IRSNotice.Value) Then Forms!foreclosuredetails!IRSNotice.Value = Null
    If Not IsNull(Forms!foreclosuredetails!Notices.Value) Then Forms!foreclosuredetails!Notices.Value = Null
    If Not IsNull(Forms!foreclosuredetails!UpdatedNotices.Value) Then Forms!foreclosuredetails!UpdatedNotices.Value = Null
    If Not IsNull(Forms!foreclosuredetails!BidReceived.Value) Then Forms!foreclosuredetails!BidReceived.Value = Null
    If Not IsNull(Forms!foreclosuredetails!BidAmount.Value) Then Forms!foreclosuredetails!BidAmount.Value = Null
    If Not IsNull(Forms!foreclosuredetails!PayoffAmount.Value) Then Forms!foreclosuredetails!PayoffAmount.Value = Null
    If Not IsNull(Forms!foreclosuredetails!Sale.Value) Then Forms!foreclosuredetails!Sale.Value = Null
    If Not IsNull(Forms!foreclosuredetails!SaleTime.Value) Then Forms!foreclosuredetails!SaleTime.Value = Null
    If Not IsNull(Forms!foreclosuredetails!Deposit.Value) Then Forms!foreclosuredetails!Deposit.Value = Null
    If Not IsNull(Forms!foreclosuredetails!SaleCert.Value) Then Forms!foreclosuredetails!SaleCert.Value = Null
    If Not IsNull(Forms!foreclosuredetails!ReviewAdProof.Value) Then Forms!foreclosuredetails!ReviewAdProof.Value = Null
      


End Sub
Public Sub RemoveSoftHold(FileNumber)
Dim rstFCDIL As Recordset
Dim strInsert As String

Set rstFCDIL = CurrentDb.OpenRecordset("SELECT * FROM FCDIL WHERE FileNumber=" & FileNumber & " And Not isNull(SoftDate) ", dbOpenDynaset, dbSeeChanges)
If Not rstFCDIL.EOF Then


DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT into SoftHoldArchive (FileNumber, DateInsert, HowInsert, SoftID, SoftDate, SoftStaffId, SoftStaffInitial) Values(" & FileNumber & ",Now(),HowInsert, SoftID, SoftDate, SoftStaffId, SoftStaffInitial)"
DoCmd.RunSQL strInsert
DoCmd.SetWarnings True


  With rstFCDIL
  .Edit
   !SoftID = Null
   !SoftDate = Null
   !SoftStaffId = Null
   !SoftStaffInitial = Null
  .Update
  End With
Set rstFCDIL = Nothing
Else
Exit Sub
End If
End Sub

Public Function CheckAttribut(FileNumber As Long, attribut As Long) As Boolean
Dim rs As Recordset
Dim sqlstr As String
sqlstr = "SELECT top 1 Warning FROM Journal where FileNumber=" & FileNumber & " order by Journaldate DESC"
Set rs = CurrentDb.OpenRecordset(sqlstr)
    If Not rs.EOF Then
    
        If attribut = rs!Warning Then
        CheckAttribut = True
        Else
        CheckAttribut = False
        End If
        Exit Function
    Else
    CheckAttribut = False
    End If
rs.Close
Set rs = Nothing
End Function


Public Function CheckIfFileWasFCFirst(FileNumber As Long) As Boolean
Dim Bfile As Recordset 'Pending
Dim Pfile As Recordset 'Bk
Dim sqlstrBK As String
Dim sqlstrP As String

sqlstrBK = "SELECT top 1 StatusDate FROM StatusList where FileNumber=" & FileNumber & " And StatusDesc = 'File type changed to BK' order by EntryTime ASC"
Set Bfile = CurrentDb.OpenRecordset(sqlstrBK)
    If Not Bfile.EOF Then
    sqlstrP = "SELECT top 1 StatusDate FROM StatusList where FileNumber=" & FileNumber & " And StatusDesc = 'File type changed to Pending' order by EntryTime ASC"
    Set Pfile = CurrentDb.OpenRecordset(sqlstrP)
    
        If Bfile!StatusDate < Pfile!StatusDate Then
        CheckIfFileWasFCFirst = True
        Exit Function
        Else
        CheckIfFileWasFCFirst = False
        End If
    Else
    CheckIfFileWasFCFirst = False
    End If
    
        
        
Pfile.Close
Bfile.Close

Set Pfile = Nothing
Set Bfile = Nothing
'vbYellow
'BackStyle

End Function

Public Sub VisibleDCForeclosureDetailsForm()

Forms!foreclosuredetails!pgNOI.Visible = False
Forms!foreclosuredetails!CourtCaseNumber.Locked = False
Forms!foreclosuredetails!OptionCoop.Visible = True
Forms!foreclosuredetails!Coop.Visible = True
Forms!foreclosuredetails!sfrmDCComplaint.Visible = True
Forms!foreclosuredetails!sfrmDCNotices.Visible = True
'Forms!ForeclosureDetails!lblCaseCaption.Visible = True
'Forms!ForeclosureDetails!sfrmCaseCaption.Visible = True
Forms!foreclosuredetails!sfrmLisPendens.Visible = True
Forms!foreclosuredetails!NOI.Enabled = False
Forms!foreclosuredetails!NOI.BackStyle = 0 'Transparent
Forms!foreclosuredetails.txtClientSentNOI.Enabled = False
Forms!foreclosuredetails.txtClientSentNOI.BackStyle = 0 'Transparent

Forms!foreclosuredetails!DocBackSOD.Visible = False
Forms!foreclosuredetails!DocBackLossMitPrelim.Visible = False
Forms!foreclosuredetails!DocBackLossMitFinal.Visible = False
Forms!foreclosuredetails!DocBackLostNote.Visible = False
Forms!foreclosuredetails!DocBackNoteOwnership.Visible = False
Forms!foreclosuredetails!DocBackAff7105.Visible = False
Forms!foreclosuredetails!StatementOfDebtDate.Visible = False
Forms!foreclosuredetails!StatementOfDebtAmount.Visible = False
Forms!foreclosuredetails!StatementOfDebtPerDiem.Visible = False
Forms!foreclosuredetails!LostNoteAffSent.Visible = False
Forms!foreclosuredetails!LostNoteNotice.Visible = False
Forms!foreclosuredetails!Docket.Visible = False
Forms!foreclosuredetails!FLMASenttoCourt.Visible = False
Forms!foreclosuredetails!LossMitFinalDate.Visible = False
Forms!foreclosuredetails!ServiceMailed.Visible = False
Forms!foreclosuredetails!SentToDocket.Visible = False

Forms!foreclosuredetails!ServiceSent.Enabled = True
Forms!foreclosuredetails!ServiceSent.Locked = False
Forms!foreclosuredetails!BorrowerServed.Enabled = True
Forms!foreclosuredetails!BorrowerServed.Locked = False

End Sub

Public Sub ExportDatabaseObjects()
On Error GoTo Err_ExportDatabaseObjects
    
    Dim db As Database
    'Dim db As DAO.Database
    Dim td As TableDef
    Dim d As Document
    Dim C As Container
    Dim i As Integer
    Dim sExportLocation As String
    
    Set db = CurrentDb()
    
    sExportLocation = "C:\Temp\" & DBVersion & "\" 'Do not forget the closing back slash! ie: C:\Temp\
    
    If Len(Dir(sExportLocation, vbDirectory)) = 0 Then
            MkDir sExportLocation
    End If
    
    'For Each td In db.TableDefs 'Tables
    '    If Left(td.Name, 4) <> "MSys" Then
    '        DoCmd.TransferText acExportDelim, , td.Name, sExportLocation & "Table_" & td.Name & ".txt", True
    '    End If
    'Next td
    
    Set C = db.Containers("Forms")
    For Each d In C.Documents
        Application.SaveAsText acForm, d.Name, sExportLocation & "Form_" & d.Name & ".txt"
    Next d
    
    Set C = db.Containers("Reports")
    For Each d In C.Documents
        Application.SaveAsText acReport, d.Name, sExportLocation & "Report_" & d.Name & ".txt"
    Next d
    
    Set C = db.Containers("Scripts")
    For Each d In C.Documents
        Application.SaveAsText acMacro, d.Name, sExportLocation & "Macro_" & d.Name & ".txt"
    Next d
    
    Set C = db.Containers("Modules")
    For Each d In C.Documents
        Application.SaveAsText acModule, d.Name, sExportLocation & "Module_" & d.Name & ".txt"
    Next d
    
    For i = 0 To db.QueryDefs.Count - 1
        Application.SaveAsText acQuery, db.QueryDefs(i).Name, sExportLocation & "Query_" & db.QueryDefs(i).Name & ".txt"
    Next i
    
    Set db = Nothing
    Set C = Nothing
    
    MsgBox "All database objects have been exported as a text file to " & sExportLocation, vbInformation
    
Exit_ExportDatabaseObjects:
    Exit Sub
    
Err_ExportDatabaseObjects:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_ExportDatabaseObjects
    
End Sub


Public Function Unlockedfiles(StaffID As Long)
Dim strUnlocked As String
strUnlocked = "Update Locks Set StaffId = 0 Where StaffID= " & StaffID
RunSQL (strUnlocked)
End Function

Public Function GetStaffConducted(StState As String, SaleConductedTrusteeID As String) As String
Select Case StState
Case "VA"

If SaleConductedTrusteeID = "" Then
'If IsNull([SaleConductedTrusteeID]) Then
    GetStaffConducted = ""
Else
    If Right([SaleConductedTrusteeID], 2) = ".5" Then
    GetStaffConducted = DLookup("[VendorName]", "[Vendors]", "ID=" & Left(SaleConductedTrusteeID, InStr(SaleConductedTrusteeID, ".") - 1))
    Else
    GetStaffConducted = DLookup("[Name]", "[Staff]", "ID=" & SaleConductedTrusteeID)
    End If
End If

Case "MD", "DC"
If SaleConductedTrusteeID = "" Then

'If IsNull([SaleConductedTrusteeID]) Then
    GetStaffConducted = ""
Else
  GetStaffConducted = DLookup("[Name]", "[Staff]", "ID=" & SaleConductedTrusteeID)
End If

End Select


End Function

Function StripAllChars(strString As String) As String
'Return only numeric values from a string
    Dim lngCtr      As Long
    Dim intChar     As Integer
 
    For lngCtr = 1 To Len(strString)
        intChar = Asc(Mid(strString, lngCtr, 1))
        If intChar >= 48 And intChar <= 57 Then
            StripAllChars = StripAllChars & Chr(intChar)
        End If
    Next lngCtr
End Function
 
Function FormatPhone(strIn As String) As Variant
On Error Resume Next
 
    strIn = StripAllChars(strIn)
 
    If InStr(1, strIn, "@") >= 1 Then
        FormatPhone = strIn
        Exit Function
    End If
 
    Select Case Len(strIn & vbNullString)
        Case 0
            FormatPhone = Null
        Case 7
            FormatPhone = Format(strIn, "@@@-@@@@")
        Case 10
            FormatPhone = Format(strIn, "(@@@) @@@-@@@@")
        Case 11
            FormatPhone = Format(strIn, "@ (@@@) @@@-@@@@")
        Case Else
            FormatPhone = strIn
    End Select
End Function


Public Function Businessdays(EnterDate As Date, DaysN As Integer, Optional A As Integer) As Integer
Dim J As Integer
J = 0
Dim DateT As Date
    If A = 1 Then
    A = 1
    Else: A = -1
    End If

Do While J < (DaysN + 1)
        DateT = (DateAdd("d", A * (J), EnterDate))
        If DatePart("w", DateT) = 1 Or DatePart("w", DateT) = 7 Or Not IsNull(DLookup("Desc", "holidays", "Holiday=#" & DateT & "#")) Then
        DaysN = DaysN + 1
        End If
J = J + 1
Loop

Businessdays = DaysN

End Function

Public Function AddToList(FileNumber As Long)
Dim TimesCount As Integer
Dim strupdate As String
Dim FileN As Long
Dim UserID As Integer
FileN = FileNumber
UserID = GetStaffID


TimesCount = DCount("[RecentID]", "Recent", "[FileNumber] = " & FileNumber & " And StaffId= " & UserID & "")

DoCmd.SetWarnings False
 If TimesCount <> 0 Then
      DoCmd.RunSQL "UPDATE Recent Set AccessTime = #" & Now() & "# WHERE StaffID =" & UserID & " And filenumber = " & FileNumber '& " ORDER By AccessTime DESC"
 Else
    strupdate = "Insert Into Recent (FileNumber, AccessTime, StaffID) Values (" & FileNumber & ", #" & Now() & "#, " & UserID & ")"
    DoCmd.RunSQL (strupdate)

End If
DoCmd.SetWarnings True


'Call ExecuteSQL(strupdate)

'Set recent = CurrentDb.OpenRecordset("SELECT * FROM Recent where StaffID=" & GetStaffID() & " AND filenumber=" & FileNumber & " ORDER BY AccessTime DESC", dbOpenDynaset, dbSeeChanges)
'
'' Attempt to find the case number in the recent list.
''
'If recent.EOF Then 'not in recent list
'    recent.AddNew
'    recent!FileNumber = FileNumber
'    recent!AccessTime = Now()
'    recent!StaffID = StaffID
'    recent.Update
'Else                    ' found in recent list, update access time
'    recent.Edit
'    recent!AccessTime = Now()
'    recent.Update
'End If

End Function

Public Function SQLRun(strSQL As String) As Long
  On Error GoTo ErrHandler
    Static db As DAO.Database
MsgBox strSQL
    If db Is Nothing Then
       Set db = CurrentDb
    End If
    db.Execute strSQL, dbFailOnError
    SQLRun = db.RecordsAffected

exitRoutine:
    Exit Function

ErrHandler:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Error in SQLRun()"
    Resume exitRoutine
  End Function


Sub ExecuteSQL(LSQL)

   Dim db As Database
  ' Dim LSQL As String
   
   Set db = CurrentDb()
  ' LSQL = "Update suppliers set supplier_name = 'IBM'"
   
   db.Execute LSQL
   
   MsgBox CStr(db.RecordsAffected) & " records were affected by this SQL statement."
   
End Sub

Public Function GetDismissalDocIDNo(ID As Long, DocNo As Long, FileNO As Long) As Long

GetDismissalDocIDNo = 0
Dim lrs As Recordset
                                
Set lrs = CurrentDb.OpenRecordset("select DocID from DocIndex " & _
                                "where FileNumber = " & FileNO & " and DocTitleID = " & DocNo & " and StaffID = " & ID & " order by DocID desc", dbOpenDynaset, dbSeeChanges)

If Not lrs.EOF Then
GetDismissalDocIDNo = lrs![DocID]
End If

lrs.Close
Set lrs = Nothing

End Function

Public Sub ForeclosureDetailsDCTabSwitchON()

'Forms!ForeclosureDetails!pgNOI.Visible = False
Forms!foreclosuredetails!txtFairDebt.Locked = True
Forms!foreclosuredetails!txtAccelerationIssued.Locked = True
Forms!foreclosuredetails!txtAccelerationLetter.Locked = True
Forms!foreclosuredetails!txtDocstoClient.Locked = True
'Forms!ForeclosureDetails!sfrmDCNotices.Visible = True
Forms!foreclosuredetails!txtTitleDue.Locked = True
Forms!foreclosuredetails!txtHUDOccLetter.Locked = True
Forms!foreclosuredetails!txt567.Locked = True
Forms!foreclosuredetails!txtDocsBack.Locked = True

'------------- done above

Forms!foreclosuredetails!txtTitleOrder.Locked = True
Forms!foreclosuredetails!txtTitleOrder.Locked = True

Forms!foreclosuredetails!DocBackSOD.Visible = False
Forms!foreclosuredetails!DocBackLossMitPrelim.Visible = False
Forms!foreclosuredetails!DocBackLossMitFinal.Visible = False
Forms!foreclosuredetails!DocBackLostNote.Visible = False
Forms!foreclosuredetails!DocBackNoteOwnership.Visible = False
Forms!foreclosuredetails!DocBackAff7105.Visible = False
Forms!foreclosuredetails!StatementOfDebtDate.Visible = False
Forms!foreclosuredetails!StatementOfDebtAmount.Visible = False
Forms!foreclosuredetails!StatementOfDebtPerDiem.Visible = False
Forms!foreclosuredetails!LostNoteAffSent.Visible = False
Forms!foreclosuredetails!LostNoteNotice.Visible = False
Forms!foreclosuredetails!Docket.Visible = False
Forms!foreclosuredetails!FLMASenttoCourt.Visible = False
Forms!foreclosuredetails!LossMitFinalDate.Visible = False
Forms!foreclosuredetails!ServiceMailed.Visible = False
Forms!foreclosuredetails!SentToDocket.Visible = False

Forms!foreclosuredetails!ServiceSent.Enabled = True
Forms!foreclosuredetails!ServiceSent.Locked = False
Forms!foreclosuredetails!BorrowerServed.Enabled = True
Forms!foreclosuredetails!BorrowerServed.Locked = False

End Sub

Public Sub ExportFile()
Dim strPath As String


DoCmd.TransferText acExportDelim, "Event_Header_Final Export Specification", "Event_Header_Final", "\\FileServer\Applications\Database\FNAMA\databaseheader.txt"
DoCmd.TransferText acExportDelim, "Event_All_F1 Export Specification", "Event_All_F1", "\\FileServer\Applications\Database\FNAMA\databaseEvent.txt"
DoCmd.TransferText acExportDelim, "Event_TRAILER_F Export Specification", "Event_TRAILER_F", "\\FileServer\Applications\Database\FNAMA\databasetrailer.txt"
''
''Dim strPath As String
'
'
'strPath = "\\FileServer\Applications\Database\FNAMA\"   'don't forget the closing \!
'
'Shell "cmd.exe copy /c """ & strPath & "*.txt"" """ & strPath & "Consolidated.txt""", 0

End Sub
Sub AppendDetails()

      Dim SourceNum As Integer
      Dim DestNum As Integer
      Dim Temp As String

      ' If an error occurs, close the files and end the macro.
      On Error GoTo ErrHandler

      ' Open the destination text file.
      DestNum = FreeFile()
      Open "C:\Users\sarab\Desktop\Header.txt" For Append As DestNum

      SourceNum = FreeFile()
'
        Open "C:\Users\sarab\Desktop\Details.txt" For Input As SourceNum

      ' Include the following line if the first line of the source
      ' file is a header row that you do now want to append to the
      ' destination file:
      ' Line Input #SourceNum, Temp

      ' Read each line of the source file and append it to the
      ' destination file.
      Do While Not EOF(SourceNum)
         Line Input #SourceNum, Temp
         Print #DestNum, Temp
      Loop

CloseFiles:

      ' Close the destination file and the source file.
      Close #DestNum
      Close #SourceNum
      Exit Sub

ErrHandler:
      MsgBox "Error # " & Err & ": " & Error(Err)
      Resume CloseFiles

   End Sub
   
Sub AppendTRAILER()

      Dim SourceNum As Integer
      Dim DestNum As Integer
      Dim Temp As String

      ' If an error occurs, close the files and end the macro.
      On Error GoTo ErrHandler

      ' Open the destination text file.
      DestNum = FreeFile()
      Open "C:\Users\sarab\Desktop\Header.txt" For Append As DestNum

      SourceNum = FreeFile()
'
        Open "C:\Users\sarab\Desktop\Details.txt" For Input As SourceNum

      ' Include the following line if the first line of the source
      ' file is a header row that you do now want to append to the
      ' destination file:
      ' Line Input #SourceNum, Temp

      ' Read each line of the source file and append it to the
      ' destination file.
      Do While Not EOF(SourceNum)
         Line Input #SourceNum, Temp
         Print #DestNum, Temp
      Loop

CloseFiles:

      ' Close the destination file and the source file.
      Close #DestNum
      Close #SourceNum
      Exit Sub

ErrHandler:
      MsgBox "Error # " & Err & ": " & Error(Err)
      Resume CloseFiles

   End Sub

