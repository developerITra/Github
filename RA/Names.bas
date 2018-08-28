Attribute VB_Name = "Names"
Option Compare Database
Option Explicit
'Mei for Prior Servicer form  9/24/15
Public strPriorServicer As String
Public bHolder As Boolean
Public bReferee As Boolean
Public bLost As Boolean
Public bPrior As Boolean
Public prntTo As Integer
Dim CapsFmt As Integer
Public Function BankoModDate() As String
Dim oFS As Object
Dim strFileName As String

    'Put your filename here
    strFileName = "\\fileserver\Applications\Database\BankruptcyDailyChecks\To_LexisNexis\RSBD.r01"


    'This creates an instance of the MS Scripting Runtime FileSystemObject class
    Set oFS = CreateObject("Scripting.FileSystemObject")
    BankoModDate = oFS.GetFile(strFileName).DateLastModified
    

Set oFS = Nothing
    
End Function
Public Function CountNames(File As Long, conditions As String) As Integer
Dim n As Recordset

If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ")", dbOpenSnapshot)
If Not n.EOF Then n.MoveLast
CountNames = n.RecordCount
n.Close

End Function

Public Function MortgagorNamesCaps(File As Long, Fmt As Integer, Caps As Integer) As String
'
' caps: 1 = upcase entire name
'       2 = upcase last name only
'
CapsFmt = Caps
MortgagorNamesCaps = MortgagorNames(File, Fmt)
CapsFmt = 0
End Function
Public Function BorrowerOwnerNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
BorrowerOwnerNames = GetNames(File, Fmt, "Noteholder = True or Owner = True")

End Function

Public Function MortgagorNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
MortgagorNames = GetNames(File, Fmt, "Mortgagor = True", NewLine, NoSSN)

End Function

Public Function DefendantNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
DefendantNames = GetNames(File, Fmt, "Defendant = True", NewLine, NoSSN)

End Function

Public Function COSNamesAddress(File As Long, Optional NewLine As String = vbNewLine) As String
COSNamesAddress = GetAddresses(File, 5, "COS=True", NewLine)
End Function

Public Function COSNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
COSNames = GetNames(File, Fmt, "COS = True", NewLine, NoSSN)

End Function

Public Function ActiveDutyNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
ActiveDutyNames = GetNames(File, Fmt, "ActiveDuty = True", NewLine, NoSSN)

End Function

Public Function MortgagorNamesLast(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
MortgagorNamesLast = GetNames(File, Fmt, "Mortgagor = True", NewLine, NoSSN)

End Function


Public Function MortgagorOwnerNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
MortgagorOwnerNames = GetNames(File, Fmt, "Mortgagor = True or Owner = True")

End Function

Public Function MortgagorNamesOneline(File As Long, Fmt As Integer) As String
MortgagorNamesOneline = GetNames(File, Fmt, "Mortgagor = True")

End Function

Public Function BorrowerMortgagorOwnerName(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
BorrowerMortgagorOwnerName = GetNames(File, Fmt, "Mortgagor = True Or Owner = True Or Noteholder = True")

End Function


Public Function MortgagorNamesIntext(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean, Optional Intext As Boolean) As String
MortgagorNamesIntext = GetNames(File, Fmt, "Mortgagor = True", NewLine, NoSSN, Intext)

End Function

Public Function DebtorNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine) As String
DebtorNames = GetNames(File, Fmt, "BKDebtor = True", NewLine)
End Function

Public Function OwnerNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine) As String
OwnerNames = GetNames(File, Fmt, "Owner = True", NewLine)
End Function

Public Function CoDebtorAndDebtorNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine) As String
CoDebtorAndDebtorNames = GetNames(File, Fmt, "BKDebtor = True or BKCoDebtor = True", NewLine)
End Function
Public Function NoteholderNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine) As String
NoteholderNames = GetNames(File, Fmt, "Noteholder = True", NewLine)
End Function


Public Function TRCIVDebtorNames(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine) As String
  TRCIVDebtorNames = GetTRCIVNames(File, Fmt, "Debtor = True", NewLine)
End Function

Public Function FairDebtBorrowerNames(File As Long) As String
' BorrowerNames = GetNames(File, Fmt, "Noteholder = True", NewLine)


Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND Noteholder = True ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop

Dim J As Integer, DataArray() As Variant, ctr As Integer, NameFlag As Boolean
ReDim DataArray(cnt, 1)
ctr = 1
For i = 1 To cnt
NameFlag = False

    For J = 1 To cnt
    If DataArray(J, 1) = Names(i) Then
    NameFlag = True
    End If
    Next J
    If NameFlag = False Then
    DataArray(ctr, 1) = Names(i)
    ctr = ctr + 1
    
    FairDebtBorrowerNames = FairDebtBorrowerNames & IIf(ctr > 2, " and ", "") & Names(i)
    
    NameFlag = True
    End If
    
    
Next i
End Function

Public Function BorrowerNames(File As Long) As String
' BorrowerNames = GetNames(File, Fmt, "Noteholder = True", NewLine)


Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND Noteholder = True ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop

Dim J As Integer

J = 0
For i = 1 To cnt
    If (InStr(1, BorrowerNames, Names(i)) = 0) Then
      J = J + 1
      BorrowerNames = BorrowerNames & IIf(J > 1, " and ", "") & Names(i)
    End If
Next i
End Function

Public Function GetAddresses(File As Long, Fmt As Integer, conditions As String, Optional NewLine As String = vbNewLine, Optional IdAddress As Integer) As String
'
' File: File number, or 0 for 'current' file
' fmt:  3 = 'normal' list, with name in caps
'       4 = captions
'       5 = 'normal' list
'       6 = list with each on one line
'
Dim n As Recordset
Dim m As Recordset
Dim cnt As Integer
Dim cntalt As Integer
Dim i As Integer
Dim J As Integer
Dim G As Boolean
Dim p As Integer
Dim l As Integer
Dim R As Integer
Dim Addr() As String
Dim namebor() As String
Dim AddrAlt() As String
Dim myCol As Collection
Dim dup As Integer
Dim ss As String
Dim Y As Integer
Dim addraltcheckN() As String
Dim newArray(20) As String


cnt = 0
If File = 0 Then
    File = Forms![Case List]!FileNumber
    Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ") ORDER BY BKCoDebtor DESC, ID, Last, First", dbOpenSnapshot)
ElseIf File = 150 Then
    File = Forms![Case List]!FileNumber
    Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " ORDER BY BKCoDebtor DESC, ID, Last, First", dbOpenSnapshot)
Else
    Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ") ORDER BY BKCoDebtor DESC, ID, Last, First", dbOpenSnapshot)
End If

If n.EOF Then
    GetAddresses = ""
    Exit Function
End If
n.MoveLast
cnt = n.RecordCount
ReDim Addr(1 To cnt) As String
ReDim addraltcheckN(1 To cnt) As String
n.MoveFirst

Select Case Fmt
    Case 3      ' normal with caps
        For i = 1 To cnt
           GetAddresses = GetAddresses & FormatName(UCase$(Nz(n("Company"))), UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), NewLine)
            If i < cnt Then GetAddresses = GetAddresses & NewLine & NewLine
            n.MoveNext
        Next i
    Case 4      ' caption
        For i = 1 To cnt
            Addr(i) = n("Address") & n("Address2") & n("City") & n("State") & n("Zip")
            n.MoveNext
        Next i
        n.MoveFirst
        For i = 1 To cnt
            If i < cnt Then
                If Addr(i) = Addr(i + 1) Then
                    If Application.CurrentProject.AllReports("45 day notice wizTest").IsLoaded = True Or Application.CurrentProject.AllReports("certificate of service").IsLoaded = True Or Application.CurrentProject.AllReports("withdraw sale order").IsLoaded = True Then
                    GetAddresses = GetAddresses & FormatName(UCase$(Nz(n("Company"))), _
                        IIf(n("Deceased") = True, "Estate of " & UCase$(Nz(n("First"))), UCase$(Nz(n("First")))), UCase$(Nz(n("Last"))), n("AKA"), _
                        Null, Null, Null, Null, Null, NewLine)
                    Else
                    GetAddresses = GetAddresses & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), _
                        Null, Null, Null, Null, Null, NewLine)
                    End If
                Else
                    GetAddresses = GetAddresses & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), n("Address"), _
                        n("Address2"), n("City"), n("State"), n("Zip"), NewLine) & "and" & NewLine
                End If
            Else
            If Application.CurrentProject.AllReports("45 day notice wizTest").IsLoaded = True Or Application.CurrentProject.AllReports("certificate of service").IsLoaded = True Then
            GetAddresses = GetAddresses & FormatName(UCase$(Nz(n("Company"))), IIf(n("Deceased") = True, "Estate of " & UCase$(Nz(n("First"))), UCase$(Nz(n("First")))), UCase$(Nz(n("Last"))), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), NewLine)
            Else
            GetAddresses = GetAddresses & FormatName(UCase$(Nz(n("Company"))), UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), NewLine)
            End If
            End If
            n.MoveNext
        Next i
    Case 5      ' normal
        For i = 1 To cnt
        If Application.CurrentProject.AllReports("withdraw sale order").IsLoaded = True Or Application.CurrentProject.AllReports("withdraw sale").IsLoaded = True Then
            GetAddresses = GetAddresses & FormatName(n("Company"), IIf(n("deceased") = True, "Estate of " & n("First"), n("First")), n("Last"), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), NewLine)
            Else
            GetAddresses = GetAddresses & FormatName(n("Company"), n("First"), n("Last"), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), NewLine)
        End If
            If i < cnt Then GetAddresses = GetAddresses & NewLine & NewLine
            n.MoveNext
        Next i
    Case 6      ' one per line
        For i = 1 To cnt
            GetAddresses = GetAddresses & FormatName(n("Company"), n("First"), n("Last"), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), ", ")
            If i < cnt Then GetAddresses = GetAddresses & NewLine & NewLine
            n.MoveNext
        Next i
    Case 7      ' one per line - add comparison for Process Service Cover Letter - MSH - 10/12/2011
        For i = 1 To cnt
            
            GetAddresses = GetAddresses & FormatName(n("Company"), IIf(n("Deceased") = True, "Estate of " & n("First"), n("First")), n("Last"), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), ", ") & " " & IIf(CheckPropertyAddress(File, n("Address"), n("City"), n("State"), n("Zip")), " PROPERTY", "")
            
            If i < cnt Then GetAddresses = GetAddresses & NewLine & NewLine
            n.MoveNext
        Next i
    
    Case 8      ' one per line - add comparison for Process Service Cover Letter - MSH - 10/12/2011
        For i = 1 To cnt
            
            GetAddresses = GetAddresses & FormatName(n("Company"), IIf(n("Deceased") = True, "Estate of " & n("First"), n("First")), n("Last"), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), ", ") & " " & IIf(CheckPropertyAddress(File, n("Address"), n("City"), n("State"), n("Zip")), " PROPERTY", "")
            
            If i < cnt Then GetAddresses = GetAddresses & NewLine
            n.MoveNext
        Next i
    
    Case 9      ' one per line - add comparison for Process Service Cover Letter - MSH - 10/12/2011
        For i = 1 To cnt
            
            GetAddresses = GetAddresses & n("Address") & " " & n("Address2") & n("City") & ", " & n("State") & " " & n("Zip")
            
            If i < cnt Then GetAddresses = GetAddresses & NewLine
            n.MoveNext
        Next i
    Case 10 ' one per line - new process service cover letter to include only property address
     For i = 1 To cnt
            
            GetAddresses = GetAddresses & " " & IIf(CheckPropertyAddress(File, n("Address"), n("City"), n("State"), n("Zip")), FormatName(n("Company"), IIf(n("Deceased") = True, "Estate of " & n("First"), n("First")), n("Last"), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), ", "), "")
            
            If i < cnt Then GetAddresses = GetAddresses & NewLine & NewLine
            n.MoveNext
        Next i
        
    Case 11
        For i = 1 To cnt
         GetAddresses = GetAddresses & "Property owner   " & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), "Mailing address  " & n("Address"), _
                        n("Address2"), "City " & n("City"), "         State " & UCase(n("State")) & "            Zip ", n("Zip"), NewLine) & _
                        NewLine & NewLine
        n.MoveNext
        Next i
    
    Case 12
        For i = 1 To cnt
         GetAddresses = GetAddresses & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), n("Address"), _
                        n("Address2"), n("City"), n("State"), n("Zip"), NewLine) & _
                        NewLine & NewLine & NewLine
        n.MoveNext
        Next i
        
    Case 13
        For i = 1 To cnt
         GetAddresses = GetAddresses & "Person authorized to maintain the property  " & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), "Mailing address  " & n("Address"), _
                        n("Address2"), "City " & n("City"), "         State " & UCase(n("State")) & "            Zip ", n("Zip"), NewLine) & _
                        NewLine & NewLine
        n.MoveNext
        Next i
        
    Case 14
        For i = 1 To cnt
         GetAddresses = GetAddresses & "Name of Property Owner:  " & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), NewLine & "Mailing address:                " & n("Address"), _
                        n("Address2"), n("City"), UCase(n("State")), n("Zip"), "    ") & NewLine & NewLine

        n.MoveNext
        Next i
        
    Case 15
        For i = 1 To cnt
         GetAddresses = GetAddresses & "Name of Person Authorized    " & NewLine & "To Maintain Property:   " & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), NewLine & "Mailing address:         " & n("Address"), _
                        n("Address2"), n("City"), UCase(n("State")), n("Zip"), "    ") & NewLine & NewLine

        n.MoveNext
        Next i

    Case 16
        For i = 1 To cnt
         GetAddresses = GetAddresses & "Last Owner of Record:   " & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), "Address:  " & n("Address"), _
                        n("Address2"), "City " & n("City"), "         State " & UCase(n("State")) & "            Zip ", n("Zip"), NewLine) & _
                        NewLine & NewLine
        n.MoveNext
        Next i
    
    Case 17
        For i = 1 To cnt
         GetAddresses = GetAddresses & "Lender Name:   " & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), "Address:  " & n("Address"), _
                        n("Address2"), "City " & n("City"), "         State " & UCase(n("State")) & "            Zip ", n("Zip"), NewLine) & _
                        NewLine & NewLine
        n.MoveNext
        Next i
    
    Case 18
        For i = 1 To cnt
         GetAddresses = GetAddresses & "Owner Name (if known)          " & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), NewLine & "Owner Address (if known):     " & n("Address"), _
                        n("Address2"), n("City"), UCase(n("State")), n("Zip"), "    ") & NewLine & NewLine

        n.MoveNext
        Next i
        
   Case 19 ' for Alternative Borrowers Address
         
    Set myCol = New Collection
    ReDim namebor(1 To cnt) As String
    
    For J = 1 To cnt
        addraltcheckN(J) = n("First") & n("Last") & n("Address") & n("Address2") & n("City") & n("State") & n("Zip")
        n.MoveNext
    Next J
        
    n.MoveFirst
       
    For J = 1 To cnt
        namebor(J) = n("First") & n("Last")
        cntalt = 0
         If File = 0 Then File = Forms![Case List]!FileNumber <> (-1)
        
        Set m = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND Noteholder=0 AND (First & Last) LIKE """ & namebor(J) & """  ORDER BY ID DESC, Last, First", dbOpenSnapshot)
        If m.EOF Then GoTo SSS
        m.MoveLast
        cntalt = m.RecordCount
        
        ReDim AddrAlt(1 To cntalt) As String
        m.MoveFirst
        
            For i = 1 To cntalt
            AddrAlt(i) = m("First") & m("Last") & m("Address") & m("Address2") & m("City") & m("State") & m("Zip")
            m.MoveNext
            Next i
        
        m.MoveFirst
        ReDim addraltcheck(1 To cntalt) As String
        ReDim Namecheck(1 To cntalt) As String
        m.MoveFirst
      
                For i = 1 To cntalt
                       Namecheck(i) = m("First") & m("Last")
                       addraltcheck(i) = m("First") & m("Last") & m("Address") & m("Address2") & m("City") & m("State") & m("Zip")
                       G = False
                    
                        For l = 1 To cnt
                          If addraltcheck(i) = addraltcheckN(l) Then
                          GoTo KKK
                          End If
                        Next
                     
                        If p = 0 Then
                             p = 1
                             newArray(p) = addraltcheck(i)
                             GoTo MMM
                        Else
                              
                            For R = 1 To p
                                    If newArray(R) = addraltcheck(i) Then
                                        GoTo KKK
                                    Else
                                        G = True
                                    End If
                             Next R
                                    
                                If G Then
                                   G = False
                                   p = p + 1
                                   newArray(p) = addraltcheck(i)
                                   GoTo MMM
                                End If
                        End If
               
                 
MMM:                    GetAddresses = GetAddresses & FormatName(UCase$(Nz(m("Company"))), UCase$(Nz(m("First"))), UCase$(Nz(m("Last"))), m("AKA"), m("Address"), m("Address2"), m("City"), m("State"), m("Zip"), NewLine) & NewLine
                
KKK:                      m.MoveNext
                  Next i
   
SSS:      n.MoveNext
 
     Next J
         
     
 Case 20 ' for Alternative ownder Address
       
       
        Set myCol = New Collection
        ReDim namebor(1 To cnt) As String
        
        For J = 1 To cnt
         addraltcheckN(J) = n("First") & n("Last") & n("Address") & n("Address2") & n("City") & n("State") & n("Zip")
         n.MoveNext
        Next J
        
        n.MoveFirst
       
        For J = 1 To cnt
      
             namebor(J) = n("First") & n("Last")
       
                  
            cntalt = 0
            If File = 0 Then File = Forms![Case List]!FileNumber <> (-1)
            Set m = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND Owner=0 AND Noteholder=0 AND Mortgagor=0 AND BKDebtor=0 AND BKCoDebtor=0 AND COLDebtor=0 AND (First & Last) LIKE """ & namebor(J) & """  ORDER BY BKCoDebtor DESC, ID, Last, First", dbOpenSnapshot)
        
            If m.EOF Then GoTo TTT
      
    
            m.MoveLast
            
            cntalt = m.RecordCount
            
           
            
     
            
            ReDim AddrAlt(1 To cntalt) As String
           
             
           
            
             m.MoveFirst
            For i = 1 To cntalt
            AddrAlt(i) = m("First") & m("Last") & m("Address") & m("Address2") & m("City") & m("State") & m("Zip")
            m.MoveNext
            Next i
            
            m.MoveFirst
            
            ReDim addraltcheck(1 To cntalt) As String
            ReDim Namecheck(1 To cntalt) As String
            
            m.MoveFirst
          
            For i = 1 To cntalt
            
                      
                        Namecheck(i) = m("First") & m("Last")
                        addraltcheck(i) = m("First") & m("Last") & m("Address") & m("Address2") & m("City") & m("State") & m("Zip")
                        G = False
                       
                        
                        
                            For l = 1 To cnt
                              If addraltcheck(i) = addraltcheckN(l) Then
                              GoTo RRR
                              End If
                            Next
                            
                
             
               If p = 0 Then
                      p = 1
                     newArray(p) = addraltcheck(i)
          
                     GoTo PPP
                Else
                      
                            For R = 1 To p
                                    If newArray(R) = addraltcheck(i) Then
                                    GoTo RRR
                                    Else
                                    G = True
                                    End If
                             Next R
                            
                            If G Then
                            
                                G = False
                                p = p + 1
                                newArray(p) = addraltcheck(i)
                                GoTo PPP
                            
                            
                            End If
                    End If
                            
                            
                    
            
          
         
PPP:      GetAddresses = GetAddresses & FormatName(UCase$(Nz(m("Company"))), UCase$(Nz(m("First"))), UCase$(Nz(m("Last"))), m("AKA"), m("Address"), m("Address2"), m("City"), m("State"), m("Zip"), NewLine) & NewLine
                           
                            
RRR:      m.MoveNext
          Next i


            
TTT:      n.MoveNext
     
          Next J
         
   
    Case Else
    
        MsgBox "Invalid format in call to GetAddresses", vbExclamation
        GetAddresses = ""
End Select
Set m = Nothing
            


End Function


Public Function GetAffadavitOfServiceAddresses(Fmt As Integer, Optional NewLine As String = vbNewLine) As String
'
' File: File number, or 0 for 'current' file
' fmt:  3 = 'normal' list, with name in caps
'       4 = captions
'       5 = 'normal' list
'       6 = list with each on one line
'
Dim n As Recordset
Dim cnt As Integer
Dim i As Integer
Dim Addr() As String

cnt = 0
Set n = CurrentDb.OpenRecordset("SELECT * FROM AffadavitOfServiceRecipients WHERE SendAffadavitOfService = -1 ORDER BY NameID, Last, First", dbOpenSnapshot)
If n.EOF Then
    GetAffadavitOfServiceAddresses = ""
    Exit Function
End If
n.MoveLast
cnt = n.RecordCount
ReDim Addr(1 To cnt) As String
n.MoveFirst

Select Case Fmt
    Case 3      ' normal with caps
        For i = 1 To cnt
            GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & FormatName(UCase$(Nz(n("Company"))), UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), NewLine)
            If i < cnt Then GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & NewLine & NewLine
            n.MoveNext
        Next i
    Case 4      ' caption
        For i = 1 To cnt
            Addr(i) = n("Address") & n("Address2") & n("City") & n("State") & n("Zip")
            n.MoveNext
        Next i
        n.MoveFirst
        For i = 1 To cnt
            If i < cnt Then
                If Addr(i) = Addr(i + 1) Then
                    GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), _
                        Null, Null, Null, Null, Null, NewLine)
                Else
                    GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & FormatName(UCase$(Nz(n("Company"))), _
                        UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), n("Address"), _
                        n("Address2"), n("City"), n("State"), n("Zip"), NewLine) & _
                        NewLine & NewLine & "and" & NewLine & NewLine
                End If
            Else
                GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & FormatName(UCase$(Nz(n("Company"))), UCase$(Nz(n("First"))), UCase$(Nz(n("Last"))), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), NewLine)
            End If
            n.MoveNext
        Next i
    Case 5      ' normal
        For i = 1 To cnt
            If n("Deceased") = True Then
            GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & FormatName(n("Company"), "Estate of " & n("First"), n("Last"), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), NewLine)
            Else
            GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & FormatName(n("Company"), n("First"), n("Last"), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), NewLine)
            End If
            If i < cnt Then GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & NewLine & NewLine
            n.MoveNext
        Next i
   Case 6      ' one per line
        For i = 1 To cnt
            GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & FormatName(n("Company"), n("First"), n("Last"), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), ", ")
            If i < cnt Then GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & NewLine & NewLine
            n.MoveNext
        Next i
    Case 7      ' one per line for PS cover
        For i = 1 To cnt
            GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & IIf(Not CheckPropertyAddress(Forms![Case List]!FileNumber, n("Address"), n("City"), n("State"), n("Zip")), FormatName(n("Company"), n("First"), n("Last"), n("AKA"), n("Address"), n("Address2"), n("City"), n("State"), n("Zip"), ", "), "")
            If GetAffadavitOfServiceAddresses <> "" Then GetAffadavitOfServiceAddresses = GetAffadavitOfServiceAddresses & NewLine & NewLine
            n.MoveNext
        Next i
    Case Else
        MsgBox "Invalid format in call to GetAffadavitOfServiceAddresses", vbExclamation
        GetAffadavitOfServiceAddresses = ""
End Select

End Function


Public Sub testGetNames()
  MsgBox GetNames(10232, 10, "Noteholder=true")
End Sub



Public Function GetNames(File As Long, Fmt As Integer, conditions As String, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean, Optional Intext As Boolean) As String
'
' File: File number, or 0 for 'current' file
' fmt:  1 = comma separated list
'       2 = comma separated list, except AND last name
'       3 = one name per line
'       4 = signature lines
'       5 = signature lines using electronic signatures
'       6 = one name per line, except AND/OR last name
'       7 = command seperated list, except AND/OR last name
'       10 = name (last, first)
'       11 = SSN, comma separated list
'       12 = SSN last 4 only, comma separated list
'       20 = comma separated list, except AND last name + No Estate of "Name
'       60 = One name per line, except AND/OR last name + No Estate of "Name
'       99 = one name and ", et al. " if multiple names
'       100 = name (last), comma separated list  'WOT
'       110 = name (last) ONLY , comma separated last
'       150 = 'Current file' All names with no conditions ** use like this **  Getnames(150,FMT,"")
        

Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
    If File = 0 Then File = Forms![Case List]!FileNumber
        If NoSSN = True Then
            Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ") AND SSN = ""999-99-9999""  ORDER BY ID, Last, First", dbOpenSnapshot)
        ElseIf File = 150 Then File = Forms![Case List]!FileNumber
            Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " ORDER BY ID, Last, First", dbOpenSnapshot)
        Else
            Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ") ORDER BY ID, Last, First", dbOpenSnapshot)
        End If


Do While Not n.EOF
    cnt = cnt + 1
    If Nz(n("Company")) <> "" Then
      Names(cnt) = n("Company")
    Else
      AKA = ""
      If Nz(n("AKA")) <> "" Then
        If Fmt = 3 Then     ' put AKA on new line if formatting one name per line
            AKA = Trim$(n("AKA"))
            If Left$(AKA, 1) = "," Then AKA = Trim$(Mid$(AKA, 2, Len(AKA) - 1))
            AKA = NewLine & AKA
        Else
            ' 2012.01.30 DaveW Blank before AKA
            AKA = " " & n("AKA")
        End If
      End If
      Select Case CapsFmt
        Case 0      ' no change
            Select Case Fmt
                Case Is <= 9
                'If Application.CurrentProject.AllReports("Order to Docket Cover").IsLoaded = True Or Application.CurrentProject.AllReports("Report of sale").IsLoaded = True Then
                
                
                If n("deceased") = True And Intext = False Then
                Names(cnt) = "Estsate of " & n("First") & " " & n("Last") & AKA
                Names(cnt) = "Estate of " & n("First") & " " & n("Last") & AKA

                Else
                Names(cnt) = n("First") & " " & n("Last") & AKA
                End If
'                Else
'                Names(cnt) = N("First") & " " & N("Last") & AKA
'                End If
                Case 10
                    Names(cnt) = n("Last") & ", " & n("First") & AKA
                Case 11
                    Names(cnt) = Nz(n("SSN"))
                Case 12
                    Names(cnt) = "xxx-xx-" & Right$(Nz(n("SSN"), "????"), 4)
                Case 20
                    Names(cnt) = n("First") & " " & n("Last") & AKA
                Case 60
                    Names(cnt) = n("First") & " " & n("Last") & AKA
                Case 99
                    Names(cnt) = n("First") & " " & n("Last") & AKA
                Case 100 ' Wells SOT 7/30 MC
                    Names(cnt) = UCase$(Nz(n("Last"))) & ", " & n("First")
                Case 110
                   Names(cnt) = Nz(n("Last"))
            End Select
        Case 1      ' entire name
            Names(cnt) = UCase$(Nz(n("First") & " " & n("Last") & AKA))
        Case 2      ' last name only
            Names(cnt) = n("First") & " " & UCase$(Nz(n("Last"))) & AKA
      End Select
    End If
    n.MoveNext
Loop

Select Case Fmt
    Case 1, 11, 12 ' comma separated list
        For i = 1 To cnt
            GetNames = GetNames & Names(i)
            If i < cnt Then GetNames = GetNames & ", "
        Next i
    Case 2      ' comma separated list, except AND last name
        For i = 1 To cnt
            GetNames = GetNames & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetNames = GetNames & ", "
                Else
                    GetNames = GetNames & " and "
                End If
            End If
        Next i
    Case 3, 10     ' one name per line
        For i = 1 To cnt
            GetNames = GetNames & Names(i) & NewLine
        Next i
    Case 4      ' signature lines
        For i = 1 To cnt
            GetNames = GetNames & "___________________________________" & NewLine & Names(i) & NewLine & NewLine
        Next i
    Case 5      ' electronic signatures
        For i = 1 To cnt
            GetNames = GetNames & "/s/ " & Names(i) & NewLine & NewLine & NewLine
        Next i
    Case 6      ' One name per line, except AND/OR last name
         For i = 1 To cnt
            If i = cnt - 1 Then
                GetNames = GetNames & Names(i) & " and/or " & NewLine
            Else
                GetNames = GetNames & Names(i) & NewLine
            End If
         Next i
    Case 7      ' comma separated list, except AND/OR last name
        For i = 1 To cnt
            GetNames = GetNames & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetNames = GetNames & ", "
                Else
                    GetNames = GetNames & " and/or "
                End If
            End If
        Next i
    Case 20 ' comma separated list, except AND last name + No Estate of "Name
          For i = 1 To cnt
            GetNames = GetNames & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetNames = GetNames & ", "
                Else
                    GetNames = GetNames & " and "
                End If
            End If
        Next i
    Case 60 ' One name per line, except AND/OR last name + No Estate of "Name
        For i = 1 To cnt
            If i = cnt - 1 Then
                GetNames = GetNames & Names(i) & " and/or " & NewLine
            Else
                GetNames = GetNames & Names(i) & NewLine
            End If
         Next i
    Case 99
         For i = 1 To 1
            GetNames = GetNames & Names(i)
            If i <= 1 Then
                If i < cnt Then
                    GetNames = GetNames & ", et al."
                Else
                    'GetNames = GetNames & " and "
                End If
            End If
        Next i
            
    Case 100
        For i = 1 To cnt
            If i = cnt - 1 Then
                GetNames = GetNames & Names(i) & NewLine
            Else
                GetNames = GetNames & Names(i) & NewLine
            End If
         Next i
    Case 110
         For i = 1 To cnt
            GetNames = GetNames & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetNames = GetNames & ", "
                Else
                    GetNames = GetNames & " and "
                End If
            End If
        Next i
        
        'For i = 1 To cnt
         '  If i = cnt - 1 Then
              ' GetNames = GetNames & Names(i) & ", "
          '  Else
               ' GetNames = GetNames & Names(i) & " and "
           ' End If
        ' Next i
    Case Else
        MsgBox "Invalid format in call to GetNames", vbExclamation
        GetNames = ""
End Select
End Function
Public Function getTrustees(columnInfo As Variant, Listcounter As Variant)

Dim i As Integer
Dim strTrustees As String

getTrustees = ""
    For i = 0 To Listcounter.ListCount
        If i < Listcounter.ListCount - 1 Then
            If i < Listcounter.ListCount - 2 Then
                getTrustees = getTrustees & columnInfo.Column(1, i) & ", "
            Else
                getTrustees = getTrustees & columnInfo.Column(1, i) & " and "
            End If
        Else
             getTrustees = getTrustees & columnInfo.Column(1, i)
        End If
    Next i

End Function ' This thing only works for listboxes right now.

Public Function GetTRCIVNames(File As Long, Fmt As Integer, conditions As String, Optional NewLine As String = vbNewLine) As String
'
' File: File number, or 0 for 'current' file
' fmt:  1 = comma separated list
'       2 = comma separated list, except AND last name
'       3 = one name per line
'       4 = signature lines
'       5 = signature lines using electronic signatures
'       11= SSN, comma separated list
'       12= SSN last 4 only, comma separated list
'
Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim CaseType As Integer
Dim namesTbl As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
CaseType = DLookup("[CaseTypeID]", "[CaseList]", "FileNumber = " & File)

    Select Case CaseType
        Case 5      ' Civil
          namesTbl = "NamesCIV"
        Case 10     ' Title Settlement
          namesTbl = "NamesTR"
        Case Else   ' should not be here
          MsgBox "Invalid case type.", vbCritical
          Exit Function
    End Select



Set n = CurrentDb.OpenRecordset("SELECT * FROM " & namesTbl & " WHERE FileNumber = " & File & " AND (" & conditions & ") ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    
    
    Select Case CapsFmt
        Case 0      ' no change
            Select Case Fmt
                Case Is <= 10
                    Names(cnt) = IIf(Not IsNull(n("Company")), n("Company"), n("First") & " " & n("Last"))
                Case 11
                    Names(cnt) = Nz(n("SSN"))
                Case 12
                    Names(cnt) = "xxx-xx-" & Right$(Nz(n("SSN"), "????"), 4)
            End Select
        Case 1      ' entire name
            Names(cnt) = IIf(Not IsNull(n("Company")), UCase$(n("Company")), UCase$(Nz(n("First") & " " & n("Last"))))
        Case 2      ' last name only
            Names(cnt) = IIf(Not IsNull(n("Company")), UCase$(n("Company")), n("First") & " " & UCase$(Nz(n("Last"))))
    End Select
    n.MoveNext
Loop
Select Case Fmt
    Case 1, 11, 12 ' comma separated list
        For i = 1 To cnt
            GetTRCIVNames = GetTRCIVNames & Names(i)
            If i < cnt Then GetTRCIVNames = GetTRCIVNames & ", "
        Next i
    Case 2      ' comma separated list, except AND last name
        For i = 1 To cnt
            GetTRCIVNames = GetTRCIVNames & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetTRCIVNames = GetTRCIVNames & ", "
                Else
                    GetTRCIVNames = GetTRCIVNames & " and "
                End If
            End If
        Next i
    Case 3      ' one name per line
        For i = 1 To cnt
            GetTRCIVNames = GetTRCIVNames & Names(i) & NewLine
        Next i
    Case 4      ' signature lines
        For i = 1 To cnt
            GetTRCIVNames = GetTRCIVNames & "___________________________________" & NewLine & Names(i) & NewLine & NewLine
        Next i
    Case 5      ' electronic signatures
        For i = 1 To cnt
            GetTRCIVNames = GetTRCIVNames & "/s/ " & Names(i) & NewLine & NewLine & NewLine
        Next i
    Case Else
        MsgBox "Invalid format in call to GetTRCIVNames", vbExclamation
        GetTRCIVNames = ""
End Select
End Function



Public Function GetCIVNames(File As Long, Fmt As Integer, conditions As String, Optional NewLine As String = vbNewLine) As String

' File: File number, or 0 for 'current' file
' fmt:  1 = comma separated list
'       2 = comma separated list, except AND last name
'       3 = one name per line
'       4 = signature lines
'       5 = signature lines using electronic signatures
'       11= SSN, comma separated list
'       12= SSN last 4 only, comma separated list
'



Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer



cnt = 0
Set n = CurrentDb.OpenRecordset("SELECT * FROM NamesCIV WHERE FileNumber = " & File & " AND (" & conditions & ") ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    Select Case CapsFmt
        
        Case 1      ' entire name
            
             Names(cnt) = IIf(Not IsNull(n("Last")), UCase$(Nz(n("First") & " " & n("Last"))), Nz(UCase$(n("Company"))))
            
        Case 2      ' last name only
            Names(cnt) = IIf(Not IsNull(n("Last")), UCase$(n("Last")), UCase$(Nz(n("Company"))))
        Case Else
            Names(cnt) = IIf(Not IsNull(n("Last")), n("Last"), Nz(n("Company")))
    End Select
    n.MoveNext
Loop

Select Case Fmt
    Case 1, 11, 12 ' comma separated list
        For i = 1 To cnt
            GetCIVNames = GetCIVNames & Names(i)
            If i < cnt Then GetCIVNames = GetCIVNames & ", "
        Next i
    Case 2      ' comma separated list, except AND last name
        For i = 1 To cnt
            GetCIVNames = GetCIVNames & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetCIVNames = GetCIVNames & ", "
                Else
                    GetCIVNames = GetCIVNames & " and "
                End If
            End If
        Next i
    Case 3      ' one name per line
        For i = 1 To cnt
            GetCIVNames = GetCIVNames & Names(i) & NewLine
        Next i
    Case 4      ' signature lines
        For i = 1 To cnt
            GetCIVNames = GetCIVNames & "___________________________________" & NewLine & Names(i) & NewLine & NewLine
        Next i
    Case 5      ' electronic signatures
        For i = 1 To cnt
            GetCIVNames = GetCIVNames & "/s/ " & Names(i) & NewLine & NewLine & NewLine
        Next i
    Case Else
        MsgBox "Invalid format in call to GetCIVNames", vbExclamation
        GetCIVNames = ""
End Select

End Function


Function FormatName(Company As Variant, First As Variant, Last As Variant, _
    AKA As Variant, addr1 As Variant, addr2 As Variant, City As Variant, _
    State As Variant, Zip As Variant, Optional NewLine As String = vbNewLine) As String
Dim n As String

If Nz(Company) <> "" Then
    n = Company & NewLine
    If Nz(First) & Nz(Last) & Nz(AKA) <> "" Then

        n = n & "Attn: "
    End If
End If

If Nz(First) & Nz(Last) & Nz(AKA) <> "" Then n = n & IIf(IsNull(First), "", First & " ") & Nz(Last) & NewLine

If Nz(AKA) <> "" Then
    If Left$(AKA, 1) = "," Then
        n = n & Trim$(Mid$(AKA, 2, Len(AKA) - 1)) & NewLine
    Else
        n = n & Trim$(AKA) & NewLine
    End If
End If

If Not IsNull(addr1) Then n = n & addr1 & NewLine
If Not IsNull(addr2) Then
n = n & addr2 & NewLine
Else
n = n
End If
n = n & City
If Not IsNull(State) Then n = n & ", " & State
If Not IsNull(Zip) Then n = n & " " & FormatZip(Zip) & NewLine
FormatName = n
End Function


Function FormatNameDear(Company As Variant, First As Variant, Last As Variant) As String
Dim n As String

If Nz(First) & Nz(Last) = "" Then
    FormatNameDear = Nz(Company)
Else
    FormatNameDear = Nz(First) & " " & Nz(Last)
End If

End Function

Public Function NoticeNames(File As Long, Optional conditions As String) As String
Dim n As Recordset
Dim sql As String

If File = 0 Then File = Forms![Case List]!FileNumber
sql = "SELECT Names.*,NoticeTypes.NoticeType FROM Names INNER JOIN NoticeTypes ON Names.NoticeType = NoticeTypes.ID WHERE Names.FileNumber = " & File
If conditions <> "" Then sql = sql & " AND " & conditions
sql = sql & " ORDER BY Last, First"

Set n = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
Do While Not n.EOF
    If Not IsNull(n("Company")) Then
        NoticeNames = NoticeNames & "Name: " & n("Company") & ", "
        ' 2012.01.30 DaveW Blank before AKA
        If Not IsNull(n("Last")) Then
            NoticeNames = NoticeNames & IIf(n("Deceased") = True, "Estate of " & n("First"), n("First")) & " " & n("Last") & " " & n("AKA") & vbNewLine
        Else
            NoticeNames = NoticeNames & vbNewLine
        End If
    Else
        ' 2012.01.30 DaveW Blank before AKA
        NoticeNames = NoticeNames & "Name: " & IIf(n("Deceased") = True, "Estate of " & n("First"), n("First")) & " " & n("Last") & " " & n("AKA") & vbNewLine
    End If
    
    NoticeNames = NoticeNames & "Address: " & n("Address") & ", "
    If Not IsNull(n("Address2")) Then NoticeNames = NoticeNames & "Address: " & n("Address2") & ", "
    NoticeNames = NoticeNames & n("City") & ", " & n("State") & " " & n("Zip") & vbNewLine
    'If NoStatus = False Then
        NoticeNames = NoticeNames & "Status: " & n("NoticeTypes.NoticeType") & vbNewLine & vbNewLine
    'Else
    '    NoticeNames = NoticeNames & vbNewLine & vbNewLine
    'End If
    n.MoveNext
Loop
n.Close
Set n = CurrentDb.OpenRecordset("SELECT * FROM qryCountyAttorney WHERE FileNumber=" & File, dbOpenSnapshot)
If Not IsNull(n("CountyAttnyName")) Then
    NoticeNames = NoticeNames & "Name: " & n("CountyAttnyName") & vbNewLine
    NoticeNames = NoticeNames & "Address: " & OneLine(n("CountyAttnyAddr")) & vbNewLine
    NoticeNames = NoticeNames & "Status: County Attorney" & vbNewLine & vbNewLine
End If
n.Close
End Function

Public Function BKService(File As Long, Optional NewLine As String = vbNewLine) As String
BKService = GetAddresses(File, 5, "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", NewLine)
End Function




Public Function NotaryNames(File As Long, conditions As String, Optional NewLine As String = vbNewLine) As String
'
' File: File number, or 0 for 'current' file
'
Dim n As Recordset
Dim i As Integer
Dim AKA As String


If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ") ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    AKA = ""
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If

  '  NotaryNames = NotaryNames & "STATE OF ____________________" & vbNewLine & vbNewLine & "CITY/COUNTY OF _______________________, to-wit:" & vbNewLine & vbNewLine
  
    NotaryNames = NotaryName & FetchNotaryLocation(Forms!Foreclosureprint!NotaryID) & vbNewLine & vbNewLine
    NotaryNames = NotaryNames & "          I, the undersigned Notary Public in and for the jurisdiction aforesaid, do hereby certify that " & UCase$(Nz(n("First") & " " & n("Last") & AKA)) & " signed the foregoing Deed in Lieu of Foreclosure on the ____ day of ____________, 20___, and acknowledged the same before me." & vbNewLine & vbNewLine
    NotaryNames = NotaryNames & "GIVEN under my hand this ____ day of ____________, 20___." & vbNewLine & vbNewLine
    
    NotaryNames = NotaryNames & "                                                                       ________________________________" & vbNewLine
   ' NotaryNames = NotaryNames & "                                                                       Notary Public" & vbNewLine
    NotaryNames = NotaryNames & "                                                                       " & FetchNotaryName(Forms!Foreclosureprint!NotaryID, True) & vbNewLine
    
    NotaryNames = NotaryNames & "                                                                       My Commission Expires __________" & vbNewLine
    
    
    NotaryNames = NotaryNames & vbNewLine & vbNewLine
    
    
    n.MoveNext
Loop

End Function

Public Function FetchNotaryName(NotaryID As Variant, Title As Boolean)

If (IsNull(NotaryID)) Then
  If (Title = True) Then
    FetchNotaryName = "Notary Public"
  Else
    FetchNotaryName = "_____________________________"
  End If
  
  Exit Function
End If

If (NotaryID = -1) Then

  If (Title = True) Then
    FetchNotaryName = "Notary Public"
  Else
    FetchNotaryName = "_____________________________"
  End If
  Exit Function
End If

FetchNotaryName = DLookup("[NotaryName]", "[Staff]", "[ID] = " & NotaryID)


End Function


Public Function FetchNotaryLocation(NotaryID As Variant)


Dim strNotaryCounty As String
Dim strNotaryState As String


If (IsNull(NotaryID)) Then
  FetchNotaryLocation = "STATE OF _____________________:" & vbNewLine & "COUNTY/CITY OF ____________________:"
  Exit Function
End If

If (NotaryID = -1) Then
  FetchNotaryLocation = "STATE OF _____________________:" & vbNewLine & "COUNTY/CITY OF ____________________:"
  Exit Function
End If

strNotaryCounty = Nz(DLookup("[NotaryCounty]", "[Staff]", "[ID] = " & NotaryID))

FetchNotaryLocation = "STATE OF " & UCase(Nz(DLookup("[NotaryState]", "[Staff]", "[ID] = " & NotaryID))) & IIf(strNotaryCounty = "", "", vbNewLine & "COUNTY/CITY OF " & UCase(strNotaryCounty))

End Function


Sub testnames()
Dim n As Recordset, File As Integer, conditions As String
File = 10232
conditions = "mortgagor = true"
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ")  AND SSN = ""999-99-9999"" ORDER BY ID, Last, First", dbOpenSnapshot)

Do While Not n.EOF
MsgBox "" & n!Last & ""
n.MoveNext
Loop

End Sub



Public Function GetNamesMD(File As Long, Fmt As Integer, conditions As String, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
'
' File: File number, or 0 for 'current' file
' fmt:  1 = comma separated list
'       2 = comma separated list, except AND last name
'       3 = one name per line
'       4 = signature lines
'       5 = signature lines using electronic signatures
'       10 = name (last, first)
'       11= SSN, comma separated list
'       12= SSN last 4 only, comma separated list
'
Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
If NoSSN = True Then
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ") AND SSN = ""999-99-9999""  ORDER BY ID, Last, First", dbOpenSnapshot)
Else
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ") ORDER BY ID, Last, First", dbOpenSnapshot)
End If
Do While Not n.EOF
    cnt = cnt + 1
    If Nz(n("Company")) <> "" Then
      Names(cnt) = n("Company")
    Else
      AKA = ""
      If Nz(n("AKA")) <> "" Then
        If Fmt = 3 Then     ' put AKA on new line if formatting one name per line
            AKA = Trim$(n("AKA"))
            If Left$(AKA, 1) = "," Then AKA = Trim$(Mid$(AKA, 2, Len(AKA) - 1))
            AKA = NewLine & AKA
        Else
            ' 2012.01.30 DaveW Blank before AKA
            AKA = " " & n("AKA")
        End If
      End If
      Select Case CapsFmt
        Case 0      ' no change
            Select Case Fmt
                Case Is <= 9
                'If Application.CurrentProject.AllReports("Order to Docket Cover").IsLoaded = True Or Application.CurrentProject.AllReports("Report of sale").IsLoaded = True Then
              '  If N("deceased") = True Then
               Names(cnt) = n("First") & " " & n("Last") & AKA
               'Names(cnt) = N("First") & " " & N("Last") & AKA

               ' Else
              '  Names(cnt) = N("First") & " " & N("Last") & AKA
              '  End If
'                Else
'                Names(cnt) = N("First") & " " & N("Last") & AKA
'                End If
                Case 10
                    Names(cnt) = n("Last") & ", " & n("First") & AKA
                Case 11
                    Names(cnt) = Nz(n("SSN"))
                Case 12
                    Names(cnt) = "xxx-xx-" & Right$(Nz(n("SSN"), "????"), 4)
            End Select
        Case 1      ' entire name
            Names(cnt) = UCase$(Nz(n("First") & " " & n("Last") & AKA))
        Case 2      ' last name only
            Names(cnt) = n("First") & " " & UCase$(Nz(n("Last"))) & AKA
      End Select
    End If
    n.MoveNext
Loop
Select Case Fmt
    Case 1, 11, 12 ' comma separated list
        For i = 1 To cnt
            GetNamesMD = GetNamesMD & Names(i)
            If i < cnt Then GetNamesMD = GetNamesMD & ", "
        Next i
    Case 2      ' comma separated list, except AND last name
        For i = 1 To cnt
            GetNamesMD = GetNamesMD & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetNamesMD = GetNamesMD & ", "
                Else
                    GetNamesMD = GetNamesMD & " and "
                End If
            End If
        Next i
    Case 3, 10     ' one name per line
        For i = 1 To cnt
            GetNamesMD = GetNamesMD & Names(i) & NewLine
        Next i
    Case 4      ' signature lines
        For i = 1 To cnt
            GetNamesMD = GetNamesMD & "___________________________________" & NewLine & Names(i) & NewLine & NewLine
        Next i
    Case 5      ' electronic signatures
        For i = 1 To cnt
            GetNamesMD = GetNamesMD & "/s/ " & Names(i) & NewLine & NewLine & NewLine
        Next i
    Case Else
        MsgBox "Invalid format in call to GetNamesMD", vbExclamation
        GetNamesMD = ""
End Select
End Function


Public Function BorrowerNamesWiz(File As Long) As String
' BorrowerNames = GetNames(File, Fmt, "Noteholder = True", NewLine)


Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![wizfairdebt]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND Noteholder = True ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop

Dim J As Integer

J = 0
For i = 1 To cnt
    If (InStr(1, BorrowerNamesWiz, Names(i)) = 0) Then
      J = J + 1
      BorrowerNamesWiz = BorrowerNamesWiz & IIf(J > 1, " and ", "") & Names(i)
    End If
Next i
End Function
Public Function BrowersNamesV2(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
BrowersNamesV2 = GetNames(File, Fmt, "Noteholder = True", NewLine, NoSSN)

End Function
Public Function AttorneyName(AId As Long) As String
'Dim AttoN As Recordset
'Set AttoN = CurrentDb.OpenRecordset("SELECT * FROM Staff WHERE ID = " & AId, dbOpenSnapshot)
'AttorneyName = AttoN!Name
'AttoN.Close

AttorneyName = DLookup("[Name]", "[Staff]", "[ID] = " & AId)
End Function

Public Function BorrowerNamesCount(File As Long) As Integer
' BorrowerNames = GetNames(File, Fmt, "Noteholder = True", NewLine)

Dim BorrowerNames As String
Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND Noteholder = True ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop
End Function



Public Function TenantNamesCount(File As Long) As Integer
' BorrowerNames = GetNames(File, Fmt, "Noteholder = True", NewLine)

Dim TenantNames As String
Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND Tenant = True ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop
Dim J As Integer

J = 0
For i = 1 To cnt
    If (InStr(1, TenantNames, Names(i)) = 0) Then
      J = J + 1
      TenantNames = TenantNames & IIf(J > 1, " and ", "") & Names(i)
    End If
Next i
TenantNamesCount = J
End Function
Public Function BorrowerNamesOne(File As Long, A As Integer) As String
' BorrowerNames = GetNames(File, Fmt, "Noteholder = True", NewLine)

Dim BorrowerNames As String
Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND Noteholder = True ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop

Dim J As Integer

J = 0
For i = 1 To cnt
    If (InStr(1, BorrowerNames, Names(i)) = 0) Then
      J = J + 1
      BorrowerNames = Names(i)
      If J = A Then BorrowerNamesOne = BorrowerNames
    End If
Next i

End Function


'AND SSN = ""999-99-9999""
Public Function MortgagorNamesCount(File As Long) As Integer
'AND SSN = ""999-99-9999""
Dim MortgagorNames As String
Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND Mortgagor = True  ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop

Dim J As Integer

J = 0
For i = 1 To cnt
    If (InStr(1, MortgagorNames, Names(i)) = 0) Then
      J = J + 1
      MortgagorNames = MortgagorNames & IIf(J > 1, " and ", "") & Names(i)
    End If
Next i
MortgagorNamesCount = J
End Function

Public Function MortgagorNamesOne(File As Long, A As Integer) As String
'AND SSN = ""999-99-9999""
Dim MortgagorNames As String
Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND Mortgagor = True  ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop

Dim J As Integer

J = 0
For i = 1 To cnt
    If (InStr(1, MortgagorNames, Names(i)) = 0) Then
      J = J + 1
      MortgagorNames = Names(i)
      If J = A Then MortgagorNamesOne = MortgagorNames
    End If
Next i

End Function



Public Function BorrowerMorgagorNamesCountNoSSN(File As Long) As Integer
' BorrowerNames = GetNames(File, Fmt, "Noteholder = True", NewLine)

Dim BorrowerNames As String
Dim n As Recordset
Dim Names(1 To 60) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (Noteholder = True Or Mortgagor = True) AND (Nz(SSN) = 0 OR SSN = ""999999999"" )ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop

Dim J As Integer

J = 0
For i = 1 To cnt
    If (InStr(1, BorrowerNames, Names(i)) = 0) Then
      J = J + 1
      BorrowerNames = BorrowerNames & IIf(J > 1, " and ", "") & Names(i)
    End If
Next i
BorrowerMorgagorNamesCountNoSSN = J
End Function
Public Function BorrowerMorgagorNamesOneNoSSN(File As Long, A As Integer) As String
' BorrowerNames = GetNames(File, Fmt, "Noteholder = True", NewLine)

Dim BorrowerNames As String
Dim n As Recordset
Dim Names(1 To 60) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (Noteholder = True Or Mortgagor = True) AND (Nz(SSN) = 0 OR SSN = ""999999999"") ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop

Dim J As Integer

J = 0
For i = 1 To cnt
    If (InStr(1, BorrowerNames, Names(i)) = 0) Then
      J = J + 1
      BorrowerNames = Names(i)
      If J = A Then BorrowerMorgagorNamesOneNoSSN = BorrowerNames
    End If
Next i

End Function

Public Function BorrowerMorgagorNamesOneSSN(File As Long, A As Integer) As String
' BorrowerNames = GetNames(File, Fmt, "Noteholder = True", NewLine)

Dim BorrowerNames As String
Dim n As Recordset
Dim Names(1 To 60) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (Noteholder = True Or Mortgagor = True) AND (Nz(SSN) <> 0 AND SSN <> ""999999999"") ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop

Dim J As Integer

J = 0
For i = 1 To cnt
    If (InStr(1, BorrowerNames, Names(i)) = 0) Then
      J = J + 1
      BorrowerNames = Names(i)
      If J = A Then BorrowerMorgagorNamesOneSSN = BorrowerNames
    End If
Next i

End Function

Public Function BorrowerMorgagorNamesCountSSN(File As Long) As Integer
' BorrowerNames = GetNames(File, Fmt, "Noteholder = True", NewLine)

Dim BorrowerNames As String
Dim n As Recordset
Dim Names(1 To 60) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber
Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (Noteholder = True Or Mortgagor = True) AND (Nz(SSN) <> 0 AND SSN <> ""999999999"" )ORDER BY ID, Last, First", dbOpenSnapshot)
Do While Not n.EOF
    cnt = cnt + 1
    AKA = ""
    ' 2012.01.30 DaveW Blank before AKA
    If Nz(n("AKA")) <> "" Then
      AKA = " " & n("AKA")
    End If
    
    Names(cnt) = n("First") & " " & n("Last") & AKA
    
    n.MoveNext
Loop

Dim J As Integer

J = 0
For i = 1 To cnt
    If (InStr(1, BorrowerNames, Names(i)) = 0) Then
      J = J + 1
      BorrowerNames = BorrowerNames & IIf(J > 1, " and ", "") & Names(i)
    End If
Next i
BorrowerMorgagorNamesCountSSN = J
End Function

Function FormatName2(Company As Variant, First As Variant, Last As Variant, _
    AKA As Variant, addr1 As Variant, addr2 As Variant, City As Variant, _
    State As Variant, Zip As Variant, Optional NewLine As String = vbNewLine) As String
Dim n As String

If Nz(Company) <> "" Then
    n = Company & NewLine
    If Nz(First) & Nz(Last) & Nz(AKA) <> "" Then

        n = n
    End If
End If

If Nz(First) & Nz(Last) & Nz(AKA) <> "" Then n = n & IIf(IsNull(First), "", First & " ") & Nz(Last) & NewLine

If Nz(AKA) <> "" Then
    If Left$(AKA, 1) = "," Then
        n = n & Trim$(Mid$(AKA, 2, Len(AKA) - 1)) & NewLine
    Else
        n = n & Trim$(AKA) & NewLine
    End If
End If

If Not IsNull(addr1) Then n = n & addr1 & NewLine
If Not IsNull(addr2) Then n = n & addr2 & NewLine
n = n & City
If Not IsNull(State) Then n = n & ", " & State
If Not IsNull(Zip) Then n = n & " " & FormatZip(Zip)
FormatName2 = n
End Function



Sub remove_Duplicates()
'dim strarray(5)
Dim newArray(20) As String
Dim myCol As Collection
Dim i As Long
Dim dup As Integer

Set myCol = New Collection
newArray(0) = "bbb"
newArray(1) = "bbb"
newArray(2) = "ccc"
newArray(3) = "ddd"
newArray(4) = "ddd"
On Error Resume Next
For i = LBound(newArray) To UBound(newArray)
myCol.Add 0, CStr(newArray(i))
If Err Then
newArray(i) = Empty
dup = dup + 1
Err.Clear
ElseIf dup Then
newArray(i - dup) = newArray(i)
newArray(i) = Empty
End If
Next
For i = LBound(newArray) To UBound(newArray)
MsgBox (newArray(i))
Debug.Print newArray(i)
Next

End Sub

Public Function GetNames1(File As Long, Fmt As Integer, conditions As String, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean, Optional Intext As Boolean) As String
'
' File: File number, or 0 for 'current' file
' fmt:  1 = comma separated list
'       2 = comma separated list, except AND last name
'       3 = one name per line
'       4 = signature lines
'       5 = signature lines using electronic signatures
'       6 = one name per line, except AND/OR last name
'       7 = command seperated list, except AND/OR last name
'       10 = name (last, first)
'       11 = SSN, comma separated list
'       12 = SSN last 4 only, comma separated list
'       20 = comma separated list, except AND last name + No Estate of "Name
'       60 = One name per line, except AND/OR last name + No Estate of "Name
'       99 = one name and ", et al. " if multiple names
'       100 = name (last), comma separated list  'WOT
'       110 = name (last) ONLY , comma separated last
'       150 = 'Current file' All names with no conditions ** use like this **  Getnames(150,FMT,"")
        

Dim n As Recordset
Dim Names(1 To 50) As String
Dim cnt As Integer
Dim i As Integer
Dim AKA As String

cnt = 0
    If File = 0 Then File = Forms![Case List]!FileNumber
        If NoSSN = True Then
            Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ") AND SSN = ""999-99-9999""  ORDER BY ID, Last, First", dbOpenSnapshot)
        ElseIf File = 150 Then File = Forms![Case List]!FileNumber
            Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " ORDER BY ID, Last, First", dbOpenSnapshot)
        Else
            Set n = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE FileNumber = " & File & " AND (" & conditions & ") ORDER BY ID, Last, First", dbOpenSnapshot)
        End If


Do While Not n.EOF
    cnt = cnt + 1
    If Nz(n("Company")) <> "" Then
      Names(cnt) = n("Company")
    Else
      AKA = ""
      If Nz(n("AKA")) <> "" Then
        If Fmt = 3 Then     ' put AKA on new line if formatting one name per line
            AKA = Trim$(n("AKA"))
            If Left$(AKA, 1) = "," Then AKA = Trim$(Mid$(AKA, 2, Len(AKA) - 1))
            AKA = NewLine & AKA
        Else
            ' 2012.01.30 DaveW Blank before AKA
            AKA = " " & n("AKA")
        End If
      End If
      Select Case CapsFmt
        Case 0      ' no change
            Select Case Fmt
                Case Is <= 9
                'If Application.CurrentProject.AllReports("Order to Docket Cover").IsLoaded = True Or Application.CurrentProject.AllReports("Report of sale").IsLoaded = True Then
                
                
                'If n("deceased") = True And Intext = False Then
                'Names(cnt) = "Estsate of " & n("First") & " " & n("Last") & AKA
                'Names(cnt) = "Estate of " & n("First") & " " & n("Last") & AKA

                'Else
                Names(cnt) = n("First") & " " & n("Last") & AKA
                'End If
'                Else
'                Names(cnt) = N("First") & " " & N("Last") & AKA
'                End If
                Case 10
                    Names(cnt) = n("Last") & ", " & n("First") & AKA
                Case 11
                    Names(cnt) = Nz(n("SSN"))
                Case 12
                    Names(cnt) = "xxx-xx-" & Right$(Nz(n("SSN"), "????"), 4)
                Case 20
                    Names(cnt) = n("First") & " " & n("Last") & AKA
                Case 60
                    Names(cnt) = n("First") & " " & n("Last") & AKA
                Case 99
                    Names(cnt) = n("First") & " " & n("Last") & AKA
                Case 100 ' Wells SOT 7/30 MC
                    Names(cnt) = UCase$(Nz(n("Last"))) & ", " & n("First")
                Case 110
                   Names(cnt) = Nz(n("Last"))
            End Select
        Case 1      ' entire name
            Names(cnt) = UCase$(Nz(n("First") & " " & n("Last") & AKA))
        Case 2      ' last name only
            Names(cnt) = n("First") & " " & UCase$(Nz(n("Last"))) & AKA
      End Select
    End If
    n.MoveNext
Loop

Select Case Fmt
    Case 1, 11, 12 ' comma separated list
        For i = 1 To cnt
            GetNames1 = GetNames1 & Names(i)
            If i < cnt Then GetNames1 = GetNames1 & ", "
        Next i
    Case 2      ' comma separated list, except AND last name
        For i = 1 To cnt
            GetNames1 = GetNames1 & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetNames1 = GetNames1 & ", "
                Else
                    GetNames1 = GetNames1 & " and "
                End If
            End If
        Next i
    Case 3, 10     ' one name per line
        For i = 1 To cnt
            GetNames1 = GetNames1 & Names(i) & NewLine
        Next i
    Case 4      ' signature lines
        For i = 1 To cnt
            GetNames1 = GetNames1 & "___________________________________" & NewLine & Names(i) & NewLine & NewLine
        Next i
    Case 5      ' electronic signatures
        For i = 1 To cnt
            GetNames1 = GetNames1 & "/s/ " & Names(i) & NewLine & NewLine & NewLine
        Next i
    Case 6      ' One name per line, except AND/OR last name
         For i = 1 To cnt
            If i = cnt - 1 Then
                GetNames1 = GetNames1 & Names(i) & " and/or " & NewLine
            Else
                GetNames1 = GetNames1 & Names(i) & NewLine
            End If
         Next i
    Case 7      ' comma separated list, except AND/OR last name
        For i = 1 To cnt
            GetNames1 = GetNames1 & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetNames1 = GetNames1 & ", "
                Else
                    GetNames1 = GetNames1 & " and/or "
                End If
            End If
        Next i
    Case 20 ' comma separated list, except AND last name + No Estate of "Name
          For i = 1 To cnt
            GetNames1 = GetNames1 & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetNames1 = GetNames1 & ", "
                Else
                    GetNames1 = GetNames1 & " and "
                End If
            End If
        Next i
    Case 60 ' One name per line, except AND/OR last name + No Estate of "Name
        For i = 1 To cnt
            If i = cnt - 1 Then
                GetNames1 = GetNames1 & Names(i) & " and/or " & NewLine
            Else
                GetNames1 = GetNames1 & Names(i) & NewLine
            End If
         Next i
    Case 99
         For i = 1 To 1
            GetNames1 = GetNames1 & Names(i)
            If i <= 1 Then
                If i < cnt Then
                    GetNames1 = GetNames1 & ", et al."
                Else
                    'GetNames = GetNames & " and "
                End If
            End If
        Next i
            
    Case 100
        For i = 1 To cnt
            If i = cnt - 1 Then
                GetNames1 = GetNames1 & Names(i) & NewLine
            Else
                GetNames1 = GetNames1 & Names(i) & NewLine
            End If
         Next i
    Case 110
         For i = 1 To cnt
            GetNames1 = GetNames1 & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    GetNames1 = GetNames1 & ", "
                Else
                    GetNames1 = GetNames1 & " and "
                End If
            End If
        Next i
        
        'For i = 1 To cnt
         '  If i = cnt - 1 Then
              ' GetNames = GetNames & Names(i) & ", "
          '  Else
               ' GetNames = GetNames & Names(i) & " and "
           ' End If
        ' Next i
    Case Else
        MsgBox "Invalid format in call to GetNames1", vbExclamation
        GetNames1 = ""
End Select
End Function

Public Function MortgagorNames1(File As Long, Fmt As Integer, Optional NewLine As String = vbNewLine, Optional NoSSN As Boolean) As String
MortgagorNames1 = GetNames1(File, Fmt, "Mortgagor = True", NewLine, NoSSN)

End Function

