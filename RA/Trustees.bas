Attribute VB_Name = "Trustees"
Option Compare Database
Option Explicit
 
Public TrusteeWordFile As Long
Dim TrusteeWordPlural As Integer
Dim TrusteeWordValue As String
Public Function SymbCollateral(sy As Long) As String

Dim kk As Long
kk = Forms![Print Affidavit Collateral file]!ts
Select Case sy
Case 1

If kk = 1 Or kk = 2 Or kk = 3 Or kk = 4 Or kk = 5 Or kk = 6 Or kk = 7 Or kk = 8 Then
    SymbCollateral = "R"
    Exit Function
    Else
    SymbCollateral = ""
    Exit Function
End If

Case 2
If kk = 2 Or kk = 3 Or kk = 4 Or kk = 5 Or kk = 6 Or kk = 7 Or kk = 8 Then
    SymbCollateral = "R"
    Else
    SymbCollateral = ""
    Exit Function
End If

Case 3
If kk = 3 Or kk = 4 Or kk = 5 Or kk = 6 Or kk = 7 Or kk = 8 Then
    SymbCollateral = "R"
    Else
    SymbCollateral = ""
    Exit Function
End If

Case 4
If kk = 4 Or kk = 5 Or kk = 6 Or kk = 7 Or kk = 8 Then
    SymbCollateral = "R"
    Else
    SymbCollateral = ""
    Exit Function
End If

Case 5
If kk = 5 Or kk = 6 Or kk = 7 Or kk = 8 Then
    SymbCollateral = "R"
    Else
    SymbCollateral = ""
    Exit Function
End If

Case 6
If kk = 6 Or kk = 7 Or kk = 8 Then
    SymbCollateral = "R"
    Else
    SymbCollateral = ""
    Exit Function
End If

Case 7
If kk = 7 Or kk = 8 Then
    SymbCollateral = "R"
    Else
    SymbCollateral = ""
    Exit Function
End If

Case 8
If kk = 8 Then
    SymbCollateral = "R"
    Else
    SymbCollateral = ""
    Exit Function
End If


    

End Select
End Function

Public Function CopyCollateral(ss As Long) As String

    
    Dim Chlis(1 To 8) As String
    Dim Chbox(1 To 8) As String
    Dim Chtex(1 To 8) As String
 
    Dim lngPositionf As Long
    Dim lngPosition As Long
    Dim lngPositionS As Long
    Dim i, J, K As Long
    K = 0
    Dim Otext As String
    Otext = Forms![Print Affidavit Collateral file].OtherDocText
       
    
    Chbox(1) = Forms![Print Affidavit Collateral file].N1
    Chbox(2) = Forms![Print Affidavit Collateral file].N2
    Chbox(3) = Forms![Print Affidavit Collateral file].N3
    Chbox(4) = Forms![Print Affidavit Collateral file].N4
    Chbox(5) = Forms![Print Affidavit Collateral file].N5
    Chbox(6) = Forms![Print Affidavit Collateral file].N6
    Chbox(7) = Forms![Print Affidavit Collateral file].N7
    Chbox(8) = Forms![Print Affidavit Collateral file].N8
    
    Chtex(1) = "Deed"
    Chtex(2) = "DOT"
    Chtex(3) = "Note"
    Chtex(4) = "Allonge"
    Chtex(5) = "Assignment"
    Chtex(6) = "Title Policy"
    Chtex(7) = "Title Commitment"
    Chtex(8) = Otext
    
    
    For i = 1 To 8
    If Chbox(i) = 1 Then
    K = K + 1
    Chlis(K) = Chtex(i)
    End If
    Next i
 
    If ss = 1 Then
    CopyCollateral = Chlis(1)
    Exit Function
    Else
    If ss = 2 Then
    CopyCollateral = Chlis(2)
    Exit Function
    Else
    If ss = 3 Then
    CopyCollateral = Chlis(3)
    Exit Function
    Else
    If ss = 4 Then
    CopyCollateral = Chlis(4)
    Exit Function
    Else
    If ss = 5 Then
    CopyCollateral = Chlis(5)
    Exit Function
    Else
    If ss = 6 Then
    CopyCollateral = Chlis(6)
    Exit Function
    Else
    If ss = 7 Then
    CopyCollateral = Chlis(7)
    Exit Function
    Else
    If ss = 8 Then
    CopyCollateral = Chlis(8)
    Exit Function
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If

    
    
    
 
End Function


Public Function TrusteeWord(File As Long, Plural As Integer) As String
'
' Return correct wording for Trustee/Substitute Trustee
' (Additional logic is needed to support mortgages)
'
' Parameters:
'   file    file number, or 0 for current file
'   plural  singular or plural form desired
'           0 = figure it out
'           1 = singular
'           2 = plural
'
Dim FC As Recordset, t As Recordset
Dim cnt As Integer

If File = 0 Then File = Forms![Case List]!FileNumber
'
' If the parameters are the same as last time, then used cached information,
' which will be much quicker than opening the database again.
'
If File = TrusteeWordFile And Plural = TrusteeWordPlural Then
    TrusteeWord = TrusteeWordValue
    Exit Function
End If

If Plural = 0 Then      ' need to figure out plural
    Set t = CurrentDb.OpenRecordset("SELECT * FROM qryTrustees WHERE FileNumber = " & File & ";", dbOpenSnapshot)
    If t.EOF Then
        cnt = 0
    Else
        t.MoveLast
        cnt = t.RecordCount
    End If
    t.Close
End If

Set FC = CurrentDb.OpenRecordset("SELECT SubstituteTrustees FROM FCdetails WHERE Current=True AND FileNumber = " & File & ";", dbOpenSnapshot)
If FC("SubstituteTrustees") Then
    Select Case Plural
        Case 0
            If cnt = 1 Then
                TrusteeWordValue = "Substitute Trustee"
            Else
                TrusteeWordValue = "Substitute Trustees"
            End If
        Case 1
            TrusteeWordValue = "Substitute Trustee"
        Case 2
            TrusteeWordValue = "Substitute Trustees"
    End Select
Else
    Select Case Plural
        Case 0
            If cnt = 1 Then
                TrusteeWordValue = "Trustee"
            Else
                TrusteeWordValue = "Trustees"
            End If
        Case 1
            TrusteeWordValue = "Trustee"
        Case 2
            TrusteeWordValue = "Trustees"
    End Select
End If
FC.Close

TrusteeWordFile = File
TrusteeWordPlural = Plural
TrusteeWord = TrusteeWordValue
End Function

Public Function MonitorTrusteeNames(File As Long, Fmt As Integer, Optional SignatureSuffix As String) As String
Dim TName As Variant

TName = DLookup("MonitorTrusteeName", "FCDetails", "Filenumber=" & File & " AND Current=True")
If IsNull(TName) Then
    MsgBox "No trustee is assigned to this file.  Document will not print correctly.", vbExclamation
    MonitorTrusteeNames = ""
    Exit Function
End If

Select Case Fmt
    Case 1, 2
        MonitorTrusteeNames = TName
    Case 3
        MonitorTrusteeNames = TName & vbNewLine
    Case 4      ' signature lines
        MonitorTrusteeNames = "___________________________________" & vbNewLine & TName
        If Not IsMissing(SignatureSuffix) Then
            MonitorTrusteeNames = MonitorTrusteeNames & vbNewLine & SignatureSuffix
        End If
        MonitorTrusteeNames = MonitorTrusteeNames & vbNewLine & vbNewLine
    Case Else
        MsgBox "Invalid format in call to MonitorTrusteeNames", vbExclamation
        MonitorTrusteeNames = ""
End Select
End Function

Public Function trusteeNames(File As Long, Fmt As Integer, Optional SignatureSuffix As String, Optional NewLine As String = vbNewLine) As String
'
' File: File number, or 0 for 'current' file
' Fmt:  1 = comma separated list
'       2 = comma separated list, except AND last name
'       3 = one name per line
'       4 = signature lines
'
Dim t As Recordset
Dim Names(1 To 50) As String
Dim SigPrefix(1 To 50) As String
Dim cnt As Integer
Dim i As Integer

cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber

If DLookup("CaseTypeID", "CaseList", "Filenumber=" & File) = 8 Then    ' monitor
    trusteeNames = MonitorTrusteeNames(File, Fmt, SignatureSuffix)
    Exit Function
End If

Set t = CurrentDb.OpenRecordset("SELECT * FROM qryTrustees WHERE FileNumber = " & File & " ORDER BY Assigned;", dbOpenSnapshot)
Do While Not t.EOF
    cnt = cnt + 1
    Names(cnt) = t("Name")
    SigPrefix(cnt) = Nz(t("SignatureLine"))
    t.MoveNext
Loop
t.Close

If cnt = 0 Then
    MsgBox "No trustees are assigned to this file.  Document will not print correctly.", vbExclamation
    trusteeNames = ""
    Exit Function
End If

Select Case Fmt
    Case 1      ' comma separated list
        For i = 1 To cnt
            trusteeNames = trusteeNames & Names(i)
            If i < cnt Then trusteeNames = trusteeNames & ", "
        Next i
    Case 2      ' comma separated list, except AND last name
        For i = 1 To cnt
            trusteeNames = trusteeNames & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    trusteeNames = trusteeNames & ", "
                Else
                    trusteeNames = trusteeNames & " and "
                End If
            End If
        Next i
    Case 3      ' one name per line
        For i = 1 To cnt
            trusteeNames = trusteeNames & Names(i) & NewLine
        Next i
    Case 4      ' signature lines
        For i = 1 To cnt
            trusteeNames = trusteeNames & "___________________________________" & NewLine
            If SigPrefix(cnt) <> "" Then
                trusteeNames = trusteeNames & SigPrefix(cnt) & NewLine
            End If
            trusteeNames = trusteeNames & Names(i)
            If Not IsMissing(SignatureSuffix) Then
                trusteeNames = trusteeNames & NewLine & SignatureSuffix
            End If
            trusteeNames = trusteeNames & NewLine & NewLine
        Next i
    Case Else
        MsgBox "Invalid format in call to TrusteeNames", vbExclamation
        trusteeNames = ""
End Select
End Function

Public Function CopyExhibits(ss As Long) As String

    
    Dim Chlis(1 To 6) As String
    Dim Chbox(1 To 6) As String
    Dim Chtex(1 To 6) As String
 
    Dim lngPositionf As Long
    Dim lngPosition As Long
    Dim lngPositionS As Long
    Dim i, J, K As Long
    K = 0
    
       
    
    Chbox(1) = Forms![Print Exhibit 9].N1
    Chbox(2) = Forms![Print Exhibit 9].N2
    Chbox(3) = Forms![Print Exhibit 9].N3
    Chbox(4) = Forms![Print Exhibit 9].N4
    Chbox(5) = Forms![Print Exhibit 9].N5
    Chbox(6) = Forms![Print Exhibit 9].N6
        
        
        
    Chtex(1) = "Attached hereto as Exhibit " & Exhibits9R(1) & " is a true and correct copy of the Debt Agreement."
    Chtex(2) = "Attached hereto as Exhibit " & Exhibits9R(2) & " is a true and correct copy of the Dees of Trust"
    Chtex(3) = "Attached hereto as Exhibit " & Exhibits9R(3) & " is a true and correct copy of the Loan Modification Agreement."
    Chtex(4) = "Attached hereto as Exhibit " & Exhibits9R(4) & " is a post-petition payment history."
    Chtex(5) = "Attached hereto as Exhibit " & Exhibits9R(5) & " is an addendum listing all fees and charges assessed to the account of the Debtor(s) post-petition."
    Chtex(6) = "Attached hereto as Exhibit " & Exhibits9R(6) & " is an addendum listing all post-petition taxes and insurance advances"
        
    
    
    
    
    
    For i = 1 To 6
    If Chbox(i) = 1 Then
    K = K + 1
    Chlis(K) = Chtex(i)
    End If
    Next i
 
    If ss = 1 Then
    CopyExhibits = Chlis(1)
    Exit Function
    Else
    If ss = 2 Then
    CopyExhibits = Chlis(2)
    Exit Function
    Else
    If ss = 3 Then
    CopyExhibits = Chlis(3)
    Exit Function
    Else
    If ss = 4 Then
    CopyExhibits = Chlis(4)
    Exit Function
    Else
    If ss = 5 Then
    CopyExhibits = Chlis(5)
    Exit Function
    Else
    If ss = 6 Then
    CopyExhibits = Chlis(6)
    Exit Function
    End If
    End If
    End If
    End If
    End If
    End If
    
    
    End Function
    

Public Function SymbExhibits(SE As Long) As String

Dim Exh As Long
Exh = Forms![Print Exhibit 9]!ts
Select Case SE
Case 1

If Exh = 1 Or Exh = 2 Or Exh = 3 Or Exh = 4 Or Exh = 5 Or Exh = 6 Then
    SymbExhibits = "(a)"
    Exit Function
    Else
    SymbExhibits = ""
    Exit Function
End If

Case 2
If Exh = 2 Or Exh = 3 Or Exh = 4 Or Exh = 5 Or Exh = 6 Then
    SymbExhibits = "(b)"
    Else
    SymbExhibits = ""
    Exit Function
End If

Case 3
If Exh = 3 Or Exh = 4 Or Exh = 5 Or Exh = 6 Then
    SymbExhibits = "(c)"
    Else
    SymbExhibits = ""
    Exit Function
End If

Case 4
If Exh = 4 Or Exh = 5 Or Exh = 6 Then
    SymbExhibits = "(D)"
    Else
    SymbExhibits = ""
    Exit Function
End If

Case 5
If Exh = 5 Or Exh = 6 Then
    SymbExhibits = "(E)"
    Else
    SymbExhibits = ""
    Exit Function
End If
  
Case 6
If Exh = 6 Then
    SymbExhibits = "(F)"
    Else
    SymbExhibits = ""
    Exit Function
End If


End Select
End Function

Public Function Exhibits9R(SE As Long) As String
Dim Exh As Long
Exh = Forms![Print Exhibit 9]!ts



Select Case SE

Case 1
    If Forms![Print Exhibit 9]!A = True Then
      
    
    Select Case Exh
    Case 1, 2, 3, 4, 5, 6
    Exhibits9R = "A"
    Exit Function
    End Select
    Else: Exhibits9R = ""
    Exit Function
    End If

Case 2
    If Forms![Print Exhibit 9]!B = True Then
    Dim C2 As Integer
    C2 = 0
    If Forms![Print Exhibit 9]!A = True Then C2 = C2 + 1
        Select Case Exh
        Case 1
        Exhibits9R = "A"
        Exit Function
        
        Case 2, 3, 4, 5, 6
        If C2 = 1 Then
        Exhibits9R = "B"
        Exit Function
        Else
        
        Exhibits9R = "A"
        Exit Function
        End If
        End Select
    Else
    Exhibits9R = ""
    Exit Function
    End If
    

Case 3
    If Forms![Print Exhibit 9]!C = True Then
    Dim c3 As Integer
    c3 = 0
    If Forms![Print Exhibit 9]!A = True Then c3 = c3 + 1
    If Forms![Print Exhibit 9]!B = True Then c3 = c3 + 1
    
    Select Case Exh
        Case 1
        Exhibits9R = "A"
        Exit Function
        
        Case 2
        If c3 = 1 Then
        Exhibits9R = "B"
        Exit Function
        Else
        Exhibits9R = "A"
        Exit Function
        End If
        
        Case 3, 4, 5, 6
            Select Case c3
            Case 0
            Exhibits9R = "A"
            Exit Function
            Case 1
            Exhibits9R = "B"
            Exit Function
            Case 2
            Exhibits9R = "C"
            Exit Function
            End Select
            
           
        End Select
    Else
    Exhibits9R = ""
    Exit Function
    End If


Case 4

If Forms![Print Exhibit 9]!d = True Then

Dim C4 As Integer
C4 = 0

If Forms![Print Exhibit 9]!A = True Then C4 = C4 + 1
If Forms![Print Exhibit 9]!B = True Then C4 = C4 + 1
If Forms![Print Exhibit 9]!C = True Then C4 = C4 + 1
    Select Case Exh
        Case 1
        Exhibits9R = "A"
        Exit Function
        
        Case 2
        If C4 > 0 Then
        Exhibits9R = "B"
        Exit Function
        Else
        Exhibits9R = "A"
        Exit Function
        End If

        Case 3, 4, 5, 6
        
        Select Case C4
         Case 0
         Exhibits9R = "A"
         Exit Function
         Case 1
         Exhibits9R = "B"
         Exit Function
         Case 2
         Exhibits9R = "C"
         Exit Function
         Case 3
         Exhibits9R = "D"
         Exit Function
        End Select
                            
            End Select
            Else
            Exhibits9R = ""
            Exit Function
            End If



Case 5
If Forms![Print Exhibit 9]!E = True Then
Dim C5 As Integer
C5 = 0

If Forms![Print Exhibit 9]!A = True Then C5 = C5 + 1
If Forms![Print Exhibit 9]!B = True Then C5 = C5 + 1
If Forms![Print Exhibit 9]!C = True Then C5 = C5 + 1
If Forms![Print Exhibit 9]!d = True Then C5 = C5 + 1

Select Case Exh
        Case 1
        Exhibits9R = "A"
        Exit Function
        
        Case 2
        If C5 > 0 Then
        Exhibits9R = "B"
        Exit Function
        Else
        Exhibits9R = "A"
        Exit Function
        End If

        Case 3, 4, 5, 6
        Select Case C5
        Case 0
        Exhibits9R = "A"
        Exit Function
        Case 1
        Exhibits9R = "B"
        Exit Function
        Case 2
        Exhibits9R = "C"
        Exit Function
        Case 3
        Exhibits9R = "D"
        Exit Function
        Case 4
        Exhibits9R = "E"
        Exit Function
        End Select
        
                      
                   
            End Select
            Else
            Exhibits9R = ""
            Exit Function
            End If


Case 6
If Forms![Print Exhibit 9]!F = True Then
Select Case Exh
        Case 1
        Exhibits9R = "A"
        Exit Function
        Case 2
        Exhibits9R = "B"
        Exit Function
        Case 3
        Exhibits9R = "C"
        Exit Function
        Case 4
        Exhibits9R = "D"
        Exit Function
        Case 5
        Exhibits9R = "E"
        Exit Function
        Case 6
        Exhibits9R = "F"
        Exit Function
        End Select
        
            Else
            Exhibits9R = ""
            Exit Function
            End If

End Select
  


End Function
