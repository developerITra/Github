Attribute VB_Name = "WordDocuments"
Option Compare Database
Option Explicit
Dim NextNumber, newNextNumber, lettercounter As Integer

Private Function GetNextLetter() As String
Dim strAlpha As String
Dim t As String

strAlpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'If LetterCounter = 0 Then
'    LetterCounter = 1
'Else
'End If
 
    t = Mid(strAlpha, lettercounter, 1)
   
    GetNextLetter = t
    lettercounter = lettercounter + 1

End Function
Private Function GetNextNumber() As String

GetNextNumber = Space$(10) & NextNumber & ".  "
NextNumber = NextNumber + 1

End Function


Private Function GetNewNextNumber() As String

GetNewNextNumber = Space$(10) & newNextNumber & ".  "
newNextNumber = newNextNumber + 1

End Function

'Private Sub FormattedTextToWord()
'Dim objWord As Object  '' Word.Application
'Dim fso As Object  '' FileSystemObject
'Dim f As Object  '' TextStream
'Dim myHtml As String, tempFileSpec As String
'
''' grab some formatted text from a Memo field
'myHtml = DLookup("LineDescription", "Line", "FileNumber=31600")
'
'Set fso = CreateObject("Scripting.FileSystemObject")  '' New FileSystemObject
'tempFileSpec = fso.GetSpecialFolder(2) & "\" & fso.GetTempName & ".htm"
'
''' write to temporary .htm file
'Set f = fso.CreateTextFile(tempFileSpec, True)
'f.Write "<html>" & myHtml & "</html>"
'f.Close
'Set f = Nothing
'
'Set objWord = CreateObject("Word.Application")  '' New Word.Application
'objWord.Documents.Add
'objWord.Selection.InsertFile tempFileSpec
'fso.DeleteFile tempFileSpec
''' the Word document now contains formatted text
'
''objWord.ActiveDocument.SaveAs2 "C:\Users\mcross\zzzTest.rtf", 6  '' 6 = wdFormatRTF
'objWord.ActiveDocument.saveas2 FileName:="Test.rtf", FileFormat:=wdFormatRTF
'objWord.Quit
'Set objWord = Nothing
'Set fso = Nothing
'End Sub

Private Function AssignmentInfo() As String
Select Case Forms!BankruptcyDetails!AssignBy
    Case 1              ' DOT
        Select Case Forms!BankruptcyDetails!AssignByDOT
            Case 1      ' assignment
                AssignmentInfo = Forms!BankruptcyDetails!OriginalBeneficiary & " assigned its interest to " & Forms![Case List]!Investor
            Case 2      ' merger
                AssignmentInfo = Trim$(Forms!BankruptcyDetails!OMergerInfo) & " is now the beneficiary of the Deed of Trust due to merger."
        End Select
    Case 2              ' note
        AssignmentInfo = "The Promissory Note has been transferred from " & Forms!BankruptcyDetails!OriginalBeneficiary & " to " & Forms![Case List]!Investor
        
    
        

If Right$(AssignmentInfo, 1) <> "." Then AssignmentInfo = AssignmentInfo & "."


End Select

If IsNull(Right$(AssignmentInfo, 1)) Then AssignmentInfo = ""

End Function



Private Sub FillField(WordDocument As Word.Document, FieldName As String, Replacement As String)
Dim ReplacementData As String, ReplacementString As String, FindData As String, ChunkBegins As Integer, ChunkSize As Integer
'test
ReplacementString = RemoveLF(Replacement)

If Len(ReplacementString) > 200 Then    ' divide into chunks
    FindData = FieldName
    ChunkBegins = 1
    Do While ChunkBegins < Len(ReplacementString)
        ChunkSize = 200
        If ChunkBegins + ChunkSize > Len(ReplacementString) Then
            ChunkSize = Len(ReplacementString) - ChunkBegins + 1
        End If
        Do While Mid$(ReplacementString, ChunkSize, 1) = "^"    ' don't split a 'symbol'
            ChunkSize = ChunkSize - 1
        Loop
        ReplacementData = Mid$(ReplacementString, ChunkBegins, ChunkSize) & "<<<<More>>>>"
        GoSub Replace
        FindData = "<<More>>"
        ChunkBegins = ChunkBegins + ChunkSize
    Loop
    ReplacementData = ""    ' remove last <<More>> tag
    GoSub Replace
Else
    FindData = FieldName
    ReplacementData = ReplacementString
    GoSub Replace
End If
Exit Sub

Replace:
Dim rngStory As Word.Range, lngJunk As Long, oShp As Shape

' This may be needed work around a bug with certain headers/footers.
' But the side effect is to create headers and footers that take up space (making the margins look taller.)
''lngJunk = WordDocument.Sections(1).Headers(1).Range.StoryType

'Iterate through all story types in the current document
For Each rngStory In WordDocument.StoryRanges
    'Iterate through all linked stories
    Do
        With rngStory.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "<<" & FindData & ">>"
            .Replacement.Text = ReplacementData
            .Execute Replace:=wdReplaceAll
        End With
        On Error Resume Next
        If rngStory.ShapeRange.Count > 0 Then
            For Each oShp In rngStory.ShapeRange
                If oShp.TextFrame.HasText Then
                    With oShp.TextFrame.TextRange.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .Text = "<<" & FindData & ">>"
                        .Replacement.Text = ReplacementData
                        .Execute Replace:=wdReplaceAll
                    End With
                End If
            Next
        End If
        On Error GoTo 0
        'Get next linked story (if any)
        Set rngStory = rngStory.NextStoryRange
    Loop Until rngStory Is Nothing
Next

Return
End Sub

Private Sub FillField_Old(WordDocument As Word.Document, FieldName As String, ReplacementData As String)
With WordDocument.Content.Find
    .ClearFormatting
    With .Replacement
        .ClearFormatting
    End With
    .Execute FindText:="<<" & FieldName & ">>", ReplaceWith:=ReplacementData, Format:=True, Replace:=wdReplaceAll
End With
End Sub

Private Sub SaveDoc(WordDoc As Word.Document, FileNumber As Long, FileName As String)
Dim Filespec As String

Filespec = DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & FileName
Do Until Dir$(Filespec) = ""
    FileName = InputBox$(FileName & " already exists, please enter another name:", , FileName)
    If Len(FileName) > 0 Then
        If UCase$(Right$(FileName, 4)) <> ".DOC" Then FileName = FileName & ".doc"
        Filespec = DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & FileName
    Else
        Exit Sub
    End If
Loop
WordDoc.SaveAs Filespec

End Sub

Public Sub Doc_OwnershipAffidavitMDHC(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit MDHC.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])



FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "Investor2", IIf(d![ClientID] = 385, d![Investor2], d![Investor])
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "NoteOwner", FetchNoteOwner(d!LoanType, IIf(d![ClientID] = 385, d![Investor2], d![Investor]), d![FCdetails.State])
FillField WordDoc, "Noteholders", GetNamesMD(0, 2, "Noteholder=True")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & d!Investor
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "FHLMCWording", IIf(d!LoanType = 5 Or d!LoanType = 4, ", and " & IIf(d![ClientID] = 385, d![Investor2], d![Investor]) & " is the holder of the Note having been transferred to " & IIf(d![ClientID] = 385, d![Investor2], d![Investor]) & " for the purposes of enforcement and conducting this foreclosure action", "")
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 464, "")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ____________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ____________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: __________________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Note Ownership Affidavit MDHC.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Note Ownership Affidavit MDHC.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_StatementOfDebtWithFiguresOcwen(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim i As Integer, itemCount As Integer, totalItems As Currency, itemsFields As String
Dim InterestFrom As String
Dim InterestTo As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Statement of Debt Figures Ocwen"
templateName = templateName & ".dot"

InterestFrom = InputBox("Interest From")
InterestTo = InputBox("Interest To")


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


WordObj.Visible = False
If MsgBox("Is There A Loan Mod? ", vbYesNo) = vbYes Then
FillField WordDoc, "Mod", ", MODIFIED by Agreement effective " & Format(InputBox(" Effective Date?   Format mm/dd/yyyy"), "mmmm d, yyyy") & " with an amended principal balance of " & Format(InputBox(" Amended Principal Balance? "), "Currency")
Else
FillField WordDoc, "Mod", ""
End If
WordObj.Visible = True

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber

FillField WordDoc, "Investor", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "") & d!Investor

FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
'FillField WordDoc, "455Only", IIf(D![ClientID] = 455, ", and continuing each month thereafter with probable future advancements made by the mortgagee, and that the Plaintiff(s) has\have the right to foreclose;", ", and continuing each month thereafter, and that the Plaintiff(s) has\have the right to foreclose;")
FillField WordDoc, "Liber", d![Liber]
FillField WordDoc, "Folio", d![Folio]
FillField WordDoc, "LPIdate", Format$(d![LPIDate], "mmmm d, yyyy")
FillField WordDoc, "LPIdate+1", Format$(d![LPIDate] + 1, "mmmm d, yyyy")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
'FillField WordDoc, "PaidStr", IIf((Nz(D![RemainingPBal], 0) > Nz(D![OriginalPBal], 0)), "Additional Interest", "Paid on principal")
'FillField WordDoc, "Paid", Format$(D!OriginalPBal - D!RemainingPBal, "Currency")
FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")
'FillField WordDoc, "diem", IIf(D![LoanType] = 3, IIf(D![ClientID] <> 531, "Per Monthly Interest: ", "Per Diem Interest: "), "Per Diem Interest: ")
'FillField WordDoc, "txtbalanc", IIf(D![ClientID] = 361, "Unpaid Principal Balance", "Remaining Balance Due")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

totalItems = 0
itemsFields = ""

Set dd = CurrentDb.OpenRecordset("SELECT Desc, Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber] & " ORDER BY StatementOfDebt.Sort_Desc DESC;", dbOpenSnapshot)
If dd.EOF Then      ' no extra lines  'More like Not an empty recordset you mean
    FillField WordDoc, "Line_Items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items", itemsFields
    dd.MoveFirst
    i = 1
    Do While Not dd.EOF
        FillField WordDoc, "Item" & i, IIf(dd!Desc = "Interest", "Interest at " & Format$(Forms![Print Statement of Debt]!InterestRate, "#0.000%") & " for " & [InterestFrom] & " to " & [InterestTo] & vbTab & Format$(Nz(dd!Amount, 0), "Currency"), dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency"))
        'FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        totalItems = totalItems + Nz(dd!Amount, 0)
        dd.MoveNext
        i = i + 1
       
    Loop
End If
dd.Close

FillField WordDoc, "BalDueDate", IIf(IsNull([Forms]![Print Statement of Debt]![txtDueDate]), "______________", Format$([Forms]![Print Statement of Debt]![txtDueDate], "mmmm d"", ""yyyy"))
FillField WordDoc, "BalanceDue", Format$(d!RemainingPBal + totalItems, "Currency")
FillField WordDoc, "PerDiemInterest", IIf(IsNull(d!PerDiem), "$_____________", Format$(d!PerDiem, "Currency"))
FillField WordDoc, "InterestRate", IIf(IsNull(d!InterestRate), "____________ %", Format(d!InterestRate, "#0.000") & "%")
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", IIf(IsNull(d!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", d!NotaryLocation)

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Statement of Debt with Figures Ocwen.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Statement of Debt with Figures Ocwen.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_Ad_MD(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Ad MD.dot", False, 0, True)
WordObj.Visible = True

FillField WordDoc, "TrusteeWord", UCase$(TrusteeWord(d![CaseList.FileNumber], 0))
FillField WordDoc, "Property", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & vbCr & d!City & ", " & d![FCdetails.State] & ", " & FormatZip(d!ZipCode)
FillField WordDoc, "Mortgagors", IIf(IsNull(d!OriginalMortgagors), MortgagorNames(d![CaseList.FileNumber], 2), d!OriginalMortgagors & " assumed by " & MortgagorNames(d![CaseList.FileNumber], 2))
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "LiberFolio", LiberFolio(d!Liber, d!Folio, d![FCdetails.State])
FillField WordDoc, "Jurisdiction", d!Jurisdiction & ", " & d![FCdetails.State]
FillField WordDoc, "SaleLocation", d!SaleLocation
FillField WordDoc, "SaleDate", Format$(d!Sale, "mmmm d, yyyy")
FillField WordDoc, "SaleTime", Format$(d!SaleTime, "h:nn AM/PM")
FillField WordDoc, "Ownership", IIf(d!Leasehold = 1, "LEASEHOLD", "FEE-SIMPLE")
FillField WordDoc, "GroundRent", IIf(d!Leasehold = 1, "subject to annual ground rent of " & Format$(d!GroundRentAmount, "Currency") & ", payable " & d!GroundRentPayable & ", ", "")
FillField WordDoc, "DOTWord", DOTWord(d!DOT)
FillField WordDoc, "PriorLien", IIf(d!LienPosition <= 1, "", "          The property will be sold subject to a prior mortgage, the amount to be announced at the time of sale." & vbCr)
FillField WordDoc, "IRSLiens", IIf(d!IRSLiens, "The property will be sold subject to a 120 day right of redemption by the Internal Revenue Service.  ", "")
FillField WordDoc, "Deposit", Format$(d!Deposit, "Currency")
FillField WordDoc, "Trustees", trusteeNames(d![CaseList.FileNumber], 2)

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Ad MD.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Ad MD.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub
Public Sub doc_SOTAffidavitMD(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qrySOTAffFCDocsLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "SOT Affidavit MD.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName] '& vbct StopBySarab
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(IsNull(d![Fair Debt]), " ", ", " & d![Fair Debt] & ", ") & d!City & ", " & d![FCdetails.State] & " " & Format(d!ZipCode)

'StppedBySarab 'WordDoc.Bookmarks("PropertyAddress").Range.Text = vbct & d![PropertyAddress] & IIf(IsNull(d![Fair Debt]), " ", ", " & d![Fair Debt] & ", ") & d!City & ", " & d![FCdetails.State] & " " & Format(d!ZipCode)
'WordDoc.Bookmarks("APTNum").Select
'WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d![Jurisdiction] & ", " & d![LongState])
FillField WordDoc, "CourtCaseNumber", IIf(IsNull(d!CourtCaseNumber), "", d!CourtCaseNumber)
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3)
FillField WordDoc, "FirmShortAddress", FirmShortAddress()
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(IsNull(d![Fair Debt]), "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "TrusteeNames2", trusteeNames(0, 3)


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "SOT Affidavit MD.dot.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "SOT Affidavit MD.dot.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub
Public Sub Doc_Ad_VA(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Ad VA.dot", False, 0, True)
WordObj.Visible = True

FillField WordDoc, "Property", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d!City & ", " & d![FCdetails.State] & " " & FormatZip(d!ZipCode)
FillField WordDoc, "DOTWord", DOTWord(d!DOT)
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "LiberFolio", LiberFolio(d!Liber, d!Folio, d![FCdetails.State])
FillField WordDoc, "Jurisdiction", d!Jurisdiction & ", " & d![FCdetails.State]
FillField WordDoc, "OriginalBalance", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "TrusteeWord", UCase$(TrusteeWord(d![CaseList.FileNumber], 0))
FillField WordDoc, "Trustees", trusteeNames(d![CaseList.FileNumber], 2)
FillField WordDoc, "SaleLocation", d!SaleLocation
FillField WordDoc, "SaleDate", Format$(d!Sale, "mmmm d, yyyy")
FillField WordDoc, "SaleTime", Format$(d!SaleTime, "h:nn AM/PM")
FillField WordDoc, "ShortLegal", IIf(d![UseFullLegal] = True, d![LegalDescription], d![ShortLegal])
FillField WordDoc, "PriorLien", IIf(d!LienPosition <= 1, "", "The property will be sold subject to " & IIf(d!LienPosition = 2, "a prior mortgage, the amount", "prior mortgages, the amounts") & " to be announced at the time of sale." & vbCr)
FillField WordDoc, "IRSLiens", IIf(d!IRSLiens, "The property will be sold subject to a 120 day right of redemption by the Internal Revenue Service.  ", "")
FillField WordDoc, "Deposit", Format$(d!Deposit, "Currency")
FillField WordDoc, "Contact", DLookup("sValue", "DB", "Name='ContactPhone'")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Ad VA.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Ad VA.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_Ad_DC(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
'Set WordDoc = WordObj.Documents.Add(TemplatePath & "Ad DC.dot", False, 0, True)
Set WordDoc = WordObj.Documents.Add(TemplatePath & "DC Sale ad.dot", False, 0, True)
WordObj.Visible = True



FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "TrusteeWord", UCase$(TrusteeWord(d![CaseList.FileNumber], 0))
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d!City & ", " & d![FCdetails.State] & ", " & FormatZip(d!ZipCode)
'FillField WordDoc, "DOTRecorded", Format$(d!DOTrecorded, "mmmm d, yyyy")
FillField WordDoc, "DoTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "Liber", d!Liber
'FillField WordDoc, "Docket", Format$(d!Docket, "mmmm d, yyyy")
FillField WordDoc, "SaleDate", Format$(d!Sale, "dddd, mmmm d, yyyy")
FillField WordDoc, "SaleTime", Format$(d!SaleTime, "h:nn AM/PM")
'FillField WordDoc, "TaxID", d!TaxID
'FillField WordDoc, "IRSLiens", IIf(d!IRSLiens, "The property will be sold subject to a 120 day right of redemption by the Internal Revenue Service.  ", "")
'FillField WordDoc, "PriorLien", IIf(d!LienPosition <= 1, "", "The property will be sold subject to a prior mortgage, the amount to be announced at the time of sale.  ")
FillField WordDoc, "Deposit", Format$(d!Deposit, "Currency")
FillField WordDoc, "TrusteeName", trusteeNames(d![CaseList.FileNumber], 2)
FillField WordDoc, "Month", Format(Date, "mmmm")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Ad DC.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Ad DC.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub
 Public Sub doc_SCRANotice90Day(Keepopen As Boolean)
 
 Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
 Set WordObj = CreateObject("Word.Application")
 Set WordDoc = WordObj.Documents.Add(TemplatePath & "SCRA NOTICE 90-Day.dot", False, 0, True)
 WordObj.Visible = True
 
 FillField WordDoc, "loginName", GetLoginName()
 End Sub

Public Sub Doc_DeedApp(Keepopen As Boolean) ' Deed of Appointment VA/MD
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Deed of Appointment"
If (d![FCdetails.State] = "VA") Then
  templateName = templateName & " VA"
End If


If d![ClientID] = 345 Then
  templateName = templateName & " Kondaur"
ElseIf d!ClientID = 466 Then 'SELECT
    templateName = templateName & " Select"

ElseIf (d![ClientID] = 334 Or d![ClientID] = 477) Then
  templateName = templateName & " Saxon"
'ElseIf d![ClientID] = 87 Then
'  templateName = templateName & " PNC"
ElseIf (d![ClientID] = 523 Or d![ClientID] = 258) And d![FCdetails.State] = "MD" Then
  templateName = templateName & " Green Tree"
ElseIf d![ClientID] = 404 Then
  templateName = templateName & " Bogman"
ElseIf (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "VA") Then
    templateName = templateName & " Dove LPP"
ElseIf (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "MD") Then
    templateName = templateName & " Dove LPP"
ElseIf (d![ClientID] = 451 And d![FCdetails.State] = "VA") Then
    templateName = templateName & " Dove"
ElseIf d!ClientID = 451 Then
    templateName = templateName & " Dove"
End If



templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")

Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

If d!ClientID = 451 Then
    WordDoc.Bookmarks("LastName").Select
    WordDoc.Bookmarks("LastName").Range.Text = d![PrimaryLastName]
End If

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])



If d!JurisdictionID = 56 And d![FCdetails.State] = "VA" Then
    FillField WordDoc, "Tax", "Tax Map Number:"
Else
    FillField WordDoc, "Tax", "Tax ID #:"
End If

If d!ClientID = 466 Then
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Liber " & d![Liber2] & " at Folio " & d![Folio2])

ElseIf (d![FCdetails.State] = "VA") Then
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", IIf(IsNull(d![Folio2]), ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & _
    " at Instrument Number " & d![Liber2], ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Book " & IIf(IsNull(d![Liber2]), " ", d![Liber2]) & ", Page " & d![Folio2]))
ElseIf d!JurisdictionID = 6 Then  'Calvert County
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Book " & d![Liber2] & ", Page " & d![Folio2])
Else 'MD or DC
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Liber " & d![Liber2] & ", Folio " & d![Folio2])
End If

FillField WordDoc, "LoanNumber", d!LoanNumber

If d!JurisdictionID = 153 Then 'Accomack VA
    FillField WordDoc, "OriginalTrustee", UCase$(d!OriginalTrustee)
    'FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & UCase$(d!Investor)
    FillField WordDoc, "InvestorAIF", IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf((d!AIF = True And d!ClientID = 532), UCase$(d![Investor]) & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for " & UCase$(d![Investor]), IIf(d![AIF] = True And d!ClientID <> 73, d![LongClientName] & " as Attorney in Fact for " & UCase$(d![Investor]), _
                                UCase$(d![Investor])))))) & IIf(d!ClientID = 73, ", By Carrington Mortgage Services, LLC as Attorney in Fact and Servicing Agent", "")
    FillField WordDoc, "OriginalBeneficiary", UCase$(d!OriginalBeneficiary)
    FillField WordDoc, "AssumedBy", IIf(IsNull(d!OriginalMortgagors), "", "assumed by " & UCase$(MortgagorNamesCaps(0, 2, 2))) & " "
    FillField WordDoc, "MortgagorNames", IIf(IsNull(d!OriginalMortgagors), UCase$(MortgagorNamesCaps(0, 2, 2)), UCase$(IIf(IsNull(d!OriginalMortgagors), " ", d!OriginalMortgagors)))
    FillField WordDoc, "TrusteeNames", UCase$(trusteeNames(0, 2))
    FillField WordDoc, "Investor", UCase$(d!Investor)
Else
    FillField WordDoc, "OriginalTrustee", d!OriginalTrustee
    
    FillField WordDoc, "InvestorAIF", IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", _
    IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", _
    IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", _
    IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for " & d![Investor], _
    IIf((d![AIF] = True And d!ClientID = 466), d![Investor] & ", by Select Portfolio Servicing, Inc., as attorney-in-fact", _
    IIf((d![AIF] = True And d!ClientID = 73 And d![FCdetails.State] = "MD"), d![Investor] & ", By Carrington Mortgage Services, LLC as Attorney in Fact and Servicing Agent", _
    IIf(d![AIF] = True And d!ClientID <> 73, d![LongClientName] & " as Attorney in Fact for " & d![Investor], _
    IIf(d!AIF = True And d!ClientID = 73, d![Investor] & ", By Carrington Mortgage Services, LLC as Attorney in Fact and Servicing Agent", d!Investor))))))))
    
    FillField WordDoc, "OriginalBeneficiary", d!OriginalBeneficiary & IIf((d!ClientID = 451 And d!MERS = -1), ", its successors and assigns", "")
    FillField WordDoc, "AssumedBy", IIf(IsNull(d!OriginalMortgagors), "", "assumed by " & MortgagorNamesCaps(0, 2, 2)) & " "
    FillField WordDoc, "MortgagorNames", IIf(IsNull(d!OriginalMortgagors), MortgagorNamesCaps(0, 2, 2), d!OriginalMortgagors)
    FillField WordDoc, "TrusteeNames", trusteeNames(0, 2)
    FillField WordDoc, "Investor", d!Investor
End If

FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))

FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "DotRecorded", Format$(d!DOTrecorded, "mmmm d, yyyy")

FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "LongState", Nz(d!LongState)
FillField WordDoc, "LiberFolio", LiberFolio(d!Liber, d!Folio, d![FCdetails.State], d![JurisdictionID])
FillField WordDoc, "Liber", d!Liber
FillField WordDoc, "Folio", Nz(d!Folio)
FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "TaxID", d!TaxID


FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")

FillField WordDoc, "MDInstrPrepared", IIf(d![FCdetails.State] = "VA", "", "This instrument was prepared under the supervision of " & d!AttorneyName & ", an attorney admitted to practice before the Court of Appeals of Maryland.")
FillField WordDoc, "MDSignLine", IIf(d![FCdetails.State] = "VA", "", "_________________________________")
FillField WordDoc, "MDAttorney", IIf(d![FCdetails.State] = "VA", "", d!AttorneyName)
    
    


FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "NotaryName", FetchNotaryName(Forms!Foreclosureprint!NotaryID, False)
FillField WordDoc, "FirmAddress", IIf(d![FCdetails.State] = "VA", "Commonwealth Trustees, LLC" & vbCr & "c/o Rosenberg & Associates, LLC" & vbCr & "8601 Westwood Center Drive, Suite 255" & vbCr & "Vienna, VA 22182", FirmAddress(vbCr))
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 226, "")
FillField WordDoc, "boa1", IIf(d![ClientID] = 446, "", "and, being duly sworn")
'FillField WordDoc, "boa2", IIf(d![ClientID] = 446, "National Association.", IIf(d![ClientID] = 385 Or d![ClientID] = 567, "company.", IIf([d!clientID] = 87, "association.", "corporation.")))
FillField WordDoc, "BOA2", IIf(d![ClientID] = 446, " National Association.", IIf(d![ClientID] = 385 Or d![ClientID] = 567, " company.", IIf(d![ClientID] = 87, " association.", " corporation.")))
'FillField WordDoc, "MERS", IIf(d![FCdetails.State] = "MD" And d![MERS] = -1, "Mortgage Electronic Registration Systems Inc. (MERS) solely as nominee for ", "")

If d!ClientID = 466 Then
    FillField WordDoc, "MERS", IIf(d![MERS] = -1, "Mortgage Electronic Registration Systems Inc. (MERS) solely as nominee for " & d![OriginalBeneficiary], d![OriginalBeneficiary])
    Else
    FillField WordDoc, "MERS", IIf(d![MERS] = -1, "Mortgage Electronic Registration Systems Inc. (MERS) solely as nominee for ", "")
End If
FillField WordDoc, "FinalLanguage", "WHEREAS, " & _
IIf(Forms!foreclosuredetails!LoanType = 5, "Federal Home Loan Mortgage Corporation is the owner of the note secured by said Deed of Trust and appointed the party of the first part with authority to hold, collect and enforce the note; and ", _
IIf(Forms!foreclosuredetails!LoanType = 4, "Federal National Mortgage Association is the owner of the note secured by said Deed of Trust and appointed the party of the first part with authority to hold, collect and enforce the note; and ", _
 "the party of the first part is the holder of the Note secured by said Deed of Trust; and,"))





WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.ActiveWindow.View = wdPrintView
WordDoc.SaveAs EMailPath & "Deed of Appointment " & d![CaseList.FileNumber] & ".doc"

Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Deed of Appointment " & d![CaseList.FileNumber] & ".doc ")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_LostNoteAffidavit(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")


If (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*") Then
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Lost Note Affidavit Dove.dot", False, 0, True)
Else
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Lost Note Affidavit.dot", False, 0, True)

End If
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
'FillField WordDoc, "OriginalMortgagors", IIf(IsNull([OriginalMortgagors]), MortgagorNames(0, 2), D![OriginalMortgagors] & ", assumed by " & MortgagorNames(0, 2))
                                        'IIf(IsNull(D!OriginalMortgagors), "", "Original Mortgagors " & D!OriginalMortgagors)
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "DOTStr", IIf(d!DOT, "Deed of Trust", "Mortgage")
FillField WordDoc, "MortgagorsStr", IIf(IsNull(d!OriginalMortgagors), MortgagorNames(0, 2), d!OriginalMortgagors)
FillField WordDoc, "MortgagorNames", IIf(IsNull(d![OriginalMortgagors]), MortgagorNames(0, 2), d![OriginalMortgagors] & ", assumed by " & MortgagorNames(0, 2))
FillField WordDoc, "Investor", d!Investor


If d!ClientID = 532 Then 'SElene
    FillField WordDoc, "InvestorAIF", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", d![Investor])

ElseIf d!ClientID = 523 Then 'Greentree
    FillField WordDoc, "InvestorAIF", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", d![Investor])
ElseIf d!ClientID = 73 Then 'And d![FCdetails.State] = "VA" Then
    FillField WordDoc, "InvestorAIF", IIf(d!AIF = True And d![FCdetails.State] = "VA", d!Investor & ", By Carrington Mortgage Services, LLC as Attorney in Fact and Servicing Agent", d!Investor)
Else
    FillField WordDoc, "InvestorAIF", IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for ", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")))) & d![Investor])

End If
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "LongState", d!LongState
FillField WordDoc, "LiberFolio", LiberFolio(d!Liber, d!Folio, d![FCdetails.State], d![JurisdictionID])
FillField WordDoc, "OriginalPBalWords", CurrencyWords(d!OriginalPBal)
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "InterestRate", Format(d!InterestRate, "#0.000")
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID) ' & " ss.IIf(IsNull(D!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", D!NotaryLocation)
FillField WordDoc, "FirmAddress", FirmAddress(vbCr)
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 154, "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Lost Note Affidavit.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Lost Note Affidavit.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub
Public Sub Doc_AssignmentSOTVASpecialized(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCdocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "AssignmentSOT VA Specialized.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "TaxID", IIf(IsNull(d!TaxID), "", d!TaxID)
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "TrusteeNames", trusteeNames(0, 2)
FillField WordDoc, "InvestorAddress", RemoveLF(d!InvestorAddress)
FillField WordDoc, "PropertyAddress", d![PropertyAddress] & IIf(IsNull(d![Fair Debt]), " ", ", " & d![Fair Debt] & ", ") & d![City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "MortgagorNames", IIf(IsNull(d![OriginalMortgagors]), MortgagorNamesCaps(0, 2, 2), d![OriginalMortgagors])
FillField WordDoc, "OriginalBeneficiary", d!OriginalBeneficiary
FillField WordDoc, "OriginalTrustee", IIf(IsNull(d![OriginalTrustee]), "", d![OriginalTrustee])

FillField WordDoc, "LiberFolio", LiberFolio(d![Liber], d![Folio], d![FCdetails.State], d![JurisdictionID])
FillField WordDoc, "rerecorded", IIf(IsNull(d![Rerecorded]), "", IIf((Not IsNull(d!Liber2) And IsNull(d!Folio2)), ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d"", ""yyyy") & " at Instrument Number " & d![Liber2], ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d"", ""yyyy") & " at Liber " & d![Liber2] & ", Folio " & d![Folio2]))

FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "LongState", d!LongState
FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d![Investor] & " by Specialized Loan Servicing LLC, its attorney-in-fact", d![Investor])
FillField WordDoc, "DOTdate", Format$(d![DOTdate], "mmmm d"", ""yyyy")
FillField WordDoc, "OriginalMortgagors", IIf(IsNull([OriginalMortgagors]), "", "assumed by " & MortgagorNamesCaps(0, 2, 2))

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "SOT VA Specialized.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "SOT VA Specialized.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_Assignment(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Assignment.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])


If Not IsNull(Forms![Case List]!MINnumber) Then
    FillField WordDoc, "MIN", Forms![Case List]!MINnumber
    Else
    FillField WordDoc, "MIN", ""
End If


If (d![FCdetails.State] = "VA") Then
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", IIf(IsNull(d![Folio2]), ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & _
    " at Instrument Number " & d![Liber2] & ", ", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Book " & d![Liber2] & ", Page " & d![Folio2] & ","))
'ElseIf d!Jurisdiction = 6 Then  'Calvert County
'    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Page " & d![Liber2] & ", Book " & d![Folio2])
Else 'MD or DC
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Book " & d![Liber2] & ", Page " & d![Folio2])
End If

FillField WordDoc, "Grantor", Forms![Print Assignment]!Grantor
FillField WordDoc, "Grantee", Forms![Print Assignment]!Grantee
FillField WordDoc, "GrantorAddress", Forms![Print Assignment]!GrantorAddress
FillField WordDoc, "GranteeAddress", Forms![Print Assignment]!GranteeAddress
FillField WordDoc, "Investor", d![OriginalTrustee]
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "MoneyText", IIf(d![ClientID] = 97, ".", ", all sums of money due and to become due thereon.")
FillField WordDoc, "TaxID", d!TaxID
FillField WordDoc, "LiberFolio", LiberFolio(d!Liber, d!Folio, d![FCdetails.State], d!JurisdictionID)
FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "DOTRecorded", d!DOTrecorded

FillField WordDoc, "MortgagorNames", IIf(IsNull(d!OriginalMortgagors), MortgagorNamesCaps(0, 2, 2), d!OriginalMortgagors)
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "LongState", d!LongState
FillField WordDoc, "OriginalPBalWords", CurrencyWords(d!OriginalPBal)
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "tax", IIf(d![FCdetails.State] = "VA", "Tax Map Number:", "Tax ID Number:")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Assignment.doc"

If (d![FCdetails.State] = "VA") Then
    Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Assignment VA.doc")
Else
    Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Assignment.doc")
End If

If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_AssignmentRecordingCoverLetter(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT CaseList.FileNumber, FCdetails.PropertyAddress, FCdetails.City, FCdetails.State, FCdetails.ZipCode, FCdetails.TaxID, JurisdictionList.State AS JurState, JurisdictionList.Jurisdiction, JurisdictionList.CourtAddress " & _
"FROM (CaseList LEFT JOIN FCdetails ON CaseList.FileNumber = FCdetails.FileNumber) INNER JOIN JurisdictionList ON CaseList.JurisdictionID = JurisdictionList.JurisdictionID WHERE [FCdetails.Current]=True and CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Assignment Recording Cover Letter.dot", False, 0, True)
'Set WordDoc = WordObj.Documents.Add(TemplatePath & "Military Affidavit Dove.dot", False, 0, True)
WordObj.Visible = True

FillField WordDoc, "date", Format(Now(), "mmmm d"", ""yyyy")
FillField WordDoc, "JurAddress", IIf(d![JurState] <> "DC", "Circuit Court for " & d![Jurisdiction], d![Jurisdiction] & " Government ") & vbNewLine & d![CourtAddress]
FillField WordDoc, "FileNumber", Forms![Case List]!FileNumber
FillField WordDoc, "TaxID", IIf(IsNull(d!TaxID), "", d!TaxID)
FillField WordDoc, "Borrowers", GetNames(d![FileNumber], 10, "Noteholder=true")
FillField WordDoc, "chkAmt", Format(DLookup("Amount", "Fees", "FeeType='FC-Assign' and State='" & d![State] & "'"), "Currency")
FillField WordDoc, "LoginName", Forms!Main!txtLoginName
FillField WordDoc, "FirmName", FirmName("MD")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Assignment Recording CoverLetter.doc"
Call SaveDoc(WordDoc, d![FileNumber], "Assignment Recording CoverLetter.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


'Marco
Public Sub Doc_MilitaryAffidavitMD(Keepopen As Boolean, ActiveDuty As Boolean) ' Military Affidavit MD
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim fName As String


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If


Set WordObj = CreateObject("Word.Application")

If (ActiveDuty = True) Then
       
   fName = "Military Affidavit Active"
   
Else
    If (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "MD" And Forms!Foreclosureprint!txtDesignator <> 3) Then
        fName = "Military Affidavit Dove"
    ElseIf d!ClientID = 404 Then
        fName = "Military Affidavit Bogman"
    ElseIf (d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3) Then  'A Designated attoryney was selected.  #1272
        fName = "Military Affidavit"
    ElseIf d!ClientID = 328 Then
        fName = "Military Affidavit SPLS"
    Else
        fName = "Military Affidavit"
    End If
End If

Set WordDoc = WordObj.Documents.Add(TemplatePath & fName & " MD.dot", False, 0, True) ' oh why you little....
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


FillField WordDoc, "designatorDate", IIf([Forms]![Foreclosureprint]![txtDesignator] = 3, "Date: " & Format(Date, "mmmm d, yyyy"), "")

FillField WordDoc, "MortgagorBorrowerNames", BorrowerMorgagorNamesOneSSN(d![CaseList.FileNumber], CopyNoR())
FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber

'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(D![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), Forms![Case List]!Investor)
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")) & IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, " " & Forms![Case List]!Investor)
FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for " & d!Investor, IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for " & Forms![Case List]!Investor, Forms![Case List]!Investor)))))))
'Polo

FillField WordDoc, "ActiveDuty", ActiveDutyNames(0, 20)
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", IIf(Split(ReportArgs, "|")(2) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), (Split(ReportArgs, "|")(1)))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 152, "")

FillField WordDoc, "only328client", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, d![Investor] & " has been proven to be the real party in interest and has the right to foreclose the subject property.", ""))



WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & fName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], fName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_MilitaryAffidavitActive(Keepopen As Boolean, strState As String) ' Military Affidavit Active VA
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim fName As String


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")

If (strState = "MD") Then
    If d!ClientID = 404 Then
        fName = "Military Affidavit Active Bogman"
    ElseIf (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "MD" And Forms!Foreclosureprint!txtDesignator <> 3) Then
        fName = "Military Affidavit Active Dove MD"
    Else
        fName = "Military Affidavit Active MD"
    End If
ElseIf (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "VA" And Forms!Foreclosureprint!txtDesignator <> 3) Then
  fName = "Military Affidavit Active Dove"
Else
  fName = "Military Affidavit Active"

End If

Set WordDoc = WordObj.Documents.Add(TemplatePath & fName & ".dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


FillField WordDoc, "designatorDate", IIf([Forms]![Foreclosureprint]![txtDesignator] = 3, "Date: " & Format(Date, "mmmm d, yyyy"), "")
FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, "Substitute Trustee", D!Investor)
FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "") & IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, " " & Forms![Case List]!Investor)))))

FillField WordDoc, "ActiveDuty", ActiveDutyNames(0, 20)
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 152, "")
FillField WordDoc, "only328client", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, d![Investor] & " has been proven to be the real party in interest and has the right to foreclose the subject property.", ""))
FillField WordDoc, "only328client2", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, "Subscribed and affirmed before me in the county of Douglas, State of Colorado, this _____ day of ________, 20____ .", ""))
FillField WordDoc, "only328client3", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, "____________________________", ""))
FillField WordDoc, "only328client4", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, "(Notary's official Signature)", ""))
FillField WordDoc, "only328client5", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, "____________________________", ""))
FillField WordDoc, "only328client6", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, "(Commission Expiration)", ""))


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & fName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Military Affidavit Active" & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_MilitaryAffidavit(Keepopen As Boolean) 'VA Military Affidavits
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
If (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "VA" And Forms!Foreclosureprint!txtDesignator <> 3) Then
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Military Affidavit Dove.dot", False, 0, True)
Else
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Military Affidavit.dot", False, 0, True)
End If
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


FillField WordDoc, "designatorDate", IIf([Forms]![Foreclosureprint]![txtDesignator] = 3, "Date: " & Format(Date, "mmmm d, yyyy"), "")
FillField WordDoc, "MortgagorBorrowerNames", BorrowerMorgagorNamesOneSSN(d![CaseList.FileNumber], CopyNoR())
FillField WordDoc, "OriginalMortgagors", IIf(IsNull(d!OriginalMortgagors), "", "Original Mortgagors " & d!OriginalMortgagors)
FillField WordDoc, "LoanNumber", d!LoanNumber
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(Forms![ForeclosureDetails]!State = "DC", "Substitute Trustee", "Commonwealth Trustees, LLC"), Forms![Case List]!Investor)
'AIF Changes
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(Forms![ForeclosureDetails]!State = "DC", "Substitute Trustee", "Commonwealth Trustees, LLC"), IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")) & IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, Forms![Case List]!Investor)
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(Forms![ForeclosureDetails]!state = "DC", "Substitute Trustee", "Commonwealth Trustees, LLC"), IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")) & IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, Forms![Case List]!Investor))
FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(Forms![foreclosuredetails]!State = "DC", "Substitute Trustee", "Commonwealth Trustee, LLC"), IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for " & d!Investor, IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for " & Forms![Case List]!Investor, Forms![Case List]!Investor)))))))


FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", IIf(IsNull(d!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", d!NotaryLocation)
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "only328client", IIf(d![ClientID] = 328, d![Investor] & " has been proven to be the real party in interest and has the right to foreclose the subject property.", "")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 152, "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Military Affidavit.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Military Affidavit.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_MilitaryAffidavitDC(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Military Affidavit DC.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


FillField WordDoc, "designatorDate", IIf([Forms]![Foreclosureprint]![txtDesignator] = 3, "Date: " & Format(Date, "mmmm d, yyyy"), "")
FillField WordDoc, "MortgagorBorrowerNames", BorrowerMorgagorNamesOneSSN(d![CaseList.FileNumber], CopyNoR())
FillField WordDoc, "OriginalMortgagors", IIf(IsNull(d!OriginalMortgagors), "", "Original Mortgagors " & d!OriginalMortgagors)
FillField WordDoc, "LoanNumber", d!LoanNumber
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(D![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), Forms![Case List]!Investor)
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")) & Forms![Case List]!Investor
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")) & Forms![Case List]!Investor)
FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for " & Forms![Case List]!Investor, Forms![Case List]!Investor))))



FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", IIf(Split(ReportArgs, "|")(2) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), (Split(ReportArgs, "|")(1)))
FillField WordDoc, "NotaryLocation", IIf(IsNull(d!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", d!NotaryLocation)
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "only328client", IIf(d![ClientID] = 328, d![Investor] & " has been proven to be the real party in interest and has the right to foreclose the subject property.", "")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 152, "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Military Affidavit DC.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Military Affidavit.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_NoteOwnershipAffidavit(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

If d!ClientID = 404 Then
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit Bogman.dot", False, 0, True)
    WordObj.Visible = True
ElseIf (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "MD") Then
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit Dove LPP.dot", False, 0, True)
    WordObj.Visible = True
ElseIf d!ClientID = 451 Then
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit Dove.dot", False, 0, True)
    WordObj.Visible = True
    
    

ElseIf d!ClientID = 328 Then 'SPLS
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit SPLS.dot", False, 0, True)
    WordObj.Visible = True
Else
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit.dot", False, 0, True)
    WordObj.Visible = True
End If

If d!ClientID = 451 Then
    WordDoc.Bookmarks("LastName").Select
    WordDoc.Bookmarks("LastName").Range.Text = d![PrimaryLastName]
End If

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])



FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "Investor2", IIf(d![ClientID] = 385, d![Investor2], d![Investor])
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "NoteOwner", FetchNoteOwner(d!LoanType, IIf(d![ClientID] = 385, d![Investor2], d![Investor]), d![FCdetails.State])
FillField WordDoc, "Noteholders", GetNamesMD(0, 2, "Noteholder=True")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
'FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & d!Investor

If d!ClientID = 328 Then 'SPLS
    FillField WordDoc, "InvestorAIF", "Specialized Loan Servicing LLC, as servicer for secured party"
ElseIf d!ClientID = 605 Then
    FillField WordDoc, "InvestorAIF", IIf(d![AIF] = True, d![LongClientName] & " as Servicer for ", "") & d![Investor]
ElseIf d!ClientID = 73 Then
    FillField WordDoc, "InvestorAIF", IIf(d![AIF] = True And d![FCdetails.State] = "MD", d![Investor] & ", By Carrington Mortgage Services, LLC as Attorney in Fact and Servicing Agent", d![Investor])
Else
    FillField WordDoc, "InvestorAIF", IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "") & d![Investor])))
End If

FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "FHLMCWording", IIf(d!LoanType = 5 Or d!LoanType = 4, ", and " & IIf(d![ClientID] = 385, d![Investor2], d![Investor]) & " is the holder of the Note having been transferred to " & IIf(d![ClientID] = 385, d![Investor2], d![Investor]) & " for the purposes of enforcement and conducting this foreclosure action", "")
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 464, "")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ____________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ____________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: __________________", "")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Note Ownership Affidavit.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Note Ownership Affidavit.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_MD7_105Affidavit(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

If Forms![Case List]!ClientID = 531 Then
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "MD 7-105 Affidavit MDHC.dot", False, 0, True)
ElseIf Forms![Case List]!ClientID = 404 Then
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "MD 7-105 Affidavit Bogman.dot", False, 0, True)
ElseIf (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "MD" And Forms!Foreclosureprint!txtDesignator <> 3) Then
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "MD 7-105 Affidavit Dove.dot", False, 0, True)
Else
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "MD 7-105 Affidavit.dot", False, 0, True)
End If
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProjectName").Select
WordDoc.Bookmarks("ProjectName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])



FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber


If d!ClientID = 532 Then 'Selene
    FillField WordDoc, "Investor", IIf([Forms]![Foreclosureprint]![txtDesignator] = 3, "", IIf((d![AIF] = True And d![ClientID] = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", d!Investor))
ElseIf d!ClientID = 523 Then 'GreenTree
    FillField WordDoc, "Investor", IIf([Forms]![Foreclosureprint]![txtDesignator] = 3, "", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", d![Investor]))
ElseIf d!ClientID = 73 Then
    FillField WordDoc, "Investor", IIf([Forms]![Foreclosureprint]![txtDesignator] = 3, "", IIf((d![AIF] = True And d![FCdetails.State] = "MD"), d![Investor] & ", By Carrington Mortgage Services, LLC as Attorney in Fact and Servicing Agent", d![Investor]))
Else
    FillField WordDoc, "Investor", IIf([Forms]![Foreclosureprint]![txtDesignator] = 3, "", IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf((d![AIF] = True And d![ClientID] = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for ", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")))) & d![Investor]))
End If

FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "NOI", Format$(d!NOI, "mmmm d, yyyy")
FillField WordDoc, "DateofDefault", Format$(d!DateOfDefault, "mmmm d, yyyy")
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "DefaultType", IIf(d!TypeOfDefault = 2, "The mortgage loan is in default because " & d!OtherDefault, "Said defendant(s) did not make the monthly mortgage payments and are in default under the terms of the " & DOTWord(d!DOT))
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "MD 7-105 Affidavit.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "MD 7-105 Affidavit.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_LossMitigationPre(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim textfinal As String
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
Dim templateName As String
Dim BankName As String

If (d![ClientID] = 6 Or d![ClientID] = 556) Then
  templateName = "Loss Mitigation - Preliminary Wells"
ElseIf d!ClientID = 446 Then
  templateName = "Loss Mitigation - Preliminary BOA"
ElseIf d!ClientID = 451 Then
    templateName = "Loss Mitigation - Preliminary Dove"
ElseIf d!ClientID = 97 Then
 templateName = "Loss Mitigation - Preliminary JP"
Else
templateName = "Loss Mitigation - Preliminary"
End If

If d![ClientID] = 6 Then
    BankName = "Loss Mitigation Preliminary Wells Fargo Bank"
ElseIf d![ClientID] = 556 Then
    BankName = "Loss Mitigation Preliminary Wells Fargo Home Mortgage"
ElseIf d!ClientID = 446 Then
    BankName = "Loss Mitigation Preliminary Bank of America"
ElseIf d!ClientID = 451 Then
    BankName = "Loss Mitigation Preliminary Dove"
ElseIf d!ClientID = 97 Then  'jp
    BankName = "Loss Mitigation - Preliminary JP"
    
Else
    BankName = templateName
End If


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

If d!ClientID = 451 Then
    WordDoc.Bookmarks("LastName").Select
    WordDoc.Bookmarks("LastName").Range.Text = d![PrimaryLastName]
End If

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "fillingdate", IIf(Not IsNull(d![Docket]), Format(d![Docket], "mm/d/yyyy"), "_____________________ ")
'FillField WordDoc, "borrower", MortgagorNamesOneline(d![CaseList.FileNumber], 2)
FillField WordDoc, "Borrower", MortgagorNames(0, 20)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", IIf(Not IsNull(d![CourtCaseNumber]), d![CourtCaseNumber], " _____________________ ")
FillField WordDoc, "NameAffianit", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "Print Name and Title of Affiant", "Name")
FillField WordDoc, "TitleAffiaitOrInvestor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "Title of Affiant")
FillField WordDoc, "Line", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "______________________________")
FillField WordDoc, "Investor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, d!Investor, "")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1484, "")
FillField WordDoc, "Resurgent", IIf(d![ClientID] = 543 Or d![ClientID] = 605, "By: New Penn Financial, LLC d/b/a Shellpoint Mortgage Servicing", "")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1484, "")


'FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))

If d!ClientID = 446 Then
    FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ____________________________", "")
    FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               " & "____________________________", "")
    FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: _________________________", "")
Else
    FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: 4200 Amon Carter Blvd", "")
    FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "                " & "Fort Worth, Texas 76155", "")
    FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone Number: 1-866-467-8090", "")
End If

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Loss Mitigation Preliminary" & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Loss Mitigation Preliminary" & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_LossMitigationFinal(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim templateName As String
Dim BankName As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

If (d![ClientID] = 6 Or d![ClientID] = 556) Then
  templateName = "Loss Mitigation - Final Wells"
ElseIf d!ClientID = 446 Then
    templateName = "Loss Mitigation - Final BOA"
ElseIf d!ClientID = 328 Then
    templateName = "Loss Mitigation - Final SPLS"
ElseIf d!ClientID = 451 Then
    templateName = "Loss Mitigation - Final Dove"
Else
templateName = "Loss Mitigation - Final"
End If

If d![ClientID] = 6 Then
    BankName = "Loss Mitigation Final Wells Fargo Bank"
ElseIf d![ClientID] = 556 Then
    BankName = "Loss Mitigation Final Wells Fargo Home Mortgage"
ElseIf d!ClientID = 446 Then
    BankName = "Loss Mitigation Final Bank of America"
ElseIf d!ClientID = 328 Then
    BankName = "Loss Mitigation Final SPLS"
ElseIf d!ClientID = 451 Then
    BankName = "Loss Mitigation Final Dove"
Else
BankName = "Loss Mitigation - Final"
End If



Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

If d!ClientID = 451 Then
    WordDoc.Bookmarks("LastName").Select
    WordDoc.Bookmarks("LastName").Range.Text = d!PrimaryLastName
End If

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", IIf(Not IsNull(d![CourtCaseNumber]), d![CourtCaseNumber], " _____________________ ")
FillField WordDoc, "NameAffianit", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "Print Name and Title of Affiant", "Name")
FillField WordDoc, "TitleAffiaitOrInvestor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "Title of Affiant")
FillField WordDoc, "Line", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "______________________________")
FillField WordDoc, "Investor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, d![Investor], "")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1345, "")
FillField WordDoc, "Resurgent", IIf(d![ClientID] = 543 Or d![ClientID] = 605, "By: New Penn Financial, LLC d/b/a Shellpoint Mortgage Servicing", "")

FillField WordDoc, "fillingdate", IIf(Not IsNull(d![Docket]), Format(d![Docket], "mm/d/yyyy"), "_____________________ ")
'FillField WordDoc, "borrower", MortgagorNamesOneline(d![CaseList.FileNumber], 2)
FillField WordDoc, "Borrower", MortgagorNames(0, 20)

'If Not d!ClientID = 446 Then FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))

If d!ClientID = 446 Then
    FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ____________________________", "")
    FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               " & "____________________________", "")
    FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: _________________________", "")
Else
    FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: 4200 Amon Carter Blvd", "")
    FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "                " & "Fort Worth, Texas 76155", "")
    FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone Number: 1-866-467-8090", "")
  '  FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
End If



WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & BankName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], BankName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_StatementOfDebt(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Statement of Debt"

If (d![ClientID] = 157) Then
  templateName = templateName & " Cenlar"
ElseIf d!ClientID = 404 Then
  templateName = templateName & " Bogman"
ElseIf (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "MD") Then
  templateName = templateName & " Dove"
ElseIf d![ClientID] = 451 Then
    templateName = templateName & " Dove"
End If
templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)



WordObj.Visible = True
WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


WordObj.Visible = False
If MsgBox("Is There A Loan Mod? ", vbYesNo) = vbYes Then
FillField WordDoc, "Mod", ", MODIFIED by Agreement effective " & Format(InputBox(" Effective Date?   Format mm/dd/yyyy"), "mmmm d, yyyy") & " with an amended principal balance of " & Format(InputBox(" Amended Principal Balance? "), "Currency")
Else
FillField WordDoc, "Mod", ""
End If

If d!JurisdictionID = 6 Then  'Calvert County
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Book " & d![Liber2] & ", Page " & d![Folio2])
Else 'MD or DC
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Liber " & d![Liber2] & ", Folio " & d![Folio2])
End If

WordObj.Visible = True
FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber


If d!ClientID = 532 Then 'SELENE Finance
    FillField WordDoc, "Investor", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", d![Investor])
ElseIf d!ClientID = 523 Then 'GreenTree
    FillField WordDoc, "Investor", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", d![Investor])
ElseIf d!ClientID = 73 Then
    FillField WordDoc, "Investor", IIf(d![AIF] = True And d![FCdetails.State] = "MD", d![Investor] & ", By Carrington Mortgage Services, LLC as Attorney in Fact and Servicing Agent", d![Investor])
Else
    FillField WordDoc, "Investor", IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for ", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")))) & d![Investor])
End If

FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "BalanceDue", IIf((d![ClientID] = 451 And d![FCdetails.State] = "MD"), "Unpaid Principal Balance Due", "Remaining Balance Due")
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "Liber", LiberFolio(d![Liber], d![Folio], d![FCdetails.State], d![JurisdictionID])
FillField WordDoc, "LPIdate", Format$(d![LPIDate], "mmmm d, yyyy")
FillField WordDoc, "LPIdate+1", Format$(d![LPIDate] + 1, "mmmm d, yyyy")
FillField WordDoc, "LPdM", DateAdd("m", -1, d![LPIDate])
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "PaidStr", IIf((Nz(d![RemainingPBal], 0) > Nz(d![OriginalPBal], 0)), "Additional Interest", "Paid on principal")
FillField WordDoc, "Paid", Format$(d!OriginalPBal - d!RemainingPBal, "Currency")
FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "InterestRate", IIf(IsNull(d!InterestRate), "____________ %", Format$(d!InterestRate, "#.000%"))
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Statement of Debt.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Statement of Debt.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close



End Sub

Public Sub Doc_StatementOfDebtFigures(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim i As Integer, itemCount As Integer, totalItems As Currency, itemsFields As String



Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Statement of Debt Figures"

If (d![ClientID] = 157) Then
  templateName = templateName & " Cenlar"
ElseIf (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "MD") Then
    templateName = templateName & " Dove-LPP"
ElseIf (d![ClientID] = 451) Then
    templateName = templateName & " Dove"
ElseIf d!ClientID = 404 Then
    templateName = templateName & " Bogman"
End If
templateName = templateName & ".dot"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

If d!ClientID = 451 Then
WordDoc.Bookmarks("LastName").Select
WordDoc.Bookmarks("LastName").Range.Text = d!PrimaryLastName
End If
WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])



WordObj.Visible = False
If MsgBox("Is There A Loan Mod? ", vbYesNo) = vbYes Then
FillField WordDoc, "Mod", " MODIFIED by Agreement effective " & Format(InputBox(" Effective Date?   Format mm/dd/yyyy"), "mmmm d, yyyy") & " with an amended principal balance of " & Format(InputBox(" Amended Principal Balance? "), "Currency")
Else
FillField WordDoc, "Mod", ""
End If
WordObj.Visible = True

If d!JurisdictionID = 6 Then  'Calvert County
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Book " & d![Liber2] & ", Page " & d![Folio2])
Else 'MD or DC
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Liber " & d![Liber2] & ", Folio " & d![Folio2])
End If

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber  ' Test


If d!ClientID = 532 Then  'Selene
    FillField WordDoc, "Investor", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", d![Investor])
ElseIf d!ClientID = 523 Then
    FillField WordDoc, "Investor", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", d![Investor])
ElseIf d!ClientID = 73 Then
    FillField WordDoc, "Investor", IIf(d![AIF] = True And d![FCdetails.State] = "MD", d![Investor] & ", By Carrington Mortgage Services, LLC as Attorney in Fact and Servicing Agent", d!Investor)
Else
    FillField WordDoc, "Investor", IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for ", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")))) & d![Investor])
End If
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "455Only", IIf(d![ClientID] = 455, ", and continuing each month thereafter with probable future advancements made by the mortgagee, and that the Plaintiff(s) has\have the right to foreclose;", ", and continuing each month thereafter, and that the Plaintiff(s) has\have the right to foreclose;")
FillField WordDoc, "Liber", LiberFolio(d![Liber], d![Folio], d![FCdetails.State], d![JurisdictionID])
FillField WordDoc, "LPIdate", Format$(d![LPIDate], "mmmm d, yyyy")
FillField WordDoc, "LPIdate+1", Format$(d![LPIDate] + 1, "mmmm d, yyyy")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "PaidStr", IIf((Nz(d![RemainingPBal], 0) > Nz(d![OriginalPBal], 0)), "Additional Interest", "Paid on principal")
FillField WordDoc, "Paid", Format$(d!OriginalPBal - d!RemainingPBal, "Currency")
FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")
FillField WordDoc, "diem", IIf(d![LoanType] = 3, IIf(d![ClientID] <> 531, "Per Monthly Interest: ", "Per Diem Interest: "), "Per Diem Interest: ")
FillField WordDoc, "txtbalanc", IIf(d![ClientID] = 361, "Unpaid Principal Balance", "Remaining Balance Due")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

totalItems = 0
itemsFields = ""
Set dd = CurrentDb.OpenRecordset("SELECT Desc, Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber] & " ORDER BY Timestamp;", dbOpenSnapshot)
If dd.EOF Then      ' no extra lines
    FillField WordDoc, "Line_Items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items", itemsFields
    dd.MoveFirst
    i = 1
    Do While Not dd.EOF
        FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        totalItems = totalItems + Nz(dd!Amount, 0)
        dd.MoveNext
        i = i + 1
    Loop
End If
dd.Close

FillField WordDoc, "BalDueDate", IIf(IsNull([Forms]![Print Statement of Debt]![txtDueDate]), "______________", Format$([Forms]![Print Statement of Debt]![txtDueDate], "mmmm d"", ""yyyy"))
FillField WordDoc, "BalanceDue", Format$(d!RemainingPBal + totalItems, "Currency")
FillField WordDoc, "PerDiemInterest", IIf(IsNull(d!PerDiem), "$_____________", Format$(d!PerDiem, "Currency"))
'FillField WordDoc, "InterestRate", IIf(IsNull(d!InterestRate), "____________ %", Format$(d!InterestRate, "#.000%"))
FillField WordDoc, "InterestRate", IIf(IsNull(d!InterestRate), "____________ %", Format$(d!InterestRate, "#0.000") & "%")
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", IIf(IsNull(d!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", d!NotaryLocation)

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Statement of Debt.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Statement of Debt.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close


End Sub

Public Sub Doc_OrderGrantingRelief(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim FileNumber As Long, Judge As String, CoDebtor As Boolean, AffDate As String, AffInfo As String, DebtorsPlural As Boolean

FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBKDocsWord WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Order Granting Relief.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("AttorneyInfo").Select
WordDoc.Bookmarks("AttorneyInfo").Range.Text = "Diane Rosenberg" & vbCr & "VA Bar 35237"

Judge = Right$(UCase$(d!CaseNo), 3)
CoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13)
If Not IsNull(d![3rdAff]) Then
    AffDate = Format$(d![3rdAff], "mmmm d, yyyy")
Else
    If Not IsNull(d![2ndAff]) Then
        AffDate = Format$(d![2ndAff], "mmmm d, yyyy")
    Else
        If Not IsNull(d![Affidavit]) Then
            AffDate = Format$(d![Affidavit], "mmmm d, yyyy")
        End If
    End If
End If
If AffDate = "" Then
    AffInfo = ""
Else
    AffInfo = "an Affidavit of Default having been sent on " & AffDate & ", no response or funds having been received, "
End If
DebtorsPlural = (CountNames(FileNumber, "BKDebtor = True AND (Owner=True OR Mortgagor=True)") > 1)

FillField WordDoc, "Header", _
    IIf(d![Districts.State] <> "VA", vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr, "") & _
    "IN THE UNITED STATES BANKRUPTCY COURT" & vbCr & _
    "FOR THE " & UCase$(d!Name) & vbCr & _
    d!Location
FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & IIf(DebtorsPlural, "s", "")
FillField WordDoc, "InvestorAddr", UCase$(d!Investor) & vbCr & RemoveLF(d!InvestorAddress)
FillField WordDoc, "Respondents", _
    GetAddresses(0, 4, _
        IIf(d![BKdetails.Chapter] = 13, _
            "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", _
            "BKDebtor=True AND (Owner=True OR Mortgagor=True)"), vbCr) & _
    IIf(d![BKdetails.Chapter] = 7, _
        vbCr & vbCr & "and" & vbCr & vbCr & FormatName("", _
                                                        UCase$(Nz(d!First)), _
                                                        UCase$(Nz(d!Last)), _
                                                        ", CHAPTER 7 TRUSTEE", _
                                                        d!Address, _
                                                        d!Address2, _
                                                        d![BKTrustees.City], _
                                                        d![BKTrustees.State], _
                                                        d!Zip, _
                                                        vbCr), _
        "")
FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]
FillField WordDoc, "Caption", IIf(Judge = "RGM" And d!RealEstate, _
    "ORDER TERMINATING AUTOMATIC STAY AS TO " & UCase$(Nz(d!PropertyAddress)), _
    "ORDER GRANTING RELIEF FROM AUTOMATIC STAY")
If Judge = "RGM" Then
    FillField WordDoc, "RGM", IIf(IsNull(d!Modify), _
                                  "", _
                                  "a Consent Order having been entered on " & Format$(d!Modify, "mmmm d, yyyy") & ", ") & _
                              AffInfo
Else
    FillField WordDoc, "RGM", ""
End If
If d![Districts.State] = "VA" Then
    FillField WordDoc, "OrderDate", ", it is this " & OrderDate()
    FillField WordDoc, "JudgeSignature", "________________________________" & vbCr & "United States Bankruptcy Judge"
    FillField WordDoc, "Submitted", "Respectfully Submitted:"
    If Forms!BankruptcyPrint!chElectronicSignature Then
        FillField WordDoc, "ElectronicSignature", "/s/ " & Forms!BankruptcyPrint!cbxAttorney
    Else
        FillField WordDoc, "ElectronicSignature", "_______________________________"
    End If
    FillField WordDoc, "AttorneySignature", Forms!BankruptcyPrint!cbxAttorney
    FillField WordDoc, "End", ""
Else
    FillField WordDoc, "OrderDate", ""
    FillField WordDoc, "JudgeSignature", ""
    FillField WordDoc, "Submitted", ""
    FillField WordDoc, "ElectronicSignature", ""
    FillField WordDoc, "AttorneySignature", ""
    FillField WordDoc, "End", "End of Order"
End If
FillField WordDoc, "DistrictName", d!Name
FillField WordDoc, "CoDebtor", IIf(CoDebtor, " and 1301", "")
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "Action", IIf(d!RealEstate, _
                                  "the foreclosure sale against the real property and improvements known as " & d!PropertyAddress & IIf(Len(Forms!BankruptcyDetails!sfrmPropAddr!Apt & "") = 0, "", ", " & Forms!BankruptcyDetails!sfrmPropAddr!Apt) & ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode]), _
                                  "repossession and sale of the " & d!PropertyDesc)
If Not d!RealEstate Then
    FillField WordDoc, "Rule4001", "ORDERED that the ten (10) day stay of Rule 4001(a)(3) be, and it is hereby, waived and the terms of this Order are immediately enforceable; and be it further" & vbCr & vbCr & vbTab
Else
    FillField WordDoc, "Rule4001", ""
End If
FillField WordDoc, "DebtorPossessive", "Debtor" & IIf(DebtorsPlural, "'s", "s'")
FillField WordDoc, "CC1", Forms!BankruptcyPrint!cbxAttorney & vbCr & FirmAddress(vbCr) & vbCr & vbCr & BKService(0, vbCr)
FillField WordDoc, "CC2", FormatName("", _
                                    d!First, _
                                    d!Last & ", Trustee", _
                                    "", _
                                    d!Address, _
                                    d!Address2, _
                                    d![BKTrustees.City], _
                                    d![BKTrustees.State], _
                                    d!Zip, _
                                    vbCr) & _
                            vbCr & vbCr & _
                            d!FirstName & " " & d!LastName & IIf(IsNull(d!LastName), "", ", Esquire") & _
                            IIf(IsNull(d!BKAttorneyFirm), _
                                "", _
                                vbCr & d!BKAttorneyFirm & vbCr & RemoveLF(d!BKAttorneyAddress))

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Order Granting Relief.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Order Granting Relief.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_FinalOrderTerminating(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim FileNumber As Long, Judge As String, CoDebtor As Boolean, AffDate As String, AffInfo As String, DebtorsPlural As Boolean

FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBKdocsWordFinalOrderTerminating WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Final Order Terminating.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("AttorneyInfo").Select
WordDoc.Bookmarks("AttorneyInfo").Range.Text = "Diane Rosenberg" & vbCr & "VA Bar 35237"

Judge = Right$(UCase$(d!CaseNo), 3)
CoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13)
If Not IsNull(d![3rdAff]) Then
    AffDate = Format$(d![3rdAff], "mmmm d, yyyy")
Else
    If Not IsNull(d![2ndAff]) Then
        AffDate = Format$(d![2ndAff], "mmmm d, yyyy")
    Else
        If Not IsNull(d![Affidavit]) Then
            AffDate = Format$(d![Affidavit], "mmmm d, yyyy")
        End If
    End If
End If
If AffDate = "" Then
    AffInfo = ""
Else
    AffInfo = "an Affidavit of Default having been sent on " & AffDate & ", no response or funds having been received, "
End If
DebtorsPlural = (CountNames(FileNumber, "BKDebtor = True AND (Owner=True OR Mortgagor=True)") > 1)

FillField WordDoc, "Header", _
    IIf(d![Districts.State] <> "VA", vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr, "") & _
    "IN THE UNITED STATES BANKRUPTCY COURT" & vbCr & _
    "FOR THE " & UCase$(d!Name) & vbCr & _
    d!Location
FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & "(s)" 'IIf(DebtorsPlural, "s", "")
FillField WordDoc, "InvestorAddr", UCase$(d!Investor) ' & vbCr & RemoveLF(D!InvestorAddress)

FillField WordDoc, "Respondents", UCase$(GetNames(0, 3, IIf(d![BKdetails.Chapter] = 13, "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", "BKDebtor=True AND (Owner=True OR Mortgagor=True)"))) & IIf(d![BKdetails.Chapter] = 7, " AND " & UCase$(Nz(d![First])) & " " & UCase$(Nz(d![Last])) & ", CHAPTER 7 TRUSTEE", "") & ""

FillField WordDoc, "PropertyAddress", UCase$(d![PropertyAddress])
FillField WordDoc, "Modify", Format$(d![Modify], "mmmm d"", ""yyyy")
FillField WordDoc, "Districts", d![Name]
FillField WordDoc, "section", IIf(GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13, "Sections 362(d) and 1301", "Section 362(d)")
FillField WordDoc, "section2", IIf(d![RealEstate], "commence foreclosure proceeding " & IIf(d![FCdetails.State] = "MD", "in the Circuit Court for " & d![Jurisdiction] & IIf(IsNull(d![LongState]), "", ", " & d![LongState]) & ", ", "") & "against the real property and improvements " & IIf(IsNull(d![ShortLegal]), "", "with a legal description of """ & d![ShortLegal] & """ also ") & "known as " & d![PropertyAddress] & ", " & IIf(IsNull(d![Fair Debt]), "", d![Fair Debt] & ", ") & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode]) & " and be it further " & _
 vbCr & vbCr & "        ORDERED that the successful purchasers, be and they are hereby allowed to obtain possession of the Subject Property pursuant to state law;", "proceed with repossession and sale of the " & d![PropertyDesc])
FillField WordDoc, "sign1", IIf(Forms!BankruptcyPrint!chElectronicSignature, "/s/ " & Forms!BankruptcyPrint!cbxAttorney, "")
FillField WordDoc, "sign2", Forms!BankruptcyPrint!cbxAttorney & " " & vbCr & FirmAddress() & vbCr & " VA Bar No. 35237 " & vbCr & FirmPhone()
FillField WordDoc, "Part2", FormatName("", d![First], d![Last], ", Trustee", d![Address], d![Address2], d![BKTrustees.City], d![BKTrustees.State], d![Zip])
FillField WordDoc, "cbxAttorney", Forms!BankruptcyPrint!cbxAttorney
FillField WordDoc, "Part3", IIf(IsNull(d![AttorneyLastName]), "", d![AttorneyFirstName] & " " & d![AttorneyLastName] & ", Esquire " & IIf(IsNull(d![AttorneyFirm]), "", d![AttorneyFirm] & "") & d![AttorneyAddress])
FillField WordDoc, "FirmAddress", FirmAddress()
FillField WordDoc, "BKService", BKService(d![CaseList.FileNumber], vbCr)



'FillField WordDoc, "Respondents", _
'    GetAddresses(0, 4, _
'        IIf(D![BKdetails.Chapter] = 13, _
'            "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", _
'            "BKDebtor=True AND (Owner=True OR Mortgagor=True)"), vbCr) & _
'    IIf(D![BKdetails.Chapter] = 7, _
'        vbCr & vbCr & "and" & vbCr & vbCr & FormatName("", _
'                                                        UCase$(Nz(D!First)), _
'                                                        UCase$(Nz(D!Last)), _
'                                                        ", CHAPTER 7 TRUSTEE", _
'                                                        D!Address, _
'                                                        D!Address2, _
'                                                        D![BKTrustees.City], _
'                                                        D![BKTrustees.State], _
'                                                        D!Zip, _
'                                                        vbCr), _
'        "")
FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]
FillField WordDoc, "Caption", IIf(Judge = "RGM" And d!RealEstate, _
    "ORDER TERMINATING AUTOMATIC STAY AS TO " & UCase$(Nz(d!PropertyAddress)), _
    "ORDER GRANTING RELIEF FROM AUTOMATIC STAY")
If Judge = "RGM" Then
    FillField WordDoc, "RGM", IIf(IsNull(d!Modify), _
                                  "", _
                                  "a Consent Order having been entered on " & Format$(d!Modify, "mmmm d, yyyy") & ", ") & _
                              AffInfo
Else
    FillField WordDoc, "RGM", ""
End If
If d![Districts.State] = "VA" Then
    FillField WordDoc, "OrderDate", ", it is this " & OrderDate()
    FillField WordDoc, "JudgeSignature", "________________________________" & vbCr & "United States Bankruptcy Judge"
    FillField WordDoc, "Submitted", "Respectfully Submitted:"
    If Forms!BankruptcyPrint!chElectronicSignature Then
        FillField WordDoc, "ElectronicSignature", "/s/ " & Forms!BankruptcyPrint!cbxAttorney
    Else
        FillField WordDoc, "ElectronicSignature", "_______________________________"
    End If
    FillField WordDoc, "AttorneySignature", Forms!BankruptcyPrint!cbxAttorney
    FillField WordDoc, "End", ""
Else
    FillField WordDoc, "OrderDate", ""
    FillField WordDoc, "JudgeSignature", ""
    FillField WordDoc, "Submitted", ""
    FillField WordDoc, "ElectronicSignature", ""
    FillField WordDoc, "AttorneySignature", ""
    FillField WordDoc, "End", "End of Order"
End If
FillField WordDoc, "DistrictName", d!Name
FillField WordDoc, "CoDebtor", IIf(CoDebtor, " and 1301", "")
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "Action", IIf(d!RealEstate, _
                                  "the foreclosure sale against the real property and improvements known as " & d!PropertyAddress & IIf(Len(Forms!BankruptcyDetails!sfrmPropAddr!Apt & "") = 0, "", ", " & Forms!BankruptcyDetails!sfrmPropAddr!Apt) & ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode]), _
                                  "repossession and sale of the " & d!PropertyDesc)
If Not d!RealEstate Then
    FillField WordDoc, "Rule4001", "ORDERED that the ten (10) day stay of Rule 4001(a)(3) be, and it is hereby, waived and the terms of this Order are immediately enforceable; and be it further" & vbCr & vbCr & vbTab
Else
    FillField WordDoc, "Rule4001", ""
End If
FillField WordDoc, "DebtorPossessive", "Debtor" & "(s)" ' IIf(DebtorsPlural, "'s", "s'")
FillField WordDoc, "CC1", Forms!BankruptcyPrint!cbxAttorney & vbCr & FirmAddress(vbCr) & vbCr & vbCr & BKService(0, vbCr)
FillField WordDoc, "CC2", FormatName("", _
                                    d!First, _
                                    d!Last & ", Trustee", _
                                    "", _
                                    d!Address, _
                                    d!Address2, _
                                    d![BKTrustees.City], _
                                    d![BKTrustees.State], _
                                    d!Zip, _
                                    vbCr) & _
                            vbCr & vbCr & _
                            d!FirstName & " " & d!LastName & IIf(IsNull(d!LastName), "", ", Esquire") & _
                            IIf(IsNull(d!BKAttorneyFirm), _
                                "", _
                                vbCr & d!BKAttorneyFirm & vbCr & RemoveLF(d!BKAttorneyAddress))

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Final Order Terminating.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Final Order Terminating.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_ConsentModifying13(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim FileNumber As Long, Judge As String, CoDebtor As Boolean, AffDate As String, AffInfo As String, DebtorsPlural As Boolean

FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBKDocsWord WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "ConsentModifying13.dot", False, 0, True)
WordObj.Visible = True
WordObj.ScreenUpdating = False
WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("AttorneyInfo").Select
WordDoc.Bookmarks("AttorneyInfo").Range.Text = "Diane Rosenberg" & vbCr & "VA Bar 35237"

Judge = Right$(UCase$(d!CaseNo), 3)
CoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13)


'If Not IsNull(d![3rdAff]) Then
'    AffDate = Format$(d![3rdAff], "mmmm d, yyyy")
'Else
'    If Not IsNull(d![2ndAff]) Then
'        AffDate = Format$(d![2ndAff], "mmmm d, yyyy")
'    Else
'        If Not IsNull(d![Affidavit]) Then
'            AffDate = Format$(d![Affidavit], "mmmm d, yyyy")
'        End If
'    End If
'End If
'If AffDate = "" Then
'    AffInfo = ""
'Else
'    AffInfo = "an Affidavit of Default having been sent on " & AffDate & ", no response or funds having been received, "
'End If
DebtorsPlural = (CountNames(FileNumber, "BKDebtor = True AND (Owner=True OR Mortgagor=True)") > 1)

FillField WordDoc, "Header", _
    IIf(d![Districts.State] <> "VA", vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr, "") & _
    "IN THE UNITED STATES BANKRUPTCY COURT" & vbCr & _
    "FOR THE " & UCase$(d!Name) & vbCr & _
    d!Location
FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & IIf(DebtorsPlural, "s", "")
FillField WordDoc, "InvestorAddr", UCase$(d!Investor) & vbCr & RemoveLF(d!InvestorAddress)
FillField WordDoc, "Respondents", _
    GetAddresses(0, 4, _
        IIf(d![BKdetails.Chapter] = 13, _
            "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", _
            "BKDebtor=True AND (Owner=True OR Mortgagor=True)"), vbCr) & _
    IIf(d![BKdetails.Chapter] = 7, _
        vbCr & vbCr & "and" & vbCr & vbCr & FormatName("", _
                                                        UCase$(Nz(d!First)), _
                                                        UCase$(Nz(d!Last)), _
                                                        ", CHAPTER 7 TRUSTEE", _
                                                        d!Address, _
                                                        d!Address2, _
                                                        d![BKTrustees.City], _
                                                        d![BKTrustees.State], _
                                                        d!Zip, _
                                                        vbCr), _
        "")
FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]

FillField WordDoc, "Caption", IIf(d!Judge = "RGM" And d!RealEstate, "CONSENT ORDER  AS TO " & UCase$(d!PropertyAddress), "CONSENT ORDER MODIFYING AUTOMATIC STAY")

Select Case Forms![Print Consent Order Modifying]!optTrustee
    Case 1
        FillField WordDoc, "TrusteeAction", "the trustee having filed a report of no distribution, "
    Case 2
        FillField WordDoc, "TrusteeAction", "the trustee having failed to file an answer, "
End Select

FillField WordDoc, "ThisDate", IIf(d![Districts.State] <> "VA", "", ", it is this " & OrderDate())

FillField WordDoc, "District", d!Districts.Name

FillField WordDoc, "CoDebtorSection", IIf(CoDebtor, " and 1301", "")

FillField WordDoc, "Action", IIf(d!RealEstate, "commence foreclosure proceeding " & IIf(d![FCdetails.State] = "MD", "in the Circuit Court for " & d!Jurisdiction & IIf(IsNull(d!LongState), "", ", " & d!LongState) & ", ", "") & "against the real property and improvements " & IIf(IsNull(d!ShortLegal), "", "with a legal description of """ & d!ShortLegal & """ also ") & "known as " & d!PropertyAddress & ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d!ZipCode) & " and to allow the successful purchaser to obtain possession of same", "proceed with repossession and sale of the " & d!PropertyDesc)

FillField WordDoc, "OrderWording", IIf(d!Judge = "RGM", "ORDERED that the Debtor shall:", "ORDERED that the above Order be and it is hereby, stayed provided that the Debtor:")

FillField WordDoc, "ConsentPaymentAmount", Format$(d!ConsentPaymentAmount, "Currency")

FillField WordDoc, "PaymentType", IIf(d!RealEstate, "mortgage", "monthly")

FillField WordDoc, "ConsentPaymentDate", Format$(d!ConsentPaymentDate, "mmmm d, yyyy")

FillField WordDoc, "NoteType", IIf(d!RealEstate, "Promissory Note secured by the " & DOTWord(d!DOT) & " on the above referenced property", d!PropertyContract)

FillField WordDoc, "ConsentPaymentInfo", d!ConsentPaymentInfo

FillField WordDoc, "2A", IIf(IsNull(d!Consent2A), "", "^p2A. " & d!Consent2A)

FillField WordDoc, "AndAttorney", IIf(IsNull(d!AttorneyLastName), "", "and Debtor's attorney ")

FillField WordDoc, "RGM1", IIf(d!Judge = "RGM", "", ", without further order of court")

FillField WordDoc, "RGM2", IIf(d!Judge = "RGM", "^pIf any amount required in Paragraph 2 is not paid timely, Movant's attorney shall mail notice to the Debtor" & _
                        IIf(IsNull(d!AttorneyLastName), "", ", Debtor's attorney") & " and Chapter 13 Trustee, and shall file an Order of Termination of Automatic " & _
                        "Stay against the Subject Property described above; and be it further^p" & _
                        "ORDERED that a default in the payment of a regularly scheduled mortgage payment as listed in Paragraph 1 shall be governed by the attached addendum.", "")

If d![Districts.State] = "VA" Then
    FillField WordDoc, "JudgeSignature", "________________________________" & vbCr & "United States Bankruptcy Judge"
    FillField WordDoc, "Submitted", "Respectfully Submitted:"
    If Forms!BankruptcyPrint!chElectronicSignature Then
        FillField WordDoc, "ElectronicSignature", "/s/ Diane S. Rosenberg"
    Else
        FillField WordDoc, "ElectronicSignature", "_______________________________"
    End If
    FillField WordDoc, "AttorneySignature", "Diane S. Rosenberg"
    FillField WordDoc, "End", ""
Else
    FillField WordDoc, "OrderDate", ""
    FillField WordDoc, "JudgeSignature", ""
    FillField WordDoc, "Submitted", ""
    FillField WordDoc, "ElectronicSignature", ""
    FillField WordDoc, "AttorneySignature", ""
    FillField WordDoc, "End", "End of Order"
End If

WordObj.Selection.HomeKey wdStory, wdMove
WordObj.ScreenUpdating = True
WordDoc.SaveAs EMailPath & "Order Granting Relief.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Order Granting Relief.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub



Public Sub mik_test()
  Call Doc_ConsentModifyingReliefVA(True)
End Sub

Public Sub Doc_ConsentModifyingReliefVA(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim FileNumber As Long, Judge As String, CoDebtor As Boolean, CoDebtorName As String, AffDate As String, AffInfo As String, DebtorsPlural As Boolean, SumRepaymentAmount As Currency, ParagraphNum As Integer, CourtCity As String, FillYear As String

FileNumber = Forms![Case List]!FileNumber

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBKDocsWord WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Consent Modifying VA.dot", False, 0, True)
WordObj.Visible = True
WordObj.ScreenUpdating = False
WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("AttorneyInfo").Select
WordDoc.Bookmarks("AttorneyInfo").Range.Text = "Mark D. Meyer" & vbCr & "VA Bar 74290"

Judge = Right$(UCase$(d!CaseNo), 3)
CoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13)


DebtorsPlural = (CountNames(FileNumber, "BKDebtor = True AND (Owner=True OR Mortgagor=True)") > 1)

FillField WordDoc, "Header", _
    IIf(d![Districts.State] <> "VA", vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr, "") & _
    "IN THE UNITED STATES BANKRUPTCY COURT" & vbCr & _
    "FOR THE " & UCase$(d!Name) & vbCr & _
    d!Location
FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & IIf(DebtorsPlural, "s", "")
FillField WordDoc, "InvestorAddr", UCase$(d!Investor) & vbCr & RemoveLF(d!InvestorAddress)
FillField WordDoc, "Respondents", _
    GetAddresses(0, 4, _
        IIf(d![BKdetails.Chapter] = 13, _
            "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", _
            "BKDebtor=True AND (Owner=True OR Mortgagor=True)"), vbCr) & _
    IIf(d![BKdetails.Chapter] = 7, _
        vbCr & vbCr & "and" & vbCr & vbCr & FormatName("", _
                                                        UCase$(Nz(d!First)), _
                                                        UCase$(Nz(d!Last)), _
                                                        ", CHAPTER 7 TRUSTEE", _
                                                        d!Address, _
                                                        d!Address2, _
                                                        d![BKTrustees.City], _
                                                        d![BKTrustees.State], _
                                                        d!Zip, _
                                                        vbCr), _
        "")
FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]

FillField WordDoc, "Caption", "CONSENT ORDER MODIFYING AUTOMATIC STAY"
FillField WordDoc, "Hearing", Format(d![Hearing], "mmmm d, yyyy")
FillField WordDoc, "Investor", d![Investor]
FillField WordDoc, "PropertyAddress", d![PropertyAddress] & IIf(Len(Forms!BankruptcyDetails!sfrmPropAddr!Apt & "") = 0, "", ", " & Forms!BankruptcyDetails!sfrmPropAddr!Apt) & ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "ShortLegalName", Nz(d![ShortLegal])

FillField WordDoc, "ConsentPayment", Format$(Nz(d![ConsentPaymentAmount], 0), "Currency")
FillField WordDoc, "ConsentPaymentDate", Format$(d![ConsentPaymentDate], "mmmm d, yyyy")
FillField WordDoc, "LatePaymentAmount", Format$(Nz(d![ConsentLateChargesAmount], 0), "Currency")

FillField WordDoc, "RepaymentTermsDetails", RepaymentItemList(FileNumber)
FillField WordDoc, "MakePaymentsTo", d![ConsentOrderPaymentTo]
FillField WordDoc, "CurrentDate", Format$(Date, "mmmm d, yyyy")
SumRepaymentAmount = Nz(DSum("[RepaymentAmount]", "[BKRepaymentTerms]", "[FileNumber] = " & FileNumber), 0)
FillField WordDoc, "TotalRepaymentAmount", Format$(SumRepaymentAmount, "Currency")

ParagraphNum = 7

If (CoDebtor = True) Then  ' Co-debtors exists
  ParagraphNum = ParagraphNum + 1
  CoDebtorName = GetNames(0, 1, "BKCoDebtor=True")
  FillField WordDoc, "List8Text", vbCr & ParagraphNum & ".           Relief is granted as to " & CoDebtorName & ", the co-debtor, from the automatic stay " & _
                                                    "imposed by 1301(a) to the same extent and on the same terms and conditions as granted as to the debtor."
End If


If (Judge = "FJS") Or (Trim(d![Last] = "Stackhouse, Jr.")) Then
  ParagraphNum = ParagraphNum + 1
  If (ParagraphNum = 9) Then
    FillField WordDoc, "List9Text", vbCr & ParagraphNum & ".            The source of funds to make the cure payment is: " & d![SourceofFundsCurePayments] & "."
  Else
    FillField WordDoc, "List8Text", ParagraphNum & ".            The source of funds to make the cure payment is: " & d![SourceofFundsCurePayments] & "."
  End If

End If

' blank out those fields that are not needed
If (ParagraphNum = 7) Then
  FillField WordDoc, "List8Text", ""
  FillField WordDoc, "List9Text", ""
ElseIf (ParagraphNum = 8) Then
  FillField WordDoc, "List9Text", ""
End If

  

CourtCity = Mid(d![Districts.Display], InStr(1, d![Districts.Display], "VA") + 3)
FillField WordDoc, "CourtCity", CourtCity
FillYear = "___________________, " & DatePart("yyyy", Date)

FillField WordDoc, "YearDate", FillYear

FillField WordDoc, "JudgeSignature", "________________________________" & vbCr & "United States Bankruptcy Judge"
    

FillField WordDoc, "ElectronicSignatureMovantCounsel", "/s/ " & Forms!BankruptcyPrint!cbxAttorney
FillField WordDoc, "MovantCounsel", Forms!BankruptcyPrint!cbxAttorney
FillField WordDoc, "DebtorCounsel", "_______________________________" & vbCr & d![AttorneyFirstName] & " " & d![AttorneyLastName] & ", Esquire"
FillField WordDoc, "DebtorCounselName", d![AttorneyFirstName] & " " & d![AttorneyLastName] & ", Esquire"
FillField WordDoc, "Chapter13Trustee", "_______________________________" & vbCr & Trim(d![First] & " " & d![Last])
FillField WordDoc, "Chapter13TrusteeName", Trim(d![First] & " " & d![Last])
FillField WordDoc, "DebtorName", DebtorNames(0, 3, vbCr)

FillField WordDoc, "End", ""
WordObj.ScreenUpdating = True

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Order Modifying Relief.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Order Modifying Relief.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Function RepaymentItemList(FileNumber As Long) As String

Dim rstRepaymentTerms As Recordset
Dim i As Integer

i = 0
Set rstRepaymentTerms = CurrentDb.OpenRecordset("SELECT * FROM BKRepaymentTerms WHERE FileNumber=" & FileNumber & " ORDER BY RepaymentDate;", dbOpenSnapshot)
Do While Not rstRepaymentTerms.EOF
    i = i + 1
    RepaymentItemList = RepaymentItemList & Format$(rstRepaymentTerms!RepaymentAmount, "Currency") & " on or before " & Format$(rstRepaymentTerms!RepaymentDate, "mmmm d, yyyy") & vbCr
    rstRepaymentTerms.MoveNext
Loop


rstRepaymentTerms.Close

If (i = 0) Then
  RepaymentItemList = "No Repayment Terms indicated."
End If

End Function

Public Sub Doc_ConsentGrantRelief(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim FileNumber As Long, Judge As String, CoDebtor As Boolean, CoDebtorName As String, AffDate As String, AffInfo As String, DebtorsPlural As Boolean, SumRepaymentAmount As Currency, ParagraphNum As Integer, CourtCity As String

FileNumber = Forms![Case List]!FileNumber

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBKDocsWord WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Order Granting Relief VA.dot", False, 0, True)
WordObj.Visible = True
WordObj.ScreenUpdating = False
WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("AttorneyInfo").Select
WordDoc.Bookmarks("AttorneyInfo").Range.Text = "Mark D. Meyer" & vbCr & "VA Bar 74290"

Judge = Right$(UCase$(d!CaseNo), 3)
CoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13)


DebtorsPlural = (CountNames(FileNumber, "BKDebtor = True AND (Owner=True OR Mortgagor=True)") > 1)

FillField WordDoc, "Header", _
    IIf(d![Districts.State] <> "VA", vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr, "") & _
    "IN THE UNITED STATES BANKRUPTCY COURT" & vbCr & _
    "FOR THE " & UCase$(d!Name) & vbCr & _
    d!Location
FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & IIf(DebtorsPlural, "s", "")
FillField WordDoc, "InvestorAddr", UCase$(d!Investor) & vbCr & RemoveLF(d!InvestorAddress)
FillField WordDoc, "Respondents", _
    GetAddresses(0, 4, _
        IIf(d![BKdetails.Chapter] = 13, _
            "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", _
            "BKDebtor=True AND (Owner=True OR Mortgagor=True)"), vbCr) & _
    IIf(d![BKdetails.Chapter] = 7, _
        vbCr & vbCr & "and" & vbCr & vbCr & FormatName("", _
                                                        UCase$(Nz(d!First)), _
                                                        UCase$(Nz(d!Last)), _
                                                        ", CHAPTER 7 TRUSTEE", _
                                                        d!Address, _
                                                        d!Address2, _
                                                        d![BKTrustees.City], _
                                                        d![BKTrustees.State], _
                                                        d!Zip, _
                                                        vbCr), _
        "")
FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]

FillField WordDoc, "Caption", "ORDER GRANTING RELIEF FROM STAY"
FillField WordDoc, "Investor", d![Investor]
FillField WordDoc, "PropertyAddress", d![PropertyAddress] & IIf(Len(Forms!BankruptcyDetails!sfrmPropAddr!Apt & "") = 0, "", ", " & Forms!BankruptcyDetails!sfrmPropAddr!Apt) & ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d!ZipCode)
FillField WordDoc, "ShortLegal", Nz(d![ShortLegal])


FillField WordDoc, "CurrentDate", Format$(Date, "mmmm d, yyyy")
FillField WordDoc, "CurrentDateDay", DatePart("d", Date)
FillField WordDoc, "CurrentDateMonthYear", Format$(Date, "mmmm, yyyy")



CourtCity = Mid(d![Districts.Display], InStr(1, d![Districts.Display], "VA") + 3)
FillField WordDoc, "CourtCity", CourtCity

FillField WordDoc, "JudgeSignature", "________________________________" & vbCr & "United States Bankruptcy Judge"
    
FillField WordDoc, "MovantCounsel", [Forms]![BankruptcyPrint]![cbxAttorney]
FillField WordDoc, "DebtorCounsel", d!BKAttorneyFirm & ", Esquire"
If (d![BKdetails.Chapter] = 7) Then
  FillField WordDoc, "Chapter7Trustee", Trim(UCase$(Nz(d!First)) & " " & UCase$(Nz(d!Last)))
Else
  FillField WordDoc, "Chapter7Trustee", ""
End If

FillField WordDoc, "DebtorName", UCase$(DebtorNames(0, 3, vbCr))


FillField WordDoc, "ElectronicSignature", IIf([Forms]![BankruptcyPrint]![chElectronicSignature], "/s/ " & [Forms]![BankruptcyPrint]![cbxAttorney], "")
FillField WordDoc, "AttorneySignature", [Forms]![BankruptcyPrint]![cbxAttorney]
FillField WordDoc, "End", ""


WordObj.Selection.HomeKey wdStory, wdMove
WordObj.ScreenUpdating = True
WordDoc.SaveAs EMailPath & "Order Granting Relief.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Order Granting Relief.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_GreenTreeLossMitigation(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim FileNumber As Long, Judge As String, CoDebtor As Boolean, CoDebtorName As String, AffDate As String, AffInfo As String, DebtorsPlural As Boolean, SumRepaymentAmount As Currency, ParagraphNum As Integer, CourtCity As String

FileNumber = Forms![Case List]!FileNumber

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBKDocsWord WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "GreenTreeLossMitigationPackage.dot", False, 0, True)
WordObj.Visible = True

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Green Tree Loss Mitigation.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Green Tree Loss Mitigation.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_MilitaryAffidavitNoSSN(Keepopen As Boolean) 'Military Affidavit No SSN VA/DC
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim MName As String
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCdocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

If (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "VA" And Forms!Foreclosureprint!txtDesignator <> 3) Then
  MName = "Military Affidavit - No SSN Dove.dot"
ElseIf (d![FCdetails.State] = "VA") Then
  MName = "Military Affidavit - No SSN.dot"
Else
  MName = "Military Affidavit - No SSNDC.dot"
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & MName, False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])

FillField WordDoc, "MortgagorBorrowerNames", BorrowerMorgagorNamesOneNoSSN(d![CaseList.FileNumber], CopyNoR())

FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "OriginalMortgagors", IIf(IsNull(d!OriginalMortgagors), "", "Original Mortgagors " & d!OriginalMortgagors)
FillField WordDoc, "LoanNumber", d!LoanNumber
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(Forms![Foreclosuredetails]!State = "DC", IIf(D![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), "Commonwealth Trustees, LLC"), IIf(D![AIF] = True, D![LongClientName] & " as Attorney in Fact for ", "")) & Forms![Case List]!Investor
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(Forms![Foreclosuredetails]!State = "DC", IIf(D![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), "Commonwealth Trustees, LLC"), IIf(D![AIF] = True, D![LongClientName] & " as Attorney in Fact for ", "")) & IIf(D!ClientID = 531, "M&T Bank as Servicer for " & D![Investor], D![Investor])
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(Forms![Foreclosuredetails]!State = "DC", IIf(D![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), "Commonwealth Trustees, LLC"), IIf(D![AIF] = True, D![LongClientName] & " as Attorney in Fact for ", "") & IIf(D!ClientID = 531, "M&T Bank as Servicer for " & D![Investor], D![Investor]))
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(Forms![ForeclosureDetails]!State = "DC", IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), "Commonwealth Trustees, LLC"), IIf((d!AIF = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", ""))) & IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, Forms![Case List]!Investor)
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(Forms![ForeclosureDetails]!State = "DC", IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), "Commonwealth Trustees, LLC"), IIf((d!AIF = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, Forms![Case List]!Investor))))
FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(d![FCdetails.State] = "VA", "Commonwealth Trustees, LLC", IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee")), IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf((d![AIF] = True And d![ClientID] = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for " & d!Investor, IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for " & d!Investor, IIf(d![ClientID] = 531, "M&T Bank as Servicer for " & d![Investor], d![Investor]))))))) & ""


FillField WordDoc, "designatorDate", IIf([Forms]![Foreclosureprint]![txtDesignator] = 3, "Date: " & Format(Date, "mmmm d, yyyy"), "")

FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", IIf(Forms![foreclosuredetails]!State = "DC", IIf(Split(ReportArgs, "|")(3) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), (Split(ReportArgs, "|")(1))), (Split(ReportArgs, "|")(1)))


FillField WordDoc, "NotaryLocation", IIf(IsNull(d!NotaryLocation), "STATE OF______________" & vbCr & "COUNTY OF ____________ :", d!NotaryLocation)
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "only328client", IIf(d![ClientID] = 328, d![Investor] & " has been proven to be the real party in interest and has the right to foreclose the subject property.", "")


FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 152, "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & MName
Call SaveDoc(WordDoc, d![CaseList.FileNumber], MName)
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub
Public Sub Doc_MilitaryAffidavitNoSSNMD(Keepopen As Boolean) 'Military Affidavit NoSSN MD
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim fName As String


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")

If d!ClientID = 404 Then
    fName = "Military Affidavit - No SSNMD Bogman"
ElseIf (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "MD" And Forms!Foreclosureprint!txtDesignator <> 3) Then
    fName = "Military Affidavit - No SSNMD Dove"
Else
  fName = "Military Affidavit - No SSNMD"
End If

Set WordDoc = WordObj.Documents.Add(TemplatePath & fName & ".dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


FillField WordDoc, "designatorDate", IIf([Forms]![Foreclosureprint]![txtDesignator] = 3, "Date: " & Format(Date, "mmmm d, yyyy"), "")
FillField WordDoc, "MortgagorBorrowerNames", BorrowerMorgagorNamesOneNoSSN(d![CaseList.FileNumber], CopyNoR())
FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(D![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf(D![AIF] = True, D![LongClientName] & " as Attorney in Fact for ", "")) & Forms![Case List]!Investor
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")) & IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, Forms![Case List]!Investor)
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")) & IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, Forms![Case List]!Investor))
'FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for " & d!Investor, IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, Forms![Case List]!Investor))))
'it keeps changing!@!
FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(3) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for " & d![Investor], IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for " & d![Investor], IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 531, "M&T Bank as Servicer for " & [Forms]![Case List]![Investor], " " & [Forms]![Case List]![Investor]))))))))
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", IIf(Split(ReportArgs, "|")(3) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), (Split(ReportArgs, "|")(1)))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 152, "")
FillField WordDoc, "only328client", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, d![Investor] & " has been proven to be the real party in interest and has the right to foreclose the subject property.", ""))
FillField WordDoc, "only328client2", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, "Subscribed and affirmed before me in the county of Douglas, State of Colorado, this _____ day of ________, 20____ .", ""))
FillField WordDoc, "only328client3", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, "____________________________", ""))
FillField WordDoc, "only328client4", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, "(Notary's official Signature)", ""))
FillField WordDoc, "only328client5", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, "____________________________", ""))
FillField WordDoc, "only328client6", IIf((d!ClientID = 328 And Forms!Foreclosureprint!txtDesignator = 3), "", IIf(d![ClientID] = 328, "(Commission Expiration)", ""))

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & fName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], fName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_LossMitigationPreliminaryNationStar(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Loss Mitigation Prelim Nation star.dot", False, 0, True)
'Loss Mitigation - Preliminary-Nation-Star-MD.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: __________________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               __________________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: _________________________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Loss Mitigation - Preliminary-Nation-Star-MD.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Loss Mitigation - Preliminary-Nation-Star-MD.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_LossMitigationFinalNationStarMD(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

If d![ClientID] = 385 And d![FCdetails.State] = "MD" Then
    Set WordObj = CreateObject("Word.Application")
    'Set WordDoc = WordObj.Documents.Add(TemplatePath & "Loss Mitigation - Final-Nation-Star-MD.dot", False, 0, True)   Mei 10-5-15
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Loss Mitigation Final Nation Star.dot", False, 0, True)
    WordObj.Visible = True
End If

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
'Mei 10-05_15
FillField WordDoc, "fillingdate", IIf(Not IsNull(d![Docket]), Format(d![Docket], "mm/d/yyyy"), "_____________________ ")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1484, "")
FillField WordDoc, "Investor", IIf(d![ClientID] = 385 And d![Investor] = "Nationstar Mortgage LLC", "          Nationstar maintains records for the loan that is secured by the mortgage or deed of trust being foreclosed in this action. ", _
"          Nationstar services and maintains records on behalf of " & d![Investor] & ", the secured party to the mortgage or deed of trust being foreclosed in this action.")

FillField WordDoc, "Prior Servicer", IIf(Forms![Prior Servicer]!ChPrior, "           Before the servicing of this loan transferred to Nationstar, " & [Forms]![Prior Servicer]![TxtPriorServicer] & _
" (Prior Servicer) was the servicer for the loan and it maintained the loan servicing records.  When Nationstar began servicing this loan, Prior Servicer's records for the loan were integrated and boarded into Nationstar's systems," & _
" such that Prior Servicer's records, including the collateral file, payment histories, communication logs, default letters, information," & _
"and documents concerning the Loan are now integrated into Nationstar's business records.  Nationstar maintains quality control and " & _
"verification procedures as part of the boarding process to ensure the accuracy of the boarded records.  It is the regular business practice " & _
"of Nationstar to integrate prior servicers' records into Nationstar's business records and to rely upon those boarded records in providing " & _
"its loan servicing functions.  These Prior Servicer records have been integrated and are relied upon by Nationstar as part of Nationstar's " & _
"business records.", "")


FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ___________________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ___________________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: _________________________", "")

WordObj.Selection.HomeKey wdStory, wdMove
'WordDoc.SaveAs EMailPath & "Loss Mitigation - Final-Nation-Star-MD.doc"
WordDoc.SaveAs EMailPath & "Loss Mitigation Final Nationstar.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Loss Mitigation Final Nationstar.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_LossMitigationFinalPNC(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Loss Mitigation - Final PNC.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "Docket", Nz(Format(d!Docket, "mm/d/yyyy"))
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Loss Mitigation - Final PNC.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Loss Mitigation - Final PNC.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_LossMitigationPrelimPNC(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Loss Mitigation - Prelim PNC.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "Docket", Nz(Format(d!Docket, "mm/d/yyyy"))
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Loss Mitigation - Prelim PNC.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Loss Mitigation - Prelim PNC.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_StatementOfDebtFiguresMDCDMT(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim i As Integer, itemCount As Integer, totalItems As Currency, itemsFields As String



Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String

templateName = "Statement of Debt Figures MDCDMT"

'If (D![ClientID] = 157) Then
  'TemplateName = TemplateName & " Cenlar"
'End If
templateName = templateName & ".dot"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])

WordObj.Visible = False
If MsgBox("Is There A Loan Mod? ", vbYesNo) = vbYes Then
FillField WordDoc, "Mod", ", MODIFIED by Agreement effective " & Format(InputBox(" Effective Date?   Format mm/dd/yyyy"), "mmmm d, yyyy") & " with an amended principal balance of " & Format(InputBox(" Amended Principal Balance? "), "Currency")
Else
FillField WordDoc, "Mod", ""
End If
WordObj.Visible = True

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", IIf(d![ClientID] = 456, d!Investor, "M&T Bank As Servicer for Maryland Department of Housing and Community Development, Community Development Administration")


If (d![FCdetails.State] = "VA") Then
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", IIf(IsNull(d![Folio2]), ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & _
    " at Instrument Number " & d![Liber2], ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Book " & d![Liber2] & ", Page " & d![Folio2]))
ElseIf d!JurisdictionID = 6 Then  'Calvert County
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Book " & d![Liber2] & ", Page " & d![Folio2])
Else 'MD or DC
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Liber " & d![Liber2] & ", Folio " & d![Folio2])
End If

FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "Liber", LiberFolio(d![Liber], d![Folio], d![FCdetails.State], d![JurisdictionID])
FillField WordDoc, "LastPaymentApplied", Format$(d![LPIDate], "mmmm d, yyyy")
FillField WordDoc, "LastPayment", Format$(d![LPIDate] + 1, "mmmm d, yyyy")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "PaidStr", IIf((Nz(d![RemainingPBal], 0) > Nz(d![OriginalPBal], 0)), "Additional Interest", "Paid on principal")
FillField WordDoc, "Paid", Format$(d!OriginalPBal - d!RemainingPBal, "Currency")
FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

totalItems = 0
itemsFields = ""
Set dd = CurrentDb.OpenRecordset("SELECT Desc, Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber] & " ORDER BY Timestamp;", dbOpenSnapshot)
If dd.EOF Then      ' no extra lines
    FillField WordDoc, "Line_Items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items", itemsFields
    dd.MoveFirst
    i = 1
    Do While Not dd.EOF
        FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        totalItems = totalItems + Nz(dd!Amount, 0)
        dd.MoveNext
        i = i + 1
    Loop
End If
dd.Close

FillField WordDoc, "BalDueDate", IIf(IsNull([Forms]![Print Statement of Debt]![txtDueDate]), "______________", Format$([Forms]![Print Statement of Debt]![txtDueDate], "mmmm d"", ""yyyy"))
FillField WordDoc, "BalanceDue", Format$(d!RemainingPBal + totalItems, "Currency")
FillField WordDoc, "PerDiemInterest", IIf(d!LoanType = 3, "", IIf(IsNull(d!PerDiem), "$_____________", Format$(d!PerDiem, "Currency")))
FillField WordDoc, "Dime", IIf(d!LoanType = 3, "", "Per Diem Interest:")
FillField WordDoc, "InterestRate", IIf(IsNull(d!InterestRate), "____________ %", Format(d!InterestRate, "#0.000") & "%")
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", IIf(IsNull(d!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", d!NotaryLocation)





WordObj.Selection.HomeKey wdStory, wdMove
If d!ClientID = 531 Then
    WordDoc.SaveAs EMailPath & "Statement of Debt with Figures MDCDMT.doc"
    Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Statement of Debt with Figures MDCDMT.doc")
Else
    WordDoc.SaveAs EMailPath & "Statement of Debt with Figures M&T.doc"
    Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Statement of Debt with Figures M&T.doc")
End If
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close


End Sub

Public Sub Doc_StatementOfDebtWithFiguresBOA(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim i As Integer, itemCount As Integer, totalItems As Currency, amountInterest As Currency, itemsFields As String, J As Integer
Dim K As Recordset
Dim Rresult As Long
Dim InterestFrom As String
Dim InterestTo As String

'Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
Set d = CurrentDb.OpenRecordset("Select * FROM qryFCStmtofDebtBOAWord WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Statement of Debt Figures BOA"

If (d![ClientID] = 157) Then
  templateName = templateName & " Cenlar"
End If
templateName = templateName & ".dot"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProjectName").Select
WordDoc.Bookmarks("ProjectName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])

WordObj.Visible = False
If IsNull(d!LoanMod) = False Then
    FillField WordDoc, "Mod", " and modified by Loan Modification Agreement executed on " & Format$(d![LoanMod], "mmmm d, yyyy") & " and recorded on " & Format$(d!LoanModRecorded, "mmmm d, yyyy") & IIf(IsNull(d![LiberLoanMod] = True), "", " in Liber " & _
    d![LiberLoanMod] & " at Folio " & d![FolioLoanMod])
Else
    FillField WordDoc, "Mod", ""
End If

InterestFrom = InputBox("Interest From")
InterestTo = InputBox("Interest To")

WordObj.Visible = True

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor

FillField WordDoc, "owner", IIf(d![LoanType] = 4, "Federal National Mortgage Association is the owner", _
IIf(d![LoanType] = 5, "Federal Home Loan Mortgage Corporation is the owner", d![Investor] & " is the owner"))

FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "Liber", LiberFolio(d![Liber], d![Folio], d![FCdetails.State], d![JurisdictionID])
FillField WordDoc, "namefcin", IIf(d![Investor] <> d![LongClientName], "who services the loan which is the subject of this proceeding", "as an officer of BANA")
FillField WordDoc, "LastPaymentApplied", Format$(d![DateOfDefault], "mmmm d, yyyy")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "PaidStr", IIf((Nz(d![RemainingPBal], 0) > Nz(d![OriginalPBal], 0)), "Additional Interest", "Paid on principal")
FillField WordDoc, "Paid", Format$(d!OriginalPBal - d!RemainingPBal, "Currency")
'FillField WordDoc, "PBalance", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")
FillField WordDoc, "DateAsOf", Forms![Print Statement of Debt]!txtDueDate
FillField WordDoc, "DateEffective", Forms![Print Statement of Debt]!txtDueDate + 1

FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Phone Number:", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "1-866-467-8090", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Address:", "")
FillField WordDoc, "AnneArundel4", IIf(d![JurisdictionID] = 3, "4200 Amon Carter Blvd", "")
FillField WordDoc, "AnneArundel5", IIf(d![JurisdictionID] = 3, "Fort Worth, Texas 76155", "")

FillField WordDoc, "DateBalance", IIf(IsNull(Forms![Print Statement of Debt]!txtDueDate), "______________", Format$(Forms![Print Statement of Debt]!txtDueDate, "mmmm d"", ""yyyy"))
FillField WordDoc, "RT", "asdfadsf"
FillField WordDoc, "ReRecorded", IIf(IsNull(d!Rerecorded), "", "re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Liber " & d![Liber2] & " at Folio " & d![Folio2] & ", ")

totalItems = 0
amountInterest = 0
itemsFields = ""


'Set dd = CurrentDb.OpenRecordset("SELECT Desc, TempSortId, Amount FROM QryStatementOfDebtWithFigurersBOA WHERE FileNumber=" & d![CaseList.FileNumber] & " ORDER BY TempSortId DESC;", dbOpenSnapshot)
Set dd = CurrentDb.OpenRecordset("select * from statementofdebt where filenumber=" & d![CaseList.FileNumber] & " ORDER BY Sort_Desc ASC;", dbOpenDynaset, dbSeeChanges)
If dd.EOF Then      ' no extra lines
    FillField WordDoc, "Line_Items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
    
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items", itemsFields & " "
    
    dd.MoveFirst
    i = 1

     Do While Not dd.EOF
     '   FillField WordDoc, "Item" & i, IIf(dd!Desc = "Interest Due", "Interest Due from " & InterestFrom & " to " & InterestTo & _
     '   " @ " & IIf(Forms![Print Statement of Debt]!chVarRate = 0, Format(Forms![Print Statement of Debt]!InterestRate, "#.000%"), _
     '   " variable rate(s)"), dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency"))
        If dd!Desc = "Interest Due" Then
            FillField WordDoc, "Item" & i, "Interest Due from " & InterestFrom & " to " & InterestTo & " @ " & IIf(Forms![Print Statement of Debt]!chVarRate = 0, Format(Forms![Print Statement of Debt]!InterestRate, "#0.000") & "%", " variable rate(s)")
        ElseIf dd!Desc = "Escrow Balance Credit" Then
            FillField WordDoc, "Item" & i, dd!Desc & "                                            " & Format$(Nz(dd!Amount, 0), "Currency")
        ElseIf dd!Desc = "Payment Advance - Principal/Interest/Escrow" Then
            FillField WordDoc, "Item" & i, dd!Desc & "        " & Format$(Nz(dd!Amount, 0), "Currency")
        ElseIf dd!Desc = "Unapplied Funds Credit" Then
            FillField WordDoc, "Item" & i, dd!Desc & "                                           " & Format$(Nz(dd!Amount, 0), "Currency")
        ElseIf dd!Desc = "Credits" Then
            FillField WordDoc, "Item" & i, dd!Desc & vbTab & "                                           " & Format(DSum("[Amount]", "qryBOASOD", "[Credit] = True AND [FileNumber] =" & Forms![Case List]!FileNumber & ""), "Currency")
        Else
            FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        End If
        dd.MoveNext
        i = i + 1
    Loop
End If
    
 dd.Close
 
  Set K = CurrentDb.OpenRecordset("SELECT Desc, Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber], dbOpenSnapshot)
  J = 1
  Do While Not K.EOF
   If K!Desc = "Interest Amount" Then
   amountInterest = amountInterest + Format$(Nz(K!Amount, 0), "Currency")
   End If
      
   totalItems = totalItems + Nz(K!Amount, 0)
   K.MoveNext
   J = J + 1
  Loop
  K.Close

FillField WordDoc, "BalDueDate", IIf(IsNull([Forms]![Print Statement of Debt]![txtDueDate]), "______________", Format$([Forms]![Print Statement of Debt]![txtDueDate], "mmmm d"", ""yyyy"))
FillField WordDoc, "BalanceDue", Format$(d!RemainingPBal + totalItems, "Currency")
FillField WordDoc, "PerDiemInterest", IIf(IsNull(d!PerDiem), "$_____________", Format$(d!PerDiem, "Currency"))
FillField WordDoc, "InterestRate", IIf(IsNull(d!InterestRate), "____________ %", Format$(d!InterestRate, "#0.000") & "%")
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", IIf(IsNull(d!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", d!NotaryLocation)
FillField WordDoc, "amountInterest$", Format$(amountInterest, "Currency")
FillField WordDoc, "jurisdiction", d!Jurisdiction
FillField WordDoc, "dotdate#", Format$(d![DOTdate], "mmmm d, yyyy")
FillField WordDoc, "liberT", d![Liber]
FillField WordDoc, "folioT", d![Folio]
FillField WordDoc, "lpidate", d!LPIDate
FillField WordDoc, "MortgagorNamesintext", MortgagorNamesIntext(0, 2, 0, 0, -1)
'FillField WordDoc, "lpidate#", D![LPIDate]
FillField WordDoc, "lastmonth", LastMonth(d![LPIDate])


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Statement of Debt Figures BOA.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Statement of Debt Figures BOA.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_OwnershipAffidavitBOA(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit BOA.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])


FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "NoteOwner", FetchNoteOwner(d!LoanType, d!Investor, d![FCdetails.State])
FillField WordDoc, "namefcin", IIf(d![Investor] <> d![LongClientName], "who services the loan which is the subject of this proceeding", "as an officer of BANA")
FillField WordDoc, "Noteholders", GetNamesMD(0, 2, "Noteholder=True")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & d!Investor
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "FHLMCWording", IIf(d!LoanType = 5 Or d!LoanType = 4, ", and " & d!Investor & " is the holder of the Note having been transferred to " & d!Investor & " for the purposes of enforcement and conducting this foreclosure action", "")
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "jurisdiction", d!Jurisdiction
FillField WordDoc, "dotdate", Format(d![DOTdate], "mmmm d, yyyy")
FillField WordDoc, "FetchNoteOwnerT", FetchNoteOwner(d![LoanType], d![Investor], d![FCdetails.State])
FillField WordDoc, "CurrentT", IIf(d![LoanType] = 5, "", "current ")
FillField WordDoc, "Loantext", IIf(d![LoanType] = 1 Or d![LoanType] = 2 Or d![LoanType] = 3, "", "For purposes of foreclosure, the Note or other debt instrument is held by " & d![Investor] & ".")
FillField WordDoc, "Investor", d![Investor]

FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Phone Number:", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "1-866-467-8090", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Address:", "")
FillField WordDoc, "AnneArundel4", IIf(d![JurisdictionID] = 3, "4200 Amon Carter Blvd", "")
FillField WordDoc, "AnneArundel5", IIf(d![JurisdictionID] = 3, "Fort Worth, Texas 76155", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Note Ownership Affidavit BOA.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Note Ownership Affidavit BOA.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_OwnershipAffidavitOcwen(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit Ocwen.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])


FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "NoteOwner", FetchNoteOwner(d!LoanType, d!Investor, d![FCdetails.State])
FillField WordDoc, "namefcin", IIf(d![Investor] <> d![LongClientName], "who services the loan which is the subject of this proceeding", "as an officer of BANA")
FillField WordDoc, "Noteholders", GetNamesMD(0, 2, "Noteholder=True")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & d!Investor
'FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
'FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "FHLMCWording", IIf(d!LoanType = 5 Or d!LoanType = 4, ", and " & d!Investor & " is the holder of the Note having been transferred to " & d!Investor & " for the purposes of enforcement and conducting this foreclosure action", "")
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "jurisdiction", d!Jurisdiction
FillField WordDoc, "dotdate", Format(d![DOTdate], "mmmm d, yyyy")
FillField WordDoc, "FetchNoteOwnerT", FetchNoteOwner(d![LoanType], d![Investor], d![FCdetails.State])
FillField WordDoc, "CurrentT", IIf(d![LoanType] = 5, "", "current ")
FillField WordDoc, "Loantext", IIf(d![LoanType] = 1 Or d![LoanType] = 2 Or d![LoanType] = 3, "", "For purposes of foreclosure, the Note or other debt instrument is held by " & d![Investor] & ".")
FillField WordDoc, "Investor", d![Investor]
FillField WordDoc, "Loaner", IIf(d!LoanType = 5, "I HEREBY CERTIFY that Federal Home Loan Mortgage Corp. (“FHLMC” is the owner of the debt instrument secured by the Mortgage or Deed of Trust which is the subject of the instant foreclosure action, and has authorized Ocwen Loan Servicing, LLC to be the noteholder for purposes of this foreclosure action.", _
"I HEREBY CERTIFY that " & d![Investor] & " is the owner of the debt instrument secured by the Mortgage or Deed of Trust which is the subject of the instant foreclosure action.")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Note Ownership Affidavit Ocwen.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Note Ownership Affidavit Ocwen.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_BOACoverSheet(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If


    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit Cover letter BOA.dot", False, 0, True)
    WordObj.Visible = True


WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "MortgagorNameT", MortgagorNames(0, 3)
FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)

FillField WordDoc, "ContactNameT", IIf([Forms]![foreclosuredetails]![State] = "VA", GetLoginName() & vbCrLf & vbCrLf & vbTab & vbTab & "Return the executed documents to:" _
& vbCrLf & vbTab & vbTab & "Rosenberg & Associates" & vbCrLf & vbTab & vbTab & "ATTN: Sale Setting Team" & vbCrLf & vbTab & vbTab & "8601 Westwood Center Dr, Suite 255" & vbCrLf & vbTab & vbTab & "Vienna, VA 22182", _
IIf([Forms]![Foreclosureprint]![chAssignment] = -1, GetLoginName() & vbCrLf & vbCrLf & vbTab & vbTab & "Return the executed Assignment to:" _
& vbCrLf & vbTab & vbTab & "Rosenberg & Associates" & vbCrLf & vbTab & vbTab & "ATTN: Assignments Department" & vbCrLf & vbTab & vbTab & "8601 Westwood Center Dr, Suite 255" & vbCrLf & vbTab & vbTab & "Vienna, VA 22182" _
, GetLoginName() & vbCrLf & vbCrLf & vbTab & vbTab & "Return the executed documents to:" & vbCrLf & vbTab & vbTab & "Rosenberg & Associates" & vbCrLf & vbTab & vbTab & "ATTN: Docketing Team" & vbCrLf _
& vbTab & vbTab & "7910 Woodmont Avenue, Suite 750" & vbCrLf & vbTab & vbTab & "Bethesda, MD 20814"))

FillField WordDoc, "DateofUpload", Format$(Now(), "mmmm d, yyyy")

FillField WordDoc, "Disclaimer2", IIf(Forms![Case List]!State = "MD", trusteeNames(0, 2) & ", the substitute trustees listed, are employees of the attorney firm", "")
'FillField WordDoc, "Disclaimer", IIf(Not IsNull(Forms![Print Assignment]!cboBOANoteLocation), Forms![Print Assignment]!cboBOANoteLocation.Column(0), IIf(Forms!ForeclosureDetails!DocBackOrigNote = -1, "The firm is in possession of the original note", "The firm is NOT in possession of the original note"))

FillField WordDoc, "Disclaimer", IIf(Forms!foreclosuredetails!DocBackOrigNote = -1, "The firm is in possession of the original note", "The firm is NOT in possession of the original note")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Note Ownership Affidavit Cover letter.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Note Ownership Affidavit Cover letter.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_BOACoverSheetAssignment(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If


    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit Cover letter BOA.dot", False, 0, True)
    WordObj.Visible = True


WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "MortgagorNameT", MortgagorNames(0, 3)
FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)

FillField WordDoc, "ContactNameT", IIf([Forms]![foreclosuredetails]![State] = "VA", GetLoginName() & vbCrLf & vbCrLf & vbTab & vbTab & "Return the executed documents to:" _
& vbCrLf & vbTab & vbTab & "Rosenberg & Associates" & vbCrLf & vbTab & vbTab & "ATTN: Sale Setting Team" & vbCrLf & vbTab & vbTab & "8601 Westwood Center Dr, Suite 255" & vbCrLf & vbTab & vbTab & "Vienna, VA 22182", _
IIf([Forms]![Foreclosureprint]![chAssignment] = -1, GetLoginName() & vbCrLf & vbCrLf & vbTab & vbTab & "Return the executed Assignment to:" _
& vbCrLf & vbTab & vbTab & "Rosenberg & Associates" & vbCrLf & vbTab & vbTab & "ATTN: Assignments Department" & vbCrLf & vbTab & vbTab & "8601 Westwood Center Dr, Suite 255" & vbCrLf & vbTab & vbTab & "Vienna, VA 22182" _
, GetLoginName() & vbCrLf & vbCrLf & vbTab & vbTab & "Return the executed documents to:" & vbCrLf & vbTab & vbTab & "Rosenberg & Associates" & vbCrLf & vbTab & vbTab & "ATTN: Docketing Team" & vbCrLf _
& vbTab & vbTab & "7910 Woodmont Avenue, Suite 750" & vbCrLf & vbTab & vbTab & "Bethesda, MD 20814"))

FillField WordDoc, "DateofUpload", Format$(Now(), "mmmm d, yyyy")

FillField WordDoc, "Disclaimer2", IIf(Forms![Case List]!State = "MD", trusteeNames(0, 2) & ", the substitute trustees listed, are employees of the attorney firm", "")
FillField WordDoc, "Disclaimer", IIf(Not IsNull(Forms![Print Assignment]!cboBOANoteLocation), Forms![Print Assignment]!cboBOANoteLocation.Column(0), IIf(Forms!foreclosuredetails!DocBackOrigNote = -1, "The firm is in possession of the original note", "The firm is NOT in possession of the original note"))

'FillField WordDoc, "Disclaimer", IIf(Forms!ForeclosureDetails!DocBackOrigNote = -1, "The firm is in possession of the original note", "The firm is NOT in possession of the original note")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Note Ownership Affidavit Cover letter.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Note Ownership Affidavit Cover letter.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_45DayNoteOwnershipAffidavitBOA(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit 45 days BOA.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])



FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "NoteOwner", FetchNoteOwner(d!LoanType, d!Investor, d![FCdetails.State])
FillField WordDoc, "Noteholders", GetNamesMD(0, 2, "Noteholder=True")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "DateOfDefault", Format$(d!DateOfDefault, "mmmm d, yyyy")
FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & d!Investor
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "namefcin", IIf(d![Investor] <> d![LongClientName], "who services the loan which is the subject of this proceeding", "as an officer of BANA")
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "FHLMCWording", IIf(d!LoanType = 5 Or d!LoanType = 4, ", and " & d!Investor & " is the holder of the Note having been transferred to " & d!Investor & " for the purposes of enforcement and conducting this foreclosure action", "")
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 592, "")
FillField WordDoc, "ifstatmd", IIf(d![FCdetails.State] = "MD", "________________  ", Format$(d![NOI], "mmmm d, yyyy"))
FillField WordDoc, "jurisdiction", d![Jurisdiction]

FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Phone Number:", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "1-866-467-8090", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Address:", "")
FillField WordDoc, "AnneArundel4", IIf(d![JurisdictionID] = 3, "4200 Amon Carter Blvd", "")
FillField WordDoc, "AnneArundel5", IIf(d![JurisdictionID] = 3, "Fort Worth, Texas 76155", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Note Ownership Affidavit 45 days BOA.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Note Ownership Affidavit 45 days BOA.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_StatementOfDebtWithFiguresJP(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim i As Integer, itemCount As Integer, totalItems As Currency, amountInterest As Currency, itemsFields As String, J As Integer
Dim K As Recordset
Dim Rresult As Long


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Statement of Debt Figures JP"

'If (D![ClientID] = 157) Then
'  TemplateName = TemplateName & " Cenlar"
'End If
templateName = templateName & ".dot"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProjectName").Select
WordDoc.Bookmarks("ProjectName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "owner", IIf(d![LoanType] = 4, "holder", IIf(d![LoanType] = 5, "holder", "owner"))
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "BorrowerNames", BorrowerNames(0)
FillField WordDoc, "Item2", IIf(Forms!foreclosuredetails!LienPosition <> 1, "The Note and Mortgage, (and modification agreement where applicable) are collectively referred to as the " & Chr(34) & "Loan Documents." & Chr(34), "")
FillField WordDoc, "Item5", IIf(Forms!foreclosuredetails!LienPosition <> 1, "5.  As of " & " __________ (unless specified otherwise below), Plaintiff seeks to recover the following itemized sums of money that are due and owing under the Loan Documents, exclusive of attorney's fees and costs in this action:", "5. The Borrower owes, as of " & " __________ " & "," & " following itemized sums of money, exclusive of costs and expenses that the borrower owes: ")
'FillField WordDoc, "Liber", LiberFolio(D![Liber], D![Folio], D![FCdetails.State], D![JurisdictionID])
'FillField WordDoc, "LastPaymentApplied", Format$(D![DateOfDefault], "mmmm d, yyyy")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
'FillField WordDoc, "PaidStr", IIf((Nz(D![RemainingPBal], 0) > Nz(D![OriginalPBal], 0)), "Additional Interest", "Paid on principal")
'FillField WordDoc, "Paid", Format$(D!OriginalPBal - D!RemainingPBal, "Currency")
FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")
'FillField WordDoc, "DateAsOf", Forms![Print Statement of Debt]!txtDueDate
'FillField WordDoc, "DateEffective", Forms![Print Statement of Debt]!txtDueDate + 1
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ______________________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ______________________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ____________________________", "")

'FillField WordDoc, "DateBalance", IIf(IsNull(Forms![Print Statement of Debt]!txtDueDate), "______________", Format$(Forms![Print Statement of Debt]!txtDueDate, "mmmm d"", ""yyyy"))

totalItems = 0
amountInterest = 0
itemsFields = ""
Set dd = CurrentDb.OpenRecordset("SELECT Desc,  Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber], dbOpenSnapshot)
If dd.EOF Then      ' no extra lines
    FillField WordDoc, "Line_Items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
       
    
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items", itemsFields & " "
    
    dd.MoveFirst
    i = 1
    Do While Not dd.EOF
        FillField WordDoc, "Item" & i, IIf(dd!Desc = "Escrow", dd!Desc, dd!Desc & vbTab & "$___________")
        'TotalItems = TotalItems + Nz(dd!Amount, 0)
        dd.MoveNext
        i = i + 1
    Loop
End If
 dd.Close
 
  Set K = CurrentDb.OpenRecordset("SELECT Desc, Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber], dbOpenSnapshot)
  J = 1
  Do While Not K.EOF
   If K!Desc = "Interest" Then
   amountInterest = amountInterest + Format$(Nz(K!Amount, 0), "Currency")
   End If
      
   totalItems = totalItems + Nz(K!Amount, 0)
   K.MoveNext
   J = J + 1
  Loop
  K.Close

'FillField WordDoc, "BalanceDue", Format$(K!Amount, "Currency")

'FillField WordDoc, "BalDueDate", IIf(IsNull([Forms]![Print Statement of Debt]![txtDueDate]), "______________", Format$([Forms]![Print Statement of Debt]![txtDueDate], "mmmm d"", ""yyyy"))
FillField WordDoc, "BalanceDue", Format$(d!RemainingPBal + totalItems, "Currency")
FillField WordDoc, "PerDiemInterest", IIf(IsNull(d!PerDiem), "$_____________", Format$(d!PerDiem, "Currency"))
'FillField WordDoc, "InterestRate", IIf(IsNull(D!InterestRate), "____________ %", Format$(D!InterestRate, "#.000%"))
'FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
'FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
'FillField WordDoc, "NotaryLocation", IIf(IsNull(D!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", D!NotaryLocation)
FillField WordDoc, "amountInterest$", Format$(amountInterest, "Currency")
FillField WordDoc, "jurisdiction", d!Jurisdiction
FillField WordDoc, "dotdate#", d![DOTdate]
'FillField WordDoc, "liberT", D![Liber]
'FillField WordDoc, "folioT", D![Folio]
FillField WordDoc, "lpidate", d!LPIDate
FillField WordDoc, "LPIdate+1", DateAdd("d", 1, d![LPIDate])
FillField WordDoc, "MortgagorNamesintext", MortgagorNamesIntext(0, 2, 0, 0, -1)
'FillField WordDoc, "lpidate#", D![LPIDate]
FillField WordDoc, "lastmonth", LastMonth(d![LPIDate])


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Statement of Debt Figures JP.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Statement of Debt Figures JP.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_DeedOfAppointmentChase(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String

'TemplateName = "Deed of Appointment Chase MD"
If (d![FCdetails.State] = "MD") Then
 templateName = "Deed of Appointment Chase MD"
Else
templateName = "Deed of Appointment Chase VA"
End If
'End If
'If D![ClientID] = 345 Then
'  TemplateName = TemplateName & " Kondaur"
'ElseIf (D![ClientID] = 334 Or D![ClientID] = 477) Then
'  TemplateName = TemplateName & " Saxon"
'ElseIf D![ClientID] = 87 Then
'  TemplateName = TemplateName & " PNC"
'
'End If

templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")

Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

If d![FCdetails.State] = "MD" Then
    'Do nothing'
    'Ticket 823 only provide FileNumber in Margin
    WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
    WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
Else
    WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
    WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
    WordDoc.Bookmarks("ProName").Select
    WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
    WordDoc.Bookmarks("PropertyAddress").Select
    WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
    WordDoc.Bookmarks("APTNUM").Select
    WordDoc.Bookmarks("APTNUM").Range.Text = Nz(d![Fair Debt])
End If

If (d![FCdetails.State] = "VA") Then
    FillField WordDoc, "reRecorded", IIf(IsNull(d![Rerecorded]), "", IIf(IsNull(d![Folio2]), ", and re-recorded " & Format$(d![Rerecorded], "mmmm d, yyyy") & _
    " in Instrument Number " & d![Liber2] & ", ", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Book " & d![Liber2] & ", at Page " & d![Folio2] & ", "))
Else 'MD or DC
    FillField WordDoc, "reRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Deed/Book " & d![Liber2] & ", page " & d![Folio2] & ",")
End If


FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "TrusteeNames", trusteeNames(0, 2)
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "MortgagorNames", IIf(IsNull(d!OriginalMortgagors), MortgagorNamesCaps(0, 2, 2), d!OriginalMortgagors)
'FillField WordDoc, "AssumedBy", IIf(IsNull(D!OriginalMortgagors), "", "assumed by " & MortgagorNamesCaps(0, 2, 2)) & " "
FillField WordDoc, "Jurisdiction", d!Jurisdiction
'FillField WordDoc, "LongState", Nz(D!LongState)
FillField WordDoc, "Liber", d!Liber
FillField WordDoc, "Filio", IIf(IsNull(d!Folio), "", d!Folio)
'FillField WordDoc, "LegalDescription", D!LegalDescription
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "TaxID", d!TaxID
FillField WordDoc, "OriginalTrustee", d!OriginalTrustee
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "OriginalBeneficiary", d!OriginalBeneficiary
FillField WordDoc, "Book", IIf(IsNull(d!Folio), _
" in Book  N/A , at Page N/A or Instrument/Recording" & IIf(d![FCdetails.State] = "MD", " No.", " Number ") _
& d!Liber, " in Book " & d!Liber & ", at Page " & d!Folio & " or Instrument/Recording" & IIf(d![FCdetails.State] = "MD", " No. N/A", " Number N/A"))
FillField WordDoc, "DoTrecorded", Format$(d!DOTrecorded, "mmmm d"", ""yyyy")
'FillField WordDoc, "MDInstrPrepared", IIf(D![FCdetails.State] = "VA", "", "This instrument was prepared under the supervision of " & D!AttorneyName & ", an attorney admitted to practice before the Court of Appeals of Maryland.")
'FillField WordDoc, "MDSignLine", IIf(D![FCdetails.State] = "VA", "", "_________________________________")
'FillField WordDoc, "MDAttorney", IIf(D![FCdetails.State] = "VA", "", D!AttorneyName)
'FillField WordDoc, "InvestorAIF", IIf(D!AIF = True, D!LongClientName & " as Attorney in Fact for ", "") & D!Investor
'FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
'FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
'FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!ForeclosurePrint!NotaryID)
'FillField WordDoc, "NotaryName", FetchNotaryName(Forms!ForeclosurePrint!NotaryID, False)
'FillField WordDoc, "FirmAddress", IIf(D![FCdetails.State] = "VA", "Commonwealth Trustees, LLC" & vbCr & "c/o Rosenberg & Associates, LLC" & vbCr & "8601 Westwood Center Drive, Suite 255" & vbCr & "Vienna, VA 22182", FirmAddress(vbCr))
'FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!ForeclosurePrint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 226, "")

'FillField WordDoc, "FinalLanguage", "WHEREAS, " & _
IIf(Forms!foreclosuredetails!LoanType = 5, "Federal Home Loan Mortgage Corporation is the owner of the note secured by said Deed of Trust and appointed the party of the first part with authority to hold, collect and enforce the note; and", _
IIf(Forms!foreclosuredetails!LoanType = 4, "Federal National Mortgage Association is the owner of the note secured by said Deed of Trust and appointed the party of the first part with authority to hold, collect and enforce the note; and", _
 "the party of the first part is the owner and holder of the Note secured by said Deed of Trust; and,"))





WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & templateName & ".dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], templateName & ".dot")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_OwnershipAffidavitChase(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")

'#1226 MC 10/15/2014
'If d!JurisdictionID = 3 Then
'    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit JP Anne.dot", False, 0, True)
'Else
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit JP.dot", False, 0, True)
'End If
'/#1226

WordObj.Visible = True




WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])



FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)

FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 464, "")
'#1226 MC 10/15/2014
'FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: _______________________", "")
'FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               _______________________", "")
'FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: _____________________", "")
'/#1226

Dim Text5, Text1, Text3, Text4 As String
Text5 = "that  Federal Home Loan Mortgage Corporation is the owner of the loan evidenced by the Note, that  " & d![LongClientName] & " is the servicer of the loan evidenced by the Note and that the copy of the Note filed in this foreclosure case is a true and accurate copy. "
Text4 = "that  Federal National Mortgage Association is the owner of the loan evidenced by the Note, that " & d![LongClientName] & " is the servicer of the loan evidenced by the Note, and that the copy of the Note filed in this foreclosure action is a true and accurate copy."
Text1 = "that " & d![Investor] & " is the owner of the loan evidenced by the Note that is the subject of this foreclosure action and that the copy of the Note filed in this foreclosure case is a true and  accurate copy."
Text3 = "that " & d![LongClientName] & " is the owner of the loan evidenced by the Note that is the subject of this foreclosure action, as demonstrated by the fact that it is the holder of the Note."

FillField WordDoc, "budytext", IIf(Forms!foreclosuredetails!LoanType = 5, Text5, IIf(Forms!foreclosuredetails!LoanType = 4, Text4, IIf(Forms!foreclosuredetails!LoanType = 1 Or Forms!foreclosuredetails!LoanType = 2 Or Forms!foreclosuredetails!LoanType = 3, Text1, Text3)))

FillField WordDoc, "filenumber", d![CaseList.FileNumber]
WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Note Ownership Affidavit JP.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Note Ownership Affidavit JP.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_DeedofAppointmentWells(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
'Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
Dim templateName As String
Dim BankName As String

If Forms![foreclosuredetails]!State = "MD" Then
templateName = "Deed of Appointment WellsMD"
Else
templateName = "Deed of Appointment WellsVA"
End If

If Forms![Case List]!ClientID = 6 Then BankName = "Deed of Appointment-Wells Fargo Bank NA"
If Forms![Case List]!ClientID = 556 Then BankName = "Deed of Appointment-Wells Fargo Home Mortgage"

templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")

Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)

WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


FillField WordDoc, "bankName", d!LongClientName
FillField WordDoc, "BankAddress", d!BankStreet
FillField WordDoc, "BankCity", d!BankCity
FillField WordDoc, "BankState", d!BankState
FillField WordDoc, "BankZipCod", FormatZip(d!BankZipcode)
FillField WordDoc, "Tax", d!TaxID
FillField WordDoc, "Addy", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "DOTdate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 20)
FillField WordDoc, "originaltrustee", d!OriginalTrustee
FillField WordDoc, "OriginalBeneficiary", d!OriginalBeneficiary
FillField WordDoc, "DOTrecorded", Format$(d!DOTrecorded, "mmmm d, yyyy")

FillField WordDoc, "Jurisdiction", d!Jurisdiction

'FillField WordDoc, "JurisdictionMD", IIf(D![Jurisdiction] = "Baltimore City", "City of Baltimore", "County of " & D![Jurisdiction])
FillField WordDoc, "JurisdictionMD", IIf(d![Jurisdiction] = "Baltimore City", "City of Baltimore", IIf(d![JurisdictionID] <> 18, "County of " & Replace(d![Jurisdiction], " County", ""), "County of " & Replace(d![Jurisdiction], "'s County", "")))

FillField WordDoc, "Folio", d!Folio
FillField WordDoc, "Liber", d!Liber

'FillField WordDoc, "Folio2", d!Folio2
'FillField WordDoc, "Liber2", d!Liber2
FillField WordDoc, "Rerecorded", IIf(IsNull(d![Rerecorded]) = True, "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Liber " & d![Liber2] & " at Folio " & d![Folio2]) & ", "


FillField WordDoc, "N/AFoilo", IIf(IsNull(d![Folio]), ", under Instrument Number " & d![Liber] & ", in book N/A, page N/A", ", under Instrument Number N/A, in book " & d![Liber] & ", " & "page " & d![Folio])
FillField WordDoc, "FCdetails.State", d![FCdetails.State]
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "TrusteeNames", IIf(Forms![foreclosuredetails]!State = "MD", trusteeNames(0, 2), "Commonwealth Trustees, LLC")
FillField WordDoc, "investor", d![Investor]
FillField WordDoc, "investorMD", IIf(d![Investor] <> "Wells Fargo Bank, N.A.", "authorized agent of " & d![Investor], d![Investor])
FillField WordDoc, "LegalDescription", d!LegalDescription

'FillField WordDoc, "trusteeType", IIf(Forms!foreclosureDetails!Option221, "Successor Trustee ", "Substitute Trustee ")
FillField WordDoc, "trusteeList", getTrustees(Forms!foreclosuredetails!lstTrustees, Forms!foreclosuredetails!lstTrustees)

FillField WordDoc, "Text27", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & _
", " & d!City & ", " & d![FCdetails.State] & " " & FormatZip(d!ZipCode)

FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 226, "")
'If Page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 464, ""), 450, 4000, True)

FillField WordDoc, "filenumber", d![CaseList.FileNumber]
WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & BankName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], BankName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_DeedofAppointmentWellsVA(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
Dim templateName As String
Dim BankName As String

templateName = "Deed of Appointment WellsVA"

If Forms![Case List]!ClientID = 6 Then BankName = "Deed of Appointment-Wells Fargo Bank NA"
If Forms![Case List]!ClientID = 556 Then BankName = "Deed of Appointment-Wells Fargo Home Mortgage"

templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")

Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)

WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

FillField WordDoc, "addressVA", FirmShortAddressVA()
FillField WordDoc, "bankName", d!LongClientName
FillField WordDoc, "BankAddress", d!BankStreet
FillField WordDoc, "BankCity", d!BankCity
FillField WordDoc, "BankState", d!BankState
FillField WordDoc, "BankZipCod", FormatZip(d!BankZipcode)
FillField WordDoc, "Tax", d!TaxID

FillField WordDoc, "Property", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "DOTdate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 100)
FillField WordDoc, "originaltrustee", d!OriginalTrustee
FillField WordDoc, "OriginalBeneficiary", d!OriginalBeneficiary
FillField WordDoc, "DOTrecorded", Format$(d!DOTrecorded, "mmmm d, yyyy")
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "N/AFolio", IIf(IsNull(d![Folio]), ", under Instrument Number " & d![Liber] & ", in book N/A, page N/A", ", under Instrument Number N/A, in book " & d![Liber] & ", " & "page " & d![Folio])
FillField WordDoc, "FCdetails.State", d![FCdetails.State]
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "TrusteeNames", IIf(Forms![foreclosuredetails]!State = "MD", trusteeNames(0, 2), "Commonwealth Trustees, LLC")
FillField WordDoc, "investor", d![Investor]
'FillField WordDoc, "LegalDescription", D!LegalDescription

'fillField WordDoc, "Text27", D!PropertyAddress & ", " & D!City & ", " & D![FCdetails.State] & " " & FormatZip(D!ZipCode)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 226, "")
'If Page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 464, ""), 450, 4000, True)
FillField WordDoc, "filenumber", d![CaseList.FileNumber]

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & BankName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], BankName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_DeedofAppointmentWellsVASelect(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
Dim templateName As String
Dim BankName As String

templateName = "Deed of Appointment WellsVA Select"

If Forms![Case List]!ClientID = 6 Then BankName = "Deed of Appointment-Wells Fargo Bank NA"
If Forms![Case List]!ClientID = 556 Then BankName = "Deed of Appointment-Wells Fargo Home Mortgage"

templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")

Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)

WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


FillField WordDoc, "reRecorded", IIf(IsNull(d![Rerecorded]), "", IIf(IsNull(d![Folio2]), " and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & _
" under Instrument Number " & d![Liber2], ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Book " & d![Liber2] & ", Page " & d![Folio2]))
 
FillField WordDoc, "LoanMod", IIf(IsNull(d![LoanMod]), "", IIf(IsNull(d![FolioLoanMod]), " modified and recorded on " & Format$(d![LoanMod], "mmmm d, yyyy") & " as Instrument No. " & d![LiberLoanMod] & _
" in the office of the Register of Deeds of " & d![Jurisdiction] & ", " & d![FCdetails.State] & ", ", " modified and recorded on " & Format$(d![LoanMod], "mmmm d, yyyy") & " in Book " & d![LiberLoanMod] & ", Page " & d![FolioLoanMod] & ","))

FillField WordDoc, "addressVA", FirmShortAddressVA()
FillField WordDoc, "bankName", d!LongClientName
FillField WordDoc, "BankAddress", d!BankStreet
FillField WordDoc, "BankCity", d!BankCity
FillField WordDoc, "BankState", d!BankState
FillField WordDoc, "BankZipCod", FormatZip(d!BankZipcode)
'FillField WordDoc, "Tax", d!TaxID
FillField WordDoc, "Tax", IIf(d![JurisdictionID] = 56, "Tax Map Number " & d![TaxID], "Tax ID No: " & d![TaxID])
FillField WordDoc, "Property", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "DOTdate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "Mortgagors", MortgagorNames(0, 20)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 100)
FillField WordDoc, "originaltrustee", d!OriginalTrustee
FillField WordDoc, "OriginalBeneficiary", d!OriginalBeneficiary
FillField WordDoc, "DOTrecorded", Format$(d!DOTrecorded, "mmmm d, yyyy")
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "N/AFolio", IIf(IsNull(d![Folio]), ", under Instrument Number " & d![Liber] & ", in book N/A, page N/A", ", under Instrument Number N/A, in book " & d![Liber] & ", " & "page " & d![Folio])
FillField WordDoc, "FCdetails.State", d![FCdetails.State]
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "TrusteeNames", IIf(Forms![foreclosuredetails]!State = "MD", trusteeNames(0, 2), "Commonwealth Trustees, LLC")
FillField WordDoc, "investor", d![Investor]
'FillField WordDoc, "LegalDescription", D!LegalDescription

'fillField WordDoc, "Text27", D!PropertyAddress & ", " & D!City & ", " & D![FCdetails.State] & " " & FormatZip(D!ZipCode)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 226, "")
'If Page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 464, ""), 450, 4000, True)
FillField WordDoc, "filenumber", d![CaseList.FileNumber]



WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & BankName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], BankName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_NOIAffidavitWells(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
Dim MainDocument As String
Dim templateName As String



If Forms![Case List]!ClientID = 6 Then templateName = "NOI Wells Fargo Bank NA"
If Forms![Case List]!ClientID = 556 Then templateName = "NOI Wells Fargo Home Mortgage"


'TemplateName = TemplateName & ".dot"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Note Ownership Affidavit Wells.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "NoteOwner", FetchNoteOwner(d!LoanType, d!Investor, d![FCdetails.State])
FillField WordDoc, "Noteholders", GetNamesMD(0, 2, "Noteholder=True")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & d!Investor
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "FHLMCWording", IIf(d!LoanType = 5 Or d!LoanType = 4, ", and " & d!Investor & " is the holder of the Note having been transferred to " & d!Investor & " for the purposes of enforcement and conducting this foreclosure action", "")
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "jurisdiction", d!Jurisdiction
FillField WordDoc, "dotdate", Format(d![DOTdate], "mmmm d, yyyy")
FillField WordDoc, "FetchNoteOwnerT", FetchNoteOwner(d![LoanType], d![Investor], d![FCdetails.State])
FillField WordDoc, "CurrentT", IIf(d![LoanType] = 5, "", "current ")
FillField WordDoc, "Investor", d![Investor]
FillField WordDoc, "BorrowerNames", BorrowerNames(d![CaseList.FileNumber])
FillField WordDoc, "DOTIF", IIf(d!DOT = True, "Deed of trust", "Mortgage")
FillField WordDoc, "DateOfDefault", Format$(d![DateOfDefault], "mmmm d, yyyy")
FillField WordDoc, "MortgagorNamesOneline", MortgagorNames(0, 2)
FillField WordDoc, "NOI", Format$(d![NOI], "mmmm d, yyyy")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 592, "")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & templateName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], templateName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_OwnershipAffidavitWells(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
Dim MainDocument As String
Dim templateName As String
Dim Text5 As String
Dim Text4 As String
Dim Textother As String

Text5 = "Federal Home Loan Mortgage Corporation is the owner of the loan evidenced by the Note, and that Federal Home Loan Mortgage Corporation has authorized " & d![Investor] & " to be the holder of the Note for purposes of conducting this foreclosure action."
Text4 = "Federal National Mortgage Association is the owner of the loan evidenced by the Note, and that Federal National Mortgage Association has authorized " & d![Investor] & " to be the holder of the Note for purposes of conducting this foreclosure action."
Textother = d![Investor] & " is the owner and holder of the loan evidenced by the Note."

If Forms![Case List]!ClientID = 6 Then templateName = "Ownership Affidavit Wells"
If Forms![Case List]!ClientID = 556 Then templateName = "Ownership Affidavit Wells"


'TemplateName = TemplateName & ".dot"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Ownership Affidavit Wells.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])


FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "NoteOwner", FetchNoteOwner(d!LoanType, d!Investor, d![FCdetails.State])
FillField WordDoc, "Noteholders", GetNamesMD(0, 2, "Noteholder=True")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "correct", IIf(d![FCdetails.State] = "MD", "accurate", "correct")
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & d!Investor
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "FHLMCWording", IIf(d!LoanType = 5 Or d!LoanType = 4, ", and " & d!Investor & " is the holder of the Note having been transferred to " & d!Investor & " for the purposes of enforcement and conducting this foreclosure action", "")
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "jurisdiction", d!Jurisdiction
FillField WordDoc, "dotdate", Format(d![DOTdate], "mmmm d, yyyy")
FillField WordDoc, "FetchNoteOwnerT", FetchNoteOwner(d![LoanType], d![Investor], d![FCdetails.State])
FillField WordDoc, "CurrentT", IIf(d![LoanType] = 5, "", "current ")
FillField WordDoc, "Investor", d![Investor]
FillField WordDoc, "BorrowerNames", BorrowerNames(d![CaseList.FileNumber])
FillField WordDoc, "DOTIF", IIf(d!DOT = True, "Deed of trust", "Mortgage")
FillField WordDoc, "DateOfDefault", Format$(d![DateOfDefault], "mmmm d, yyyy")
FillField WordDoc, "MortgagorNamesOneline", MortgagorNames(0, 2)
FillField WordDoc, "NOI", Format$(d![NOI], "mmmm d, yyyy")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 464, "")
FillField WordDoc, "Textbudy", IIf(d![LoanType] = 5, Text5, IIf(d![LoanType] = 4, Text4, Textother))
FillField WordDoc, "InvestorIf", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "") & d![Investor]
FillField WordDoc, "Anne", IIf(d![JurisdictionID] = 3, "Address: _______________", "")
FillField WordDoc, "Anne2", IIf(d![JurisdictionID] = 3, "                " & "_______________", "")
FillField WordDoc, "Anne3", IIf(d![JurisdictionID] = 3, "Phone No.: _____________", "")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & templateName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], templateName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_StatementOfDebtWithFiguresWells(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim i As Integer, itemCount As Integer, totalItems As Currency, amountInterest As Currency, itemsFields As String, J As Integer
Dim K As Recordset
Dim Rresult As Long
Dim IntrestTo As String
Dim IntrestFrom As String

IntrestFrom = "Interest From " & InputBox("Interest From")
IntrestTo = " To " & InputBox("Interest To")


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String

If Forms![Case List]!ClientID = 6 Then templateName = "Statement of Debt Figures Wells"
If Forms![Case List]!ClientID = 556 Then templateName = "Statement of Debt Figures Wells"

'If (D![ClientID] = 157) Then
'  TemplateName = TemplateName & " Cenlar"
'End If
'TemplateName = TemplateName & ".dot"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Statement of Debt Figures Wells.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProjectName").Select
WordDoc.Bookmarks("ProjectName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "owner", IIf(d![LoanType] = 4, "holder", IIf(d![LoanType] = 5, "holder", "owner"))
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "BorrowerNames", BorrowerNames(0)
'FillField WordDoc, "Liber", LiberFolio(D![Liber], D![Folio], D![FCdetails.State], D![JurisdictionID])
'FillField WordDoc, "LastPaymentApplied", Format$(D![DateOfDefault], "mmmm d, yyyy")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
'FillField WordDoc, "PaidStr", IIf((Nz(D![RemainingPBal], 0) > Nz(D![OriginalPBal], 0)), "Additional Interest", "Paid on principal")
'FillField WordDoc, "Paid", Format$(D!OriginalPBal - D!RemainingPBal, "Currency")
FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "BrowersNamesV2", BrowersNamesV2(d![CaseList.FileNumber], 2)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")
'FillField WordDoc, "DateAsOf", Forms![Print Statement of Debt]!txtDueDate
'FillField WordDoc, "DateEffective", Forms![Print Statement of Debt]!txtDueDate + 1
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "              ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")
'FillField WordDoc, "DateBalance", IIf(IsNull(Forms![Print Statement of Debt]!txtDueDate), "______________", Format$(Forms![Print Statement of Debt]!txtDueDate, "mmmm d"", ""yyyy"))

totalItems = 0
amountInterest = 0
itemsFields = ""
Set dd = CurrentDb.OpenRecordset("SELECT Desc,  Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber] & " ORDER BY Timestamp ASC ;", dbOpenSnapshot)
If dd.EOF Then      ' no extra lines
    FillField WordDoc, "Line_Items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
       
    
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
 
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items", itemsFields & " "
    
    dd.MoveFirst
    i = 1
    Do While Not dd.EOF
    If dd!Desc = "Interest" Then
    FillField WordDoc, "Item" & i, IntrestFrom & IntrestTo & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
   i = i + 1
   dd.MoveNext
    Else
        FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        i = i + 1
        dd.MoveNext
    End If
        'TotalItems = TotalItems + Nz(dd!Amount, 0)
'        dd.MoveNext
'        i = i + 1

    Loop

End If
 dd.Close
 
  Set K = CurrentDb.OpenRecordset("SELECT Desc, Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber], dbOpenSnapshot)
  J = 1
  Do While Not K.EOF
   If K!Desc = "Interest" Then
   amountInterest = amountInterest + Format$(Nz(K!Amount, 0), "Currency")
   End If

   totalItems = totalItems + Nz(K!Amount, 0)
   K.MoveNext
   J = J + 1
  Loop
  K.Close


FillField WordDoc, "Total", Format$(totalItems, "Currency")
FillField WordDoc, "rate", IIf(d![LoanType] = 3, "Interest will continue to accrue according to the terms of the note. ", IIf(Forms![Print Statement of Debt]!cbxRateType = "Fixed", "Per diem interest in the amount of _____ will accrue on the principal from _______ ", IIf(Forms![Print Statement of Debt]!cbxRateType = "HELOC", "A daily variable per diem will accrue on the principal in accordance with the variable rate as set forth in the Note ", "Per diem interest in the amount of _______ will accrue on the principal from ______ to the next interest rate change date and accrue thereafter in accordance with the variable rate as set forth in the Note")))
FillField WordDoc, "BalDueDate", IIf(IsNull([Forms]![Print Statement of Debt]![txtDueDate]), "______________", Format$([Forms]![Print Statement of Debt]![txtDueDate], "mmmm d"", ""yyyy"))
FillField WordDoc, "BalanceDue", Format$(d!RemainingPBal + totalItems, "Currency")
FillField WordDoc, "PerDiemInterest", IIf(IsNull(d!PerDiem), "$_____________", Format$(d!PerDiem, "Currency"))

'FillField WordDoc, "PerDiemInterest", IIf(IsNull(D!PerDiem), "$_____________", Format$(D!PerDiem, "Currency"))
'FillField WordDoc, "InterestRate", IIf(IsNull(D!InterestRate), "____________ %", Format$(D!InterestRate, "#.000%"))
'FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
'FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
'FillField WordDoc, "NotaryLocation", IIf(IsNull(D!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", D!NotaryLocation)
FillField WordDoc, "amountInterest$", Format$(amountInterest, "Currency")
FillField WordDoc, "jurisdiction", d!Jurisdiction
FillField WordDoc, "dotdate#", d![DOTdate]
'FillField WordDoc, "liberT", D![Liber]
'FillField WordDoc, "folioT", D![Folio]
FillField WordDoc, "lpidate", d!LPIDate
FillField WordDoc, "LPIdate+1", DateAdd("d", 1, d![LPIDate])
FillField WordDoc, "MortgagorNamesintext", MortgagorNamesIntext(0, 2, 0, 0, -1)
'FillField WordDoc, "lpidate#", D![LPIDate]
FillField WordDoc, "lastmonth", LastMonth(d![LPIDate])
FillField WordDoc, "Anne", IIf(d![JurisdictionID] = 3, "Address: _______________", "")
FillField WordDoc, "Anne2", IIf(d![JurisdictionID] = 3, "                " & "_______________", "")
FillField WordDoc, "Anne3", IIf(d![JurisdictionID] = 3, "Phone No.: _____________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & templateName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], templateName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_ConsentModifyingAll(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim FileNumber As Long, Judge As String, CoDebtor As Boolean, AffDate As String, AffInfo As String, DebtorsPlural As Boolean

FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBKDocsWord WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If


Select Case d!District
    Case 8, 9, 10, 11, 18
        Set WordObj = CreateObject("Word.Application")
        Set WordDoc = WordObj.Documents.Add(TemplatePath & "Consent Order Modifying Western VA.dot", False, 0, True)
        WordObj.Visible = True
    Case Else
        Set WordObj = CreateObject("Word.Application")
        Set WordDoc = WordObj.Documents.Add(TemplatePath & "Consent Order Modifying.dot", False, 0, True)
        WordObj.Visible = True
End Select


'Set WordObj = CreateObject("Word.Application")
'Set WordDoc = WordObj.Documents.Add(TemplatePath & "Consent Order Modifying.dot", False, 0, True)
'WordObj.Visible = True
'WordObj.ScreenUpdating = False


WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
'WordDoc.Bookmarks("ProjectName").Select
'WordDoc.Bookmarks("ProjectName").Range.text = D![PrimaryDefName]
'WordDoc.Bookmarks("PropertyAddress").Select
'WordDoc.Bookmarks("PropertyAddress").Range.text = D![PropertyAddress]

'WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
'WordDoc.Bookmarks("FileNumber").Range.text = D![CaseList.FileNumber]
'WordDoc.Bookmarks("AttorneyInfo").Select
'WordDoc.Bookmarks("AttorneyInfo").Range.text = "Diane Rosenberg" & vbCr & "VA Bar 35237"

FillField WordDoc, "Header", IIf(d![Districts.State] <> "VA", vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine, "") & _
    "IN THE UNITED STATES BANKRUPTCY COURT" & vbNewLine & "FOR THE " & UCase$(d!Name) & vbNewLine & d!Location
'Judge = Right$(UCase$(D!CaseNo), 3)
'CoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And D![BKdetails.Chapter] = 13)


'If Not IsNull(d![3rdAff]) Then
'    AffDate = Format$(d![3rdAff], "mmmm d, yyyy")
'Else
'    If Not IsNull(d![2ndAff]) Then
'        AffDate = Format$(d![2ndAff], "mmmm d, yyyy")
'    Else
'        If Not IsNull(d![Affidavit]) Then
'            AffDate = Format$(d![Affidavit], "mmmm d, yyyy")
'        End If
'    End If
'End If
'If AffDate = "" Then
'    AffInfo = ""
'Else
'    AffInfo = "an Affidavit of Default having been sent on " & AffDate & ", no response or funds having been received, "
'End If
DebtorsPlural = (CountNames(FileNumber, "BKDebtor = True AND (Owner=True OR Mortgagor=True)") > 1)



FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & IIf(DebtorsPlural, "s", "")

FillField WordDoc, "InvestorAddr", UCase$(d![Investor])

FillField WordDoc, "Respondents", UCase$(GetNames(0, 3, IIf(d![BKdetails.Chapter] = 13, _
"(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", "BKDebtor=True AND (Owner=True OR Mortgagor=True)"))) & _
IIf(d![BKdetails.Chapter] = 7, " and " & UCase$(Nz(d![First])) & " " & _
UCase$(Nz(d![Last])) & ", CHAPTER 7 TRUSTEE", "")


'& vbCr & RemoveLF(D!InvestorAddress)
'FillField WordDoc, "RespondentsEnd", GetAddresses(0, 4, IIf(D![BKdetails.Chapter] = 13, _
'            "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", _
'            "BKDebtor=True AND (Owner=True OR Mortgagor=True)"), vbCr) & _
'    IIf(D![BKdetails.Chapter] = 7, _
'        vbCr & vbCr & "and" & vbCr & vbCr & FormatName("", _
'                                                        UCase$(Nz(D!First)), _
'                                                        UCase$(Nz(D!Last)), _
'                                                        ", CHAPTER 7 TRUSTEE", _
'                                                        D!Address, _
'                                                        D!Address2, _
'                                                        D![BKTrustees.City], _
'                                                        D![BKTrustees.State], _
'                                                        D!Zip, _
'                                                        vbCr), _
'        "")
FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]

Judge = Right$(UCase$(d![CaseNo]), 3)
FillField WordDoc, "Caption", IIf(Judge = "RGM" And d![RealEstate], "CONSENT ORDER  AS TO " & UCase$(d![PropertyAddress]), "CONSENT ORDER MODIFYING AUTOMATIC STAY")
'IIf(D!Judge = "RGM" And D!RealEstate, "CONSENT ORDER  AS TO " & UCase$(D!PropertyAddress), "CONSENT ORDER MODIFYING AUTOMATIC STAY")

Select Case Forms![Print Consent Order Modifying]!optTrustee
    Case 1
        FillField WordDoc, "TrusteeAction", "the trustee having filed a report of no distribution, "
    Case 2
        FillField WordDoc, "TrusteeAction", "the trustee having failed to file an answer, "
    Case 0
        FillField WordDoc, "TrusteeAction", ""
    Case 3
        FillField WordDoc, "TrusteeAction", ""
End Select

FillField WordDoc, "ThisDate", IIf(d![Districts.State] <> "VA", "", ", it is this " & OrderDate())

FillField WordDoc, "District", d![Name]
Dim chCoDebtor As Boolean
 If (GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13) Then chCoDebtor = True

FillField WordDoc, "CoDebtorSection", IIf(chCoDebtor, " and 1301", "")

FillField WordDoc, "Action", IIf(d!RealEstate, "commence foreclosure proceeding " & _
IIf(d![FCdetails.State] = "MD", "in the Circuit Court for " & d!Jurisdiction & _
IIf(IsNull(d!LongState), "", ", " & d!LongState) & ", ", "") & _
"against the real property and improvements " & IIf(IsNull(d!ShortLegal), "", _
"with a legal description of """ & d!ShortLegal & """ also ") & "known as " & _
d!PropertyAddress & IIf(Len(Forms!BankruptcyDetails!sfrmPropAddr!Apt & "") = 0, "", ", " & Forms!BankruptcyDetails!sfrmPropAddr!Apt) & _
 ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & _
FormatZip(d!ZipCode) & " and to allow the successful purchaser to obtain possession of same", _
"proceed with repossession and sale of the " & d!PropertyDesc)

'FillField WordDoc, "PropertyDesc", D![PropertyDesc]

FillField WordDoc, "OrderWording", IIf(Judge = "RGM", "ORDERED that the Debtor shall:", _
"ORDERED that the above Order be and it is hereby, stayed provided that the Debtor:")

FillField WordDoc, "ConsentPaymentAmount", Format$(d!ConsentPaymentAmount, "Currency")

FillField WordDoc, "PaymentType", IIf(d!RealEstate, "mortgage", "monthly")

FillField WordDoc, "ConsentPaymentDate", Format$(d!ConsentPaymentDate, "mmmm d, yyyy")

FillField WordDoc, "NoteType", IIf(d!RealEstate, "Promissory Note secured by the " & DOTWord(d!DOT) & " on the above referenced property", d!PropertyContract)

FillField WordDoc, "ConsentPaymentInfo", d!ConsentPaymentInfo

FillField WordDoc, "2A", IIf(IsNull(d!Consent2A), "", "^p2A. " & d!Consent2A)
FillField WordDoc, "ConsentPaymentTo", d![ConsentPaymentTo]

FillField WordDoc, "AndAttorney", IIf(IsNull(d!AttorneyLastName), "", "and Debtor's attorney ")

FillField WordDoc, "RGM1", IIf(Judge = "RGM", "", ", without further order of court")

FillField WordDoc, "RGM2", IIf(Judge = "RGM", "^pIf any amount required in Paragraph 2 is not paid timely, Movant's attorney shall mail notice to the Debtor" & _
                        IIf(IsNull(d!AttorneyLastName), "", ", Debtor's attorney") & " and Chapter 13 Trustee, and shall file an Order of Termination of Automatic " & _
                        "Stay against the Subject Property described above; and be it further^p" & _
                        "ORDERED that a default in the payment of a regularly scheduled mortgage payment as listed in Paragraph 1 shall be governed by the attached addendum.", "")

If d![Districts.State] = "VA" Then
    FillField WordDoc, "JudgeSignature", "___________________________" & vbCr & "United States Bankruptcy Judge"
    FillField WordDoc, "Submitted", "Respectfully Submitted:"
    
    
    FillField WordDoc, "AttorneySignature", "Diane S. Rosenberg"
    FillField WordDoc, "End", ""
Else
    FillField WordDoc, "OrderDate", ""
    FillField WordDoc, "JudgeSignature", ""
    FillField WordDoc, "Submitted", ""
  '  FillField WordDoc, "ElectronicSignature", ""
    FillField WordDoc, "AttorneySignature", ""
    FillField WordDoc, "End", "End of Order"
End If

FillField WordDoc, "ElectronicSignature", IIf([Forms]![BankruptcyPrint]![chElectronicSignature], "/s/ " & [Forms]![BankruptcyPrint]![cbxAttorney], "")
FillField WordDoc, "ElectronicSignature2", IIf([Forms]![BankruptcyPrint]![chElectronicSignature], IIf(IsNull(d![LastName]), DebtorNames(0, 5), "/s/ " & d![FirstName] & " " & d![LastName]), "")

FillField WordDoc, "AttMovment", [Forms]![BankruptcyPrint]![cbxAttorney]
FillField WordDoc, "AttDebtor", d![FirstName] & " " & d![LastName]
FillField WordDoc, "AttSign", IIf(Forms![Print Consent Order Modifying]!optTrustee = 3 And _
Forms!BankruptcyPrint!chElectronicSignature, "/s/ " & d![First] & " " & d![Last], "")
FillField WordDoc, "AttSignb", IIf(Forms![Print Consent Order Modifying]!optTrustee = 3, "_______________________" & vbCr & d![First] & " " & d![Last] & vbCr & "Chapter 13 Trustee", "")

FillField WordDoc, "ATTpage4", [Forms]![BankruptcyPrint]![cbxAttorney] & ", Esquire"
FillField WordDoc, "Attpage4Address", FirmAddress()
FillField WordDoc, "AttFirmname", IIf(IsNull(d![LastName]), "", d![FirstName] & " " & d![LastName] & ", Esquire " & vbCr & IIf(IsNull(d![BKAttorneyFirm]), "", d![BKAttorneyFirm] & vbCr) & d![BKAttorneyAddress])
FillField WordDoc, "Truss", FormatName("", d![First], d![Last], ", Trustee", d![Address], d![Address2], d![BKTrustees.City], d![BKTrustees.State], d![Zip])
FillField WordDoc, "AttFirm", IIf(IsNull(d![AttorneyLastName]), "", d![AttorneyFirstName] & " " & d![AttorneyLastName] & ", Esquire")
'FillField WordDoc, "AttFirmname", IIf(IsNull(D![BKAttorneyFirm]), "", D![BKAttorneyFirm] & "")
FillField WordDoc, "AttFirmAddress", d![BKAttorneyAddress]
FillField WordDoc, "BKService", BKService(0)
FillField WordDoc, "Copy", IIf(Forms!BankruptcyPrint!chElectronicSignature And d![Districts.State] <> "VA", _
"I HEREBY CERTIFY that the terms of the copy of the consent order submitted to the Court are identical to those set forth in the original consent order; and the signatures represented by the /s/__________ on this copy reference the signatures of consenting parties on the original consent order.", "")
FillField WordDoc, "lastSignature", [Forms]![BankruptcyPrint]![cbxAttorney]

WordObj.Selection.HomeKey wdStory, wdMove
WordObj.ScreenUpdating = True
WordDoc.SaveAs EMailPath & "Consent Order Modifying.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Consent Order Modifying.dot")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_DebtBK(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, R As Recordset, B As Recordset
Dim p As String, C As String, TotalItemsB As Integer, ItemsFieldsB As String, totalItems As Currency
Dim totalA As Integer, totalB As Integer

p = "pre"
C = "post"
Dim FileNumber As Long, Judge As String, CoDebtor As Boolean, AffDate As String, AffInfo As String, DebtorsPlural As Boolean, itemsFields As String
Dim itemCount As Integer, i As Integer, J As Integer

FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBKDocsWord WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)



FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBKDocsWord WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If



Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Consent Order Modifying Debt.dot", False, 0, True)
WordObj.Visible = True


WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "Header", IIf(d![Districts.State] <> "VA", vbNewLine, "") & _
    "IN THE UNITED STATES BANKRUPTCY COURT" & vbNewLine & "FOR THE " & UCase$(d!Name) & vbNewLine & d!Location


FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & IIf(DebtorsPlural, "s", "")

FillField WordDoc, "InvestorAddr", UCase$(d![Investor])

FillField WordDoc, "Respondents", UCase$(GetNames(0, 3, IIf(d![BKdetails.Chapter] = 13, _
"(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", "BKDebtor=True AND (Owner=True OR Mortgagor=True)"))) & _
IIf(d![BKdetails.Chapter] = 7, " and " & UCase$(Nz(d![First])) & " " & _
UCase$(Nz(d![Last])) & ", CHAPTER 7 TRUSTEE", "")


FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]

Judge = Right$(UCase$(d![CaseNo]), 3)
FillField WordDoc, "Caption", "EXHIBIT A -- DEBT"
totalA = 0
totalItems = 0
itemsFields = ""
Set R = CurrentDb.OpenRecordset("Select * FROM  BKDebt WHERE FileNumber=" & FileNumber & " AND BKDebt.PrePost = '" & p & "' ORDER BY Timestamp;", dbOpenSnapshot)
If R.EOF Then   'no extra lines
FillField WordDoc, "Line_Items", ""
Else
  R.MoveLast
  itemCount = R.RecordCount
  For i = 1 To itemCount
  itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
  Next i
  itemsFields = Left$(itemsFields, Len(itemsFields) - 1)
  FillField WordDoc, "Line_Items", itemsFields
  R.MoveFirst
  i = 1
  Do While Not R.EOF
  FillField WordDoc, "Item" & i, R!Desc & vbTab & Format$(Nz(R!Amount, 0), "Currency")
  totalItems = totalItems + Nz(R!Amount, 0)
  R.MoveNext
  i = i + 1
  Loop
    R.MoveFirst
    For J = 1 To R.RecordCount
    totalA = totalA + R!Amount
    R.MoveNext
    Next J
End If

R.Close
FillField WordDoc, "total", Format$(Nz(totalA, 0), "Currency")


totalB = 0
TotalItemsB = 0
ItemsFieldsB = ""
Set B = CurrentDb.OpenRecordset("Select * FROM  BKDebt WHERE FileNumber=" & FileNumber & " AND BKDebt.PrePost = '" & C & "' ORDER BY Timestamp;", dbOpenSnapshot)
If B.EOF Then   'no extra lines
FillField WordDoc, "Line_ItemsB", ""
Else
  B.MoveLast
  itemCount = B.RecordCount
  For i = 1 To itemCount
  ItemsFieldsB = ItemsFieldsB & "<<ItemB" & i & ">>" & vbCr
  Next i
  ItemsFieldsB = Left$(ItemsFieldsB, Len(ItemsFieldsB) - 1)
  FillField WordDoc, "Line_ItemsB", ItemsFieldsB
  B.MoveFirst
  i = 1
  Do While Not B.EOF
  FillField WordDoc, "ItemB" & i, B!Desc & vbTab & Format$(Nz(B!Amount, 0), "Currency")
  TotalItemsB = TotalItemsB + Nz(B!Amount, 0)
  B.MoveNext
  i = i + 1
  Loop
  B.MoveFirst
  For J = 1 To B.RecordCount
   totalB = totalB + B!Amount
    B.MoveNext
    Next J
End If
B.Close
FillField WordDoc, "totalB", Format$(Nz(totalB, 0), "Currency")

FillField WordDoc, "totalAll", Format$(Nz(CLng(totalA) + CLng(totalB), 0), "Currency")
  
  
WordObj.Selection.HomeKey wdStory, wdMove
WordObj.ScreenUpdating = True
WordDoc.SaveAs EMailPath & "Consent Order Modifying Debt.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Debt.dot")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_CHAMCoverSheet(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")

'FillField WordDoc, "DocType", IIf([Forms]![ForeclosurePrint]![chDeedOfApp] = True, "Substitution of Trustee", IIf([Forms]![ForeclosurePrint]![chMilitaryAffidavitActive] = True, _
'"Active Duty Military Affidavit", IIf([Forms]![ForeclosurePrint]![chMilitaryAffidavit] = True, _
'"Military Affidavit", IIf([Forms]![ForeclosurePrint]![chMilitaryAffidavitNoSSN] = True, "Military Affidavit No SSN"))))

If Forms!Foreclosureprint!chDeedOfApp = True Then
    FillField WordDoc, "DocType", "Substitution of Trustee"
ElseIf Forms!Foreclosureprint!chMilitaryAffidavitActive = True Then
    FillField WordDoc, "DocType", "Active Duty Military Affidavit"
ElseIf Forms!Foreclosureprint!chMilitaryAffidavit = True Then
    FillField WordDoc, "DocType", "Military Affidavit"
ElseIf Forms!Foreclosureprint!chMilitaryAffidavitNoSSN = True Then
    FillField WordDoc, "DocType", "Military Affidavit No SSN"
ElseIf Forms!Foreclosureprint!chSOD2 = True Then
    FillField WordDoc, "DocType", "Statement of Debt with Figures"
ElseIf Forms!Foreclosureprint!chSOD = True Then
    FillField WordDoc, "DocType", "Statement of Debt"
ElseIf Forms!Foreclosureprint!chLossMitPrelim = True Then
    FillField WordDoc, "DocType", "Loss Mitigation - Preliminary"
ElseIf Forms!Foreclosureprint!chLossMitFinal = True Then
    FillField WordDoc, "DocType", "Loss Mitigation - Final"
ElseIf Forms!Foreclosureprint!chNoteOwnership = True Then
    FillField WordDoc, "DocType", "Affidavit of Note Ownership"
ElseIf Forms!Foreclosureprint!chAffMD7105 = True Then
    FillField WordDoc, "DocType", "NOI Affidavit"

End If

FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_DILJudgmentAffidavit(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryDILword WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Do While d.EOF = False
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Deed in Lieu - Judgement Affidavit.dot", False, 0, True)
    WordObj.Visible = True
    
    WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
    WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
    
    FillField WordDoc, "First", d!First
    FillField WordDoc, "Last", d!Last
      
    If d!Leasehold = 1 Then
        FillField WordDoc, "Leasehold", "Leasehold"
    Else
        FillField WordDoc, "Leasehold", "Fee Simple"
    End If
    
    FillField WordDoc, "PropertyAddress", d!PropertyAddress
    FillField WordDoc, "City", d!City
    FillField WordDoc, "State", d!LongState
    FillField WordDoc, "ZipCode", FormatZip(d!ZipCode)
    d.MoveNext
Loop

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Deed in Lieu - Judgement Affidavit.dot"

'Why do we need a function to save a word document, the save feature is BUILT in...
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Deed in Lieu - Judgement Affidavit.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
Set d = Nothing
End Sub   '************************* This Sub handles the Word Doc for DIL Judgment Affidavit



Public Sub Doc_DILBorrowerLetter(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim backbydate As String
Dim backbydate2 As String
Dim mortgagors As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryfcdocswordlite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

backbydate = Nz(InputBox("Enter a back by date", "Back by Date"), "")
Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Deed in Lieu - Borrower Letter.dot", False, 0, True)

WordObj.Visible = True

'WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
'WordDoc.Bookmarks("FileNumber").Range.text = D![CaseList.FileNumber]
If backbydate = "" Then
Else
backbydate2 = Format(DateAdd("d", 7, backbydate), "mmmm d, yyyy")
End If

mortgagors = MortgagorOwnerNames(0, 2)

FillField WordDoc, "backbydate", Format(backbydate, "mmmm d, yyyy")
FillField WordDoc, "backbydate2", backbydate2

FillField WordDoc, "date", Date
FillField WordDoc, "Mortgagors", mortgagors
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!LongState
FillField WordDoc, "ZipCode", FormatZip(d!ZipCode)
FillField WordDoc, "Filenumber", d![CaseList.FileNumber]
FillField WordDoc, "LoginName", GetLoginName
FillField WordDoc, "STaffEmail", GetStaffEmail

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Deed in Lieu - Borrower Letter.dot"

'Why do we need a function to save a word document, the save feature is BUILT in...
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Deed in Lieu - Borrower Letter.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
Set d = Nothing
End Sub   '************************* This Sub handles the Word Doc for DIL Letter to Borrower

Public Sub Doc_DIL(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim investorAdd As String
Dim mortgagors As String
Dim sql As String
Dim dotDateFormatted As String
Dim origpBalFormat As String
Dim dotrecordedFormat As String
Dim mortgagorLine As String
Dim recordedInput As String
Dim LiberInput As String
Dim FolioInput As String
Dim consider As String
Dim assess As String
Dim marylandStr As String
Dim zipper As String

 
    Set d = CurrentDb.OpenRecordset("SELECT * FROM qryDilWord2 WHERE [CaseList.FileNumber] = " & Forms![Case List]!FileNumber, dbOpenSnapshot)
    If d.EOF Then
        MsgBox "No data found"
        Exit Sub
    End If
    
    If d!LongState = "Virginia" Then
        consider = Format(InputBox("Please enter a Consideration Amount:", "Deed in Lieu of Foreclosure - CONSIDERATION"), "Currency")
        assess = Format(InputBox("Please enter an Assessed Value Amount:", "Deed in Lieu of Foreclosure - ASSESSSED VALUE"), "Currency")
    Else ' skip it
    End If
    
    recordedInput = Format(InputBox("Deed Recorded on:", "Deed in Lieu - DATE RECORDED "), "mmmm d, yyyy")
    LiberInput = InputBox("Please enter a Liber Number: ", "Deed in Lieu - LIBER NUMBER ")
    FolioInput = InputBox("Please enter a Folio Number: ", "Deed in Lieu - FOLIO NUMBER ")
     
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Deed in Lieu.dot", False, 0, True)
    
    WordObj.Visible = True
    
    'Could/Should combine these into 1 line below
    If Not IsNull([D!InvestorAddress]) Then
          investorAdd = OneLine(d!InvestorAddress)
    Else
        investorAdd = ""
    End If
    If Not IsNull([D!DOTdate]) Then
        dotDateFormatted = Format(d!DOTdate, "mmmm d, yyyy")
    Else
        dotDateFormatted = ""
    End If
    If Not IsNull([D!OriginalPBal]) Then
         origpBalFormat = Format(d!OriginalPBal, "Currency")
    Else
        origpBalFormat = ""
    End If
    If Not IsNull([D!DOTrecorded]) Then
           dotrecordedFormat = Format(d!DOTrecorded, "mmmm d, yyyy")
    Else
        dotrecordedFormat = ""
    End If
        
    mortgagors = MortgagorOwnerNames(0, 2)
          
    If d!LongState = "Virginia" Then
        FillField WordDoc, "Consideration", "CONSIDERATION: " & consider
        FillField WordDoc, "AssessedValue", "ASSESSED VALUE: " & assess
        FillField WordDoc, "TaxID", "TAX ID NO.: " & IIf(IsNull(d!TaxID), "", d!TaxID)
       
        If Not IsNull(Forms!Foreclosureprint!Attorney.Column(0)) Then
            FillField WordDoc, "Deed", "Deed Prepared by:"
            FillField WordDoc, "Atty", GetStaffFullName(Forms!Foreclosureprint!Attorney.Column(0))
            FillField WordDoc, "Bar", FetchbarNumberPoundSign(Forms!Foreclosureprint!Attorney, d![FCdetails.State])
        Else
            FillField WordDoc, "Deed", ""
            FillField WordDoc, "Atty", ""
            FillField WordDoc, "Bar", ""
        End If
        'FillField WordDoc, "Atty", GetStaffFullName(Forms!ForeclosurePrint!Attorney.Column(0))
        'FillField WordDoc, "Bar", FetchBarNumberPoundSign(Forms!ForeclosurePrint!Attorney, D![FCdetails.State])
        FillField WordDoc, "Maryland", ""
    Else
        FillField WordDoc, "Consideration", ""
        FillField WordDoc, "AssessedValue", ""
        FillField WordDoc, "TaxID", ""
        FillField WordDoc, "Deed", ""
        FillField WordDoc, "Atty", ""
        FillField WordDoc, "Bar", ""
        marylandStr = "This instrument was prepared under the supervision of an attorney admitted to practice before the Court of Appeals of Maryland."
        marylandStr = marylandStr + vbCrLf
        marylandStr = marylandStr & "                                                                                                            __________________________"
        FillField WordDoc, "Maryland", marylandStr
    End If
    
    FillField WordDoc, "Mortgagors", mortgagors
    FillField WordDoc, "Investor", IIf(IsNull(d!Investor), "", d!Investor)
    FillField WordDoc, "InvestorAddress", IIf(IsNull(d!PropertyAddress), "", d!PropertyAddress)
    FillField WordDoc, "dotdate", dotDateFormatted
    FillField WordDoc, "OriginalBeneficiary", IIf(IsNull(d!OriginalBeneficiary), "", d!OriginalBeneficiary)
    FillField WordDoc, "OriginalPBal", origpBalFormat
    FillField WordDoc, "dotrecorded", dotrecordedFormat
    FillField WordDoc, "liber", IIf(IsNull(d!Liber), "", d!Liber)
    FillField WordDoc, "Folio", IIf(IsNull(d!Folio), "", d!Folio)
    FillField WordDoc, "Jurisdiction", IIf(IsNull(d!Jurisdiction), "", d!Jurisdiction)
    FillField WordDoc, "State", IIf(IsNull(d!LongState), "", d!LongState)
        
    If Not IsNull([D!Leasehold]) Then
        If d!Leasehold = 1 Then
            FillField WordDoc, "Leasehold", "Leasehold with an annual ground rent of " & Format(d![GroundRentAmount], "Currency")
        Else
            FillField WordDoc, "Leasehold", "Fee Simple"
        End If
    Else
        FillField WordDoc, "Leasehold", ""
    End If
    If Not IsNull([D!Zipcode]) Then
        zipper = FormatZip(d!ZipCode)
    Else
        zipper = ""
    End If
    FillField WordDoc, "Legal", IIf(IsNull(d!LegalDescription), "", d!LegalDescription)
    FillField WordDoc, "TaxID", IIf(IsNull(d!TaxID), "", d!TaxID)
    FillField WordDoc, "PropertyAddress", IIf(IsNull(d!PropertyAddress), "", d!PropertyAddress)
    FillField WordDoc, "City", IIf(IsNull(d!City), "", d!City)
    FillField WordDoc, "ZipCode", zipper
    FillField WordDoc, "RecordedInput", recordedInput
    FillField WordDoc, "LiberInput", LiberInput
    FillField WordDoc, "FolioInput", FolioInput
    FillField WordDoc, "MortgagorsLines", MortgagorOwnerNames(0, 4)
     
    FillField WordDoc, "Filenumber", IIf(IsNull(d![CaseList.FileNumber]), "", d![CaseList.FileNumber])
    FillField WordDoc, "LoginName", GetLoginName
    FillField WordDoc, "STaffEmail", GetStaffEmail
  
    WordObj.Selection.HomeKey wdStory, wdMove
    WordDoc.SaveAs EMailPath & "Deed in Lieu.dot"
    
    'Why do we need a function to save a word document, the save feature is BUILT in...
    Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Deed in Lieu.doc")
    If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
    Set WordObj = Nothing
    d.Close
    Set d = Nothing
  
End Sub   '************************* This Sub handles the Word Doc for DIL (Main)

Public Sub Doc_DILCertificate(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim investorAdd As String
Dim mortgagors As String
Dim sql As String
Dim dotDateFormatted As String
Dim origpBalFormat As String
Dim dotrecordedFormat As String
Dim mortgagorLine As String
Dim recordedInput As String
Dim vacateDate As String
Dim zipper As String

    Set d = CurrentDb.OpenRecordset("SELECT * FROM qryDilCert WHERE [CaseList.FileNumber] = " & Forms![Case List]!FileNumber, dbOpenSnapshot)
    
    If d.EOF Then
        MsgBox "No data found"
        Exit Sub
    End If
    
    vacateDate = Format(InputBox("Please enter a Vacate by Date:", "Deed in Lieu - VACATE DATE "), "mmmm d, yyyy")
    If Not IsNull([D!InvestorAddress]) Then
          investorAdd = OneLine(d!InvestorAddress)
    Else
        investorAdd = ""
    End If
    If Not IsNull([D!DOTdate]) Then
        dotDateFormatted = Format(d!DOTdate, "mmmm d, yyyy")
    Else
        dotDateFormatted = ""
    End If
    If Not IsNull([D!OriginalPBal]) Then
        origpBalFormat = Format(d!OriginalPBal, "Currency")
    Else
       origpBalFormat = ""
    End If
    If Not IsNull([D!DOTrecorded]) Then
       dotrecordedFormat = Format(d!DOTrecorded, "mmmm d, yyyy")
    Else
       dotrecordedFormat = ""
    End If
    If Not IsNull([D!Zipcode]) Then
        zipper = FormatZip(d!ZipCode)
    Else
        zipper = ""
    End If
       
     
    Do While d.EOF = False
        Set WordObj = CreateObject("Word.Application")
        Set WordDoc = WordObj.Documents.Add(TemplatePath & "Deed in Lieu - Certificate.dot", False, 0, True)
        
        WordObj.Visible = True
        mortgagors = MortgagorOwnerNames(0, 2)
                      
        FillField WordDoc, "Mortgagors", mortgagors
        FillField WordDoc, "Investor", IIf(IsNull(d!Investor), "", d!Investor)
        FillField WordDoc, "PropertyAddress", IIf(IsNull(d!PropertyAddress), "", d!PropertyAddress)
        FillField WordDoc, "City", IIf(IsNull(d!City), "", d!City)
        FillField WordDoc, "State", IIf(IsNull(d!LongState), "", d!LongState)
        FillField WordDoc, "ZipCode", zipper
        FillField WordDoc, "dotdate", dotDateFormatted
        FillField WordDoc, "OriginalBeneficiary", IIf(IsNull(d!OriginalBeneficiary), "", d!OriginalBeneficiary)
        FillField WordDoc, "OriginalPBal", origpBalFormat
        FillField WordDoc, "liber", IIf(IsNull(d!Liber), "", d!Liber)
        FillField WordDoc, "Folio", IIf(IsNull(d!Folio), "", d!Folio)
        FillField WordDoc, "Jurisdiction", IIf(IsNull(d!Jurisdiction), "", d!Jurisdiction)
        FillField WordDoc, "vacateDate", vacateDate
        FillField WordDoc, "First", IIf(IsNull(d!First), "", d!First)
        FillField WordDoc, "Last", IIf(IsNull(d!Last), "", d!Last)
        FillField WordDoc, "Filenumber", IIf(IsNull(d![CaseList.FileNumber]), "", d![CaseList.FileNumber])
    d.MoveNext
    Loop
  
    WordObj.Selection.HomeKey wdStory, wdMove
    WordDoc.SaveAs EMailPath & "Deed in Lieu - Certificate.dot"
    
    'Why do we need a function to save a word document, the save feature is BUILT in...
    Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Deed in Lieu - Certificate.doc")
    If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
    Set WordObj = Nothing
    d.Close
    Set d = Nothing
  
End Sub   '************************* This Sub handles the Word Doc for DIL (Main)

Public Sub Doc_DeedofAppointmentOcwen(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
Dim templateName As String
Dim BankName As String

'If Forms![foreclosureDetails]!State = "MD" Then
templateName = "Deed of Appointment Ocwen"
'Else
'TemplateName = "Deed of Appointment WellsVA"
'End If

'If Forms![Case List]!ClientID = 6 Then BankName = "Deed of Appointment-Wells Fargo Bank NA"
'If Forms![Case List]!ClientID = 556 Then BankName = "Deed of Appointment-Wells Fargo Home Mortgage"
BankName = "Deed of Appointment-Ocwen"

templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")

Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)

WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


FillField WordDoc, "bankName", d!LongClientName
FillField WordDoc, "BankAddress", d!BankStreet
FillField WordDoc, "BankCity", d!BankCity
FillField WordDoc, "BankState", d!BankState
FillField WordDoc, "BankZipCod", FormatZip(d!BankZipcode)
FillField WordDoc, "Tax", d!TaxID
FillField WordDoc, "Addy", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "DOTdate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "originaltrustee", d!OriginalTrustee
FillField WordDoc, "OriginalBeneficiary", d!OriginalBeneficiary
FillField WordDoc, "DOTrecorded", Format$(d!DOTrecorded, "mmmm d, yyyy")

FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "JurisdictionMD", IIf(d![Jurisdiction] = "Baltimore City", "City of Baltimore", "County of " & d![Jurisdiction])
FillField WordDoc, "Folio", d!Folio
FillField WordDoc, "Liber", d!Liber

FillField WordDoc, "N/AFoilo", IIf(IsNull(d![Folio]), ", as Instrument No. " & d![Liber] & ", in book N/A, page N/A", ", as Instrument No. N/A, in book " & d![Liber] & ", " & "page " & d![Folio])
FillField WordDoc, "FCdetails.State", d![FCdetails.State]
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "TrusteeNames", IIf(Forms![foreclosuredetails]!State = "MD", trusteeNames(0, 2), "Commonwealth Trustees, LLC")
FillField WordDoc, "investor", d![Investor]
FillField WordDoc, "investorMD", IIf(d![Investor] <> "Wells Fargo Bank, N.A.", "authorized agent of " & d![Investor], d![Investor])
FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "MERS", IIf(d![FCdetails.State] = "MD" And d![MERS] = -1, "Mortgage Electronic Registration Systems Inc. (MERS) solely as nominee for " & d![OriginalBeneficiary], "")
'FillField WordDoc, "trusteeType", IIf(Forms!foreclosureDetails!Option221, "Successor Trustee ", "Substitute Trustee ")
FillField WordDoc, "trusteeList", getTrustees(Forms!foreclosuredetails!lstTrustees, Forms!foreclosuredetails!lstTrustees)

FillField WordDoc, "Text27", d!PropertyAddress & ", " & d!City & ", " & d![FCdetails.State] & " " & FormatZip(d!ZipCode)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 226, "")
'If Page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 464, ""), 450, 4000, True)

FillField WordDoc, "filenumber", d![CaseList.FileNumber]
WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & BankName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], BankName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_NationstarCoverSheet(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Nationstar Cover Letter.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, "No Date Selected")


'FillField WordDoc, "MortgagorNameT", MortgagorNames(0, 3)

FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName()

FillField WordDoc, "Date", Format$(Date, "mmmm d, yyyy")
FillField WordDoc, "PrimaryFirst", Forms!foreclosuredetails!PrimaryFirstName
FillField WordDoc, "PrimaryLast", Forms!foreclosuredetails!PrimaryLastName
FillField WordDoc, "SecondaryFirst", IIf(IsNull(Forms!foreclosuredetails!SecondaryFirstName), "", Forms!foreclosuredetails!SecondaryFirstName & ",")
FillField WordDoc, "SecondaryLast", IIf(IsNull(Forms!foreclosuredetails!SecondaryLastName), "", Forms!foreclosuredetails!SecondaryLastName)
'FillField WordDoc, "Comma", IIf(IsNull([SecondaryFirstName]), "", ", ")
FillField WordDoc, "FirmAddress", IIf(d![FCdetails.State] = "VA", "8601 Westwood Center Dr. Suite 255, Vienna, VA 22182", "7910 Woodmont Avenue, Suite 750, Bethesda, MD 20814")
FillField WordDoc, "ATTN", IIf(d![FCdetails.State] = "VA", "VA Sale Setting team", "Edwin Orellana")
'
' If Application.CurrentProject.AllForms("Prior Servicer").IsLoaded Then
'    DoCmd.Close acForm, "Prior Servicer"
' End If

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Nationstar Cover Sheet.doc"
'WordDoc.SaveAs EMailPath & "Nationstar Cover Sheet.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Nationstar Cover Sheet.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

  
End Sub


Public Sub Doc_DeedofAppointmentMDHC(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Deed of Appointment MDHC"


templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")

Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


If d!JurisdictionID = 56 And d![FCdetails.State] = "VA" Then
    FillField WordDoc, "Tax", "Tax Map Number:"
Else
    FillField WordDoc, "Tax", "Tax ID #:"
End If


FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "TrusteeNames", trusteeNames(0, 2)
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "MortgagorNames", IIf(IsNull(d!OriginalMortgagors), MortgagorNamesCaps(0, 2, 2), d!OriginalMortgagors)
FillField WordDoc, "AssumedBy", IIf(IsNull(d!OriginalMortgagors), "", "assumed by " & MortgagorNamesCaps(0, 2, 2)) & " "
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "LongState", Nz(d!LongState)
FillField WordDoc, "LiberFolio", LiberFolio(d!Liber, d!Folio, d![FCdetails.State], d![JurisdictionID])
FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

If (d![FCdetails.State] = "VA") Then
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", IIf(IsNull(d![Folio2]), ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & _
    " at Instrument Number " & d![Liber2], ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Book " & d![Liber2] & ", Page " & d![Folio2]))
ElseIf d!JurisdictionID = 6 Then  'Calvert County
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Book " & d![Liber2] & ", Page " & d![Folio2])
Else 'MD or DC
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Liber " & d![Liber2] & ", Folio " & d![Folio2])
End If


FillField WordDoc, "TaxID", d!TaxID

FillField WordDoc, "OriginalTrustee", d!OriginalTrustee
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "OriginalBeneficiary", d!OriginalBeneficiary
FillField WordDoc, "MDInstrPrepared", IIf(d![FCdetails.State] = "VA", "", "This instrument was prepared under the supervision of " & d!AttorneyName & ", an attorney admitted to practice before the Court of Appeals of Maryland.")
FillField WordDoc, "MDSignLine", IIf(d![FCdetails.State] = "VA", "", "_________________________________")
FillField WordDoc, "MDAttorney", IIf(d![FCdetails.State] = "VA", "", d!AttorneyName)
FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & d!Investor
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "NotaryName", FetchNotaryName(Forms!Foreclosureprint!NotaryID, False)
FillField WordDoc, "FirmAddress", IIf(d![FCdetails.State] = "VA", "Commonwealth Trustees, LLC" & vbCr & "c/o Rosenberg & Associates, LLC" & vbCr & "8601 Westwood Center Drive, Suite 255" & vbCr & "Vienna, VA 22182", FirmAddress(vbCr))
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 226, "")
FillField WordDoc, "boa1", IIf(d![ClientID] = 446, "", "and, being duly sworn")
FillField WordDoc, "boa2", IIf(d![ClientID] = 446, "National Association.", "corporation.")
FillField WordDoc, "MERS", IIf(d![FCdetails.State] = "MD" And d![MERS] = -1, "Mortgage Electronic Registration Systems Inc. (MERS) solely as nominee for ", "")
FillField WordDoc, "FinalLanguage", "WHEREAS, " & _
IIf(Forms!foreclosuredetails!LoanType = 5, "Federal Home Loan Mortgage Corporation is the owner of the note secured by said Deed of Trust and appointed the party of the first part with authority to hold, collect and enforce the note; and ", _
IIf(Forms!foreclosuredetails!LoanType = 4, "Federal National Mortgage Association is the owner of the note secured by said Deed of Trust and appointed the party of the first part with authority to hold, collect and enforce the note; and ", _
 "the party of the first part is the holder of the Note secured by said Deed of Trust; and,"))

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Deed of Appointment MDHC.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Deed of Appointment MDHC.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close


End Sub

Public Sub Doc_DeedofAppointmentMT(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Deed of Appointment M&T"


templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")

Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


If d!JurisdictionID = 56 And d![FCdetails.State] = "VA" Then
    FillField WordDoc, "Tax", "Tax Map Number:"
Else
    FillField WordDoc, "Tax", "Tax ID #:"
End If


FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "TrusteeNames", trusteeNames(0, 2)
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "MortgagorNames", IIf(IsNull(d!OriginalMortgagors), MortgagorNamesCaps(0, 2, 2), d!OriginalMortgagors)
FillField WordDoc, "AssumedBy", IIf(IsNull(d!OriginalMortgagors), "", "assumed by " & MortgagorNamesCaps(0, 2, 2)) & " "
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "LongState", Nz(d!LongState)
FillField WordDoc, "LiberFolio", LiberFolio(d!Liber, d!Folio, d![FCdetails.State], d![JurisdictionID])
FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

If (d![FCdetails.State] = "VA") Then
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", IIf(IsNull(d![Folio2]), ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & _
    " at Instrument Number " & d![Liber2], ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Book " & d![Liber2] & ", Page " & d![Folio2]))
ElseIf d!JurisdictionID = 6 Then  'Calvert County
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Book " & d![Liber2] & ", Page " & d![Folio2])
Else 'MD or DC
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Liber " & d![Liber2] & ", Folio " & d![Folio2])
End If

FillField WordDoc, "TaxID", d!TaxID

FillField WordDoc, "OriginalTrustee", d!OriginalTrustee
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "OriginalBeneficiary", d!OriginalBeneficiary
FillField WordDoc, "MDInstrPrepared", IIf(d![FCdetails.State] = "VA", "", "This instrument was prepared under the supervision of " & d!AttorneyName & ", an attorney admitted to practice before the Court of Appeals of Maryland.")
FillField WordDoc, "MDSignLine", IIf(d![FCdetails.State] = "VA", "", "_________________________________")
FillField WordDoc, "MDAttorney", IIf(d![FCdetails.State] = "VA", "", d!AttorneyName)
FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & d!Investor
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "NotaryName", FetchNotaryName(Forms!Foreclosureprint!NotaryID, False)
FillField WordDoc, "FirmAddress", IIf(d![FCdetails.State] = "VA", "Commonwealth Trustees, LLC" & vbCr & "c/o Rosenberg & Associates, LLC" & vbCr & "8601 Westwood Center Drive, Suite 255" & vbCr & "Vienna, VA 22182", FirmAddress(vbCr))
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 226, "")
FillField WordDoc, "boa1", IIf(d![ClientID] = 446, "", "and, being duly sworn")
FillField WordDoc, "boa2", IIf(d![ClientID] = 446, "National Association.", "corporation.")
FillField WordDoc, "MERS", IIf(d![FCdetails.State] = "MD" And d![MERS] = -1, "Mortgage Electronic Registration Systems Inc. (MERS) solely as nominee for ", "")
FillField WordDoc, "FinalLanguage", "WHEREAS, " & _
IIf(Forms!foreclosuredetails!LoanType = 5, "Federal Home Loan Mortgage Corporation is the owner of the note secured by said Deed of Trust and appointed the party of the first part with authority to hold, collect and enforce the note; and ", _
IIf(Forms!foreclosuredetails!LoanType = 4, "Federal National Mortgage Association is the owner of the note secured by said Deed of Trust and appointed the party of the first part with authority to hold, collect and enforce the note; and ", _
 "the party of the first part is the holder of the Note secured by said Deed of Trust; and,"))

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Deed of Appointment M&T.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Deed of Appointment M&T.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_LossMitigationPrelimMDHCP(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim textfinal As String
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
Dim templateName As String
Dim BankName As String




If d![ClientID] = 456 Then
BankName = "Loss Mitigation Preliminary M&T Bank"
templateName = "Loss Mitigation - Prelim MT"

ElseIf d![ClientID] = 531 Then
BankName = "Loss Mitigation Preliminary MDHC"
templateName = "Loss Mitigation - Prelim MDHCP"
Else
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "fillingdate", IIf(Not IsNull(d![Docket]), Format(d![Docket], "mm/d/yyyy"), "_____________________ ")
FillField WordDoc, "borrower", MortgagorNamesOneline(d![CaseList.FileNumber], 2)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", IIf(Not IsNull(d![CourtCaseNumber]), d![CourtCaseNumber], " _____________________ ")
FillField WordDoc, "NameAffianit", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "Print Name and Title of Affiant", "Name")
FillField WordDoc, "TitleAffiaitOrInvestor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "Title of Affiant")
FillField WordDoc, "Line", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "______________________________")
FillField WordDoc, "Investor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, d!Investor, "")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1484, "")
'FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ____________________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "                " & "____________________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: _________________________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & BankName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], BankName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_LossMitigationPrelimNationStar(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim textfinal As String
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
'Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsLite WHERE CaseList.FileNumber=" & Forms![Case list]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
Dim templateName As String
Dim BankName As String


If d![ClientID] = 385 And d![FCdetails.State] = "MD" Then
BankName = "Nationstar Mortgage LLC"

templateName = "Loss Mitigation Prelim Nation star"
'
'ElseIf d![ClientID] = 531 Then
'BankName = "Loss Mitigation Preliminary MDHC"
'templateName = "Loss Mitigation - Prelim Nationstar"
'Else
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "fillingdate", IIf(Not IsNull(d![Docket]), Format(d![Docket], "mm/d/yyyy"), "_____________________ ")
FillField WordDoc, "borrower", MortgagorNamesOneline(d![CaseList.FileNumber], 2)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", IIf(Not IsNull(d![CourtCaseNumber]), d![CourtCaseNumber], " _____________________ ")
'FillField WordDoc, "NameAffianit", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "Print Name and Title of Affiant", "Name")    'Mei
'FillField WordDoc, "TitleAffiaitOrInvestor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "Title of Affiant")  'Mei
'FillField WordDoc, "Line", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "______________________________")  'Mei
'FillField WordDoc, "Investor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, d!Investor, "")   'Mei
FillField WordDoc, "Investor", IIf(d![ClientID] = 385 And d![Investor] = "Nationstar Mortgage LLC", "          Nationstar maintains records for the loan that is secured by the mortgage or deed of trust being foreclosed in this action. ", _
"          Nationstar services and maintains records on behalf of " & d![Investor] & ", the secured party to the mortgage or deed of trust being foreclosed in this action.")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1484, "")

FillField WordDoc, "Prior Servicer", IIf(Forms![Prior Servicer]!ChPrior, "           Before the servicing of this loan transferred to Nationstar, " & [Forms]![Prior Servicer]![TxtPriorServicer] & _
" (Prior Servicer) was the servicer for the loan and it maintained the loan servicing records.  When Nationstar began servicing this loan, Prior Servicer's records for the loan were integrated and boarded into Nationstar's systems," & _
" such that Prior Servicer's records, including the collateral file, payment histories, communication logs, default letters, information," & _
"and documents concerning the Loan are now integrated into Nationstar's business records.  Nationstar maintains quality control and " & _
"verification procedures as part of the boarding process to ensure the accuracy of the boarded records.  It is the regular business practice " & _
"of Nationstar to integrate prior servicers' records into Nationstar's business records and to rely upon those boarded records in providing " & _
"its loan servicing functions.  These Prior Servicer records have been integrated and are relied upon by Nationstar as part of Nationstar's " & _
"business records.", "")


'FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ____________________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "                " & "____________________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: _________________________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Loss Mitigation Preliminary Nationstar.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Loss Mitigation Preliminary Nationstar.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_LossMitigationFinalMDHCP(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim templateName As String
Dim BankName As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If



If d![ClientID] = 456 Then
BankName = "Loss Mitigation Final M&T Bank"
templateName = "Loss Mitigation - Final MT"
ElseIf d![ClientID] = 531 Then
BankName = "Loss Mitigation Final MDHC"
templateName = "Loss Mitigation - Final MDHCP"
Else
End If


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", IIf(Not IsNull(d![CourtCaseNumber]), d![CourtCaseNumber], " _____________________ ")
FillField WordDoc, "NameAffianit", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "Print Name and Title of Affiant", "Name")
FillField WordDoc, "TitleAffiaitOrInvestor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "Title of Affiant")
FillField WordDoc, "Line", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "______________________________")
FillField WordDoc, "Investor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, d![Investor], "")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1345, "")
'FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: _________________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "                " & "_________________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: _______________________", "")
FillField WordDoc, "fillingdate", IIf(Not IsNull(d![Docket]), Format(d![Docket], "mm/d/yyyy"), "_____________________ ")
FillField WordDoc, "borrower", MortgagorNamesOneline(d![CaseList.FileNumber], 2)

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & BankName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], BankName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_RecordingDeedCoverMD(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim sql As String
Dim templateName As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryRecorderPrint WHERE PrintInfoMD.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot, dbSeeChanges)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If


Set WordObj = CreateObject("Word.Application")

Select Case d!JurisdictionID

Case 18  'PG County
    templateName = "Deed Recording Cover PG"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
Case 4 'Baltimore City
    templateName = "Deed Recording Cover BaltCi"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    
    FillField WordDoc, "LienExpires", Nz(d!LienExpires, "________")
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "waterfees", Nz(d!WaterFees, "________")
Case 5   'Baltimore County
    templateName = "Deed Recording Cover BaltCounty"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    
    FillField WordDoc, "LienExpires", Nz(d!LienExpires, "________")
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    'FillField WordDoc, "waterfees", Nz(d!WaterFees, "________")
Case 3 ' Anne Arundel
    templateName = "Deed Recording Cover Anne"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "waterfees", Nz(d!WaterFees, "________")
Case 17 ' Montgomery County MD
    templateName = "Deed Recording Cover MoCoMD"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
Case 10 'Charles County, MD
    templateName = "Deed Recording Cover Charles"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Certificate", Nz(d!CertofTaxLien, "_________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
Case 12 ' Frederick MD
    templateName = "Deed Recording Cover Frederick"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "TaxStatus", Nz(d!TaxStatus, "_________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "WaterFeeDestination", Nz(d!WaterFeeDestination, "_____________")
Case 23 'Washington County
    templateName = "Deed Recording Cover WashCo"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    'FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "AgricultureTax", Nz(d!AgricultureTax, "___________")
Case 24 'WIcomico County
    templateName = "Deed Recording Cover Wicomico"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
 '   FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "WaterFeeDestination", Nz(d!WaterFeeDestination, "___________")
Case 25 'Worcester
    templateName = "Deed Recording Cover Worcester"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "waterFeeDestination", Nz(d!WaterFeeDestination, "________")
Case 19 ' Queen Anne's
    templateName = "Deed Recording Cover QueenAnne"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
Case 14 ' Harford
    templateName = "Deed Recording Cover Harford"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "LienExpires", Nz(d!LienExpires, "________")
Case 8 'Carroll County
    templateName = "Deed Recording Cover Carroll"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    'FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "LienExpires", Nz(d!LienExpires, "________")
    FillField WordDoc, "sewerFees", Nz(d!SewerFees, "_____________")
  
Case 7 'Caroline County
    templateName = "Deed Recording Cover Caroline"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "TownRPropertyTax", Nz(d!SewerFees, "_____________")
    FillField WordDoc, "WaterTarget", Nz(d!WaterFeeDestination, "____________")
    FillField WordDoc, "TownTaxDestination", Nz(d!TownRPropertyTaxDestination, "___________")
Case 6 'Calvert
    templateName = "Deed Recording Cover Calvert"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
   ' FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
   ' FillField WordDoc, "TownRPropertyTax", Nz(d!SewerFees, "_____________")
Case 22 'Talbot
    templateName = "Deed Recording Cover Talbot"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
   ' FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
   ' FillField WordDoc, "TownRPropertyTax", Nz(d!SewerFees, "_____________")
     FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
Case 16 'Kent
    templateName = "Deed Recording Cover Kent"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
   ' FillField WordDoc, "TownRPropertyTax", Nz(d!SewerFees, "_____________")
    ' FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
Case 20 'St. Mary's
    templateName = "Deed Recording Cover StMary"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
Case 21 'Somerset County
    templateName = "Deed Recording Cover Somerset"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "CityTax", Nz(d!CityTax, "___________")
    FillField WordDoc, "CityTaxDestination", Nz(d!CityTaxDestination, "____________")
    FillField WordDoc, "WaterFeeDestination", Nz(d!WaterFeeDestination, "_____________")
    
Case 2 'Allegany
    templateName = "Deed Recording Cover Allegany"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "WaterTarget", Nz(d!WaterFeeDestination, "_____________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "LienExpires", Nz(d!LienExpires, "________")
Case 11 'Dorchester
    templateName = "Deed Recording Cover Dorchester"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "WaterFeeDestination", Nz(d!WaterFeeDestination, "_____________")
Case 13 'Garrett
    templateName = "Deed Recording Cover Garrett"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
   ' FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
Case 9 'Cecil
    templateName = "Deed Recording Cover Cecil"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
   ' FillField WordDoc, "LienExpires", Nz(d!LienExpires, "________")
   ' FillField WordDoc, "sewerFees", Nz(d!SewerFees, "_____________")
Case 15 'Howard
    templateName = "Deed Recording Cover Howard"
    Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
    WordObj.Visible = True
    FillField WordDoc, "stateRecordationTax", Nz(d!StateRecordationTax, "________")
    FillField WordDoc, "statetransfertax", Nz(d!StateTransferTax, "_________")
    FillField WordDoc, "Propertytax", Nz(d!PropertyTax, "________")
    FillField WordDoc, "Waterfees", Nz(d!WaterFees, "___________")
    'FillField WordDoc, "cityTax", Nz(d!CityTax, "________")
    FillField WordDoc, "countyTransferTax", Nz(d!CountyTransferTax, "________")
    FillField WordDoc, "LienExpires", Nz(d!LienExpires, "________")
   ' FillField WordDoc, "sewerFees", Nz(d!SewerFees, "_____________")
Case Else
    MsgBox "No Word Document created for this County yet"
    d.Close

    Exit Sub
    
End Select

FillField WordDoc, "PrimaryDefName", d!PrimaryDefName
FillField WordDoc, "RecordingCharge", Forms![Print Deed Recording Cover Letter MD].RecordingFee
FillField WordDoc, "DeedREcordingFee", Nz(d!AbstractorFees, "________")
FillField WordDoc, "RecorderName", Nz(d!RecorderName, "________")
FillField WordDoc, "Date", Format(Now(), "mmmm d"", ""yyyy")
FillField WordDoc, "Mail", IIf(Forms![Print Deed Recording Cover Letter MD]!cboMail.Column(0) = "US Postal", "", Forms![Print Deed Recording Cover Letter MD]!cboMail.Column(0))
FillField WordDoc, "RecorderName", Forms![Print Deed Recording Cover Letter MD]!cbxAbstractorTarget.Column(1)
FillField WordDoc, "RecorderAttn", IIf(IsNull(d![RecorderATTN]) = True, "", "ATTN: " & d![RecorderATTN])
FillField WordDoc, "RecorderAddress", d!RecorderAddress & " " & Nz(d!RecorderAddress2, "")
FillField WordDoc, "RecorderAddress2", d![RecorderCity] & ", " & d![RecorderState] & " " & FormatZip(d![RecorderZip])
FillField WordDoc, "Mortgagor", MortgagorNamesOneline(0, 110)
FillField WordDoc, "Client", Forms![Case List]!ClientID.Column(1)
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!State
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "Filenumber", d!FileNumber
FillField WordDoc, "LoginName", GetLoginName()
'FillField WordDoc, "Investor", d!Investor

'
'FillField WordDoc, "RecFeeDestination", d!RecFeeDestination

'FillField WordDoc, "StateRecTaxDestination", d!StateRecTaxDestination

'FillField WordDoc, "StateTranstaxDestination", d!StateTransTaxDestination

'FillField WordDoc, "CountyTransTaxDestination", d!CountyTransTaxDestination
'FillField WordDoc, "CountyRPropertyTax", d!CountyRPropertyTax
'FillField WordDoc, "CountyRPropTaxDestination", d!CountyRPropTaxDestination
'FillField WordDoc, "TownRPropertyTax", d!TownRPropertyTax
'FillField WordDoc, "TownRPropertyTaxDestination", d!TownRPropertyTaxDestination
'FillField WordDoc, "waterfees", d!WaterFees
'FillField WordDoc, "waterFeeDestination", d!WaterFeeDestination


'FillField WordDoc, "OtherDesc", IIf(IsNull(d![OtherDescription]), "", "16.  Check in the amount of " & Format$(Nz(d![OtherAmount]), _
'"Currency") & " for " & d![OtherDescription])
'FillField WordDoc, "OtherDesc2", IIf(IsNull(d![OtherDescription2]), "", "17.  Check in the amount of " & Format$(Nz(d![OtherAmount2]), _
'"Currency") & " for " & d![OtherDescription2])
'FillField WordDoc, "OtherDesc3", IIf(IsNull(d![OtherDescription3]), "", "18.  Check in the amount of " & Format$(Nz(d![OtherAmount3]), _
'"Currency") & " for " & d![OtherDescription3])
'FillField WordDoc, "LoginName", GetLoginName()
'FillField WordDoc, "Firm", FirmName("MD")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & " " & templateName & ".doc"
Call SaveDoc(WordDoc, d![FileNumber], templateName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_EVNoticeBalti(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Dim mortgagors As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryfcdocswordlite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

'backbydate = Nz(InputBox("Enter a back by date", "Back by Date"), "")
Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Eviction Notice Baltimore City.dot", False, 0, True)

WordObj.Visible = True

'WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
'WordDoc.Bookmarks("FileNumber").Range.text = D![CaseList.FileNumber]

mortgagors = MortgagorOwnerNames(0, 2)
FillField WordDoc, "Lockout", Format([Forms]![EvictionDetails]![LockoutDate], "mmmm d, yyyy")
FillField WordDoc, "date", Format(Date, "mmmm d, yyyy")
FillField WordDoc, "Mortgagors", mortgagors
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!LongState
FillField WordDoc, "ZipCode", FormatZip(d!ZipCode)
FillField WordDoc, "Filenumber", d![CaseList.FileNumber]
'FillField WordDoc, "LoginName", GetLoginName
'FillField WordDoc, "STaffEmail", GetStaffEmail
FillField WordDoc, "Attorney", Forms![EvictionDetails]![Attorney].Column(1)

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Eviction Notice Baltimore City.dot"

'Why do we need a function to save a word document, the save feature is BUILT in...
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Eviction Notice Baltimore City.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
Set d = Nothing
End Sub   '****

Public Sub Doc_StatementOfDebtFiguresCham(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim i As Integer, itemCount As Integer, totalItems As Currency, itemsFields As String


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Statement of Debt Figures Cham"
templateName = templateName & ".dot"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


WordObj.Visible = False
If MsgBox("Is There A Loan Mod? ", vbYesNo) = vbYes Then
FillField WordDoc, "Mod", " MODIFIED by Agreement effective " & Format(InputBox(" Effective Date?   Format mm/dd/yyyy"), "mmmm d, yyyy") & " with an amended principal balance of " & Format(InputBox(" Amended Principal Balance? "), "Currency")
Else
FillField WordDoc, "Mod", ""
End If
WordObj.Visible = True

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber

FillField WordDoc, "Investor", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "") & d!Investor

FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "Reason", Forms![Print Statement of Debt]!txtOtherReason
'FillField WordDoc, "455Only", IIf(d![ClientID] = 455, ", and continuing each month thereafter with probable future advancements made by the mortgagee, and that the Plaintiff(s) has\have the right to foreclose;", ", and continuing each month thereafter, and that the Plaintiff(s) has\have the right to foreclose;")
FillField WordDoc, "Liber", LiberFolio(d![Liber], d![Folio], d![FCdetails.State], d![JurisdictionID])
'FillField WordDoc, "LPIdate", Format$(d![LPIDate], "mmmm d, yyyy")
'FillField WordDoc, "LPIdate+1", Format$(d![LPIDate] + 1, "mmmm d, yyyy")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "PaidStr", IIf((Nz(d![RemainingPBal], 0) > Nz(d![OriginalPBal], 0)), "Additional Interest", "Paid on principal")
FillField WordDoc, "Paid", Format$(d!OriginalPBal - d!RemainingPBal, "Currency")
FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")
FillField WordDoc, "ReRecorded", IIf(IsNull(d!Rerecorded), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Liber " & d![Liber2] & ", Folio " & d![Folio2])

FillField WordDoc, "Diem", ""  ' Ticket 906
'FillField WordDoc, "diem", IIf(d![LoanType] = 3, IIf(d![ClientID] <> 531, "Per Monthly Interest: ", "Per Diem Interest: "), "Per Diem Interest: ")


FillField WordDoc, "txtbalanc", IIf(d![ClientID] = 361, "Unpaid Principal Balance", "Remaining Balance Due")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

totalItems = 0
itemsFields = ""
Set dd = CurrentDb.OpenRecordset("SELECT Desc, Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber] & " ORDER BY Timestamp;", dbOpenSnapshot)
If dd.EOF Then      ' no extra lines
    FillField WordDoc, "Line_Items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items", itemsFields
    dd.MoveFirst
    i = 1
    Do While Not dd.EOF
        FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        totalItems = totalItems + Nz(dd!Amount, 0)
        dd.MoveNext
        i = i + 1
    Loop
End If
dd.Close

FillField WordDoc, "BalDueDate", IIf(IsNull([Forms]![Print Statement of Debt]![txtDueDate]), "______________", Format$([Forms]![Print Statement of Debt]![txtDueDate], "mmmm d"", ""yyyy"))
FillField WordDoc, "BalanceDue", Format$(d!RemainingPBal + totalItems, "Currency")

FillField WordDoc, "PerDiemInterest", "" 'Ticket 906
'FillField WordDoc, "PerDiemInterest", IIf(IsNull(d!PerDiem), "$_____________", Format$(d!PerDiem, "Currency"))

FillField WordDoc, "InterestRate", IIf(IsNull(d!InterestRate), "____________ %", Format$(d!InterestRate, "#0.000") & "%")
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", IIf(IsNull(d!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", d!NotaryLocation)

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Statement of Debt CHAM.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Statement of Debt CHAM.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_CHAMCoverSheetAFF(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM AFFmd7105.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")

FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter AFFMD7105.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_CHAMCoverSheetDeed(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM DeedOfApp.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")


FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter DeedOfApp.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_CHAMCoverSheetLossFinal(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM LossMitFinal.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")


FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter LossMitFinal.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_CHAMCoverSheetLossPrelim(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM LossMitPrelim.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")


FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter LossMitPrelim.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_CHAMCoverSheetMA(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM MilitaryAffidavit.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")


FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter MA.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_CHAMCoverSheetMAActive(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM MilitaryAffidavitActive.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")


FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter MAActive.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_CHAMCoverSheetMaNoSSN(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM MilitaryAffidavitNoSSN.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")


FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter MaNoSSN.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_CHAMCoverSheetNote(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM NoteOwnership.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")


FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter Ownership Affidavit.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_CHAMCoverSheetSOD(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM SOD.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")


FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter SOD.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_CHAMCoverSheetSOD2(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Cover Letter CHAM SOD2.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "State", Forms!foreclosuredetails!State
'FillField WordDoc, "FirstLegalDate", If Is Null(Forms!foreclosureDetails![567]),"",
FillField WordDoc, "FirstLegalDate", Nz(Forms!foreclosuredetails!FirstLegal, " No Date Selected")

FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)
FillField WordDoc, "ContactNameT", GetLoginName() & "     " & GetStaffEmail()
FillField WordDoc, "ReturnAddress", IIf(Forms!foreclosuredetails!State = "VA", "Commonwealth Trustees c/o Rosenberg and Associates, 8601 Westwood Center Drive, Ste 255, Vienna, VA 22182", "Rosenberg & Associates, LLC, Attn: Dockets Team, 7910 Woodmont Avenue, Ste 750, Bethesda, MD 20814")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Cover Letter CHAM.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Champion Cover Letter SoDwFigures.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_DismissCase(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim templateName As String
Dim BankName As String
Dim rsDismiss As Recordset
Dim sql As String

sql = "SELECT FormatName([Company],IIf([deceased]=Yes,""Estate of "" & [First],[First]),[Last],[AKA],[Address],[Address2],[City],[State],[Zip])"
sql = sql + " AS FmtNames FROM [Names]"
sql = sql + " WHERE  ((Names.FileNumber= " & [Forms]![Case List]![FileNumber] & ") AND Names.COS = True);"
'"WHERE (((CaseList.FileNumber)=" & [Forms]![Case List]![FileNumber] & ") AND JurisdictionList.CountyAttnyAddr Is Not Null);"

Set rsDismiss = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

templateName = "Dismiss Case"
BankName = "Dismiss Case"

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])


FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", IIf(Not IsNull(d![CourtCaseNumber]), d![CourtCaseNumber], " _____________________ ")
FillField WordDoc, "FirmName", FirmName()
FillField WordDoc, "FirmShortAddress", FirmShortAddress()
FillField WordDoc, "FirmPhone", FirmPhone()
FillField WordDoc, "Date", Format$(Date, "mmmm d, yyyy")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1345, "")
FillField WordDoc, "AttorneyName", d!AttorneyName

Dim itemCount As Integer
Dim i As Integer
Dim itemFields As String

'grumble grumble
If rsDismiss.EOF Then
    FillField WordDoc, "Fmtnames", ""
Else
    rsDismiss.MoveLast
    itemCount = rsDismiss.RecordCount
    rsDismiss.MoveFirst
    For i = 1 To itemCount
        itemFields = itemFields + rsDismiss!FmtNames + vbCrLf + vbCrLf
        rsDismiss.MoveNext
    Next i
    FillField WordDoc, "FmtNames", itemFields
End If


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & BankName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], BankName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_LandInstruments(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim templateName As String
Dim BankName As String
Dim d As Recordset
Dim sql As String
Dim strCriteria As String
Dim i As Integer
Dim strTrustees As String


Set d = CurrentDb.OpenRecordset("qryMDLandInstruments", dbOpenSnapshot)
strCriteria = "Filenumber = " & Forms!foreclosuredetails!FileNumber & ""
If Not (d.EOF And d.BOF) Then
    d.MoveFirst
    d.FindFirst (strCriteria)
    If d.NoMatch Then
        MsgBox "No Record"
        Exit Sub
    Else
    End If
Else
    MsgBox "No record"
    Exit Sub
End If
    

templateName = "MD Land Instruments"
BankName = "MD Land Instruments"

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "Property", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!State
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "SalePrice", Nz(d!SalePrice, 0#)
FillField WordDoc, "TaxID", Nz(d!TaxID, "")
FillField WordDoc, "Fsmpl", IIf(d!Leasehold = 1, " ", "X")
FillField WordDoc, "GR", IIf(d!Leasehold = 1, "X", " ")

strTrustees = ""
    For i = 0 To Forms!foreclosuredetails!lstTrustees.ListCount - 1
        If i < Forms!foreclosuredetails!lstTrustees.ListCount - 1 Then
            If i < Forms!foreclosuredetails!lstTrustees.ListCount - 2 Then
                strTrustees = strTrustees & Forms!foreclosuredetails!lstTrustees.Column(1, i) & ", "
            Else
                strTrustees = strTrustees & Forms!foreclosuredetails!lstTrustees.Column(1, i) & " and "
            End If
        Else
             strTrustees = strTrustees & Forms!foreclosuredetails!lstTrustees.Column(1, i)
        End If
    Next i

FillField WordDoc, "Attorneys", strTrustees
FillField WordDoc, "Owners", OwnerNames(0, 2)
FillField WordDoc, "Purchaser", Nz(d!Purchaser, "")
FillField WordDoc, "PurchaserAddr", Nz(d!PurchaserAddress, "")
FillField WordDoc, "LoginName", GetLoginName()
FillField WordDoc, "Recordation", Nz(d!RecordationTax, "")
FillField WordDoc, "StTransfer", Nz(d!StateTransTax, "")
FillField WordDoc, "CoTransfer", Nz(d!CountyTransTax, "")
FillField WordDoc, "District", Nz(d!District, "")
FillField WordDoc, "Map", Nz(d!Map, "")
FillField WordDoc, "Parcel", Nz(d!ParcelNo, "")
FillField WordDoc, "SubDivision", Nz(d!SubdivisionName, "")
FillField WordDoc, "Lot", Nz(d!Lot, "")
FillField WordDoc, "Block", Nz(d!Block, "")
FillField WordDoc, "Section", Nz(d!Section, "")
FillField WordDoc, "PlatRef", Nz(d!PlatRef, "")
FillField WordDoc, "SqFt", Nz(d!SqFt, "")
FillField WordDoc, "OtherProperty", Nz(d!OtherPropAddress, "")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & BankName & ".doc"
Call SaveDoc(WordDoc, d!FileNumber, BankName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing

d.Close
Set d = Nothing
End Sub

Public Sub Doc_MediationCourtNotesWells(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data



Set d = CurrentDb.OpenRecordset("SELECT * FROM qryWellsCourtMediation WHERE FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Mediation Court Notes Wells.dot", False, 0, True)
WordObj.Visible = True

FillField WordDoc, "Date", Format(Date, "mmmm, d, yyyy")
FillField WordDoc, "Loan", Nz(d![LoanNumber], "__________________")
FillField WordDoc, "Borrower", Nz(BorrowerNames(0), "____________________")
FillField WordDoc, "MedCaseNumber", Nz(d![MedCaseNumber], "__________________")
FillField WordDoc, "CourtCaseNumber", Nz(d![CourtCaseNumber], "__________________")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Wells Mediation Court Notes.dot"
Call SaveDoc(WordDoc, d![FileNumber], "Wells Mediation Court Notes.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_PayoffJPRI(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d, dd, ddd As Recordset     ' data
Dim totalItems As Currency
Dim itemsFields As String
Dim itemsfields2 As String
Dim strSQL, strSQL2 As String
Dim itemCount As Integer
Dim i As Integer
Dim itemcount2 As Integer
Dim ii As Integer
Dim templateName As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE Caselist.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

If d![FCdetails.State] = "VA" Then
    templateName = "PayOffJPRIVA"
Else
    templateName = "PayoffJPRI"
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

strSQL = "SELECT DISTINCTROW Payoff.FileNumber, Payoff.Timestamp, Payoff.Desc, Payoff.Amount, Payoff.ID"
strSQL = strSQL + " FROM Payoff"
strSQL = strSQL + " WHERE Payoff.Desc Not Like ""Accrued Interest*"" AND payoff.FileNumber = " & Forms![Case List]!FileNumber & ";"

'strSQL2 = "SELECT DISTINCTROW Payoff.FileNumber, Payoff.Timestamp, Payoff.Desc, Payoff.Amount, Payoff.ID"
'strSQL2 = strSQL2 + " FROM Payoff"
'strSQL2 = strSQL2 + " WHERE Payoff.Desc Like ""Accrued Interest*"" AND payoff.filenumber = " & Forms![case list]!FileNumber & ";"
' Maybe save this for Reinstatement?


totalItems = 0
itemsFields = ""

Set dd = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot) 'Should we order it by something?
If dd.EOF Then
    FillField WordDoc, "Line_Items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items", itemsFields
    dd.MoveFirst
    i = 1
    Do While Not dd.EOF
        FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        'FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        totalItems = totalItems + Nz(dd!Amount, 0)
        dd.MoveNext
        i = i + 1
       
    Loop
End If
dd.Close

'itemfields2 = ""
'Set ddd = CurrentDb.OpenRecordset(strSQL2, dbOpenSnapshot) 'This is kind of dumb
'If ddd.EOF Then
'    FillField WordDoc, "Line_Items2", ""
'Else
'    ddd.MoveLast
'    itemcount2 = ddd.RecordCount
'    ' Make enough lines
'    For ii = 1 To itemcount2
'        itemsfields2 = itemsfields2 & "<<Item2" & ii & ">>" & vbCr
'    Next ii
'    itemsfields2 = Left$(itemsfields2, Len(itemsfields2) - 1) ' remove trailing CR
'    FillField WordDoc, "Line_Items2", itemsfields2
'    ddd.MoveFirst
'    ii = 1
'    Do While Not ddd.EOF
'        FillField WordDoc, "Item2" & ii, ddd!Desc & vbTab & Format$(Nz(ddd!Amount, 0), "Currency")
'        'FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
'        totalItems = totalItems + Nz(ddd!Amount, 0)
'        ddd.MoveNext
'        ii = i + 1
'
'    Loop
'End If
'ddd.Close

FillField WordDoc, "Date", Format$(Date, "mmmm d, yyyy")
FillField WordDoc, "To", Forms![Print Payoff]!To
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Mortgagors", MortgagorNames(0, 2)
FillField WordDoc, "PropertyADDRESS", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "ZipCODE", FormatZip(d!ZipCode)
FillField WordDoc, "FileNumber", Forms![Case List].FileNumber
FillField WordDoc, "DueDate", Format$(Forms![Print Payoff].DutDate, "mmmm d"", ""yyyy")
FillField WordDoc, "Salutation", Forms![Print Payoff].[PayoffLetterSalutation]
FillField WordDoc, "Delinquent", Format(totalItems, "Currency")
FillField WordDoc, "GoodThrough", Format(Forms![Print Payoff]!GoodThru, "mmmm d"", ""yyyy")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & templateName & ".dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], templateName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges

Set WordObj = Nothing
d.Close
dd.Close
Set d = Nothing
Set dd = Nothing
End Sub


Public Sub Doc_ReadAtSale(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim Due As String
Dim tax As String


Dim d As Recordset      ' data
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE Caselist.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Read at Sale.dot", False, 0, True)
WordObj.Visible = True

Select Case d![FCdetails.State]
    Case "MD", "DC"
        Due = "ten (10) days after ratification of sale"
        tax = "all transfer taxes"
    Case "VA"
        Due = "fifteen (15) days of sale"
        tax = "Grantor's tax"
End Select

FillField WordDoc, "Title", UCase$(TrusteeWord(0, 1))
FillField WordDoc, "PropertyADDRESS", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "ZipCODE", FormatZip(d!ZipCode)
FillField WordDoc, "Mortgagors", IIf(IsNull(d![OriginalMortgagors]), MortgagorNames(0, 2), d![OriginalMortgagors] & " assumed by " & MortgagorNames(0, 2))
FillField WordDoc, "DOTDate", Format$(d![DOTdate], "mmmm d"", ""yyyy")
FillField WordDoc, "LiberFolio", LiberFolio(d![Liber], d![Folio], d![FCdetails.State])
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "LongState", d!LongState
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "TrusteeNames", trusteeNames(0, 2)
FillField WordDoc, "SaleLocation", d!SaleLocation
FillField WordDoc, "SaleTime", Format$(d![Sale], "dddd"", ""mmmm d"", ""yyyy") & " at " & Format$(d![SaleTime], "h:nn AM/PM")
FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "IRS", IIf(d![IRSLiens], "  The property will be sold subject to a 120 day right of redemption by the Internal Revenue Service.", "")
FillField WordDoc, "Deposit", Format$(d![Deposit], "Currency")
FillField WordDoc, "Due", Due
FillField WordDoc, "Tax", tax
FillField WordDoc, "VANotice", IIf(d![FCdetails.State] = "VA", "       Written notice of this " & TrusteeWord(0, 1) & "'s sale, as required by Section 55-59.1 of the 1950 Code of Virginia, as amended, has been sent to the property owners as their addresses appear in the records of the noteholder, and to all parties prescribed therein.", "")


WordObj.Selection.HomeKey wdStory, wdMove
'WordDoc.SaveAs EMailPath & "Read at Sale.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Read at Sale.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_ContractForSale(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim Due As String
Dim tax As String


Dim d As Recordset      ' data
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE Caselist.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Contract for Sale.dot", False, 0, True)
WordObj.Visible = True

Select Case d![FCdetails.State]
    Case "MD", "DC"
        Due = "ten (10) days after ratification of sale"
        tax = "all transfer taxes"
    Case "VA"
        Due = "fifteen (15) days of sale"
        tax = "Grantor's tax"
End Select

'FillField WordDoc, "Title", UCase$(TrusteeWord(0, 1))
FillField WordDoc, "PropertyADDRESS", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "ZipCODE", FormatZip(d!ZipCode)
FillField WordDoc, "Sale", ThisDate(d![Sale])
FillField WordDoc, "LiberFolio", LiberFolio(d![Liber], d![Folio], d![FCdetails.State])
FillField WordDoc, "Deposit", IIf(IsNull(d![Deposit]), "______________________________ Dollars", CurrencyWords(d![Deposit])) & " (" & IIf(IsNull(d![Deposit]), "$___________", Format$(d![Deposit], "Currency"))
FillField WordDoc, "FirmAddress", OneLine(FirmAddress())
FillField WordDoc, "FirmPhone", FirmPhone()

'FillField WordDoc, "Mortgagors", IIf(IsNull(d![OriginalMortgagors]), MortgagorNames(0, 2), d![OriginalMortgagors] & " assumed by " & MortgagorNames(0, 2))'''''
'FillField WordDoc, "DOTDate", Format$(d![DOTdate], "mmmm d"", ""yyyy")
'FillField WordDoc, "LiberFolio", LiberFolio(d![Liber], d![Folio], d![FCdetails.State])
'FillField WordDoc, "Jurisdiction", d!Jurisdiction
'FillField WordDoc, "LongState", d!LongState
'FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "TrusteeNames", trusteeNames(0, 2)
'FillField WordDoc, "SaleLocation", d!SaleLocation
'FillField WordDoc, "SaleTime", Format$(d![Sale], "dddd"", ""mmmm d"", ""yyyy") & " at " & Format$(d![SaleTime], "h:nn AM/PM")
'FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "IRSLiens", IIf(d![IRSLiens], "  The property will be sold subject to a 120 day right of redemption by the Internal Revenue Service.", "")

FillField WordDoc, "Due", Due
FillField WordDoc, "Tax", tax
'FillField WordDoc, "VANotice", IIf(d![FCdetails.State] = "VA", "       Written notice of this " & TrusteeWord(0, 1) & "'s sale, as required by Section 55-59.1 of the 1950 Code of Virginia, as amended, has been sent to the property owners as their addresses appear in the records of the noteholder, and to all parties prescribed therein.", "")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Contract For Sale.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Contract For Sale.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_HUDDEEDVA(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim Due As String
Dim tax As String


Dim d As Recordset      ' data
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryHUDDeedVA WHERE FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "HUD Deed VA.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("Atty").Select
WordDoc.Bookmarks("Atty").Range.Text = d!NameVA
WordDoc.Bookmarks("VABar").Select
WordDoc.Bookmarks("VABar").Range.Text = d!VABar
WordDoc.Bookmarks("FileNumber").Select
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("Property").Select
WordDoc.Bookmarks("Property").Range.Text = d![PropertyAddress] & ", " & d!City & ", VA " & FormatZip(d!ZipCode)


FillField WordDoc, "TaxID", d!TaxID
FillField WordDoc, "TrusteeNames", UCase$(trusteeNames(0, 2))
FillField WordDoc, "Mortgagors1", IIf(IsNull(d![OriginalMortgagors]), MortgagorNamesCaps(0, 2, 2), UCase$(d![OriginalMortgagors]) & " assumed by " & _
MortgagorNamesCaps(0, 2, 2))
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "HUDAddress", Nz(d!HUDAddress1Line, "")
FillField WordDoc, "DOTDate", Format$(d![DOTdate], "mmmm d"", ""yyyy")
FillField WordDoc, "LiberFolio", LiberFolio(d![Liber], d![Folio], "VA")
FillField WordDoc, "State", IIf(d![State] = "VA", "Clerk's Office, Circuit Court of ", "")
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "Mortgagors2", IIf(IsNull(d![OriginalMortgagors]), MortgagorNames(0, 2), d![OriginalMortgagors])
FillField WordDoc, "OriginalTrustee", d!OriginalTrustee
FillField WordDoc, "OriginalPBal", Format$(d![OriginalPBal], "Currency")
FillField WordDoc, "OriginalBene", d!OriginalBeneficiary
FillField WordDoc, "DeedRec", Format(d![DeedAppRecorded], "mmmm d"", ""yyyy")
FillField WordDoc, "LiberDeed", LiberFolio(d![DeedAppLiber], d![DeedAppFolio], d![FCdetails.State])
FillField WordDoc, "Consideration", Format$(d![SalePrice], "Currency")
FillField WordDoc, "Assessed", Format$(d![AssessedValue], "Currency")
FillField WordDoc, "SaleTime", IIf(IsNull(d![SaleTime]), "__________", Format$(d![SaleTime], "h:nn AM/PM"))
FillField WordDoc, "Sale", IIf(IsNull(d![Sale]), "_________________________", Format$(d![Sale], "mmmm d"", ""yyyy"))
FillField WordDoc, "SaleLocation", d!SaleLocation
FillField WordDoc, "LongState", d!LongState 'JurisdictionList.Longstate
FillField WordDoc, "SalePrice", IIf(IsNull(d![SalePrice]), "$_____________", Format$(d![SalePrice], "Currency"))
FillField WordDoc, "Lien", IIf(d![LienPosition] <= 1, "", ", subject to paying off senior lien(s)")
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "AttorneyName", d!AttorneyName
FillField WordDoc, "CommonwealthTitle", d!CommonwealthTitle
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "NotaryName", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "NotaryExpires", NotaryExpires()
FillField WordDoc, "PropertyAddress", d![PropertyAddress] & ", " & d!City & ", VA " & FormatZip(d!ZipCode)


WordObj.Selection.HomeKey wdStory, wdMove
WordObj.PrintPreview = True
WordDoc.SaveAs EMailPath & "HUD Deed VA.dot"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "HUD Deed VA.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_AuditFiling(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim templateName As String
Dim BankName As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If


templateName = "Line Notifying of Audit Filing"
Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
'FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", IIf(Not IsNull(d![CourtCaseNumber]), d![CourtCaseNumber], " _____________________ ")
FillField WordDoc, "Attorney", Forms![Audit - MD].Attorney.Column(1)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1345, "")
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "CosNames", COSNamesAddress(0)


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Line Notifying of Audit Filing.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Line Notifying of Audit Filing.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_ReturnDocCoverLTR(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim sqlString As String
Dim itemsFields As String
Dim dd As Recordset
Dim itemCount As Integer
Dim i As Integer
Dim num As Integer

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryReturnDocs WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Return Doc Cover LTR.dot", False, 0, True)
WordObj.Visible = True
FillField WordDoc, "Date", Date
FillField WordDoc, "ReturnName", IIf(IsNull(d![COLLName] = True), d![LongClientName], d![COLLName])
FillField WordDoc, "CollAddress", IIf(IsNull(d!COLLAddress = True), "", d!COLLAddress)      'Mei doc can't open if CollAddress is null. 09/10/15
FillField WordDoc, "CollAddress2", IIf(d![COLLAddress2] = Null, d![COLLCity] & ", " & d![COLLState] & " " & d![COLLZip], d![COLLAddress2] & vbCrLf _
 & d![COLLCity] & ", " & d![COLLState] & " " & d![COLLZip])

FillField WordDoc, "PrimaryDefName", d!PrimaryDefName
FillField WordDoc, "PropertyAddress", d![PropertyAddress] & " " & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt] & ", ") & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "FileNumber", d!FileNumber
FillField WordDoc, "LoginName", GetLoginName()

Dim itemFields As String
'totalItems = 0
itemsFields = ""
num = 1

Set dd = CurrentDb.OpenRecordset("SELECT DocDescription FROM ReturnedDocs WHERE FileNumber=" & d![FileNumber] & ";", dbOpenSnapshot)

If dd.EOF Then
    FillField WordDoc, "line_items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
    dd.MoveFirst
    For i = 1 To itemCount
        itemFields = itemFields & i & ". " & dd!DocDescription + vbCrLf + vbCrLf
        dd.MoveNext
    Next i
    FillField WordDoc, "line_items", itemFields
End If





WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Return Doc Cover LTR.dot"
Call SaveDoc(WordDoc, d![FileNumber], "Return Doc Cover LTR.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_LisPendens(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryLisPendens WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Notice of Lis Pendens.dot", False, 0, True)
WordObj.Visible = True


FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
FillField WordDoc, "ComplaintFiled", Format(d!ComplaintFiled, "mmmm d, yyyy")
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "PropertyAddress", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Owners", OwnerNames(0, 2)
FillField WordDoc, "CaseCaption", d![Investor] & " v. " & DefendantNames(0, 99)
FillField WordDoc, "CaseNumber", d!CourtCaseNumber
FillField WordDoc, "OriginalPBal", Format(d!OriginalPBal, "Currency")
FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "AttyName", GetStaffFullName(Forms!Foreclosureprint!Attorney.Column(0)) & " (" & FetchbarNumberPoundSign(Forms!Foreclosureprint!Attorney, d![State]) & ")"




WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Notice of Lis Pendens.dot"
Call SaveDoc(WordDoc, Forms![Case List].FileNumber, "Notice of Lis Pendens.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_ConsentTerminatingWD13(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim FileNumber As Long, Judge As String, CoDebtor As Boolean, AffDate As String, AffInfo As String, DebtorsPlural As Boolean

FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBKDocsWord WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "ConsentTerminatingWD13.dot", False, 0, True)
WordObj.Visible = True
WordObj.ScreenUpdating = False

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("AttorneyInfo").Select

'RE-VISIT
WordDoc.Bookmarks("AttorneyInfo").Range.Text = "Diane Rosenberg" & vbCr & "VA Bar 35237"


Judge = Right$(UCase$(d!CaseNo), 3)
CoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13)

DebtorsPlural = (CountNames(FileNumber, "BKDebtor = True AND (Owner=True OR Mortgagor=True)") > 1)

FillField WordDoc, "Header", _
    IIf(d![Districts.State] <> "VA", vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr, "") & _
    "IN THE UNITED STATES BANKRUPTCY COURT" & vbCr & _
    "FOR THE " & UCase$(d!Name) & vbCr & _
    d!Location

FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & IIf(DebtorsPlural, "s", "")

FillField WordDoc, "InvestorAddr", UCase$(d!Investor) & vbCr & RemoveLF(d!InvestorAddress)

FillField WordDoc, "Respondents", _
    GetAddresses(0, 4, _
        IIf(d![BKdetails.Chapter] = 13, _
            "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", _
            "BKDebtor=True AND (Owner=True OR Mortgagor=True)"), vbCr)
    
    'Why is there chapter 7 anything in a chapter 13 Sub???????  !>.<!
  '  IIf(d![BKdetails.Chapter] = 7, _
  '      vbCr & vbCr & "and" & vbCr & vbCr & FormatName("", _
   '                                                     UCase$(Nz(d!First)), _
   '                                                     UCase$(Nz(d!Last)), _
   '                                                     ", CHAPTER 7 TRUSTEE", _
   '                                                     d!Address, _
   '                                                     d!Address2, _
   '                                                     d![BKTrustees.City], _
   '                                                     d![BKTrustees.State], _
   '                                                     d!Zip, _
    '                                                    vbCr), _
    '    "")
    
FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]

'FillField WordDoc, "Caption", IIf(d!Judge = "RGM" And d!RealEstate, "CONSENT ORDER  AS TO " & UCase$(d!PropertyAddress), "CONSENT ORDER MODIFYING AUTOMATIC STAY")

Select Case Forms![Print Consent Order Terminating 13]!optTrustee
    Case 0
        FillField WordDoc, "TrusteeAction", "the parties having reached an agreement,"
    Case 1
        FillField WordDoc, "TrusteeAction", "the trustee having filed a report of no distribution, the parties having reached an agreement, "
    Case 2
        FillField WordDoc, "TrusteeAction", "the trustee having failed to file an answer, the debtor(s) and movant having reached an agreement, "
    Case 3
        FillField WordDoc, "TrusteeAction", "the parties having reached an agreement,"
End Select

'FillField WordDoc, "ThisDate", IIf(d![Districts.State] <> "VA", "", ", it is this " & OrderDate())

FillField WordDoc, "District", d!Name 'd!Districts.Name
FillField WordDoc, "CoDebtorSection", IIf(CoDebtor, "Sections 362(d) and 1301", "Section 362(d)")
FillField WordDoc, "Action", IIf(d!RealEstate, "commence foreclosure proceeding " & IIf(d![FCdetails.State] = "MD", "in the Circuit Court for " & d!Jurisdiction & IIf(IsNull(d!LongState), "", ", " & d!LongState) & ", ", "") & "against the real property and improvements " & IIf(IsNull(d!ShortLegal), "", "with a legal description of """ & d!ShortLegal & """ also ") & "known as " & d!PropertyAddress & ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d!ZipCode) & " and to allow the successful purchaser to obtain possession of same", "proceed with repossession and sale of the " & d!PropertyDesc)

'FillField WordDoc, "OrderWording", IIf(d!Judge = "RGM", "ORDERED that the Debtor shall:", "ORDERED that the above Order be and it is hereby, stayed provided that the Debtor:")
'FillField WordDoc, "ConsentPaymentAmount", Format$(d!ConsentPaymentAmount, "Currency")
'FillField WordDoc, "PaymentType", IIf(d!RealEstate, "mortgage", "monthly")
'FillField WordDoc, "ConsentPaymentDate", Format$(d!ConsentPaymentDate, "mmmm d, yyyy")
'FillField WordDoc, "NoteType", IIf(d!RealEstate, "Promissory Note secured by the " & DOTWord(d!DOT) & " on the above referenced property", d!PropertyContract)
'FillField WordDoc, "ConsentPaymentInfo", d!ConsentPaymentInfo
'FillField WordDoc, "2A", IIf(IsNull(d!Consent2A), "", "^p2A. " & d!Consent2A)

'FillField WordDoc, "AndAttorney", IIf(IsNull(d!AttorneyLastName), "", "and Debtor's attorney ")

'FillField WordDoc, "RGM1", IIf(d!Judge = "RGM", "", ", without further order of court")

'FillField WordDoc, "RGM2", IIf(d!Judge = "RGM", "^pIf any amount required in Paragraph 2 is not paid timely, Movant's attorney shall mail notice to the Debtor" & _
'                        IIf(IsNull(d!AttorneyLastName), "", ", Debtor's attorney") & " and Chapter 13 Trustee, and shall file an Order of Termination of Automatic " & _
'                        "Stay against the Subject Property described above; and be it further^p" & _
'                        "ORDERED that a default in the payment of a regularly scheduled mortgage payment as listed in Paragraph 1 shall be governed by the attached addendum.", "")

If d![Districts.State] = "VA" Then FillField WordDoc, "JudgeSignature", "________________________________" & vbCr & "United States Bankruptcy Judge"
    
FillField WordDoc, "MovantAtty", [Forms]![BankruptcyPrint]![cbxAttorney]
FillField WordDoc, "DebtorAtty", IIf(IsNull(d![AttorneyLastName]), DebtorNames(0, 4), "________________________________" & d![AttorneyFirstName] & " " & d![AttorneyLastName] & "Attorney for Debtor")
FillField WordDoc, "DebtorColumn", IIf([Forms]![BankruptcyPrint]![chElectronicSignature], IIf(IsNull(d![AttorneyLastName]), DebtorNames(0, 5), "/s/ " & d![AttorneyFirstName] & " " & d![AttorneyLastName]), "")
FillField WordDoc, "MovantColumn", IIf([Forms]![Print Consent Order Terminating 13]![optTrustee] = 3 And [Forms]![BankruptcyPrint]![chElectronicSignature], "/s/ " & d![First] & " " & d![Last], "")
FillField WordDoc, "MovantColumn2", IIf([Forms]![Print Consent Order Terminating 13]![optTrustee] = 3, "_______________________________" & d![First] & " " & d![Last] & "Chapter 13 Trustee", "")
FillField WordDoc, "AttyName", Forms!BankruptcyPrint!cbxAttorney
FillField WordDoc, "FirmAddress", FirmAddress()
FillField WordDoc, "TrusteeName", FormatName("", d![First], d![Last], ", Trustee", d![Address], d![Address2], d![BKTrustees.City], d![BKTrustees.State], d![Zip])
FillField WordDoc, "AttySig", IIf(IsNull(d![AttorneyLastName]), "", d![AttorneyFirstName] & " " & d![AttorneyLastName] & ", Esquire" & IIf(IsNull(d![AttorneyFirm]), "", d![AttorneyFirm] & "") & d![AttorneyAddress])
FillField WordDoc, "BkServ", BKService(0)
    
FillField WordDoc, "WordyParagraph", IIf(Forms!BankruptcyPrint!chElectronicSignature And d![Districts.State] <> "VA", "          I HEREBY CERTIFY that the terms of the copy of the consent order submitted to the Court are identical to those set forth in the original consent order; and the signatures represented by the /s/__________ on this copy reference the signatures of consenting parties on the original consent order.", "")
FillField WordDoc, "End", IIf(d![Districts.State] <> "VA", "End of Order", "")
    
FillField WordDoc, "CertDate", Format$(Date, "mmmm d"", ""yyyy")
FillField WordDoc, "ESig", IIf([Forms]![BankruptcyPrint]![chElectronicSignature], "/s/ " & [Forms]![BankruptcyPrint]![cbxAttorney], "")
    
    
    'FillField WordDoc, "Submitted", "Respectfully Submitted:"
'    If Forms!BankruptcyPrint!chElectronicSignature Then
 '       FillField WordDoc, "ElectronicSignature", "/s/ Diane S. Rosenberg"
 '   Else
 '       FillField WordDoc, "ElectronicSignature", "_______________________________"
 '   End If
 '   FillField WordDoc, "AttorneySignature", "Diane S. Rosenberg"
 '   FillField WordDoc, "End", ""
'Else
 '   FillField WordDoc, "OrderDate", ""
 '   FillField WordDoc, "JudgeSignature", ""
 '   FillField WordDoc, "Submitted", ""
 '   FillField WordDoc, "ElectronicSignature", ""
 '   FillField WordDoc, "AttorneySignature", ""
'    FillField WordDoc, "End", "End of Order"
'End If

WordObj.Selection.HomeKey wdStory, wdMove
WordObj.ScreenUpdating = True
WordDoc.SaveAs EMailPath & "ConsentTerminatingWD13.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "ConsentTerminatingWD13.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_DCAccounting(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim sqlString As String
Dim TotalProceeds As Currency
Dim TotalExpenses As Currency
Dim NoticeAmount As Currency
Dim TotalDebt As Currency
Dim availableToPay As Currency
Dim FinalBalance As Currency

Set dd = CurrentDb.OpenRecordset("Select * From qryDCAccounting WHERE Caselist.FileNumber=" & Forms![Case List]!FileNumber & "AND AuditID =" & [Forms]![Audit - DC]![AuditID], dbOpenSnapshot)

If dd.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
NoticeAmount = [Forms]![Audit - DC]![NoticesAmount]
TotalExpenses = Nz(dd![FilingFee]) + Nz(dd![DeedApp]) + Nz(dd![Bond]) + Nz(dd![AdvertisingPreSale]) + Nz(dd![AttorneyFee]) + Nz(dd![TitleReport]) + Nz(dd![AuctioneerFee]) + Nz(NoticeAmount) + Nz(dd![OtherAmount]) + Nz(dd![Other2Amount]) + Nz(dd![Other3Amount]) + Nz(dd![Other4Amount]) + Nz(dd![Other5Amount]) + Nz(dd![Other6Amount]) + Nz(dd![Other7Amount]) + Nz(dd![Other8Amount]) + Nz(dd![Other9Amount]) + Nz(dd![Other10Amount]) + Nz(dd![LisPendensFee])

TotalProceeds = Nz(dd![Proceeds]) + Nz(dd![Interest]) + Nz(dd![PropertyTaxes]) + Nz(dd![OtherProceeds1Amount]) + Nz(dd![OtherProceeds2Amount])
TotalDebt = Nz(dd![StatementofDebt]) + Nz(dd![Interest])
availableToPay = TotalProceeds - TotalExpenses
FinalBalance = availableToPay - TotalDebt


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "DC Accounting.dot", False, 0, True)
WordObj.Visible = True

FillField WordDoc, "Defendants", DefendantNames(0, 99)
FillField WordDoc, "JudgmentEntered", Format$(d!JudgmentEntered, "mmmm d, yyyy")
FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
'FillField WordDoc, "LegalDescription", d!LegalDescription
'FillField WordDoc, "TrusteeWord", UCase$(TrusteeWord(d![CaseList.FileNumber], 0))
FillField WordDoc, "PropertyAddress", d!PropertyAddress & vbCr & "" & d!City & ", " & d![FCdetails.State] & ", " & FormatZip(d!ZipCode)
'FillField WordDoc, "DOTRecorded", Format$(d!DOTrecorded, "mmmm d, yyyy")
'FillField WordDoc, "DoTDate", Format$(d!DOTdate, "mmmm d, yyyy")
'FillField WordDoc, "Liber", d!Liber
'FillField WordDoc, "Docket", Format$(d!Docket, "mmmm d, yyyy")
FillField WordDoc, "SaleDate", Format$(d!Sale, "dddd, mmmm d, yyyy")
FillField WordDoc, "SaleTime", Format$(d!SaleTime, "h:nn AM/PM")
'FillField WordDoc, "TaxID", d!TaxID
'FillField WordDoc, "IRSLiens", IIf(d!IRSLiens, "The property will be sold subject to a 120 day right of redemption by the Internal Revenue Service.  ", "")
'FillField WordDoc, "PriorLien", IIf(d!LienPosition <= 1, "", "The property will be sold subject to a prior mortgage, the amount to be announced at the time of sale.  ")
'FillField WordDoc, "Deposit", Format$(d!Deposit, "Currency")
FillField WordDoc, "Trustees", trusteeNames(d![CaseList.FileNumber], 2)
FillField WordDoc, "TrusteeList", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "PrintDate", Format(Date, "mmmm, d, yyyy")
FillField WordDoc, "Parties", NoticeNames(d![CaseList.FileNumber], "")
'FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
'FillField WordDoc, "ComplaintFiled", Format(d!ComplaintFiled, "mmmm d, yyyy")
FillField WordDoc, "Investor", d!Investor
'FillField WordDoc, "PropertyAddress", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & d![ZipCode]
'FillField WordDoc, "Owners", OwnerNames(0, 2)
'FillField WordDoc, "CaseCaption", d![Investor] & " v. " & DefendantNames(0, 99)
FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
'FillField WordDoc, "OriginalPBal", Format(d!OriginalPBal, "Currency")
'FillField WordDoc, "LegalDescription", d!LegalDescription
'FillField WordDoc, "AttyName", GetStaffFullName(Forms!ForeclosurePrint!Attorney.Column(0)) & " (" & FetchBarNumberPoundSign(Forms!ForeclosurePrint!Attorney, d![State]) & ")"
FillField WordDoc, "Interest", IIf(Nz(dd![Interest]) = 0, "", "Interest " & Format$(dd![Interest3From], "m/d/yyyy" & " - " & Format$(dd![Interest3To], "m/d/yyyy")))
FillField WordDoc, "PropertyTaxesDesc", Nz(dd!PropertyTaxesDesc, "")
FillField WordDoc, "PropertyTaxes", Format$(dd![PropertyTaxes], "Currency")
'FillField WordDoc, "LeaseHold", IIf(dd![Leasehold] And dd![CaseList.JurisdictionID] = 4, "Leasehold Real Estate subject to annual ground rent of " & Format$(dd![GroundRentAmount], "Currency") & " payable " & dd![GroundRentPayable], "Real Estate")
FillField WordDoc, "Proceeds", Format$(dd![Proceeds], "Currency")
FillField WordDoc, "InterestAmount", IIf(Nz(dd![Interest]) = 0, "", Format$(dd![Interest], "Currency"))
FillField WordDoc, "TotalProceeds", Format$(TotalProceeds, "Currency")

'Expenses
FillField WordDoc, "FilingFee", Format$(dd![FilingFee], "Currency")
FillField WordDoc, "DeedApp", Format$(dd![DeedApp], "Currency")
FillField WordDoc, "Bond", Format$(dd![Bond], "Currency")
FillField WordDoc, "NewspaperPreSale", dd!NewspaperPreSale
FillField WordDoc, "AdvertisingPreSale", Format$(dd![AdvertisingPreSale], "Currency")
FillField WordDoc, "FirmName", FirmName()
FillField WordDoc, "AttorneyFee", Format$(dd![AttorneyFee], "Currency")
FillField WordDoc, "Abstractor", dd!Abstractor
FillField WordDoc, "TitleReport", Format$(dd![TitleReport], "Currency")
FillField WordDoc, "Auctioneer", Nz(dd!Auctioneer, "")
FillField WordDoc, "AuctioneerFee", Format$(dd![AuctioneerFee], "Currency")
FillField WordDoc, "Notices", Format(Forms![Audit - DC]!NoticesAmount, "Currency")
FillField WordDoc, "OtherDesc", IIf(dd![OtherDesc] <> "", dd![OtherDesc], "")
FillField WordDoc, "Other2Desc", IIf(dd![Other2Desc] <> "", dd![Other2Desc], "")
FillField WordDoc, "Other3Desc", IIf(dd![Other3Desc] <> "", dd![Other3Desc], "")
FillField WordDoc, "Other4Desc", IIf(dd![Other4Desc] <> "", dd![Other4Desc], "")
FillField WordDoc, "Other5Desc", IIf(dd![Other5Desc] <> "", dd![Other5Desc], "")
FillField WordDoc, "OtherAmount", IIf(dd![OtherDesc] <> "", Format$(dd![OtherAmount], "Currency"), "")
FillField WordDoc, "Other2Amount", IIf(dd![Other2Desc] <> "", Format$(dd![Other2Amount], "Currency"), "")
FillField WordDoc, "Other3Amount", IIf(dd![Other3Desc] <> "", Format$(dd![Other3Amount], "Currency"), "")
FillField WordDoc, "Other4Amount", IIf(dd![Other4Desc] <> "", Format$(dd![Other4Amount], "Currency"), "")
FillField WordDoc, "Other5Amount", IIf(dd![Other5Desc] <> "", Format$(dd![Other5Amount], "Currency"), "")

FillField WordDoc, "Other6Desc", IIf(dd![OtherDesc] <> "", dd![Other6Desc], "")
FillField WordDoc, "Other7Desc", IIf(dd![Other2Desc] <> "", dd![Other7Desc], "")
FillField WordDoc, "Other8Desc", IIf(dd![Other3Desc] <> "", dd![Other8Desc], "")
FillField WordDoc, "Other9Desc", IIf(dd![Other4Desc] <> "", dd![Other9Desc], "")
FillField WordDoc, "Other10Desc", IIf(dd![Other5Desc] <> "", dd![Other10Desc], "")
FillField WordDoc, "Other6Amount", IIf(dd![OtherDesc] <> "", Format$(dd![Other6Amount], "Currency"), "")
FillField WordDoc, "Other7Amount", IIf(dd![Other2Desc] <> "", Format$(dd![Other7Amount], "Currency"), "")
FillField WordDoc, "Other8Amount", IIf(dd![Other3Desc] <> "", Format$(dd![Other8Amount], "Currency"), "")
FillField WordDoc, "Other9Amount", IIf(dd![Other4Desc] <> "", Format$(dd![Other9Amount], "Currency"), "")
FillField WordDoc, "Other10Amount", IIf(dd![Other5Desc] <> "", Format$(dd![Other10Amount], "Currency"), "")

FillField WordDoc, "TotalExpenses", Format(TotalExpenses, "Currency")

'Section 3
FillField WordDoc, "LisPendens", Format(Nz(dd!LisPendensFee, 0), "Currency")
FillField WordDoc, "StatementOfDebt", Format$(dd![StatementofDebt], "Currency")
FillField WordDoc, "InterestFrom", Format$(dd![InterestFrom], "m/d/yyyy")
FillField WordDoc, "InterestTo", Format$(dd![InterestTo], "m/d/yyyy")
FillField WordDoc, "DateDiff", Nz(DateDiff("d", dd![InterestFrom], dd![InterestTo]), "")
FillField WordDoc, "PerDiem", Nz(Format$(dd![InterestPerDiem], "Currency"), "")
FillField WordDoc, "PerDiemTotal", Nz(Format$(DateDiff("d", dd![InterestFrom], dd![InterestTo]) * dd![InterestPerDiem], "Currency"), "")
FillField WordDoc, "TotalDebt", Format(TotalDebt, "currency")
FillField WordDoc, "AvailableTOPay", Format(availableToPay, "Currency")
FillField WordDoc, "FinalBalance", Format(FinalBalance, "Currency")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "DC Accounting.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "DC Accounting.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
dd.Close

End Sub

Public Sub Doc_DCTrusteeAffidavit(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryDCTrusteeAFF WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "DC Trustee Affidavit.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

FillField WordDoc, "Pub1", Format(InputBox("Please enter the next Publication Date.  (2 of 4)  Format mm/dd/yyyy"), "mmmm d, yyyy")
FillField WordDoc, "Pub2", Format(InputBox("Please enter the next Publication Date.  (3 of 4)  Format mm/dd/yyyy"), "mmmm d, yyyy")
FillField WordDoc, "Pub3", Format(InputBox("Please enter the next Publication Date.  (4 of 4)  Format mm/dd/yyyy"), "mmmm d, yyyy")

FillField WordDoc, "AttyName", GetStaffFullName(Forms!Foreclosureprint!Attorney.Column(0))
FillField WordDoc, "Defendants", DefendantNames(0, 99)
FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "PropertyAddress", d!PropertyAddress & vbCr & "" & d!City & ", " & d![State] & ", " & FormatZip(d!ZipCode)
FillField WordDoc, "Sale", Format$(d!Sale, "dddd, mmmm d, yyyy")
FillField WordDoc, "SaleTime", Format$(d!SaleTime, "h:nn AM/PM")
FillField WordDoc, "Notices", Format(d!Notices, "mmmm d, yyyy")
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "firstPub", Format(d!FirstPub, "mmmm d, yyyy")
FillField WordDoc, "BondPosted", Format(d!BondPosted, "mmmm d, yyyy")
FillField WordDoc, "Purchaser", d!Purchaser
FillField WordDoc, "SalePrice", Format(d!SalePrice, "Currency")
FillField WordDoc, "SelectedTrustee", GetStaffFullName(Forms!Foreclosureprint!Attorney.Column(0)) & " (" & FetchbarNumberPoundSign(Forms!Foreclosureprint!Attorney, d![State]) & ")"
FillField WordDoc, "Month", Format(Date, "mmmm")
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "NoticePosted", Format(d!NoticePosted, "mmmm d, yyyy")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "DC Trustee Affidavit.doc"
Call SaveDoc(WordDoc, d![FileNumber], "DC Trustee Affidavit.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close


End Sub

Public Sub Doc_BaltCityIntakeSheet(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBaltCityIntake WHERE FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "BaltCity Intake Sheet.dot", False, 0, True)
WordObj.Visible = True


FillField WordDoc, "PropertyAddress", Nz(d!PropertyAddress) & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "TrusteeNames", trusteeNames(d![FileNumber], 2)
FillField WordDoc, "Purchaser", Nz(d!Purchaser)
FillField WordDoc, "SalePrice", Nz(Format(d!SalePrice, "Currency"))
FillField WordDoc, "GetloginName", GetLoginName()
FillField WordDoc, "GetStaffEmail", GetStaffEmail()
FillField WordDoc, "PurchaserAddress", Nz(d!PurchaserAddress)


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Baltimore City Intake Sheet.doc"
Call SaveDoc(WordDoc, d![FileNumber], "Baltimore City Intake Sheet.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close


End Sub

Public Sub Doc_DCOrderGrantingDefaultWells(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range, DefendantsAddress As String
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryDCTrusteeAFF WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "DC Order Granting Default - Wells.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & d![Fair Debt]

'DefendantsAddress = GetAddresses(d!FileNumber, 5, "Defendant=True", NewLine)
'FillField WordDoc, "AttyName", GetStaffFullName(Forms!ForeclosurePrint!Attorney.Column(0))

FillField WordDoc, "Defendants", DefendantNames(0, 99)
FillField WordDoc, "DefendantNames", DefendantNames(0, 99)
FillField WordDoc, "DefendantAddress", GetAddresses(0, 5, "Defendant = True") ', vbCr)
FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "PropertyAddress", d!PropertyAddress & d![Fair Debt] & vbCr & "" & d!City & ", " & d![State] & ", " & FormatZip(d!ZipCode)
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![Fair Debt] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Month", Format(Date, "mmmm")
FillField WordDoc, "Year", Format(Date, "yyyy")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "DC Order Granting Default Wells.doc"
Call SaveDoc(WordDoc, d![FileNumber], "DC Order Granting Default Wells.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close


End Sub


Public Sub Doc_DeedOfAppointmentSPLS(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String


If (d![FCdetails.State] = "MD") Then
    templateName = "Deed of Appointment SPLS MD"
Else
    templateName = "Deed of Appointment SPLS VA"
End If

templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")

Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


If d!JurisdictionID = 56 And d![FCdetails.State] = "VA" Then
    FillField WordDoc, "Tax", "Tax Map Number:"
Else
    FillField WordDoc, "Tax", "Tax ID #:"
End If

If (d![FCdetails.State] = "VA") Then
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", IIf(IsNull(d![Folio2]), ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & _
    " at Instrument Number " & d![Liber2], ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Book " & d![Liber2] & ", Page " & d![Folio2]))
ElseIf d!JurisdictionID = 6 Then  'Calvert County
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Book " & d![Liber2] & ", Page " & d![Folio2])
Else 'MD or DC
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Liber " & d![Liber2] & ", Page " & d![Folio2])
End If

FillField WordDoc, "LoanNumber", d!LoanNumber

If d!JurisdictionID = 153 Then ' Accomack VA
    FillField WordDoc, "OriginalTrustee", UCase$(d!OriginalTrustee)
    'FillField WordDoc, "InvestorAIF", IIf(d!AIF = True, d!LongClientName & " as Attorney in Fact for ", "") & UCase$(d!Investor)
    FillField WordDoc, "InvestorAIF", IIf((d!AIF = True And d!ClientID = 532), UCase$(d![Investor]) & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "") & UCase$(d![Investor]))
    FillField WordDoc, "OriginalBeneficiary", UCase$(d!OriginalBeneficiary)
    FillField WordDoc, "AssumedBy", IIf(IsNull(d!OriginalMortgagors), "", "assumed by " & UCase$(MortgagorNamesCaps(0, 2, 2))) & " "
    FillField WordDoc, "MortgagorNames", IIf(IsNull(d!OriginalMortgagors), UCase$(MortgagorNamesCaps(0, 2, 2)), UCase$(d!OriginalMortgagors))
    FillField WordDoc, "TrusteeNames", UCase$(trusteeNames(0, 2))
    FillField WordDoc, "Investor", UCase$(d!Investor)
Else
    FillField WordDoc, "OriginalTrustee", d!OriginalTrustee
    FillField WordDoc, "InvestorAIF", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "") & d![Investor])
    FillField WordDoc, "OriginalBeneficiary", d!OriginalBeneficiary
    'FillField WordDoc, "AssumedBy", IIf(IsNull(d!OriginalMortgagors), "", "assumed by " & MortgagorNamesCaps(0, 2, 2)) & " "
    FillField WordDoc, "AssumedBy", " "
    FillField WordDoc, "MortgagorNames", IIf(IsNull(d!OriginalMortgagors), MortgagorNamesCaps(0, 2, 2), d!OriginalMortgagors)
    'FillField WordDoc, "TrusteeNames", d!OriginalTrustee
    FillField WordDoc, "TrusteeNames", trusteeNames(0, 2)
    FillField WordDoc, "Investor", d!Investor
    'FillField WordDoc, "InvestorSPLS", d![Investor] & IIf((d![AIF] = True And d![Investor] <> "Specialized Loan Servicing, LLC"), ", by Specialized Loan Servicing, LLC, as servicer for secured party", "")
    FillField WordDoc, "InvestorSPLS", d![Investor] & IIf((d![AIF] = True And Not Trim(UCase$(d![Investor])) Like "SPECIALIZED LOAN SERVICING*"), ", by Specialized Loan Servicing, LLC, as servicer for secured party", "")
 
End If

FillField WordDoc, "Filenumber", d![CaseList.FileNumber]
FillField WordDoc, "DOTDate", Format$(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "LongState", Nz(d!LongState)
FillField WordDoc, "LiberFolio", LiberFolio(d!Liber, d!Folio, d![FCdetails.State], d![JurisdictionID])
FillField WordDoc, "Liber", Nz(d!Liber)
FillField WordDoc, "Folio", Nz(d!Folio)
FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "TaxID", d!TaxID
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")

FillField WordDoc, "MDInstrPrepared", IIf(d![FCdetails.State] = "VA", "", "This instrument was prepared under the supervision of " & d!AttorneyName & ", an attorney admitted to practice before the Court of Appeals of Maryland.")
FillField WordDoc, "MDSignLine", IIf(d![FCdetails.State] = "VA", "", "_________________________________")
FillField WordDoc, "MDAttorney", IIf(d![FCdetails.State] = "VA", "", d!AttorneyName)
    
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", FetchNotaryLocation(Forms!Foreclosureprint!NotaryID)
FillField WordDoc, "NotaryName", FetchNotaryName(Forms!Foreclosureprint!NotaryID, False)
FillField WordDoc, "FirmAddress", IIf(d![FCdetails.State] = "VA", "Commonwealth Trustees, LLC" & vbCr & "c/o Rosenberg & Associates, LLC" & vbCr & "8601 Westwood Center Drive, Suite 255" & vbCr & "Vienna, VA 22182", FirmAddress(vbCr))
FillField WordDoc, "NotaryNameSignLine", FetchNotaryName(Forms!Foreclosureprint!NotaryID, True)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 226, "")
FillField WordDoc, "boa1", IIf(d![ClientID] = 446, "", "and, being duly sworn")
'FillField WordDoc, "boa2", IIf(d![ClientID] = 446, "National Association.", IIf(d![ClientID] = 385 Or d![ClientID] = 567, "company.", IIf([d!clientID] = 87, "association.", "corporation.")))
FillField WordDoc, "BOA2", IIf(d![ClientID] = 446, " National Association.", IIf(d![ClientID] = 385 Or d![ClientID] = 567, " company.", IIf(d![ClientID] = 87, " association.", " corporation.")))
'FillField WordDoc, "MERS", IIf(d![FCdetails.State] = "MD" And d![MERS] = -1, "Mortgage Electronic Registration Systems Inc. (MERS) solely as nominee for ", "")
FillField WordDoc, "MERS", IIf(d![MERS] = -1, "Mortgage Electronic Registration Systems Inc. (MERS) solely as nominee for ", "")

FillField WordDoc, "FinalLanguage", "WHEREAS, " & _
IIf(Forms!foreclosuredetails!LoanType = 5, "Federal Home Loan Mortgage Corporation is the owner of the note secured by said Deed of Trust and appointed the party of the first part with authority to hold, collect and enforce the note; and ", _
IIf(Forms!foreclosuredetails!LoanType = 4, "Federal National Mortgage Association is the owner of the note secured by said Deed of Trust and appointed the party of the first part with authority to hold, collect and enforce the note; and ", _
 "the party of the first part is the holder of the Note secured by said Deed of Trust; and,"))





WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Deed of Appointment SPLS.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Deed of Appointment SPLS.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_EvictionTenant_NTQ_SPS(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qry90DayNoticeSPSWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Eviction Print MD Tenant NTQ SPS.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

'DefendantsAddress = GetAddresses(d!FileNumber, 5, "Defendant=True", NewLine)
FillField WordDoc, "AttySignature", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
FillField WordDoc, "FileNumber", d!FileNumber
'FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!State
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "Purchaser", d!Purchaser
FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Eviction Print MD Tenant NTQ SPS.doc"
Call SaveDoc(WordDoc, d![FileNumber], "Eviction Print MD Tenant NTQ SPS.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_EvictionOwner_NTQ_SPS(Keepopen As Boolean)
'Call MsgBox("This section is under going testing", vbExclamation + vbAbortRetryIgnore, "Don't test this yet")

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qry90DayNoticeSPSWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Eviction Print MD Owner NTQ SPS.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

'DefendantsAddress = GetAddresses(d!FileNumber, 5, "Defendant=True", NewLine)
FillField WordDoc, "AttySignature", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
FillField WordDoc, "FileNumber", d!FileNumber
'FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!State
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "Purchaser", d!Purchaser
FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Eviction Print MD Owner NTQ SPS.doc"
Call SaveDoc(WordDoc, d![FileNumber], "Eviction Print MD Owner NTQ SPS.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_Eviction_MD_NTQ_53(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qry90DayNoticeSPSWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Eviction Print MD NTQ Fifth Third.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

FillField WordDoc, "Sale", Format(d!Sale, "mmmm d, yyyy")
'DefendantsAddress = GetAddresses(d!FileNumber, 5, "Defendant=True", NewLine)
FillField WordDoc, "AttySignature", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
FillField WordDoc, "FileNumber", d!FileNumber
'FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!State
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "BrokerNM", d!BrokerNm
FillField WordDoc, "BrokerPh", FormatPhone(d!BrokerPh)
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "Purchaser", d!Purchaser
FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
FillField WordDoc, "LastDays", Format(LastDay90(Date), "mmmm d, yyyy")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Eviction Print MD NTQ Fifth Third.doc"
Call SaveDoc(WordDoc, d![FileNumber], "Eviction Print MD NTQ Fifth Third.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_Eviction_VA_Owner_NTQ_53(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qry90DayNoticeSPSWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Eviction Print VA Owner NTQ Fifth Third.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

FillField WordDoc, "Sale", Format(d!Sale, "mmmm d, yyyy")
'DefendantsAddress = GetAddresses(d!FileNumber, 5, "Defendant=True", NewLine)
FillField WordDoc, "AttySignature", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
FillField WordDoc, "FileNumber", d!FileNumber
'FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!State
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "BrokerNM", d!BrokerNm
FillField WordDoc, "BrokerPh", FormatPhone(d!BrokerPh)
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "Purchaser", d!Purchaser
FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
FillField WordDoc, "LastDays", Format(LastDay90(Date), "mmmm d, yyyy")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Eviction Print VA Owner NTQ Fifth Third.doc"
Call SaveDoc(WordDoc, d![FileNumber], "Eviction Print VA Owner NTQ Fifth Third.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_Eviction_VA_Tenant_NTQ_53(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qry90DayNoticeSPSWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Eviction Print VA Tenant NTQ Fifth Third.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

FillField WordDoc, "Sale", Format(d!Sale, "mmmm d, yyyy")
'DefendantsAddress = GetAddresses(d!FileNumber, 5, "Defendant=True", NewLine)
FillField WordDoc, "AttySignature", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
FillField WordDoc, "FileNumber", d!FileNumber
'FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!State
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "BrokerNM", d!BrokerNm
FillField WordDoc, "BrokerPh", FormatPhone(d!BrokerPh)
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "Purchaser", d!Purchaser
FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
'FillField WordDoc, "LastDays", Format(LastDay90(Date), "mmmm d, yyyy")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Eviction Print VA Tenant NTQ Fifth Third.doc"
Call SaveDoc(WordDoc, d![FileNumber], "Eviction Print VA Tenant NTQ Fifth Third.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_LossMitigationPrelimSPLS(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim textfinal As String
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
Dim templateName As String
Dim BankName As String


BankName = "Loss Mitigation Preliminary SPLS"
templateName = "Loss Mitigation Prelim - SPLS"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "fillingdate", IIf(Not IsNull(d![Docket]), Format(d![Docket], "mm/d/yyyy"), "_____________________ ")
FillField WordDoc, "borrower", MortgagorNamesOneline(d![CaseList.FileNumber], 2)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", IIf(Not IsNull(d![CourtCaseNumber]), d![CourtCaseNumber], " _____________________ ")
FillField WordDoc, "NameAffianit", "Print Name and Title of Affiant"
'FillField WordDoc, "TitleAffiaitOrInvestor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "Title of Affiant")
'FillField WordDoc, "Line", IIf(d![ClientID] = 6 Or d![ClientID] = 556, "", "______________________________")
FillField WordDoc, "Investor", IIf(d![ClientID] = 6 Or d![ClientID] = 556, d!Investor, "")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1484, "")
'FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ____________________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "                " & "____________________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: _________________________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & BankName & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], BankName & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_Eviction_MD14102b(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

'Set d = CurrentDb.OpenRecordset("SELECT * FROM qryMD14102b WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND EVDetails.Current = TRUE;", dbOpenSnapshot)
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryMD14102bWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE AND EVdetails.current = True;", dbOpenSnapshot)


If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "MD14-102b.dot", False, 0, True)
WordObj.Visible = True

FillField WordDoc, "Case", Nz(d!CourtCaseNumber, "_________")
FillField WordDoc, "TrusteeNames", trusteeNames(0, 2)
FillField WordDoc, "Defendants", MortgagorOwnerNames(0, 2)
FillField WordDoc, "Broker", Nz(d!BrokerNm)

FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "MD14-102b.doc"
Call SaveDoc(WordDoc, d![FileNumber], "MD14-102b.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

'#1219
Public Sub Doc_StatementOfDebtWithFiguresSPLS(Keepopen As Boolean)

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim i As Integer, itemCount As Integer, totalItems As Currency, amountInterest As Currency, itemsFields As String, J As Integer
Dim K As Recordset
Dim Rresult As Long
Dim InterestFrom As String
Dim InterestTo As String

Set d = CurrentDb.OpenRecordset("Select * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Statement of Debt Figures Spls"
templateName = templateName & ".dot"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


WordObj.Visible = False
If MsgBox("Is There A Loan Mod? ", vbYesNo) = vbYes Then
FillField WordDoc, "Mod", ", MODIFIED by Agreement effective " & Format(InputBox(" Effective Date?   Format mm/dd/yyyy"), "mmmm d, yyyy") & " with an amended principal balance of " & Format(InputBox(" Amended Principal Balance? "), "Currency")
Else
FillField WordDoc, "Mod", ""
End If


InterestFrom = InputBox("Interest From")
InterestTo = InputBox("Interest To")

WordObj.Visible = True

FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "Investor", d!Investor

FillField WordDoc, "owner", IIf(d![LoanType] = 4, "Federal National Mortgage Association is the owner", _
IIf(d![LoanType] = 5, "Federal Home Loan Mortgage Corporation is the owner", d![Investor] & " is the owner"))
FillField WordDoc, "InvestorSPLS", IIf(d![Investor] = "Specialized Loan Servicing, LLC", d![Investor], d![Investor] & " by its servicer, Specialized Loan Servicing, LLC")
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "LastPaymentApplied", Format$(d![DateOfDefault], "mmmm d, yyyy")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
'FillField WordDoc, "PBalance", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")


FillField WordDoc, "DateBalance", IIf(IsNull(Forms![Print Statement of Debt]!txtDueDate), "______________", Format$(Forms![Print Statement of Debt]!txtDueDate, "mmmm d"", ""yyyy"))
FillField WordDoc, "ReRecorded", IIf(IsNull(d!Rerecorded), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Liber " & d![Liber2] & ", Folio " & d![Folio2])

totalItems = 0
amountInterest = 0
itemsFields = ""


'Set dd = CurrentDb.OpenRecordset("SELECT Desc, TempSortId, Amount FROM QryStatementOfDebtWithFigurersBOA WHERE FileNumber=" & d![CaseList.FileNumber] & " ORDER BY TempSortId DESC;", dbOpenSnapshot)
Set dd = CurrentDb.OpenRecordset("select * from statementofdebt where filenumber=" & d![CaseList.FileNumber] & " ORDER BY Sort_Desc ASC;", dbOpenDynaset, dbSeeChanges)
If dd.EOF Then      ' no extra lines
    FillField WordDoc, "Line_Items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
    
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items", itemsFields & " "
    
    dd.MoveFirst
    i = 1

     Do While Not dd.EOF
     '   FillField WordDoc, "Item" & i, IIf(dd!Desc = "Interest Due", "Interest Due from " & InterestFrom & " to " & InterestTo & _
     '   " @ " & IIf(Forms![Print Statement of Debt]!chVarRate = 0, Format(Forms![Print Statement of Debt]!InterestRate, "#.000%"), _
     '   " variable rate(s)"), dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency"))
        If dd!Desc = "Interest" Then
           ' FillField WordDoc, "Item" & i, "Interest Due from " & InterestFrom & " to " & InterestTo & " @ " & IIf(Forms![Print Statement of Debt]!chVarRate = 0, Format(Forms![Print Statement of Debt]!InterestRate, "#0.000") & "%", " variable rate(s)")
            FillField WordDoc, "Item" & i, "Interest at " & IIf([Forms]![Print Statement of Debt]![chVarRate] = 0, Format([Forms]![Print Statement of Debt]![InterestRate], "0.000") & "%", " variable rate(s)") & " From " & InterestFrom & " to " & InterestTo & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        Else
            FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        End If
        dd.MoveNext
        i = i + 1
    Loop
End If
    
 dd.Close
 
  Set K = CurrentDb.OpenRecordset("SELECT Desc, Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber], dbOpenSnapshot)
  J = 1
  Do While Not K.EOF
   If K!Desc = "Interest" Then
   amountInterest = amountInterest + Format$(Nz(K!Amount, 0), "Currency")
   End If
      
   totalItems = totalItems + Nz(K!Amount, 0)
   K.MoveNext
   J = J + 1
  Loop
  K.Close

FillField WordDoc, "BalDueDate", IIf(IsNull([Forms]![Print Statement of Debt]![txtDueDate]), "______________", Format$([Forms]![Print Statement of Debt]![txtDueDate], "mmmm d"", ""yyyy"))
FillField WordDoc, "BalanceDue", Format$(d!RemainingPBal + totalItems, "Currency")
FillField WordDoc, "diem", IIf(d![LoanType] = 3, "Monthly", "Diem")
FillField WordDoc, "PerDiemInterest", IIf(IsNull(d!PerDiem), "$_____________", Format$(d!PerDiem, "Currency"))
FillField WordDoc, "InterestRate", IIf(IsNull(d!InterestRate), "____________ %", Format$(d!InterestRate, "#0.000") & "%")
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "NotaryLocation", IIf(IsNull(d!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", d!NotaryLocation)
FillField WordDoc, "amountInterest$", Format$(amountInterest, "Currency")
FillField WordDoc, "jurisdiction", d!Jurisdiction
FillField WordDoc, "liber", d![Liber]
FillField WordDoc, "folio", d![Folio]

'FillField WordDoc, "MortgagorNamesintext", MortgagorNamesIntext(0, 2, 0, 0, -1)



WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Statement of Debt Figures SPLS.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Statement of Debt Figures SPLS.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub
'/#1219

Public Sub Doc_AffCoverSheet(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If


    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Affidavit Cover letter.dot", False, 0, True)
    WordObj.Visible = True


WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]

FillField WordDoc, "ClientLoanT", Forms!foreclosuredetails!LoanNumber
FillField WordDoc, "MortgagorNameT", MortgagorNames(0, 3)
FillField WordDoc, "PropertyAddressT", Forms!foreclosuredetails!PropertyAddress & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt]) & ", " & Forms!foreclosuredetails!City & ", " & Forms!foreclosuredetails!State & " " & FormatZip(Forms!foreclosuredetails!ZipCode)

FillField WordDoc, "ContactNameT", GetLoginName()

FillField WordDoc, "DateofUpload", Format$(Now(), "mmmm d, yyyy")

'FillField WordDoc, "Disclaimer2", IIf(Forms![Case list]!State = "MD", trusteeNames(0, 2) & ", the substitute trustees listed, are employees of the attorney firm", "")
'FillField WordDoc, "Disclaimer", IIf(Not IsNull(Forms![Print Assignment]!cboBOANoteLocation), Forms![Print Assignment]!cboBOANoteLocation.Column(0), IIf(Forms!ForeclosureDetails!DocBackOrigNote = -1, "The firm is in possession of the original note", "The firm is NOT in possession of the original note"))

'FillField WordDoc, "Disclaimer", IIf(Forms!ForeclosureDetails!DocBackOrigNote = -1, "The firm is in possession of the original note", "The firm is NOT in possession of the original note")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Affidavit Cover letter.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Affidavit Cover letter.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub DOC_AffidavitofLienInstrumentNationStar(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim FileNumber  As Long


FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Affidavit of Lien Instrument Nationtar.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
'WordDoc.Bookmarks("ProName").Select
'WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)

FillField WordDoc, "Investor", IIf(d![ClientID] = 385 And d![Investor] = "Nationstar Mortgage LLC", "Nationstar maintains records for the loan that is secured by the mortgage or deed of trust being foreclosed in this action. ", _
"Nationstar services and maintains records on behalf of " & d![Investor] & ", the secured party to the mortgage or deed of trust being foreclosed in this action.")

FillField WordDoc, "Prior Servicer", IIf(Forms![Prior Servicer]!ChPrior, "           Before the servicing of this loan transferred to Nationstar, " & [Forms]![Prior Servicer]![TxtPriorServicer] & _
" (Prior Servicer) was the servicer for the loan and it maintained the loan servicing records.  When Nationstar began servicing this loan, Prior Servicer's records for the loan were integrated and boarded into Nationstar's systems," & _
" such that Prior Servicer's records, including the collateral file, payment histories, communication logs, default letters, information," & _
"and documents concerning the Loan are now integrated into Nationstar's business records.  Nationstar maintains quality control and " & _
"verification procedures as part of the boarding process to ensure the accuracy of the boarded records.  It is the regular business practice " & _
"of Nationstar to integrate prior servicers' records into Nationstar's business records and to rely upon those boarded records in providing " & _
"its loan servicing functions.  These Prior Servicer records have been integrated and are relied upon by Nationstar as part of Nationstar's " & _
"business records.", "")


FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "loan servicing", IIf(d!Investor = "Nationstar Mortgage LLC", "", " loan servicing")
FillField WordDoc, "DOTdate", Format(d!DOTdate, "mmmm dd, yyyy")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Affidavit Lien Instrument Nationtar.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Affidavit Lien Instrument Nationtar.doc")

If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
 
'
'If Application.CurrentProject.AllForms("Prior Servicer").IsLoaded Then
'    DoCmd.Close acForm, "Prior Servicer"
'End If

End Sub

Public Sub Doc_BaileeLTR(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim dd As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryGreenTreeBailee WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Green Tree Bailee Letter.dot", False, 0, True)
WordObj.Visible = True


FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "PropertyAddress", d![PropertyAddress]
FillField WordDoc, "PropertyAddress2", d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "ProjectName", d!PrimaryDefName
FillField WordDoc, "GetLoginName", GetLoginName()

Set dd = CurrentDb.OpenRecordset("SELECT * FROM BaileeDocs WHERE FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If dd.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

FillField WordDoc, "OrigMort", IIf(dd!OriginalMortgage = True, "X", "  ")
FillField WordDoc, "CopyMort", IIf(dd!CopyMortgage = True, "X", "  ")
FillField WordDoc, "OrigNote", IIf(dd!OriginalNote = True, "X", "  ")
FillField WordDoc, "CopyNote", IIf(dd!CopyNote = True, "X", "  ")
FillField WordDoc, "OrigLNA", IIf(dd!OriginalLNA = True, "X", "  ")
FillField WordDoc, "CopyLNA", IIf(dd!CopyLNA = True, "X", "  ")
FillField WordDoc, "OrigMod", IIf(dd!OriginalModification = True, "X", "  ")
FillField WordDoc, "CopyMod", IIf(dd!CopyModification = True, "X", "  ")
FillField WordDoc, "OrigOther", IIf(dd!OriginalOther = True, "X", "  ")
FillField WordDoc, "CopyOther", IIf(dd!CopyMortgage = True, "X", "  ")



WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Green Tree Bailee Letter.dot"
Call SaveDoc(WordDoc, Forms![Case List].FileNumber, "Green Tree Bailee Letter.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
dd.Close
End Sub

Public Sub Doc_BKBOAMotionForRelief(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim dd As Recordset
Dim ddd As Recordset
Dim FileNumber As Long, DebtorsPlural As Boolean
Dim i As Integer, itemCount As Integer, totalItems As Currency, itemsFields As String
Dim delinquent As Currency
Dim delinquenttotal, encumbrances As Currency
newNextNumber = 1
NextNumber = 1

FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBOAMFr WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Motion for Relief BOA.dot", False, 0, True)
WordObj.Visible = True
WordObj.ScreenUpdating = False
WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
'WordDoc.Bookmarks("AttorneyInfo").Select
'WordDoc.Bookmarks("AttorneyInfo").Range.Text = "Diane Rosenberg" & vbCr & "VA Bar 35237"

'Judge = Right$(UCase$(d!CaseNo), 3)
'CoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13)



encumbrances = Nz((InputBox("Sum of known encumbrances?")), 0)
DebtorsPlural = (CountNames(FileNumber, "BKDebtor = True AND (Owner=True OR Mortgagor=True)") > 1)

FillField WordDoc, "Header", _
    "IN THE UNITED STATES BANKRUPTCY COURT" & vbCr & _
    "FOR THE " & UCase$(d!Name) & vbCr & _
    d!Location

FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & IIf(DebtorsPlural, "s", "")

FillField WordDoc, "InvestorAddr", UCase$(d!Investor) & vbCr & RemoveLF(d!InvestorAddress)
FillField WordDoc, "Respondents", _
    GetAddresses(0, 4, _
        IIf(d![BKdetails.Chapter] = 13, _
            "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", _
            "BKDebtor=True AND (Owner=True OR Mortgagor=True)"), vbCr) & _
    IIf(d![BKdetails.Chapter] = 7, _
        vbCr & vbCr & "and" & vbCr & vbCr & FormatName("", _
                                                        UCase$(Nz(d!First)), _
                                                        UCase$(Nz(d!Last)), _
                                                        ", CHAPTER 7 TRUSTEE", _
                                                        d!Address, _
                                                        d!Address2, _
                                                        d![BKTrustees.City], _
                                                        d![BKTrustees.State], _
                                                        d!Zip, _
                                                        vbCr), _
        "")

FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]
FillField WordDoc, "Date", Format(Date, "mmmm d yyyy")
FillField WordDoc, "AndAttorney", IIf(IsNull(d!AttorneyLastName), "", "and Debtor's attorney ")
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "GetNextNumber1", GetNextNumber()
FillField WordDoc, "GetNextNumber2", GetNextNumber()
FillField WordDoc, "GetNextNumber3", GetNextNumber()
FillField WordDoc, "GetNextNumber4", GetNextNumber()

FillField WordDoc, "OtherCostsParagraph", IIf(d![otherCosts] > 0 Or IsNull(d![otherCosts]) = False, GetNextNumber() & "The foregoing Other Costs consist of the following:", "")
FillField WordDoc, "OtherCostsTable", IIf(d![otherCosts] > 0 Or IsNull(d![otherCosts]) = False, "Pre-Petition Or Post-Petition   | Transaction Date | Fee Description | Amount", "")
If d!otherCosts > 0 Or IsNull(d!otherCosts) = False Then
' Do nothing
Else
'Other costs Blank
NextNumber = NextNumber - 2
End If

FillField WordDoc, "GetNextNumber7", GetNextNumber()
FillField WordDoc, "GetNextNumber8", GetNextNumber()
FillField WordDoc, "GetNextNumber9", GetNextNumber()

If MsgBox("Is Debtor being evaluated for Loss Mit?", vbYesNo) = vbYes Then
    FillField WordDoc, "LossMit", GetNextNumber() & "As of the date of this Motion, the Debtor is being evaluated for a loss mitigation option.   Additional information regarding such evaluation is available upon request."
Else
    FillField WordDoc, "LossMit", ""
End If

FillField WordDoc, "GetNextNumber10", GetNextNumber()
FillField WordDoc, "GetNextNumber11", GetNextNumber()
FillField WordDoc, "ConvChapter", IIf(IsNull(d![ConvChapter]) = True, d![BKdetails.Chapter], d![ConvChapter])
FillField WordDoc, "FilingDate", Format(d![DateofFiling], "mmmm d"", ""yyyy")
FillField WordDoc, "ConvDate", IIf(IsNull(d![ConvDate]) = True, "", "An orderF converting the case to a case under chapter 7 was entered on " & Format(d![ConvDate], "mmmm d"", ""yyyy") & ".")
FillField WordDoc, "OriginalPbal", Format(d![OriginalPBal], "Currency")
FillField WordDoc, "SecurityDoc", Nz(d!SecurityDoc)
FillField WordDoc, "Jurisdiction", Forms![Case List]!JurisdictionID.Column(1)
FillField WordDoc, "SecurityParagraph", IIf((IsNull(d![SecurityDoc]) = False And d![AssignByDOT] = 1), GetNextNumber() & "All rights and remedies under the Security Instrument have been assigned to the Movant pursuant to that certain assignment of Security Instrument , a copy of which is attached hereto as Exhibit C.", "")
FillField WordDoc, "boaAsOfDate", Format(d![boaAsOfDate], "mmmm d"", ""yyyy")
FillField WordDoc, "PrincipalBalance", Format(d!principalBalance, "Currency")
FillField WordDoc, "AccruedInterest", Format(d!accruedInterest, "Currency")
FillField WordDoc, "LateCharges", Format(d!lateCharges, "Currency")
FillField WordDoc, "InsurancePremiums", Format(d!insurancePremiums, "Currency")
FillField WordDoc, "DebtorTaxesInsurance", Format(d!debtorTaxesInsurance, "Currency")
FillField WordDoc, "OtherCosts", Format(d!otherCosts, "Currency")

FillField WordDoc, "PartialPayments", Format(d!partialPayments, "Currency")
FillField WordDoc, "OutstandingObligations", Format(d!outstandingObligations, "Currency")
FillField WordDoc, "BOAAsofDate", Format(d![boaAsOfDate], "mmmm d"", ""yyyy")
FillField WordDoc, "Fee", Format(d![Fee], "Currency")
FillField WordDoc, "MarketValue", Format(d![MarketValue], "Currency")
FillField WordDoc, "Source", Nz(d!Source)
FillField WordDoc, "Encumbrances", Format(encumbrances, "Currency")
FillField WordDoc, "GetNewNextNumber1", GetNewNextNumber()
FillField WordDoc, "GetNewNextNumber2", GetNewNextNumber()
FillField WordDoc, "GetNewNextNumber3", GetNewNextNumber()
FillField WordDoc, "GetnewNextNumber4", GetNewNextNumber()


totalItems = 0
itemsFields = ""
If (d![otherCosts] = 0 Or IsNull(d![otherCosts]) = True) Then
    
      FillField WordDoc, "Line_Items", ""
      FillField WordDoc, "PreTitle", ""
      FillField WordDoc, "PostTitle", ""
      FillField WordDoc, "PreTotal", ""
      FillField WordDoc, "PostTotal", ""
'skip other costs
Else
    Set dd = CurrentDb.OpenRecordset("SELECT PrePost, TimeStamp,  Desc, Amount FROM BKDebt WHERE FileNumber=" & d![FileNumber] & " ORDER BY Timestamp;", dbOpenSnapshot)
    If dd.EOF Then      ' no extra lines
        FillField WordDoc, "Line_Items", ""
    Else
        dd.MoveLast
        itemCount = dd.RecordCount
        ' Make enough lines
        For i = 1 To itemCount
            itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
        Next i
        itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
        FillField WordDoc, "Line_Items", itemsFields
        dd.MoveFirst
        i = 1
        Do While Not dd.EOF
            FillField WordDoc, "Item" & i, dd!PrePost & vbTab & Format(dd!Timestamp, "mm/dd/yyyy") & vbTab & dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
     
            totalItems = totalItems + Nz(dd!Amount, 0)
            dd.MoveNext
            i = i + 1
        Loop
        FillField WordDoc, "PreTitle", "Pre-Petition Total:"
        FillField WordDoc, "PostTitle", "Post-Petition Total:"
        FillField WordDoc, "PreTotal", Format(Nz(DSum("Amount", "BKDebt", "Prepost = 'Pre-Petition' AND FileNumber=" & [Forms]![Case List]![FileNumber] & ""), 0), "Currency")
        FillField WordDoc, "PostTotal", Format(Nz(DSum("Amount", "BKDebt", "Prepost = 'Post-Petition' AND FileNumber=" & [Forms]![Case List]![FileNumber] & ""), 0), "Currency")
        'FillField WordDoc, "Total", Format(totalItems, "Currency")
    End If
    dd.Close

End If
totalItems = 0
itemsFields = ""
Set ddd = CurrentDb.OpenRecordset("SELECT numMissedPayments, missedPaymentsTo,  missedPaymentsFrom, MonthlyAmount FROM BKMissedPayments WHERE FileNumber=" & d![FileNumber] & " ORder BY MissedPaymentsFrom;", dbOpenSnapshot)
If ddd.EOF Then      ' no extra lines
    FillField WordDoc, "Line_Items", ""
Else
    ddd.MoveLast
    itemCount = ddd.RecordCount
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items2", itemsFields
    ddd.MoveFirst
    i = 1
    Do While Not ddd.EOF
        delinquent = ddd!numMissedPayments * ddd!MonthlyAmount
        FillField WordDoc, "Item" & i, ddd!numMissedPayments & vbTab & ddd!missedPaymentsFrom & _
        vbTab & ddd!missedPaymentsTo & vbTab & Format$(Nz(ddd!MonthlyAmount, 0), "Currency") & vbTab & Format(delinquent, "Currency")
        
 
        totalItems = totalItems + Nz(ddd!MonthlyAmount, 0)
        delinquenttotal = delinquent + delinquenttotal
        ddd.MoveNext
        i = i + 1
    Loop
    FillField WordDoc, "Total", Format(delinquenttotal + d!partialPayments, "Currency")

End If
ddd.Close

FillField WordDoc, "TrusteeName", FormatName("", d![First], d![Last] & IIf(d![BKdetails.Chapter] = 11, "", ", Trustee"), "", d![Address], d![Address2], d![BKTrustees.City], d![BKTrustees.State], d![Zip])
FillField WordDoc, "AttorneyName", d![AttorneyFirstName] & " " & d![AttorneyLastName] & IIf(IsNull(d![AttorneyLastName]), "", ", Esquire")
FillField WordDoc, "AttorneyFirmName", IIf(IsNull(d![AttorneyFirm]), "", d![AttorneyFirm]) & vbCr & d![AttorneyAddress]





    If Forms!BankruptcyPrint!chElectronicSignature Then
        FillField WordDoc, "ElectronicSignature", "/s/ " & Forms!BankruptcyPrint!cbxAttorney
    Else
        FillField WordDoc, "ElectronicSignature", "_______________________________"
    End If
    FillField WordDoc, "AttorneySignature", Forms!BankruptcyPrint!cbxAttorney
FillField WordDoc, "BkService", BKService(0)

WordObj.Selection.HomeKey wdStory, wdMove
WordObj.ScreenUpdating = True
WordDoc.SaveAs EMailPath & "BOA Motion For Relief.doc"
Call SaveDoc(WordDoc, d![FileNumber], "BOA Motion for Relief.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_ClerkCoverLTR(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim dd As Recordset




Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Clerk Cover Letter.dot", False, 0, True)
WordObj.Visible = True


FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
FillField WordDoc, "PropertyAddress", [Forms]![foreclosuredetails]![PropertyAddress] & IIf(Len(Forms!foreclosuredetails![Fair Debt] & "") = 0, "", ", " & Forms!foreclosuredetails![Fair Debt])
FillField WordDoc, "CIty", [Forms]![foreclosuredetails]![City]
FillField WordDoc, "State", [Forms]![foreclosuredetails]![State]
FillField WordDoc, "Zip", FormatZip([Forms]![foreclosuredetails]![ZipCode])
FillField WordDoc, "CaseNumber", Forms![foreclosuredetails]![CourtCaseNumber]
FillField WordDoc, "Filenumber", [Forms]![Case List]![FileNumber]
FillField WordDoc, "Jurisdiction", Forms![Case List]![JurisdictionID].Column(1)

FillField WordDoc, "GetLoginName", GetLoginName()



WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Clerk Cover Letter.dot"
Call SaveDoc(WordDoc, Forms![Case List].FileNumber, "Clerk Cover Letter.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing

End Sub
Public Sub Doc_BKChaseMotionForRelief(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset
Dim dd As Recordset
Dim ddd As Recordset
Dim FileNumber As Long, DebtorsPlural As Boolean
Dim response As Integer
Dim rightToForecloseLanguage As String
Dim i As Integer, itemCount As Integer, totalItems As Currency, itemsFields As String
Dim delinquent As Currency
Dim delinquenttotal, encumbrances As Currency

newNextNumber = 1
NextNumber = 1
lettercounter = 1
FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryMFRChaseWord WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Motion for Relief Chase.dot", False, 0, True)
WordObj.Visible = True
WordObj.ScreenUpdating = False
'WordDoc.Bookmarks("FileNumber").Select
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]

FillField WordDoc, "Header", "IN THE UNITED STATES BANKRUPTCY COURT" & vbCr & _
    "FOR THE " & UCase$(d!Name) & vbCr & d!Location
FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]
FillField WordDoc, "Date", Format(Date, "mmmm d yyyy")
FillField WordDoc, "AndAttorney", IIf(IsNull(d!AttorneyLastName), "", "and Debtor's attorney ")
FillField WordDoc, "Investor", d!Investor & IIf(d![Districts.ID] = 7 Or d![Districts.ID] = 6, " ", _
    " its successors and/or assigns,") & " movant, by its attorneys, " & [Forms]![BankruptcyPrint]![cbxAttorney] & ", and " & FirmName() & ", and respectfully represents as follows:"
FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & IIf(DebtorsPlural, "s", "")
FillField WordDoc, "InvestorAddr", UCase$(d!Investor) & vbCr & RemoveLF(d!InvestorAddress)
FillField WordDoc, "Respondents", _
    GetAddresses(0, 4, _
        IIf(d![BKdetails.Chapter] = 13, _
            "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", _
            "BKDebtor=True AND (Owner=True OR Mortgagor=True)"), vbCr)
 '-----------------------------------------------------------
FillField WordDoc, "GetNextNumber1", GetNextNumber()
Dim chCoDebtor As Boolean
 If GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13 Then chCoDebtor = True
FillField WordDoc, "chCoDebtor", IIf(chCoDebtor, " and 1301", "")
'2.
FillField WordDoc, "GetNextNumber2", GetNextNumber()
FillField WordDoc, "DateOFFiling", Format$(d![DateofFiling], "mmmm d"", ""yyyy")

If CountNames([FileNumber], "BKDebtor = True And (Owner=True Or Mortgagor=True)") > 1 Then DebtorsPlural = True

FillField WordDoc, "Debtor", GetNames(0, 2, "BKDebtor=True AND (Owner=True OR Mortgagor=True)") & " (""Debtor" & IIf(DebtorsPlural, "s", "") & """)"
FillField WordDoc, "chapter", IIf(IsNull(d![ConvChapter]), d![BKdetails.Chapter], d![ConvChapter])
FillField WordDoc, "Value2", IIf(IsNull(d![ConvDate]), "", ", which case later converted to Chapter " & d![BKdetails.Chapter] & " on ") & Format$(d!ConvDate, "mmmm d"", ""yyyy")
'3.
FillField WordDoc, "GetNextNumber3", IIf(d![BKdetails.Chapter] = 11, "", GetNextNumber() & d![First] _
    & " " & d![Last] & " is the Chapter " & d![BKdetails.Chapter] & " trustee of the Debtor" & IIf([DebtorsPlural], "s'", "'s") & " estate.")
'4.
FillField WordDoc, "GetNextNumber4", GetNextNumber()
FillField WordDoc, "value4", GetNames(0, 2, "BKCoDebtor=True") & " (""Co-Debtor" & IIf(CountNames(0, "BKCoDebtor=True") > 1, "s", "") & """) " & IIf(CountNames(0, "BKCoDebtor=True") > 1, "are co-debtors", "is a co-debtor")
'5.
FillField WordDoc, "GetNextNumber5", GetNextNumber()
FillField WordDoc, "value5", IIf([DebtorsPlural], "s", "")
FillField WordDoc, "value6", IIf(d!RealEstate, "parcel of " & IIf(d![Leasehold], "leasehold", "fee simple"), "")
FillField WordDoc, "value7", IIf(IsNull(d![LegalDescription]), "", "with a legal description of """ & d![LegalDescription] & """ also ")
FillField WordDoc, "PropertyAddress", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "value8", IIf(IsNull(d![PropertyDesc]), "", d![PropertyDesc])
'6.
FillField WordDoc, "GetNextNumber6", GetNextNumber()
FillField WordDoc, "AssignmentInfo", AssignmentInfo() & IIf([Forms]![Case List]![ClientID] = 385, "", " The documents evidencing the movant's security interest are attached hereto. ")
'FillField WordDoc, "AssignmentInfo", IIf([Forms]![Case List]![ClientID] = 385, "", " The documents evidencing the movant's security interest are attached hereto. ")

If [Forms]![Print Motion for Relief]![Text33] = "Yes" Then
    If [Forms]![Print Motion for Relief]![Text31] = "Yes" Then
        FillField WordDoc, "checktxt33", GetNextNumber()
        FillField WordDoc, "txtOf33", " JPMorgan Chase Bank, N.A., services the loan on the property referenced in this Motion for Relief. In the event the automatic stay in this case is lifted/set aside, this case dismisses, and/or the debtor obtains a discharge and a foreclosure action is commenced on the mortgaged property, the foreclosure will be conducted in the name of " & d![Investor] & "."
    ElseIf [Forms]![Print Motion for Relief]![Text34] = "Yes" Then
        FillField WordDoc, "checktxt33", GetNextNumber()
        FillField WordDoc, "txtOf33", " JPMorgan Chase Bank, N.A., services the loan on the property referenced in this Motion for Relief. In the event the automatic stay in this case is lifted/set aside, this case dismisses, and/or the debtor obtains a discharge and a foreclosure action is commenced on the mortgaged property, the foreclosure will be conducted in the name of " & d![Investor] & "."
    End If
Else
    If [Forms]![Print Motion for Relief]![Text34] = "No" Then
        FillField WordDoc, "checktxt33", GetNextNumber()
        FillField WordDoc, "txtOf33", " JPMorgan Chase Bank, N.A., services the loan on the property referenced in this Motion for Relief. In the event the automatic stay in this case is lifted/set aside, this case dismisses, and/or the debtor obtains a discharge and a foreclosure action is commenced on the mortgaged property, the foreclosure will be conducted in the name of " & d![Investor] & "."
    End If
End If

FillField WordDoc, "BeforeNextNo", IIf([Forms]![Print Motion for Relief]![Text31] = "Yes", "Said entity is unable to find the promissory note and will seek to prove the promissory note using a lost note affidavit.", _
IIf([Forms]![Print Motion for Relief]![Text31] = "No" And [Forms]![Print Motion for Relief]![Text34] = "Yes", "Said entity, directly or through an agent, has possession of the promissory note. The promissory note is either made payable to said entity or has been duly endorsed.", _
IIf([Forms]![Print Motion for Relief]![Text31] = "No" And [Forms]![Print Motion for Relief]![Text34] = "No", "Said entity, directly or through an agent, has possession of the promissory note. Said entity will enforce the promissory note as transferee in possession.", "")))

FillField WordDoc, "GetNextNumber8", GetNextNumber()
FillField WordDoc, "Value9", IIf(d![RealEstate], DOTWord(d![DOT]), IIf(IsNull(d![PropertyContract]), "", d![PropertyContract]))
FillField WordDoc, "DueDate", Format$([Forms]![Print Motion for Relief]![DueDate])
FillField WordDoc, "amount", Format$([Forms]![Print Motion for Relief]![Amount], "Currency")

FillField WordDoc, "GetNextNumber9", GetNextNumber()
FillField WordDoc, "Value10", IIf(DebtorsPlural, "s are", " is")
FillField WordDoc, "Value11", IIf(d![RealEstate], "Note and Mortgage", IIf(IsNull(d![PropertyContract]), "", d![PropertyContract]))

FillField WordDoc, "GetNextNumber10", GetNextNumber()
FillField WordDoc, "PaymentMonth", [Forms]![Print Motion for Relief]![Payment]
FillField WordDoc, "Value13", IIf(d![RealEstate], "Debtor" & IIf(DebtorsPlural, "s'", "'s") & " residence", "subject property") & " is dissipating."

FillField WordDoc, "GetNextNumber11", GetNextNumber()
FillField WordDoc, "GetNextNumber12", GetNextNumber()
FillField WordDoc, "Section", IIf(chCoDebtor, "Sections 362 and 1301", "Section 362")
FillField WordDoc, "NoteDeedTrust", IIf(d![RealEstate], "Note and " & DOTWord(d!DOT), IIf(IsNull(d![PropertyContract]), "", d![PropertyContract]))
FillField WordDoc, "GetNextNumber13", GetNextNumber()
FillField WordDoc, "Section1", IIf([chCoDebtor], "Sections 362 and 1301", "Section 362")
FillField WordDoc, "Section2", IIf(d![RealEstate], "Note and " & DOTWord(d!DOT), IIf(IsNull(d![PropertyContract]), "", d![PropertyContract])) & "."

If [Forms]![Print Motion for Relief]![chReorganization] Then
    FillField WordDoc, "GetNextNumber14", GetNextNumber()
    FillField WordDoc, "txtOf14", "The subject property is not necessary for an effective reorganization."
Else
    FillField WordDoc, "GetNextNumber14", ""
    FillField WordDoc, "txtOf14", ""
End If

FillField WordDoc, "investor2", d![Investor] & IIf(d![Districts.ID] = 7 Or d![Districts.ID] = 6, " ", " its successors and/or assigns,")
FillField WordDoc, "value14", IIf(d![RealEstate], "a foreclosure sale, accept a deed in lieu or agree to a short sale of the real property and improvements located at " & d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode]), "repossession and sale of the " & IIf(IsNull(d![PropertyDesc]), "", d![PropertyDesc]) & ";")

FillField WordDoc, "ElectronicSignature", IIf(Forms!BankruptcyPrint!chElectronicSignature, "/s/ " & Forms!BankruptcyPrint!cbxAttorney, "")
FillField WordDoc, "AttorneySignature", [Forms]![BankruptcyPrint]![cbxAttorney]
FillField WordDoc, "Date", Format$(Date, "mmmm d, yyyy")
FillField WordDoc, "BKService", BKService(0)
'FillField WordDoc, "TrusteeName", FormatName("", d![First], d![Last] & IIf(d![Chapter] = 11, "", ", Trustee"), "", d!Address, d![Address2], d![City], d![State], d![Zip])
FillField WordDoc, "TrusteeName", FormatName("", d![First], d![Last] & IIf(d![BKdetails.Chapter] = 11, "", ", Trustee"), "", d![Address], d![Address2], d![BKTrustees.City], d![BKTrustees.State], d![Zip])
FillField WordDoc, "AttorneyName", d![AttorneyFirstName] & " " & d![AttorneyLastName] & IIf(IsNull(d![AttorneyLastName]), "", ", Esquire")
FillField WordDoc, "AttorneyFirmName", IIf(IsNull(d![AttorneyFirm]), "", d![AttorneyFirm])
FillField WordDoc, "AttAddr", d!AttorneyAddress

WordObj.Selection.HomeKey wdStory, wdMove
WordObj.ScreenUpdating = True
End Sub

Public Sub Doc_BKBOAMotionForReliefCh13(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim dd As Recordset
Dim ddd As Recordset
Dim FileNumber As Long, DebtorsPlural As Boolean
Dim response As Integer
Dim rightToForecloseLanguage As String
Dim i As Integer, itemCount As Integer, totalItems As Currency, itemsFields As String
Dim delinquent As Currency
Dim delinquenttotal, encumbrances As Currency

newNextNumber = 1
NextNumber = 1
lettercounter = 1
FileNumber = Forms![Case List]!FileNumber
Set d = CurrentDb.OpenRecordset("SELECT * FROM qryBOAMFr13 WHERE CaseList.FileNumber=" & FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Motion for Relief BOA CH13.dot", False, 0, True)
WordObj.Visible = True
WordObj.ScreenUpdating = False
WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
'WordDoc.Bookmarks("AttorneyInfo").Select
'WordDoc.Bookmarks("AttorneyInfo").Range.Text = "Diane Rosenberg" & vbCr & "VA Bar 35237"

'Judge = Right$(UCase$(d!CaseNo), 3)
'CoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And d![BKdetails.Chapter] = 13)



encumbrances = Nz((InputBox("Sum of known encumbrances?")), 0)

rightToForecloseLanguage = "Enter a Numeric option (1, 2, or 3)" & vbCr & vbCr
rightToForecloseLanguage = rightToForecloseLanguage + "Option 1. The Note is either made payable to Movant or has been duly endorsed." & vbCr & vbCr
rightToForecloseLanguage = rightToForecloseLanguage + "Option 2. Movant will enforce the Note as transferee in possession." & vbCr & vbCr
rightToForecloseLanguage = rightToForecloseLanguage + "Option 3. Movant is unable to find the Note and will seek to prove the Note using a lost note affidavit."

response = InputBox(rightToForecloseLanguage, "Which Right to Foreclose Language?")

Select Case response

    Case 1
        FillField WordDoc, "Response", "The Note is either made payable to Movant or has been duly endorsed."
    Case 2
        FillField WordDoc, "Response", "Movant will enforce the Note as transferee in possession."
    Case 3
        FillField WordDoc, "Response", "Movant is unable to find the Note and will seek to prove the Note using a lost note affidavit."
End Select

DebtorsPlural = (CountNames(FileNumber, "BKDebtor = True AND (Owner=True OR Mortgagor=True)") > 1)

FillField WordDoc, "Header", _
    "IN THE UNITED STATES BANKRUPTCY COURT" & vbCr & _
    "FOR THE " & UCase$(d!Name) & vbCr & _
    d!Location

FillField WordDoc, "Debtors", UCase$(DebtorNames(0, 3, vbCr)) & vbCr & _
    "     Debtor" & IIf(DebtorsPlural, "s", "")

FillField WordDoc, "InvestorAddr", UCase$(d!Investor) & vbCr & RemoveLF(d!InvestorAddress)
FillField WordDoc, "Respondents", _
    GetAddresses(0, 4, _
        IIf(d![BKdetails.Chapter] = 13, _
            "(BKDebtor=True OR BKCoDebtor=True) AND (Owner=True OR Mortgagor=True)", _
            "BKDebtor=True AND (Owner=True OR Mortgagor=True)"), vbCr) & _
    IIf(d![BKdetails.Chapter] = 7, _
        vbCr & vbCr & "and" & vbCr & vbCr & FormatName("", _
                                                        UCase$(Nz(d!First)), _
                                                        UCase$(Nz(d!Last)), _
                                                        ", CHAPTER 7 TRUSTEE", _
                                                        d!Address, _
                                                        d!Address2, _
                                                        d![BKTrustees.City], _
                                                        d![BKTrustees.State], _
                                                        d!Zip, _
                                                        vbCr), _
        "")

FillField WordDoc, "CaseNo", d!CaseNo
FillField WordDoc, "Chapter", d![BKdetails.Chapter]
FillField WordDoc, "Date", Format(Date, "mmmm d yyyy")
FillField WordDoc, "AndAttorney", IIf(IsNull(d!AttorneyLastName), "", "and Debtor's attorney ")
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![FCdetails.City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "GetNextNumber1", GetNextNumber()
FillField WordDoc, "GetNextNumber2", GetNextNumber()
FillField WordDoc, "GetNextNumber3", GetNextNumber()
FillField WordDoc, "GetNextNumber4", GetNextNumber()
FillField WordDoc, "SecurityParagraph", IIf((IsNull(d![SecurityDoc]) = False And d![AssignByDOT] = 1), GetNextNumber() & "All rights and remedies under the Deed of Trust have been assigned to the Movant pursuant to that certain assignment of Deed of Trust, a copy of which is attached hereto as Exhibit C.", "")

FillField WordDoc, "GetNextNumber5", GetNextNumber()
FillField WordDoc, "GetNextNumber6", GetNextNumber()


FillField WordDoc, "OtherCostsParagraph", IIf(d![otherCosts] > 0, GetNextNumber() & "The foregoing Other Costs consist of the following:", "")
FillField WordDoc, "OtherCostsTable", IIf(d![otherCosts] > 0, "Pre-Petition Or   | Transaction Date | Fee Description                          | Amount", "")
FillField WordDoc, "OtherCostsTable2", IIf(d![otherCosts] > 0, "Post-Petition", "")
FillField WordDoc, "OtherCostsTable3", IIf(d![otherCosts] > 0, "__________________________________________________________________", "")
If d!otherCosts > 0 Then
' Do nothing
Else
'Other costs Blank
NextNumber = NextNumber - 1
End If

FillField WordDoc, "GetNextNumber7", GetNextNumber()
FillField WordDoc, "GetNextNumber8", GetNextNumber()
FillField WordDoc, "GetNextNumber9", GetNextNumber()


FillField WordDoc, "GetNextNumber10", GetNextNumber()


If MsgBox("Is Debtor being evaluated for Loss Mit?", vbYesNo) = vbYes Then
    FillField WordDoc, "LossMit", GetNextNumber() & "As of the date of this Motion, the Debtor is being evaluated for a loss mitigation option.   Additional information regarding such evaluation is available upon request."
Else
    FillField WordDoc, "LossMit", ""
End If
FillField WordDoc, "GetNextNumber11", GetNextNumber()
FillField WordDoc, "GetNextNumber12", GetNextNumber()


FillField WordDoc, "ConvChapter", IIf(IsNull(d![ConvChapter]) = True, d![BKdetails.Chapter], d![ConvChapter])
FillField WordDoc, "FilingDate", Format(d![DateofFiling], "mmmm d"", ""yyyy")
FillField WordDoc, "ConvDate", IIf(IsNull(d![ConvDate]) = True, "", "An order converting the case to a case under chapter 7 was entered on " & Format(d![ConvDate], "mmmm d"", ""yyyy") & ".")
FillField WordDoc, "OriginalPbal", Format(d![OriginalPBal], "Currency")
FillField WordDoc, "SecurityDoc", Nz(d!SecurityDoc)
FillField WordDoc, "Jurisdiction", Forms![Case List]!JurisdictionID.Column(1)
FillField WordDoc, "boaAsOfDate", Format(d![boaAsOfDate], "mmmm d"", ""yyyy")
FillField WordDoc, "PrincipalBalance", Format(d!principalBalance, "Currency")
FillField WordDoc, "AccruedInterest", Format(d!accruedInterest, "Currency")
FillField WordDoc, "LateCharges", Format(d!lateCharges, "Currency")
FillField WordDoc, "InsurancePremiums", Format(d!insurancePremiums, "Currency")
FillField WordDoc, "DebtorTaxesInsurance", Format(d!debtorTaxesInsurance, "Currency")
FillField WordDoc, "OtherCosts", Format(d!otherCosts, "Currency")

FillField WordDoc, "PostPetitionPartialPayments", Format(d!postPetitionPartialPayments, "Currency")
FillField WordDoc, "PartialPayments", Format(d!partialPayments, "Currency")
FillField WordDoc, "OutstandingObligations", Format(d!outstandingObligations, "Currency")
FillField WordDoc, "BOAAsofDate", Format(d![boaAsOfDate], "mmmm d"", ""yyyy")
FillField WordDoc, "Fee", Format(d![Fee], "Currency")
FillField WordDoc, "MarketValue", Format(d![MarketValue], "Currency")
FillField WordDoc, "Source", Nz(d!Source)
FillField WordDoc, "Encumbrances", Format(encumbrances, "Currency")
FillField WordDoc, "GetNewNextNumber1", GetNewNextNumber()
FillField WordDoc, "GetNewNextNumber2", GetNewNextNumber()
FillField WordDoc, "GetNewNextNumber3", GetNewNextNumber()
FillField WordDoc, "GetnewNextNumber4", GetNewNextNumber()
FillField WordDoc, "CoDebtor1", IIf(DebtorsPlural = True, "AND CO-DEBTOR STAY", "")
FillField WordDoc, "CoDebtor2", IIf(DebtorsPlural = True, "AND 1301", "")
FillField WordDoc, "CoDebtor3", IIf(DebtorsPlural = True, "and co-debtor stay", "")
FillField WordDoc, "PlanCOnfirmDate", Format(d![PlanConfirmDate], "mmmm d, yyyy")
FillField WordDoc, "GetNextLetter", GetNextLetter()
FillField WordDoc, "GetNextLetter2", GetNextLetter()
FillField WordDoc, "GetNextLetter3", GetNextLetter()
FillField WordDoc, "GetNextLetter4", GetNextLetter()
FillField WordDoc, "GetNextLetter5", GetNextLetter()



totalItems = 0
itemsFields = ""
If (d![otherCosts] = 0 Or IsNull(d![otherCosts]) = True) Then
    
      FillField WordDoc, "Line_Items", ""
      FillField WordDoc, "PreTitle", ""
      FillField WordDoc, "PostTitle", ""
      FillField WordDoc, "PreTotal", ""
      FillField WordDoc, "PostTotal", ""
'skip other costs
Else
    Set dd = CurrentDb.OpenRecordset("SELECT PrePost, TimeStamp,  Desc, Amount FROM BKDebt WHERE FileNumber=" & d![FileNumber] & " ORDER BY Timestamp;", dbOpenSnapshot)
    If dd.EOF Then      ' no extra lines
        FillField WordDoc, "Line_Items", ""
    Else
        dd.MoveLast
        itemCount = dd.RecordCount
        ' Make enough lines
        For i = 1 To itemCount
            itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
        Next i
        itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
        FillField WordDoc, "Line_Items", itemsFields
        dd.MoveFirst
        i = 1
        Do While Not dd.EOF
            FillField WordDoc, "Item" & i, dd!PrePost & vbTab & Format(dd!Timestamp, "mm/dd/yyyy") & vbTab & dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
     
            totalItems = totalItems + Nz(dd!Amount, 0)
            dd.MoveNext
            i = i + 1
        Loop
        FillField WordDoc, "PreTitle", "Pre-Petition Total:"
        FillField WordDoc, "PostTitle", "Post-Petition Total:"
        FillField WordDoc, "PreTotal", Format(Nz(DSum("Amount", "BKDebt", "Prepost = 'Pre-Petition' AND FileNumber=" & [Forms]![Case List]![FileNumber] & ""), 0), "Currency")
        FillField WordDoc, "PostTotal", Format(Nz(DSum("Amount", "BKDebt", "Prepost = 'Post-Petition' AND FileNumber=" & [Forms]![Case List]![FileNumber] & ""), 0), "Currency")
        'FillField WordDoc, "Total", Format(totalItems, "Currency")
    End If
    dd.Close

End If
totalItems = 0
itemsFields = ""
Set ddd = CurrentDb.OpenRecordset("SELECT numMissedPayments, missedPaymentsTo,  missedPaymentsFrom, MonthlyAmount FROM BKMissedPayments WHERE FileNumber=" & d![FileNumber] & " ORder BY MissedPaymentsFrom;", dbOpenSnapshot)
If ddd.EOF Then      ' no extra lines
    FillField WordDoc, "Line_Items", ""
Else
    ddd.MoveLast
    itemCount = ddd.RecordCount
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items2", itemsFields
    ddd.MoveFirst
    i = 1
    Do While Not ddd.EOF
        delinquent = ddd!numMissedPayments * ddd!MonthlyAmount
        FillField WordDoc, "Item" & i, ddd!numMissedPayments & vbTab & ddd!missedPaymentsFrom & _
        vbTab & ddd!missedPaymentsTo & vbTab & Format$(Nz(ddd!MonthlyAmount, 0), "Currency") & vbTab & Format(delinquent, "Currency")
        
 
        totalItems = totalItems + Nz(ddd!MonthlyAmount, 0)
        delinquenttotal = delinquent + delinquenttotal
        ddd.MoveNext
        i = i + 1
    Loop
    FillField WordDoc, "Total", Format(delinquenttotal + d!postPetitionPartialPayments, "Currency")

End If
ddd.Close

FillField WordDoc, "TrusteeName", FormatName("", d![First], d![Last] & IIf(d![BKdetails.Chapter] = 11, "", ", Trustee"), "", d![Address], d![Address2], d![BKTrustees.City], d![BKTrustees.State], d![Zip])
FillField WordDoc, "AttorneyName", d![AttorneyFirstName] & " " & d![AttorneyLastName] & IIf(IsNull(d![AttorneyLastName]), "", ", Esquire")
FillField WordDoc, "AttorneyFirmName", IIf(IsNull(d![AttorneyFirm]), "", d![AttorneyFirm]) & vbCr & d![AttorneyAddress]


    If Forms!BankruptcyPrint!chElectronicSignature Then
        FillField WordDoc, "ElectronicSignature", "/s/ " & Forms!BankruptcyPrint!cbxAttorney
    Else
        FillField WordDoc, "ElectronicSignature", "_______________________________"
    End If
    FillField WordDoc, "AttorneySignature", Forms!BankruptcyPrint!cbxAttorney
FillField WordDoc, "BkService", BKService(0)

WordObj.Selection.HomeKey wdStory, wdMove
WordObj.ScreenUpdating = True
WordDoc.SaveAs EMailPath & "BOA Motion For Relief CH13.doc"
Call SaveDoc(WordDoc, d![FileNumber], "BOA Motion for Relief CH13.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_CourtesyEvictionLetter(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range

Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryEVCourtesy WHERE FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Eviction Courtesy Letter.dot", False, 0, True)
WordObj.Visible = True

FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
FillField WordDoc, "FullName", IIf(IsNull(d!Company) = True, d!First & " " & d!Last, d!Company)
FillField WordDoc, "Address", d![PropertyAddress]
FillField WordDoc, "Address2", d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Sale", Format(d!Sale, "mmmm d, yyyy")
FillField WordDoc, "State", d!State
'FillField WordDoc, "Attorney", GetStaffFullName([Forms]![EvictionPrint]![cboAttorney].[Column](0))
FillField WordDoc, "Attorney", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
FillField WordDoc, "Purchaser", d!Purchaser

Select Case Forms![Case List]!ClientID
    Case 466
        FillField WordDoc, "RelocationPara", d![Purchaser] & " provides relocation assistance programs to occupants of its foreclosed properties. To discuss these programs and your options under them, please contact the Select Portfolio Servicing, Inc. relocation assistance hotline @ 1-800-962-6010. Select Portfolio Servicing, Inc. is the servicing agent for " & d![Purchaser] & "."
    Case Else
        FillField WordDoc, "RelocationPara", d![Purchaser] & " may provide relocation assistance programs to occupants of its foreclosed properties. To discuss these programs and your options under them, please contact " & d![BrokerName] & ", or " & GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0)) & " @ " & FormatPhone(d![BrokerPhone]) & ", or (301) 907-8000."
End Select

'FillField WordDoc, "RelocationPara", IIf([Forms]![Case List]![ClientID] = 466, d![Purchaser] & " provides relocation assistance programs to occupants of its foreclosed properties. To discuss these programs and your options under them, please contact the Select Portfolio Servicing, Inc. relocation assistance hotline @ 1-800-962-6010. Select Portfolio Servicing, Inc. is the servicing agent for " & d![Purchaser] & ".", _
'd![Purchaser] & " may provide relocation assistance programs to occupants of its foreclosed properties. To discuss these programs and your options under them, please contact " & d![BrokerName] & ", or " & GetStaffFullName([Forms]![EvictionPrint]![cboAttorney].[Column](0)) & " @ " & FormatPhone(d![BrokerPhone]) & ", or (301) 907-8000.")



WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Eviction Courtesy Letter.dot"
Call SaveDoc(WordDoc, Forms![Case List].FileNumber, "Eviction Courtesy Letter.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_EvictionTenant_NTQ_SPS_PTFA(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qry90DayNoticeSPSWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Eviction Print PTFA Tenant NTQ SPS.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

'DefendantsAddress = GetAddresses(d!FileNumber, 5, "Defendant=True", NewLine)
FillField WordDoc, "AttySignature", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
FillField WordDoc, "FileNumber", d!FileNumber
'FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!State
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "Purchaser", d!Purchaser
FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Eviction Print MD Tenant NTQ SPS.doc"
Call SaveDoc(WordDoc, d![FileNumber], "Eviction Print PTFA Tenant NTQ SPS.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_EvictionOwner_NTQ_SPS_PTFA(Keepopen As Boolean)
'Call MsgBox("This section is under going testing", vbExclamation + vbAbortRetryIgnore, "Don't test this yet")

Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qry90DayNoticeSPSWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)


If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Eviction Print PTFA Owner NTQ SPS.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

'DefendantsAddress = GetAddresses(d!FileNumber, 5, "Defendant=True", NewLine)
FillField WordDoc, "AttySignature", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
FillField WordDoc, "FileNumber", d!FileNumber
'FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!State
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "Purchaser", d!Purchaser
FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Eviction Print PTFA Owner NTQ SPS.doc"
Call SaveDoc(WordDoc, d![FileNumber], "Eviction Print PTFA Owner NTQ SPS.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_Eviction_PTFA_NTQ_53(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qry90DayNoticeSPSWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Eviction Print PTFA NTQ Fifth Third.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]

FillField WordDoc, "Sale", Format(d!Sale, "mmmm d, yyyy")
'DefendantsAddress = GetAddresses(d!FileNumber, 5, "Defendant=True", NewLine)
FillField WordDoc, "AttySignature", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
FillField WordDoc, "FileNumber", d!FileNumber
'FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
FillField WordDoc, "PropertyAddress", d!PropertyAddress
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d!State
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "BrokerNM", d!BrokerNm
FillField WordDoc, "BrokerPh", FormatPhone(d!BrokerPh)
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "Purchaser", d!Purchaser
FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
FillField WordDoc, "LastDays", Format(LastDay90(Date), "mmmm d, yyyy")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Eviction Print PTFA NTQ Fifth Third.doc"
Call SaveDoc(WordDoc, d![FileNumber], "Eviction Print PTFA NTQ Fifth Third.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_QuitClaimDeed(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset
Dim GrantorName As String
Dim GrantorAddress As String
Dim GranteeName As String, GranteeAddress As String
Dim exemptText As String


Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsConventionalDeedWord WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If
GrantorName = InputBox("Grantor Name?")
GrantorAddress = InputBox("Grantor Address?")
GranteeName = InputBox("Grantee Name?")
GranteeAddress = InputBox("Grantee Address?")


If d![FCdetails.State] = "MD" Then
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Quit Claim Deed MD.dot", False, 0, True)
    WordObj.Visible = True
    
    WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
    WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
    WordDoc.Bookmarks("ProName").Select
    WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
    WordDoc.Bookmarks("PropertyAddress").Select
    WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt] & ", ") & d!City & ", " & "MD " & FormatZip(d!ZipCode)
Else
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Quit Claim Deed VA.dot", False, 0, True)
    WordObj.Visible = True

    
    If MsgBox("Is the file Transter Tax Exempt?", vbYesNo) = vbYes Then
        FillField WordDoc, "exemptText", InputBox("Enter the section reference", "Transfer Tax Exempt")
        FillField WordDoc, "Tax", "TAX EXEMPT PURSUANT TO CODE OF"
    Else
        FillField WordDoc, "exemptText", ""
        FillField WordDoc, "Tax", ""
    End If
    WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
    WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
    WordDoc.Bookmarks("Property").Select
    WordDoc.Bookmarks("Property").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d!City & ", VA " & FormatZip(d!ZipCode)
    WordDoc.Bookmarks("Atty").Select
    WordDoc.Bookmarks("Atty").Range.Text = d!NameVA
    WordDoc.Bookmarks("VABar").Select
    WordDoc.Bookmarks("VABar").Range.Text = d!VABar
    'WordDoc.Bookmarks("FileNumber").Select
    'WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
    FillField WordDoc, "FileNumber", d![CaseList.FileNumber]

End If

FillField WordDoc, "GrantorName", GrantorName
FillField WordDoc, "GrantorAddress", GrantorAddress
FillField WordDoc, "GranteeName", GranteeName
FillField WordDoc, "GranteeAddress", GranteeAddress


FillField WordDoc, "Attorney", GetStaffFullName(Forms!Foreclosureprint!Attorney.Column(0))
FillField WordDoc, "LegalDescription", d!LegalDescription
FillField WordDoc, "TaxID", d!TaxID
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Jurisdiction", d!Jurisdiction


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Quit Claim Deed " & d![CaseList.FileNumber] & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Quit Claim Deed " & d![CaseList.FileNumber] & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub


Public Sub Doc_SpecialWarranty(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset
Dim GrantorName As String ' Add by sarab to fix the invalid error 2/10
Dim GrantorAddress As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsConventionalDeedWord WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

If d![FCdetails.State] = "MD" Then
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Special Warranty Deed MD.dot", False, 0, True)
    WordObj.Visible = True

    WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
    WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
    
    FillField WordDoc, "Purchaser", IIf(IsNull(d!Purchaser), "", d!Purchaser)
    FillField WordDoc, "PurchaserAddress", OneLine(d!PurchaserAddress)
    FillField WordDoc, "LegalDescription", d!LegalDescription
    FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
    FillField WordDoc, "Leasehold", IIf(d![Leasehold] = 1, "Leasehold", IIf(d![Leasehold] = 0, "Fee Simple", ""))
    FillField WordDoc, "Attorney", GetStaffFullName(Forms!Foreclosureprint!Attorney.Column(0))
    FillField WordDoc, "TAXID", d!TaxID


Else
    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "Special Warranty Deed VA.dot", False, 0, True)
    WordObj.Visible = True

    WordDoc.Bookmarks("Property").Select
    WordDoc.Bookmarks("Property").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
    WordDoc.Bookmarks("PropertyAdd").Select
    WordDoc.Bookmarks("PropertyAdd").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
    WordDoc.Bookmarks("File").Select      ' use this method for bookmarks in header/footer
    WordDoc.Bookmarks("File").Range.Text = d![CaseList.FileNumber]
    WordDoc.Bookmarks("AssessedValue").Select
    WordDoc.Bookmarks("AssessedValue").Range.Text = Nz(Format(d!AssessedValue, "Currency"))
    WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
    WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
    WordDoc.Bookmarks("Atty").Select
    WordDoc.Bookmarks("Atty").Range.Text = d!NameVA
    WordDoc.Bookmarks("VABar").Select
    WordDoc.Bookmarks("VABar").Range.Text = d!VABar
    
    
    
    
    FillField WordDoc, "TaxID", d!TaxID
    FillField WordDoc, "FileNumber", d![CaseList.FileNumber]
    FillField WordDoc, "Purchaser", d!Purchaser
    FillField WordDoc, "PurchaserAddress", Nz(OneLine(d!PurchaserAddress))
    FillField WordDoc, "LegalDescription", d!LegalDescription
    FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![City] & ", " & d![FCdetails.State] & " " & FormatZip(d![ZipCode])
    FillField WordDoc, "Leasehold", IIf(d![Leasehold] = 1, "Leasehold", IIf(d![Leasehold] = 0, "Fee Simple", ""))
    FillField WordDoc, "Attorney", GetStaffFullName(Forms!Foreclosureprint!Attorney.Column(0))
End If

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Special Warranty Deed " & d![CaseList.FileNumber] & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Special Warranty Deed " & d![CaseList.FileNumber] & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_Line(Keepopen As Boolean)

'FormattedTextToWord
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim templateName As String
Dim BankName As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryLineWord WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If


templateName = "Line"
Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName & ".dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & (vbCr) & d!City & ", " & d![FCdetails.State]


FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
'FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", IIf(Not IsNull(d![CourtCaseNumber]), d![CourtCaseNumber], " _____________________ ")
FillField WordDoc, "Attorney", Forms!Foreclosureprint.Attorney.Column(1)
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 1345, "")
FillField WordDoc, "LineDescription", PlainText(d!LineDescription)
FillField WordDoc, "Year", Format(Date, "yyyy")
FillField WordDoc, "CosNames", COSNamesAddress(0)
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", IIf(Split(ReportArgs, "|")(2) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), (Split(ReportArgs, "|")(1)))
FillField WordDoc, "Investor", IIf(Split(ReportArgs, "|")(2) = 3, IIf(d![SubstituteTrustees] = True, "Substitute Trustee", "Trustee"), IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf(d!ClientID = 531, "M&T Bank as Servicer for " & Forms![Case List]!Investor, IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for " & d!Investor, IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for " & Forms![Case List]!Investor, Forms![Case List]!Investor)))))))
FillField WordDoc, "ThisDate", ThisDate(Date)

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Line.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Line.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_MDExpiredLease(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim Tenant As Integer

Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryexpiredLeaseWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE AND Names.Tenant = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Tenant = TenantNamesCount(d!FileNumber)

Do While Tenant <> 0



    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "MD Expired Lease.dot", False, 0, True)
    WordObj.Visible = True

    'WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
    'WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
    'WordDoc.Bookmarks("ProName").Select
    'WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
    'WordDoc.Bookmarks("PropertyAddress").Select
    'WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

    FillField WordDoc, "Owner", OwnerNames(0, 2)
    FillField WordDoc, "Occupants", d!FullName
    FillField WordDoc, "Sale", Format(d!Sale, "mmmm d, yyyy")
    FillField WordDoc, "Attorney", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
    FillField WordDoc, "FileNumber", d!FileNumber
    FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
    FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
    FillField WordDoc, "PropertyAddress2", d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
   
    FillField WordDoc, "FirmPhone", FirmPhone()
    FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
    'FillField WordDoc, "Year", Format(Date, "yyyy")
    FillField WordDoc, "Purchaser", d!Purchaser
    FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
    FillField WordDoc, "LastDays", Format(LastDay90(Date), "mmmm d, yyyy")

    WordObj.Selection.HomeKey wdStory, wdMove
    WordDoc.SaveAs EMailPath & "MD Expired Lease " & d!FileNumber & "(" & Tenant & ").doc"
    Call SaveDoc(WordDoc, d![FileNumber], "MD Expired Lease " & d!FileNumber & "(" & Tenant & ").doc")
    If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges

d.MoveNext
Tenant = Tenant - 1

Loop
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_MDLeaseTermination(Keepopen As Boolean)
'Call MsgBox("This section is under construction", vbExclamation + vbAbortRetryIgnore, "Word Documents arent finished yet")
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim Tenant As Integer

Dim d As Recordset

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryexpiredLeaseWord WHERE FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE AND Names.Tenant = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Tenant = TenantNamesCount(d!FileNumber)

Do While Tenant <> 0



    Set WordObj = CreateObject("Word.Application")
    Set WordDoc = WordObj.Documents.Add(TemplatePath & "MD Lease Termination.dot", False, 0, True)
    WordObj.Visible = True

    'WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
    'WordDoc.Bookmarks("FileNumber").Range.Text = d![FileNumber]
    'WordDoc.Bookmarks("ProName").Select
    'WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
    'WordDoc.Bookmarks("PropertyAddress").Select
    'WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

    FillField WordDoc, "Owner", OwnerNames(0, 2)
    FillField WordDoc, "Occupants", d!FullName
    FillField WordDoc, "Sale", Format(d!Sale, "mmmm d, yyyy")
    FillField WordDoc, "Attorney", GetStaffFullName(Forms!EvictionPrint!cboAttorney.Column(0))
    FillField WordDoc, "FileNumber", d!FileNumber
    FillField WordDoc, "CourtCaseNumber", d!CourtCaseNumber
    FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
    FillField WordDoc, "PropertyAddress2", d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
   
    FillField WordDoc, "FirmPhone", FirmPhone()
    FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
    'FillField WordDoc, "Year", Format(Date, "yyyy")
    FillField WordDoc, "Purchaser", d!Purchaser
    FillField WordDoc, "Date", Format(Date, "mmmm d, yyyy")
    FillField WordDoc, "LastDays", Format(LastDay90(Date), "mmmm d, yyyy")

    WordObj.Selection.HomeKey wdStory, wdMove
    WordDoc.SaveAs EMailPath & "MD Lease Termination " & d!FileNumber & "(" & Tenant & ").doc"
    Call SaveDoc(WordDoc, d![FileNumber], "MD Lease Termination " & d!FileNumber & "(" & Tenant & ").doc")
    If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges

d.MoveNext
Tenant = Tenant - 1

Loop
Set WordObj = Nothing
d.Close
End Sub

Public Sub Doc_LossMitigationFinalSPS(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Loss Mitigation - Final SPS.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])


FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "InvestorAIF", IIf([Forms]![Case List]![AIF], d![Investor] & ", by Select Portfolio Servicing, Inc., as attorney-in-fact", d![Investor])
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber, "_________________")
FillField WordDoc, "Docket", Nz(Format(d!Docket, "mm/d/yyyy"), "___________________")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Loss Mitigation - Final Select " & d![CaseList.FileNumber] & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Loss Mitigation - Final Select " & d![CaseList.FileNumber] & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub



Public Sub Doc_LossMitigationPrelimSPS(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Loss Mitigation - Prelim SPS.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])


FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "InvestorAIF", IIf([Forms]![Case List]![AIF], d![Investor] & ", by Select Portfolio Servicing, Inc., as attorney-in-fact", d![Investor])
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber, "_________________")
FillField WordDoc, "Docket", Nz(Format(d!Docket, "mm/d/yyyy"), "___________________")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Loss Mitigation - Prelim Select " & d![CaseList.FileNumber] & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Loss Mitigation - Prelim Select " & d![CaseList.FileNumber] & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub
Public Sub doc_CombinedAffidavitofComplianceNationStar(Keepopen As Boolean)     'Mei 10/7/15
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim i As Integer, itemCount As Integer, totalItems As Currency, itemsFields As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Combined Affidavit of Compliance NationstarMD.dot", False, 0, True)
WordObj.Visible = True


WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])



FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "Jurisdiction", d!Jurisdiction
'FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "InvestorAIF", IIf([Forms]![Case List]![AIF], d![Investor] & ", by Select Portfolio Servicing, Inc., as attorney-in-fact", d![Investor])
FillField WordDoc, "DotDate", Format([Forms]![foreclosuredetails]![DOTdate], "mmmm d"", ""yyyy")

FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "Mortgagors", MortgagorNames(0, 20)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "Investor", IIf(d![ClientID] = 385 And d![Investor] = "Nationstar Mortgage LLC", "Nationstar maintains records for the loan that is secured by the mortgage or deed of trust being foreclosed in this action. ", _
"Nationstar services and maintains records on behalf of " & d![Investor] & " the secured party to the mortgage or deed of trust being foreclosed in this action.")

FillField WordDoc, "Prior Servicer", IIf(bPrior = True, "           Before the servicing of this loan transferred to Nationstar, " & strPriorServicer & _
" (Prior Servicer) was the servicer for the loan and it maintained the loan servicing records.  When Nationstar began servicing this loan, Prior Servicer's records for the loan were integrated and boarded into Nationstar's systems," & _
" such that Prior Servicer's records, including the collateral file, payment histories, communication logs, default letters, information," & _
"and documents concerning the Loan are now integrated into Nationstar's business records.  Nationstar maintains quality control and " & _
"verification procedures as part of the boarding process to ensure the accuracy of the boarded records.  It is the regular business practice " & _
"of Nationstar to integrate prior servicers' records into Nationstar's business records and to rely upon those boarded records in providing " & _
"its loan servicing functions.  These Prior Servicer records have been integrated and are relied upon by Nationstar as part of Nationstar's " & _
"business records.", "")

DoCmd.Close acForm, "Prior Servicer"                '10/9/15

If bReferee = True Then
    FillField WordDoc, "HolderTransLost", (d!Investor & " directly or through an agent, has possession of the promissory note and held the note at the time of filing the foreclosure complaint. " & d!Investor & " will enforce the promissory note as transferee in possession.")
ElseIf bHolder Then
    FillField WordDoc, "HolderTransLost", (d!Investor & " directly or through an agent, has possession of the promissory note and held the note at the time of filing the foreclosure complaint. The promissory note is made payable to " & d!Investor & " OR The promissory note has been duly indorsed.")
ElseIf bLost = True Then
    FillField WordDoc, "HolderTransLost", (d!Investor & " is unable to find the promissory note and will seek to prove the promissory note using a lost note affidavit.")
End If

FillField WordDoc, "BorrowerNames", BorrowerNames(d![CaseList.FileNumber])
FillField WordDoc, "DOTdate", Format([Forms]![foreclosuredetails]![DOTdate], "mmmm d"", ""yyyy")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
FillField WordDoc, "propertyaddress2", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d!City & ", " & d![FCdetails.State] & " " & FormatZip(d!ZipCode)

FillField WordDoc, "AsOfDate", Format([Forms]![Print Statement of Debt]![txtDueDate], "mmmm dd"", ""yyyy")

'FillField WordDoc, "AsOfDate", Format([Forms]![Print Statement of Debt]![txtDueDate], "mmmm dd"", ""yyyy")
If d!LoanType = 5 Then
    FillField WordDoc, "chkLoanType", "Federal Home Loan Mortgage Corporation "
ElseIf d!LoanType = 4 Then
    FillField WordDoc, "chkLoanType", "Federal National Mortgage Association "
Else
    FillField WordDoc, "chkLoanType", d!Investor
End If

If d!LoanType = 5 Or d!LoanType = 4 Then
    FillField WordDoc, "chkLoanType2", "For the purposes of foreclosure, the note or other debt instrument is held by " & d!Investor & "."
Else
    FillField WordDoc, "chkLoanType2", ""
End If

FillField WordDoc, "DateOfDefault", Format$(d!DateOfDefault, "mmmm d, yyyy")
FillField WordDoc, "NOI", Format$(d!NOI, "mmmm d, yyyy")

FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "FromToDates", IIf([Forms]![Print Statement of Debt]!Check86 = 0, "From  " & [Forms]![Print Statement of Debt]![Text100] & "  To  " & [Forms]![Print Statement of Debt]![Text102] & "  at  " & [Forms]![Print Statement of Debt]![Text104] & "%", "")
'FillField WordDoc, "Rate", " " & [Forms]![Print Statement of Debt]![Text104] & "%"
FillField WordDoc, "IntrestAmt", Format$([Forms]![Print Statement of Debt]![Text94], "Currency")

'if there're adj rates:
If [Forms]![Print Statement of Debt]!Check86 Then    'the adj rate check box is checked
    Set dd = CurrentDb.OpenRecordset("SELECT * FROM StatementOfDebtAdjRate WHERE FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
        If Not dd.EOF Then
                dd.MoveLast
                 For i = 1 To dd.RecordCount
                    itemsFields = itemsFields & "<<AdjRate" & i & ">>" & vbCr
                 Next i
                
                 FillField WordDoc, "AdjRate", itemsFields
                
                dd.MoveFirst
                i = 1
                
                Do While Not dd.EOF
                
                FillField WordDoc, "AdjRate" & i, "        From   " & dd!DateFrom & "   to  " & dd!DateTo & "  at  " & dd!ADJRate & "%" & vbTab & Format(dd!Amount, "currency")
                dd.MoveNext
                i = i + 1
                Loop
        End If
Else
    FillField WordDoc, "AdjRate", ""
End If

Dim rs As Recordset
Dim itemTags As String
'Dim totalItems As Currency

Set rs = CurrentDb.OpenRecordset("SELECT StatementOfDebt.ID, StatementOfDebt.FileNumber, StatementOfDebt.Desc, StatementOfDebt.Amount, StatementOfDebt.Timestamp, StatementOfDebt.Sort_Desc FROM StatementOfDebt WHERE FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)

If rs.EOF Then
    MsgBox "No data found from StatementOfDebt table."
    Exit Sub
Else
    rs.MoveLast
        For i = 1 To rs.RecordCount
            itemTags = itemTags & "<<Line_Items" & i & ">>" & vbCr
        Next i
         
         FillField WordDoc, "Line_Items", itemTags
                
                rs.MoveFirst
                i = 1
                
                Do While Not rs.EOF
                
                FillField WordDoc, "Line_Items" & i, rs!Desc & vbTab & Format(rs!Amount, "currency")
                totalItems = totalItems + Nz(rs!Amount, 0)
                rs.MoveNext
                i = i + 1
                Loop
End If

FillField WordDoc, "Total", "Total" & vbTab & Format$(totalItems + [Forms]![Print Statement of Debt]![Text94] + d!RemainingPBal, "Currency")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Combined Affidavit of Compliance Nationstar " & d![CaseList.FileNumber] & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Combined Affidavit of Compliance Nationstar " & d![CaseList.FileNumber] & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub


Public Sub Doc_ComplianceAffidavit(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset, dd As Recordset      ' data
Dim i As Integer, itemCount As Integer, totalItems As Currency, itemsFields As String



Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Combined Affidavit of Compliance"

If (d![ClientID] = 466) Then
    templateName = templateName & " Select"
End If
templateName = templateName & ".dot"


Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)
WordObj.Visible = True


WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])



FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "Investor", d!Investor
FillField WordDoc, "InvestorAIF", IIf([Forms]![Case List]![AIF], d![Investor] & ", by Select Portfolio Servicing, Inc., as attorney-in-fact", d![Investor])
FillField WordDoc, "DotDate", Format([Forms]![foreclosuredetails]![DOTdate], "mmmm d"", ""yyyy")

FillField WordDoc, "TrusteeNames", trusteeNames(0, 3, , vbCr)
FillField WordDoc, "FirmShortAddress", IIf(d!CaseTypeID = 8, d!MonitorTrusteeAddress, FirmShortAddress(vbCr))
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "Mortgagors", MortgagorNames(0, 20)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)

FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber, "______________")
'FillField WordDoc, "455Only", IIf(d![ClientID] = 455, ", and continuing each month thereafter with probable future advancements made by the mortgagee, and that the Plaintiff(s) has\have the right to foreclose;", ", and continuing each month thereafter, and that the Plaintiff(s) has\have the right to foreclose;")
FillField WordDoc, "Liber", Nz(d!Liber, "________")
FillField WordDoc, "Folio", Nz(d!Folio, "________")
FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " in Liber " & d![Liber2] & " at Folio " & d![Folio2])

FillField WordDoc, "LPIdate", Format$(d![LPIDate], "mmmm d, yyyy")
FillField WordDoc, "LPIdate+1", Format$(d![LPIDate] + 1, "mmmm d, yyyy")
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")

'FillField WordDoc, "PaidStr", IIf((Nz(d![RemainingPBal], 0) > Nz(d![OriginalPBal], 0)), "Additional Interest", "Paid on principal")
'FillField WordDoc, "Paid", Format$(d!OriginalPBal - d!RemainingPBal, "Currency")
FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")

'May have to put these in later
'FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
'FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
'FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

totalItems = 0
itemsFields = ""
Set dd = CurrentDb.OpenRecordset("SELECT Desc, Amount FROM StatementOfDebt WHERE FileNumber=" & d![CaseList.FileNumber] & " ORDER BY Timestamp;", dbOpenSnapshot)
If dd.EOF Then      ' no extra lines
    FillField WordDoc, "Line_Items", ""
Else
    dd.MoveLast
    itemCount = dd.RecordCount
    ' Make enough lines
    For i = 1 To itemCount
        itemsFields = itemsFields & "<<Item" & i & ">>" & vbCr
    Next i
    itemsFields = Left$(itemsFields, Len(itemsFields) - 1) ' remove trailing CR
    FillField WordDoc, "Line_Items", itemsFields
    dd.MoveFirst
    i = 1
    Do While Not dd.EOF
        FillField WordDoc, "Item" & i, dd!Desc & vbTab & Format$(Nz(dd!Amount, 0), "Currency")
        totalItems = totalItems + Nz(dd!Amount, 0)
        dd.MoveNext
        i = i + 1
    Loop
End If
dd.Close

FillField WordDoc, "BalDueDate", IIf(IsNull([Forms]![Print Statement of Debt]![txtDueDate]), "______________", Format$([Forms]![Print Statement of Debt]![txtDueDate], "mmmm d"", ""yyyy"))
FillField WordDoc, "BalanceDue", Format$(d!RemainingPBal + totalItems, "Currency")

FillField WordDoc, "InterestRate", IIf(IsNull(d!InterestRate), "____________ %", Format$(d!InterestRate, "#0.000") & "%")
FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
'FillField WordDoc, "NotaryLocation", IIf(IsNull(d!NotaryLocation), "STATE OF ______________:" & vbCr & "COUNTY OF ____________ :", d!NotaryLocation)

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Combined Affidavit of Compliance Select " & d![CaseList.FileNumber] & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Combined Affidavit of Compliance Select " & d![CaseList.FileNumber] & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close


End Sub



Public Sub Doc_MotionToIntervene(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryMotionToInterveneWord WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Motion to Intervene.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])


FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
'FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "MonitorTrusteeName", Nz(d!MonitorTrusteeName)
FillField WordDoc, "Client", Forms![Case List]!ClientID.Column(1) '[Forms]![Case List]![ClientID].[Column](1)

FillField WordDoc, "HeaderNames", d![PrimaryFirstName] & " " & d![PrimaryLastName]
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
'FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "DotDate", Format(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "Names", d![PrimaryFirstName] & " " & d![PrimaryLastName] & IIf(Len(d![SecondaryFirstName] & " " & d![SecondaryLastName]) = 1, "", " and " & d![SecondaryFirstName] & " " & d![SecondaryLastName])
FillField WordDoc, "DotRecorded", Format(d!DOTrecorded, "mmmm d, yyyy")
FillField WordDoc, "Liber", d!Liber
FillField WordDoc, "Folio", Nz(d!Folio)
FillField WordDoc, "Jurisdiction", d!Jurisdiction
FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Attorney", Forms!MonitorPrint!Attorney.Column(1) '[Forms]![MonitorPrint]![Attorney].[Column](1)
FillField WordDoc, "FirmName", FirmName()
FillField WordDoc, "FirmAddress", FirmAddress()
'FillField WordDoc, "GetAddress", GetAddresses(150, 5, "")
FillField WordDoc, "GetAddress", GetAddresses(0, 5, "Owner = true and (Mortgagor = true or Noteholder = true)")
FillField WordDoc, "ThisDate", ThisDate(Date)
'FillField WordDoc, "City", d!City
FillField WordDoc, "LongState", d![State]
'FillField WordDoc, "InvestorAIF", IIf([Forms]![Case List]![AIF], d![Investor] & ", by Select Portfolio Servicing, Inc., as attorney-in-fact", d![Investor])
'FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber, "_________________")
'FillField WordDoc, "Docket", Nz(Format(d!Docket, "mm/d/yyyy"), "___________________")
'FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
'FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
'FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Motion to Intervene " & d![CaseList.FileNumber] & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Motion to Intervene " & d![CaseList.FileNumber] & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_OrderGrantingIntervention(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryMotionToInterveneWord WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Order Granting Intervention.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])


FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
'FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "MonitorTrusteeName", d!MonitorTrusteeName
FillField WordDoc, "Client", Forms![Case List]!ClientID.Column(1) '[Forms]![Case List]![ClientID].[Column](1)

FillField WordDoc, "HeaderNames", d![PrimaryFirstName] & " " & d![PrimaryLastName] & vbCr & _
IIf(Len(d![SecondaryFirstName] & " " & d![SecondaryLastName]) = 1, "", d![SecondaryFirstName] & " " & d![SecondaryLastName])

'FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
'FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
'FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
'FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
'FillField WordDoc, "DotDate", Format(d!DOTdate, "mmmm d, yyyy")
'FillField WordDoc, "Names", d![PrimaryFirstName] & " " & d![PrimaryLastName] & IIf(Len(d![SecondaryFirstName] & " " & d![SecondaryLastName]) = 1, "", " and " & d![SecondaryFirstName] & " " & d![SecondaryLastName])
'FillField WordDoc, "DotRecorded", Format(d!DOTrecorded, "mmmm d, yyyy")
'FillField WordDoc, "Liber", d!Liber
'FillField WordDoc, "Folio", d!Folio
'FillField WordDoc, "Jurisdiction", d!Jurisdiction
'FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
'FillField WordDoc, "Attorney", Forms!MonitorPrint!Attorney.Column(1) '[Forms]![MonitorPrint]![Attorney].[Column](1)
'FillField WordDoc, "FirmName", FirmName()
'FillField WordDoc, "FirmAddress", FirmAddress()
'FillField WordDoc, "GetAddresses", GetAddresses(150, 5, "")
FillField WordDoc, "GetAddress", GetAddresses(0, 5, "Owner = true and (Mortgagor = true or Noteholder = true)")

'FillField WordDoc, "ThisDate", ThisDate(Date)
'FillField WordDoc, "City", d!City
FillField WordDoc, "LongState", d![State]
'FillField WordDoc, "InvestorAIF", IIf([Forms]![Case List]![AIF], d![Investor] & ", by Select Portfolio Servicing, Inc., as attorney-in-fact", d![Investor])
'FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber, "_________________")
'FillField WordDoc, "Docket", Nz(Format(d!Docket, "mm/d/yyyy"), "___________________")
'FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
'FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
'FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")
FillField WordDoc, "year", Format(Date, "yyyy")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Order Granting Intervention " & d![CaseList.FileNumber] & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Order Granting Intervention " & d![CaseList.FileNumber] & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_StatementOfDebtMonitor(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryFCDocsWordLite WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber, dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Dim templateName As String
templateName = "Statement of Debt Monitor"

'If (d![ClientID] = 157) Then
'  templateName = templateName & " Cenlar"
'ElseIf d!ClientID = 404 Then
'  templateName = templateName & " Bogman"
'ElseIf (d![ClientID] = 451 And Trim(UCase$(d!Investor)) Like "LPP MORTGAGE*" And d![FCdetails.State] = "MD") Then
'  templateName = templateName & " Dove"
'ElseIf d![ClientID] = 451 Then
'    templateName = templateName & " Dove"
'End If
templateName = templateName & ".dot"

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & templateName, False, 0, True)



WordObj.Visible = True
WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress]
WordDoc.Bookmarks("APTNum").Select
WordDoc.Bookmarks("APTNum").Range.Text = Nz(d![Fair Debt])


WordObj.Visible = False
If MsgBox("Is There A Loan Mod? ", vbYesNo) = vbYes Then
FillField WordDoc, "Mod", ", MODIFIED by Agreement effective " & Format(InputBox(" Effective Date?   Format mm/dd/yyyy"), "mmmm d, yyyy") & " with an amended principal balance of " & Format(InputBox(" Amended Principal Balance? "), "Currency")
Else
FillField WordDoc, "Mod", ""
End If

If d!JurisdictionID = 6 Then  'Calvert County
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Book " & d![Liber2] & ", Page " & d![Folio2])
Else 'MD or DC
    FillField WordDoc, "ReRecorded", IIf(IsNull(d![Rerecorded]), "", ", and re-recorded on " & Format$(d![Rerecorded], "mmmm d, yyyy") & " at Liber " & d![Liber2] & ", Folio " & d![Folio2])
End If

WordObj.Visible = True
FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
FillField WordDoc, "LoanNumber", d!LoanNumber


'If d!ClientID = 532 Then 'SELENE Finance
'    FillField WordDoc, "Investor", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", d![Investor])
'ElseIf d!ClientID = 523 Then 'GreenTree
'    FillField WordDoc, "Investor", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", d![Investor])
'Else
'    FillField WordDoc, "Investor", IIf((d!AIF = True And d!ClientID = 373), d!Investor & ", by BSI Financial Services as attorney in fact", IIf((d![AIF] = True And d![ClientID] = 523), d![Investor] & " by Green Tree Servicing LLC as servicer", IIf((d![AIF] = True And d!ClientID = 532), d![Investor] & ", by and through Selene Finance LP, its Servicer and Attorney-in-Fact", IIf((d![AIF] = True And d![ClientID] = 605), d![LongClientName] & " as Servicer for ", IIf(d![AIF] = True, d![LongClientName] & " as Attorney in Fact for ", "")))) & d![Investor])
'End If

FillField WordDoc, "MonitorTrusteeName", Nz(d!MonitorTrusteeName)
FillField WordDoc, "FirmShortAddress", Nz(d!MonitorTrusteeAddress)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
'FillField WordDoc, "MortgagorNames", MortgagorNames(0, 3, vbCr)
FillField WordDoc, "Names", d![PrimaryFirstName] & " " & d![PrimaryLastName] & vbCr & _
IIf(Len(d![SecondaryFirstName] & " " & d![SecondaryLastName]) = 1, "", d![SecondaryFirstName] & " " & d![SecondaryLastName])

FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "LongClientName", d!LongClientName

FillField WordDoc, "City", d!City
FillField WordDoc, "State", d![FCdetails.State]
FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)
FillField WordDoc, "Liber", LiberFolio(d![Liber], d![Folio], d![FCdetails.State], d![JurisdictionID])
'FillField WordDoc, "LPIdate", Format$(d![LPIDate], "mmmm d, yyyy")
'FillField WordDoc, "LPIdate+1", Format$(d![LPIDate] + 1, "mmmm d, yyyy")
'FillField WordDoc, "LPdM", DateAdd("m", -1, d![LPIDate])
FillField WordDoc, "OriginalPBal", Format$(d!OriginalPBal, "Currency")
'FillField WordDoc, "PaidStr", IIf((Nz(d![RemainingPBal], 0) > Nz(d![OriginalPBal], 0)), "Additional Interest", "Paid on principal")
'FillField WordDoc, "Paid", Format$(d!OriginalPBal - d!RemainingPBal, "Currency")
'FillField WordDoc, "RemainingPBal", Format$(d!RemainingPBal, "Currency")
FillField WordDoc, "InterestRate", IIf(IsNull(d!InterestRate), "____________ %", Format$(d!InterestRate, "#.000%"))
'FillField WordDoc, "InvestorName", (Split(ReportArgs, "|")(0))
'FillField WordDoc, "InvestorTitle", (Split(ReportArgs, "|")(1))
FillField WordDoc, "Barcode", AddDocPreIndex(d![CaseList.FileNumber], 223, "")
FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Statement of Debt.doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Statement of Debt.doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub Doc_MotionToReleaseFunds(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d As Recordset      ' data
Dim MotiontoDepositSurplus, MotionToIntervene, OrderGranting, OrderGrantingIntervention, CheckDeposited, CheckAmount As String

Set d = CurrentDb.OpenRecordset("SELECT * FROM qryMotionToInterveneWord WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Motion to Release Funds.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

MotiontoDepositSurplus = InputBox("Motion to Deposit Surplus Monies filed date")
'MotionToIntervene = InputBox("Motion to Intervene filed")
OrderGranting = InputBox("Order Granting the Motion to Deposit Surplus Monies filed")
'OrderGrantingIntervention = InputBox("Order Granting the Motion to Intervene filed")
CheckDeposited = InputBox("Check Deposited date")
CheckAmount = InputBox("Check Amount")


FillField WordDoc, "MotionToDepositSurplus", Format(MotiontoDepositSurplus, "mmmm dd, yyyy")
'FillField WordDoc, "MotionToIntervene", Format(MotionToIntervene, "mmmm dd, yyyy")
FillField WordDoc, "OrderGranting", Format(OrderGranting, "mmmm dd, yyyy")
'FillField WordDoc, "ORderGrantingIntervention", Format(OrderGrantingIntervention, "mmmm dd, yyyy")
FillField WordDoc, "CheckDeposited", Format(CheckDeposited, "mmmm dd, yyyy")
FillField WordDoc, "CheckAmount", Format(CheckAmount, "Currency")
'added on 6/9/15
FillField WordDoc, "MonitorMotionSurplusFiled", Format(d!MonitorMotionSurplusFiled, "mmmm d, yyyy")
'StatementOfDebtDate
'StatementOfDebtAmount
FillField WordDoc, "StatementOfDebtDate", Format(d!StatementOfDebtDate, "mmmm d, yyyy")
FillField WordDoc, "StatementOfDebtAmount", Format(d!StatementOfDebtAmount, "Currency")
'MonitorOrderSurplus
FillField WordDoc, "MonitorOrderSurplus", Format(d!MonitorOrderSurplus, "mmmm d, yyyy")
'CourtCaseNumber
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber)



FillField WordDoc, "UpperJurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
'FillField WordDoc, "LoanNumber", d!LoanNumber
FillField WordDoc, "MonitorTrusteeName", Nz(d!MonitorTrusteeName)
FillField WordDoc, "Client", Forms![Case List]!ClientID.Column(1) '[Forms]![Case List]![ClientID].[Column](1)

FillField WordDoc, "HeaderNames", d![PrimaryFirstName] & " " & d![PrimaryLastName]
FillField WordDoc, "FirmShortAddress", FirmShortAddress(vbCr)
FillField WordDoc, "TrusteeWord", TrusteeWord(0, 0)
'FillField WordDoc, "MortgagorNames", MortgagorNames(0, 2)
FillField WordDoc, "PropertyAddress", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])
FillField WordDoc, "DotDate", Format(d!DOTdate, "mmmm d, yyyy")
'FillField WordDoc, "DotRecorded", Format(d!DOTrecorded, "mmmm d, yyyy")
FillField WordDoc, "Liber", d!Liber
FillField WordDoc, "Folio", Nz(d!Folio)
FillField WordDoc, "Jurisdiction", d!Jurisdiction
'FillField WordDoc, "PropertyAddressOneLine", d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d![City] & ", " & d![State] & " " & FormatZip(d![ZipCode])
FillField WordDoc, "Attorney", Forms!MonitorPrint!Attorney.Column(1) '[Forms]![MonitorPrint]![Attorney].[Column](1)
FillField WordDoc, "FirmName", FirmName()
FillField WordDoc, "FirmAddress", FirmAddress()
'FillField WordDoc, "GetAddress", GetAddresses(150, 5, "")
FillField WordDoc, "GetAddress", GetAddresses(0, 5, "Owner = true and (Mortgagor = true or Noteholder = true)")

FillField WordDoc, "ThisDate", ThisDate(Date)
'FillField WordDoc, "City", d!City
FillField WordDoc, "LongState", d![State]
'FillField WordDoc, "InvestorAIF", IIf([Forms]![Case List]![AIF], d![Investor] & ", by Select Portfolio Servicing, Inc., as attorney-in-fact", d![Investor])
'FillField WordDoc, "Zip", FormatZip(d!ZipCode)
FillField WordDoc, "CourtCaseNumber", Nz(d!CourtCaseNumber, "_________________")
'FillField WordDoc, "Docket", Nz(Format(d!Docket, "mm/d/yyyy"), "___________________")
'FillField WordDoc, "AnneArundel1", IIf(d![JurisdictionID] = 3, "Address: ________________", "")
'FillField WordDoc, "AnneArundel2", IIf(d![JurisdictionID] = 3, "               ________________", "")
'FillField WordDoc, "AnneArundel3", IIf(d![JurisdictionID] = 3, "Phone No.: ______________", "")

WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Motion to Release Funds " & d![CaseList.FileNumber] & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Motion to Release Funds " & d![CaseList.FileNumber] & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close

End Sub

Public Sub DOC_DeedofAppointmentNationstarMD(Keepopen As Boolean)
'Added on the 10/6/15

    Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
    
    Dim d, rs As Recordset

    Set d = CurrentDb.OpenRecordset("SELECT * FROM SOT_VA_Nations WHERE CaseList.FileNumber=" & [Forms]![Case List]![FileNumber], dbOpenSnapshot)
    If d.EOF Then
        MsgBox "No data found"
        Exit Sub
    End If
    
Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Deed of Appointment NationstarMD.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select

FillField WordDoc, "DOTDate", Format(d!DOTdate, "mmmm d, yyyy")
FillField WordDoc, "mortgagorNames", MortgagorNames1(0, 3, vbNewLine)
FillField WordDoc, "Mers", IIf(d!MERS = True, "Mortgage Electronic Registration Systems, Inc.(MERS), solely as nominee for " & d!OriginalBeneficiary & " and its successors and assigns", d!OriginalBeneficiary)
FillField WordDoc, "Client", Nz(d!LongClientName)
'ClientAddress
FillField WordDoc, "ClientAddress", d!StreetAddress
FillField WordDoc, "CityStZip", d!City & ", " & d!State & " " & FormatZip(d!ZipCode)
FillField WordDoc, "Investor", Nz(d!Investor)
FillField WordDoc, "chkInvestor", IIf(d!Investor = "Nationstar Mortgage LLC", "present", "")
FillField WordDoc, "Jurisdiction", UCase$(d!Jurisdiction & ", ")
'FillField WordDoc, "Word2", IIf(d!Investor = "Nationstar Mortgage LLC", "undersigned", d!Investor)
FillField WordDoc, "trusteeNames", trusteeNames(0, 2)
FillField WordDoc, "RecordingInformation", IIf(IsNull(d!Liber) = True And IsNull(d!Folio) = True, "", "Recording Information")
FillField WordDoc, "Instrument", IIf(IsNull(d!Liber) = False And IsNull(d!Folio) = True, "Instrument: " & d!Liber & " ", IIf(IsNull(d!Liber) = False And IsNull(d!Folio) = False, "Book: " & d!Liber & " , Page: " & d!Folio & "", ""))
FillField WordDoc, "Re-recordingInfo", IIf(IsNull(d!Liber2) = True And IsNull(d!Folio2) = True, "", "Re-Recording Information")
FillField WordDoc, "Instrument2", IIf(IsNull(d!Liber2) = False And IsNull(d!Folio2) = True, "Instrument: " & d!Liber2 & " ", IIf(IsNull(d!Liber2) = False And IsNull(d!Folio2) = False, "Book: " & d!Liber2 & " , Page: " & d!Folio2 & "", ""))

FillField WordDoc, "FirmShortAddressOneLine", FirmShortAddressOneLine()
FillField WordDoc, "AIF", IIf(d!AIF = True, "Nationstar Mortgage LLC, attorney-in-fact for " & d!Investor & " ", d!Investor)
'FillField WordDoc, "Recording", IIf(IsNull(d!Liber) = False And IsNull(d!Folio) = True, "Instrument: " & d!Liber & " ", IIf(IsNull(d!Liber) = False And IsNull(d!Folio) = False, "Book: " & d!Liber & " , Page: " & d!Folio & ""))

'FillField WordDoc, "Recording", IIf(d!Liber <> "" And d!Folio = "", "Instrument: " & d!Liber & " ", IIf(d!Liber <> "" And d!Folio <> "", "Book: " & d!Liber & " , Page: " & d!Folio & "", ""))


WordObj.Selection.HomeKey wdStory, wdMove
WordDoc.SaveAs EMailPath & "Substitute Trustee MD Nationstar" & ".doc"
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Substitute Trustee MD Nationstar" & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing
d.Close


End Sub
'Added on the 6/18/15
'Doc_DeedofAppointmentNationStarVA
Public Sub Doc_DeedofAppointmentNationStarVA(Keepopen As Boolean)
Dim WordObj As Word.Application, WordDoc As Word.Document, WordRange As Word.Range
Dim d, rs As Recordset     ' data
'Dim MotiontoDepositSurplus, MotionToIntervene, OrderGranting, OrderGrantingIntervention, CheckDeposited, CheckAmount As String

'Set d = CurrentDb.OpenRecordset("SELECT * FROM qrySOT_VA_Nations WHERE CaseList.FileNumber=" & Forms![Case List]!FileNumber & " AND FcDetails.Current = TRUE;", dbOpenSnapshot)
Set d = CurrentDb.OpenRecordset("SELECT * FROM SOT_VA_Nations WHERE CaseList.FileNumber=" & [Forms]![Case List]![FileNumber], dbOpenSnapshot)
Set rs = CurrentDb.OpenRecordset("SELECT Name FROM qryTrustees WHERE FileNumber=" & [Forms]![Case List]![FileNumber], dbOpenSnapshot)

If d.EOF Then
    MsgBox "No data found"
    Exit Sub
End If

Set WordObj = CreateObject("Word.Application")
Set WordDoc = WordObj.Documents.Add(TemplatePath & "Deed of Appointment NationStar VA.dot", False, 0, True)
WordObj.Visible = True

WordDoc.Bookmarks("FileNumber").Select      ' use this method for bookmarks in header/footer
WordDoc.Bookmarks("FileNumber").Range.Text = d![CaseList.FileNumber]
WordDoc.Bookmarks("ProName").Select
WordDoc.Bookmarks("ProName").Range.Text = d![PrimaryDefName]
WordDoc.Bookmarks("PropertyAddress").Select
WordDoc.Bookmarks("PropertyAddress").Range.Text = d![PropertyAddress] & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt])

'new added
'TrusteeName
'DOTDate
'Grantors
'OriginalBeneficiary
'CurrentBeneficiary
'Recording'Re-Recording
'Client
'ClientAddress
'City
'Recording
'Re-Recording
'Investor
'word1
'Jurisdiction
'JurisdictionState
'Word2
'Trustee
'ClientInfo

FillField WordDoc, "DOTDate", Format(d!DOTdate, "mmmm d, yyyy")


If d!Jurisdiction = "Accomack County" Then      'Mei 10/20/15  Accormack cnty wants ALL capitals
  FillField WordDoc, "OriginalBeneficiary", UCase(IIf(d!MERS = True, "Mortgage Electronic Registration Systems, Inc.(MERS), solely as nominee for " & d!OriginalBeneficiary & " and its successors and assigns", d!OriginalBeneficiary))
  FillField WordDoc, "Grantors", UCase(MortgagorNames1(0, 3, vbNewLine)) 'GetNames(d![CaseList.FileNumber], 3, "Mortgagor = True")
  FillField WordDoc, "TrusteeName", UCase(trusteeNames(0, 3))
  FillField WordDoc, "CurrentBeneficiary", UCase(Nz(d!Investor))
  FillField WordDoc, "Client", UCase(Nz(d!LongClientName))
  FillField WordDoc, "Investor", UCase(Nz(d!Investor))
  FillField WordDoc, "Word2", UCase(IIf(d!Investor = "Nationstar Mortgage LLC", "undersigned", d!Investor))
  FillField WordDoc, "Trustee", UCase(Nz(rs!Name))
  FillField WordDoc, "ClientInfo", UCase(IIf(d!AIF = True, "Nationstar Mortgage LLC, attorney-in-fact for " & d!Investor & " ", d!Investor))
Else
  FillField WordDoc, "OriginalBeneficiary", IIf(d!MERS = True, "Mortgage Electronic Registration Systems, Inc.(MERS), solely as nominee for " & d!OriginalBeneficiary & " and its successors and assigns", d!OriginalBeneficiary)
  FillField WordDoc, "Grantors", MortgagorNames1(0, 3, vbNewLine) 'GetNames(d![CaseList.FileNumber], 3, "Mortgagor = True")
  FillField WordDoc, "TrusteeName", trusteeNames(0, 3)
  FillField WordDoc, "CurrentBeneficiary", Nz(d!Investor)
  FillField WordDoc, "Client", Nz(d!LongClientName)
  FillField WordDoc, "Investor", Nz(d!Investor)
  FillField WordDoc, "Word2", IIf(d!Investor = "Nationstar Mortgage LLC", "undersigned", d!Investor)
  FillField WordDoc, "Trustee", Nz(rs!Name)
  FillField WordDoc, "ClientInfo", IIf(d!AIF = True, "Nationstar Mortgage LLC, attorney-in-fact for " & d!Investor & " ", d!Investor)
End If

FillField WordDoc, "OriginalTrustee", d!OriginalTrustee
'ClientAddress
FillField WordDoc, "ClientAddress", d!StreetAddress & " " & d!StreetAddr2
FillField WordDoc, "City", d!City & ", " & d!State & " " & FormatZip(d!ZipCode)
FillField WordDoc, "word1", IIf(d!Investor = "Nationstar Mortgage LLC", "present", "")
FillField WordDoc, "Jurisdiction", UCase$(d!Jurisdiction & ", " & d!LongState)
'FillField WordDoc, "Recording", IIf(IsNull(d!Liber) = False And IsNull(d!Folio) = True, "Instrument: " & d!Liber & " ", IIf(IsNull(d!Liber) = False And IsNull(d!Folio) = False, "Book: " & d!Liber & " , Page: " & d!Folio & ""))
FillField WordDoc, "Recording", IIf(IsNull(d!Liber) = False And IsNull(d!Folio) = True, "Instrument: " & d!Liber & " ", IIf(IsNull(d!Liber) = False And IsNull(d!Folio) = False, "Book: " & d!Liber & " , Page: " & d!Folio & "", ""))
'FillField WordDoc, "Recording", IIf(d!Liber <> "" And d!Folio = "", "Instrument: " & d!Liber & " ", IIf(d!Liber <> "" And d!Folio <> "", "Book: " & d!Liber & " , Page: " & d!Folio & "", ""))

FillField WordDoc, "TaxmapNo", d!TaxID
'Mei 10-19-15 shows only if re-recording info is not null
If IsNull(d![Folio2]) And IsNull(d![Liber2]) Then
    FillField WordDoc, "Re-Recording Information:", ""
    FillField WordDoc, "Re-Recording", ""
Else
    FillField WordDoc, "Re-Recording Information:", "Re-Recording Information:"
    FillField WordDoc, "Re-Recording", IIf(IsNull(d!Liber2) = False And IsNull(d!Folio2) = True, "Instrument: " & d!Liber2 & " ", IIf(IsNull(d!Liber2) = False And IsNull(d!Folio2) = False, "Book: " & d!Liber2 & " , Page: " & d!Folio2 & "", ""))
End If
'FillField WordDoc, "Recordinginfo", IIf(IsNull(d!Liber) = True And IsNull(d!Folio) = True, "", "Recording Information")

FillField WordDoc, "Property Address", d!PropertyAddress & IIf(Len(d![Fair Debt] & "") = 0, "", ", " & d![Fair Debt]) & ", " & d!FCCity & ", " & d!FCState & " " & FormatZip(d!FCZipCode)


WordObj.Selection.HomeKey wdStory, wdMove
'WordDoc.SaveAs EMailPath & "Deed of Appointment NationStar VA " & d![CaseList.FileNumber] & ".doc"
WordDoc.SaveAs EMailPath & "Substitute Trustee Nationstar VA" & d![CaseList.FileNumber] & ".doc"
'Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Deed of Appointment NationStar VA " & d![CaseList.FileNumber] & ".doc")
Call SaveDoc(WordDoc, d![CaseList.FileNumber], "Substitute Trustee Nationstar VA " & d![CaseList.FileNumber] & ".doc")
If Not Keepopen Then WordObj.Quit SaveChanges:=wdDoNotSaveChanges
Set WordObj = Nothing

d.Close
rs.Close

Set d = Nothing
Set rs = Nothing

End Sub

