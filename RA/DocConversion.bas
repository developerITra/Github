Attribute VB_Name = "DocConversion"
' Temporary module for conversion of documents

Option Compare Database
Option Explicit
Const GroupDelimiter = ";"

Public Sub ConvertDocuments()
Dim rstfiles As Recordset

Set rstfiles = CurrentDb.OpenRecordset("SELECT FileNumber FROM CaseList ORDER BY FileNumber", dbOpenSnapshot)
Do While Not rstfiles.EOF
    Debug.Print rstfiles!FileNumber
    Call ConvertDocumentsFile(rstfiles!FileNumber)
    rstfiles.MoveNext
    DoEvents
Loop
rstfiles.Close
Debug.Print "[Done]"
End Sub

Public Sub ConvertDocumentsFile(FileNumber As Long)
Const MaxTitles = 100
Dim Titles(1 To MaxTitles) As String      ' document titles
Dim rstDoc As Recordset, Filespec As String, DisplayFilename As String, GroupName As String

On Error GoTo ConvErr

'Commented by JAE 10-30-2014 'Document Speed'
'Set rstDoc = CurrentDb.OpenRecordset("DocIndex", dbOpenDynaset, dbSeeChanges)
Filespec = Dir$(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\*.*")
GoSub GetGroup

Do While Filespec <> ""
    'Debug.Print FileNumber, Filespec
    'Commented by JAE 10-30-2014 'Document Speed'
'    rstDoc.AddNew
'    rstDoc!FileNumber = FileNumber
'    rstDoc!DocTitleID = 0
'    rstDoc!DocGroup = GroupName
'    rstDoc!StaffID = 0
'    rstDoc!DateStamp = GetDateStamp(Filespec)
'    rstDoc!Filespec = Filespec
'    rstDoc!Notes = Filespec
'    rstDoc.Update
    DoCmd.SetWarnings False
    Dim strSQLValues As String: strSQLValues = ""
    Dim strSQL As String: strSQL = ""
    strSQL = ""
    strSQLValues = FileNumber & "," & 0 & ",'" & GroupName & "'," & 0 & ",'" & GetDateStamp(Filespec) & "','" & Replace(Filespec, "'", "''") & "','" & Replace(Filespec, "'", "''") & "'"
    'Debug.Print strSQLValues
    strSQL = "Insert Into DocIndex (FileNumber,DocTitleID,DocGroup,StaffID,DateStamp,Filespec,Notes) VALUES (" & strSQLValues & ")"
    'Debug.Print strSQL
    DoCmd.RunSQL (strSQL)
    DoCmd.SetWarnings True
    
    Filespec = Dir$()
    GoSub GetGroup
Loop
rstDoc.Close
'Debug.Print "[Done]"
Exit Sub

ConvErr:
    MsgBox Err.Description, vbCritical
    Exit Sub
    
GetGroup:
    If Left$(Filespec, 1) = GroupDelimiter Then
        GroupName = Split(Filespec, GroupDelimiter)(1)
        DisplayFilename = Mid$(Filespec, Len(GroupDelimiter) * 2 + Len(GroupName) + 1)
    Else
        GroupName = ""
        DisplayFilename = Filespec
    End If
Return

End Sub

Private Function GetDateStamp(FileName As String) As Variant

If Left(FileName, 2) = GroupDelimiter & "I" Then
    GetDateStamp = DateSerial(Mid$(FileName, 4, 4), Mid$(FileName, 8, 2), Mid$(FileName, 10, 2))
    GetDateStamp = DateAdd("h", Mid$(FileName, 13, 2), GetDateStamp)
    GetDateStamp = DateAdd("n", Mid$(FileName, 15, 2), GetDateStamp)
    GetDateStamp = DateAdd("s", Mid$(FileName, 17, 2), GetDateStamp)
Else
    If Left$(Right$(FileName, 9), 1) = " " And IsNumeric(Left$(Right$(FileName, 8), 4)) And IsNumeric(Left$(Right$(FileName, 17), 8)) Then
        GetDateStamp = DateSerial(Left$(Right$(FileName, 17), 4), Left$(Right$(FileName, 13), 2), Left$(Right$(FileName, 11), 2))
        GetDateStamp = DateAdd("h", Left$(Right$(FileName, 8), 2), GetDateStamp)
        GetDateStamp = DateAdd("n", Left$(Right$(FileName, 6), 2), GetDateStamp)
    Else
        GetDateStamp = Null
    End If
End If
End Function
