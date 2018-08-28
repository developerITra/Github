Attribute VB_Name = "SQL"
Option Compare Database
Option Explicit

Const ODBC_ADD_SYS_DSN = 4
Const ODBC_ADD_DSN = 1
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
          (ByVal hwndParent As Long, ByVal fRequest As Long, _
          ByVal lpszDriver As String, ByVal lpszAttributes As String) _
          As Long

Public Sub CheckSQL()
Dim rstTest As Recordset, ErrCnt As Integer, TestValue As Integer

On Error GoTo CheckSQLErr

Do While True
    Set rstTest = CurrentDb.OpenRecordset("SELECT TestConnection FROM SQL;", dbOpenSnapshot)
    TestValue = rstTest!TestConnection
    rstTest.Close
    Set rstTest = Nothing
    
    If TestValue = 3 Then Exit Sub  ' success, get out of here
    ErrCnt = ErrCnt + 1
    If ErrCnt > 3 Then GoTo CheckSQLErr
    Call CreateDSN          ' re-create the DSN, maybe we need a different database
Loop

CheckSQLErr:
If ErrCnt > 3 Then
    MsgBox "Repeated attempts to access database have failed", vbExclamation
    Exit Sub
End If
If Err.Number = 3151 Then
    ErrCnt = ErrCnt + 1
    If CreateDSN() Then
        Resume
    Else
        MsgBox "Cannot create connection to database", vbCritical
    End If
Else
    MsgBox "Error " & Err.Number & ": " & Err.Description
End If
End Sub

Public Function CreateDSN(Optional Developer As Boolean) As Boolean

Dim strDriver As String, strAttributes As String

strDriver = "SQL Server"

If Developer Then
    strAttributes = "SERVER=SQLServer" & Chr$(0)
    strAttributes = strAttributes & "DESCRIPTION=Rosie" & Chr$(0)
    strAttributes = strAttributes & "DSN=RosenbergDB" & Chr$(0)
    strAttributes = strAttributes & "DATABASE=RosenbergTest" & Chr$(0)
    strAttributes = strAttributes & "Trusted_Connection=Yes" & Chr$(0)

Else
    strAttributes = "SERVER=SQLServer" & Chr$(0)
    strAttributes = strAttributes & "DESCRIPTION=Rosie" & Chr$(0)
    strAttributes = strAttributes & "DSN=RosenbergDB" & Chr$(0)
    strAttributes = strAttributes & "DATABASE=Rosenberg" & Chr$(0)
    strAttributes = strAttributes & "Trusted_Connection=Yes" & Chr$(0)

End If
CreateDSN = SQLConfigDataSource(0, ODBC_ADD_DSN, strDriver, strAttributes)
'MsgBox "Created DSN", vbInformation
End Function

Public Function DeveloperInfo() As String
Dim dbConnect() As String, Info As String, i As Integer

dbConnect = Split(CurrentDb.TableDefs("SQL").Connect, ";")
For i = 0 To UBound(dbConnect) - 1
    If Left$(dbConnect(i), 4) = "DSN=" Then Info = Info & dbConnect(i) & vbNewLine
    If Left$(dbConnect(i), 9) = "DATABASE=" Then Info = Info & dbConnect(i) & vbNewLine
Next
Info = Info & "INFO=" & DLookup("Info", "SQL")
DeveloperInfo = Info

End Function
Public Function GetCountOfQuery(QueryName As String) As Integer
' DaveW 2011.12.20
Dim rstFromQuery As Recordset, cntr As Integer

Set rstFromQuery = CurrentDb.OpenRecordset("Select * FROM " & QueryName, dbOpenDynaset, dbSeeChanges)
Do Until rstFromQuery.EOF
cntr = cntr + 1
rstFromQuery.MoveNext
Loop

Set rstFromQuery = Nothing
GetCountOfQuery = cntr
End Function
Public Function ShowAllProperties(obj As Object) As String
'DaveW:  2012.01.05 For analysis purposes.
Dim s, t As String, prp As Property
s = ""
For Each prp In obj.Properties
    t = prp.Name
    Debug.Print t
    s = t & vbNewLine
Next prp
ShowAllProperties = s
End Function

Public Sub RunSQL(str_SQLCommand As String)
    On Error GoTo Err_Handler
        DoCmd.SetWarnings False
        DoCmd.RunSQL str_SQLCommand
Exit_Proc:
        DoCmd.SetWarnings True
        Exit Sub
Err_Handler:
        Debug.Print Err.Description
        'display some error
        Resume Exit_Proc
End Sub
