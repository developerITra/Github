Attribute VB_Name = "Version"
Option Compare Database
Option Explicit
 
Public Const DBVersion = 2772 ' Update versiontracking table with new version
'Public Const DBVersionTest = 1

Private Sub ClearRecent()
' If you have worked on reports, also run the subroutine 'UpdateReportRecordSources' below !!!!!!
End Sub


Private Sub UpdateReportRecordSources()
Dim rstDoc As Recordset, sql As String, QueryWarning As Boolean

'' Update the saved record sources used by TestDocument.
Debug.Print "Updating RecordSources..."
Set rstDoc = CurrentDb.OpenRecordset("DocumentList", dbOpenDynaset, dbSeeChanges)
Do While Not rstDoc.EOF
    DoCmd.OpenReport rstDoc!DocName, acViewDesign, , , acHidden
    sql = Application.Reports(rstDoc!DocName).RecordSource
    DoCmd.Close acReport, rstDoc!DocName
    If Nz(rstDoc!RecordSource) <> sql Then
        If InStr(1, sql, "qryFCDocs.") > 0 Then     ' can't reference this query
            Debug.Print "Invalid RecordSource in '" & rstDoc!DocName & "', cannot reference qryFCDocs"
            QueryWarning = True
            rstDoc.Edit
            rstDoc!RecordSource = Null
            rstDoc.Update
        Else
            Debug.Print "Updating RecordSource for '" & rstDoc!DocName & "'"
            rstDoc.Edit
            rstDoc!RecordSource = sql
            rstDoc.Update
        End If
    End If
    rstDoc.MoveNext
Loop
rstDoc.Close
Debug.Print "RecordSources have been updated"
If QueryWarning Then MsgBox "One or more records could not be updated, see debug log", vbCritical
End Sub

Public Sub CheckVersion(ForceUpgrade As Boolean)
'
' See if a more recent version of the database is available on the server.
'
Dim ver As String
Dim retval



On Error GoTo CheckVersionErr

'If CurrentProject.Name = "RA.accdb" Or CurrentProject.Name = "RA.accde" Or CurrentProject.Name = "RAtest.accdb" Then

    Open dbLocation & "version.dat" For Input Access Read As #1
    Input #1, ver
    Close #1
    If DBVersion <> ver And UCase$(Right$(CurrentProject.Name, 5)) <> "accdb" Then
        If ForceUpgrade Then
            'Open Environ("temp") & "\UpdateRA.bat" For Output As #1
            'Print #1, "If Not ""%localappdata%"" == """" Goto VarOK"
            'Print #1, "Set localappdata=%userprofile%\Local Settings\Application Data"
            'Print #1, ": VarOK"
            'Print #1, "Set /A Counter=1"
            'Print #1, ": Test"
            'Print #1, "If Not Exist ""%localappdata%\Programs\RA.laccdb"" Goto Continue"
            'Print #1, "Set /A Counter=%Counter%+1"
            'Print #1, "If %Counter% GEQ 500 Goto Stuck"
            'Print #1, "Goto Test"
            'Print #1, ": Continue"
            'Print #1, "If Exist ""%localappdata%\Programs"" Goto DirOK"
            'Print #1, "MkDir ""%localappdata%\Programs"""
            'Print #1, ": DirOK"
            'Print #1, "Call %systemroot%\setlocalserver.bat"
            'Print #1, "If Not ""%copyfromserver%"" == """" Goto ServerOK"
            'Print #1, "Set copyfromserver=FileServer"
            'Print #1, ": ServerOK"
            'Print #1, "copy \\%copyfromserver%\Applications\Database\RA.accde ""%localappdata%\Programs\RA.accde"""
            'Print #1, "cscript \\%copyfromserver%\Applications\Database\CreateShortcut.vbs"
            'Print #1, "Start """ & SysCmd(acSysCmdAccessDir) & "msaccess.exe"" ""%localappdata%\Programs\RA.accde"""
            'Print #1, "Exit"
            'Print #1, ": Stuck"
            'Print #1, "Del ""%localappdata%\Programs\RA.laccdb"""
            'Print #1, "If Not Exist ""%localappdata%\Programs\RA.laccdb"" Goto Continue"
            'Print #1, "Pause ""Cannot update Rosenberg & Associates because it is still running.  Contact support if you need more help."""
            'Close #1
            'retval = Shell(Environ("temp") & "\UpdateRA.bat", vbNormalFocus)
            MsgBox "A newer version of this system is available.  Please exit the system and start it again to get the latest version.  Rosie will now close to force the update, please restart Rosie!  If this message appears after you have just started Rosie, please open a ticket.", vbExclamation
            DoCmd.Quit
        Else
    
            MsgBox "A newer version of this system is available.  Please exit the system and start it again to get the latest version.  In 15 minutes, the system will upgrade automatically!", vbExclamation
        End If
    End If
    
    'If Dir$(Environ("ProgramFiles") & "\Chilkat Software Inc\Chilkat FTP\ChilkatFTP.dll") = "" Then
    '    MsgBox "Chilkat FTP Module is missing, this program may not run correctly.  Please report this message to technical support.", vbExclamation
    'End If
    
    Exit Sub

'End If


CheckVersionErr:
'MsgBox "Network error: Unable to verify that you are running the latest version of this database. And the time now is: " & Now(), vbExclamation
'DoCmd.Quit
Exit Sub

End Sub
