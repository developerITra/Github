VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Permissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function GetStaffNames(Priv As String) As String
Dim rsStaff As Recordset

Set rsStaff = CurrentDb.OpenRecordset("SELECT Name FROM Staff WHERE active=true and " & Priv & " ORDER BY Name", dbOpenSnapshot)
If rsStaff.EOF Then
    GetStaffNames = ""
Else
    Do While Not rsStaff.EOF
        GetStaffNames = GetStaffNames & rsStaff!Name & ", "
        rsStaff.MoveNext
    Loop
    GetStaffNames = Left$(GetStaffNames, Len(GetStaffNames) - 2)
End If
rsStaff.Close
End Function


Private Function GetPermissionNames(Priv As String) As String
Dim rsPermission As Recordset

Set rsPermission = CurrentDb.OpenRecordset("SELECT Description FROM PermissionRules WHERE  " & Priv & " ORDER BY Name", dbOpenSnapshot)
If rsPermission.EOF Then
    GetPermissionNames = ""
Else
    Do While Not rsPermission.EOF
        GetPermissionNames = GetPermissionNames & rsPermission!Description & ", "
        rsPermission.MoveNext
    Loop
    GetPermissionNames = GetPermissionNames
End If
rsPermission.Close
End Function

