Attribute VB_Name = "Tasks"
Option Compare Database
Option Explicit

Public Function CreateTask(FullName As String, Subject As String, Message As String) As Boolean
Dim Username As Variant

Username = DLookup("Username", "Staff", "Fullname=""" & FullName & """")
If IsNull(Username) Then
    MsgBox "Cannot create task: Cannot find username for " & FullName
    CreateTask = False
    Exit Function
End If

End Function
