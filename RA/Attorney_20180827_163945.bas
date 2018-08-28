Attribute VB_Name = "Attorney"
Option Compare Database
Option Explicit




Public Function AssignmentNames(File As Long, Fmt As Integer) As String
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
Dim ls_state As String



cnt = 0
If File = 0 Then File = Forms![Case List]!FileNumber


ls_state = Nz(DLookup("State", "FCDetails", "Filenumber=" & File & " and Current=true"))
If (Len(ls_state) = 0) Then
  MsgBox "Please fill in the foreclosure state before continuing.", vbExclamation
  Exit Function
End If
Set t = CurrentDb.OpenRecordset("SELECT Name FROM Staff WHERE Active=true and attorney=true and Practice" & ls_state & "=true", dbOpenSnapshot)

Do While Not t.EOF
    cnt = cnt + 1
    Names(cnt) = t("Name") & ", Esq."
    t.MoveNext
Loop
t.Close

Select Case Fmt
    Case 1      ' comma separated list
        For i = 1 To cnt
            AssignmentNames = AssignmentNames & Names(i)
            If i < cnt Then AssignmentNames = AssignmentNames & ", "
        Next i
    Case 2      ' comma separated list, except AND last name
        For i = 1 To cnt
            AssignmentNames = AssignmentNames & Names(i)
            If i < cnt Then
                If i < cnt - 1 Then
                    AssignmentNames = AssignmentNames & ", "
                Else
                    AssignmentNames = AssignmentNames & " and "
                End If
            End If
        Next i
    Case 3      ' one name per line
        For i = 1 To cnt
            AssignmentNames = AssignmentNames & Names(i) & vbCrLf
            
        Next i
    Case 4      ' signature lines
        For i = 1 To cnt
            AssignmentNames = AssignmentNames & "___________________________________" & vbNewLine
            AssignmentNames = AssignmentNames & vbNewLine
            AssignmentNames = AssignmentNames & Names(i)
            AssignmentNames = AssignmentNames & vbNewLine & vbNewLine
        Next i
    Case Else
        MsgBox "Invalid format in call to AssignmentNames", vbExclamation
        AssignmentNames = ""
End Select

End Function
