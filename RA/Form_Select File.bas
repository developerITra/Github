VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Select File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdOK_Click()
Dim files As Recordset
Dim FN As Long              ' file number

On Error GoTo Err_cmdOK_Click
If IsNull(FileNum) And IsNull(List.Value) Then
    MsgBox "Enter a file number, or select a file from the list", vbExclamation
    Exit Sub
End If

If IsNull(FileNum) Then
    FN = List.Value
Else
    FN = FileNum
End If

Set files = CurrentDb.OpenRecordset("SELECT FileNumber FROM Caselist WHERE FileNumber=" & FN, dbOpenSnapshot)
If files.EOF Then
    MsgBox "No such file number: " & FN
    Me!FileNum = ""
    Me!FileNum.SetFocus
Else
    OpenCase FN
    DoCmd.Close acForm, "Select File"
    On Error GoTo Err_cmdOK_DontCare
    Forms![Case List].SetFocus
    On Error GoTo Err_cmdOK_Click
End If
files.Close

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
Err_cmdOK_DontCare:
    Resume Next
    
End Sub

Private Sub FileNum_Change()
List.Value = Null
End Sub

Private Sub Form_Open(Cancel As Integer)
FileNum = List.Column(0, 0)
End Sub

Private Sub List_Click()
FileNum = List.Value
End Sub

Private Sub List_DblClick(Cancel As Integer)
Call cmdOK_Click
End Sub

