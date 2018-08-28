VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TestOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()

Dim rs As Recordset
Dim sqlstr As String
sqlstr = "SELECT top 1 Warning FROM Journal where FileNumber=" & 31600 & " order by Journaldate DESC"
Set rs = CurrentDb.OpenRecordset(sqlstr)

If Not rs.EOF Then

MsgBox (rs!Warning)
 
 
End If

rs.Close
Set rs = Nothing
End Sub


'Warning

