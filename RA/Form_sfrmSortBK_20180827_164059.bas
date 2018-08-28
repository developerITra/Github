VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmSortBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmTable As String
' 2012.03.08 DaveW Turned-off disable

Private Sub Form_Open(Cancel As Integer)
Dim rstqueue As Recordset, fld As Field
Dim s As String
Dim u, qry, sorter As String
Dim i, J As Integer

' Set frmTable for this subform.
frmTable = Me.Parent.lstFiles.RowSource
'Set tbl = Application.CurrentData.AllTables(frmTable)
'ShowAllProperties (tbl)

'Set rstqueue = CurrentDb.OpenRecordset("Select * FROM " & frmTable, dbOpenDynaset, dbSeeChanges)

s = ""
On Error Resume Next
s = Me.Parent.txtInitialSort
Resume Next

With Me.sfrmSort_cmbSortBy
.RowSource = frmTable
.Requery
If "" = s _
Then
Else
    Call RunSort(s)
    On Error Resume Next
    .Value = s
    On Error GoTo 0
End If
.Enabled = True
End With

End Sub

Private Sub SortTheList()
Dim s As String, sFld As String

With Me.sfrmSort_cmbSortBy
    .Requery
    On Error Resume Next
    sFld = .Value
    On Error GoTo 0
End With


'sFld = ""
If sFld <> "" _
Then
    Call RunSort(sFld)
End If

End Sub
Private Sub RunSort(sFld As String)
Dim s As String

If sFld <> "" _
Then
    s = "select * from " & frmTable & " order by"
    s = s & " [" & sFld & "]"
    If 2 = sfrmSort_frmSortOrder.Value Then s = s & " desc"
    s = s & ";"
    With Me.Parent
        .lstFiles.RowSource = s
        '.Refresh
    End With
End If

End Sub
Private Sub sfrmSort_cmbSortBy_Change()
Me.sfrmSort_frmSortOrder.Enabled = True
Call SortTheList
End Sub

Private Sub sfrmSort_cmbSortBy_Click()
Me.sfrmSort_frmSortOrder.Enabled = True
End Sub

'Private Sub lstSortBy_AfterUpdate()
'Debug.Print "lstSortBy_AfterUpdate"
'    Call SortTheList
'End Sub

Private Sub sfrmSort_frmSortOrder_Click()
 'Debug.Print "sfrmSort_frmSortOrder_Click"
 Call SortTheList
End Sub

Function listQueryFields(strQryName As String) As String
On Error GoTo listQueryFields_Error
    Dim db As DAO.Database
    Dim qryfld As DAO.QueryDef
    Dim fld As Field
 
    Set db = CurrentDb()
    Set qryfld = db.QueryDefs(strQryName)
    For Each fld In qryfld.Fields    'loop through all the fields of the Query
        Debug.Print fld.Name
    Next
    
Error_Handler_Exit:
    Set qryfld = Nothing
    Set db = Nothing
    Exit Function
 
listQueryFields_Error:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.Number & vbCrLf & "Error Source: listQueryFields" & vbCrLf & _
    "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function
