VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MissingDocsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
Dim rstdocs As Recordset, DocsFlag As Boolean
'Remove document record from table once received
DocsFlag = True
FileNbr = Forms![Case List]!FileNumber
Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DocumentMissing where filenbr=" & FileNbr & " AND ID=" & lstFiles.Value, dbOpenDynaset, dbSeeChanges)

With rstdocs
.Edit
!DocRecd = Now
!docrecdby = StaffID
.Update
.Close
End With

Set rstdocs = CurrentDb.OpenRecordset("Select * FROM DocumentMissing where filenbr=" & FileNbr, dbOpenDynaset, dbSeeChanges)

With rstdocs
Do Until .EOF
If IsNull(!DocRecd) Then
DocsFlag = False
End If
.MoveNext
Loop
End With
Set rstdocs = Nothing

If DocsFlag = True Then
'Note file in DocsMissing queue as Received
Dim rstqueue As Recordset
Set rstqueue = CurrentDb.OpenRecordset("Select * FROM wizardqueuestats where filenumber=" & FileNbr & " and current=true", dbOpenDynaset, dbSeeChanges)

With rstqueue
.Edit
!DocsRecdFlag = True
.Update
End With

End If

DoCmd.Close
    
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click


DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub



Private Sub lstFiles_DblClick(Cancel As Integer)

Call cmdOK_Click


End Sub
