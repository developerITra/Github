VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmScans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
Dim rstScans As Recordset, Filespec As String

On Error GoTo OpenErr

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM Scans;"
DoCmd.SetWarnings True

Set rstScans = CurrentDb.OpenRecordset("Scans", dbOpenDynaset)

Filespec = Dir$(ClosedScanLocation & Me.OpenArgs & "*.pdf")
Do While Filespec <> ""
    rstScans.AddNew
    rstScans!Filespec = Filespec
    rstScans.Update
    Filespec = Dir$()
Loop
rstScans.Close
lstFiles.Requery

Exit Sub

OpenErr:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub cmdAll_Click()
Dim i As Long

On Error GoTo Err_cmdAll_Click

For i = 0 To lstFiles.ListCount - 1
    lstFiles.Selected(i) = True
Next i

Exit_cmdAll_Click:
    Exit Sub

Err_cmdAll_Click:
    MsgBox Err.Description
    Resume Exit_cmdAll_Click
    
End Sub

Private Sub cmdInvert_Click()
Dim i As Long

On Error GoTo Err_cmdInvert_Click

For i = 0 To lstFiles.ListCount - 1
    If lstFiles.Selected(i) Then
        lstFiles.Selected(i) = False
    Else
        lstFiles.Selected(i) = True
    End If
Next i

Exit_cmdInvert_Click:
    Exit Sub

Err_cmdInvert_Click:
    MsgBox Err.Description
    Resume Exit_cmdInvert_Click
    
End Sub

Private Sub cmdView_Click()
Dim i As Long

On Error GoTo Err_cmdView_Click

For i = 0 To lstFiles.ListCount - 1
    If lstFiles.Selected(i) Then StartDoc (lstFiles.Column(1, i))
Next i

Exit_cmdView_Click:
    Exit Sub

Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click
    
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
StartDoc (lstFiles.Column(1))
End Sub
