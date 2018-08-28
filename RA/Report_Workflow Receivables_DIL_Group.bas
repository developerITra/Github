VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Workflow Receivables_DIL_Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim NoData As Boolean

Private Sub Report_NoData(Cancel As Integer)
NoData = True
End Sub

Private Function GetRecordCount() As String
If NoData Then
    GetRecordCount = "No files"
Else
    If txtRC = 1 Then
        GetRecordCount = "1 file"
    Else
        GetRecordCount = txtRC & " files"
    End If
End If
End Function

Private Sub Report_Open(Cancel As Integer)
  Me.GroupLevel(0).ControlSource = Forms!ReportsWorkflow!cboARRecAll
End Sub

