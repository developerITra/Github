VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdOpen_Click()
On Error GoTo Err_cmdOpen_Click

If IsNull(List) Then
    If List.ListCount > 1 Then
        MsgBox "Select a file to open.", vbExclamation
    Else
        MsgBox "Type a few characters to begin the search.", vbExclamation
    End If
    Exit Sub
End If
DoCmd.Minimize
OpenCase Me!List

Exit_cmdOpen_Click:
    Exit Sub

Err_cmdOpen_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpen_Click
    
End Sub

Private Sub UpdateList()
Dim Active As String

List.RowSource = ""
UserText.SetFocus
If Nz(UserText.Text) = "" Then
    Status = "Type a few letters to begin the search"
    Exit Sub
End If

Status = "Searching..."

If optActive Then
    Active = "Active = True AND "
Else
    Active = ""
End If

Select Case SearchBy
    
    Case 1      ' names
'        List.ColumnCount = 3
'        List.ColumnWidths = ".5 in; 1.25 in; 1 in"
'        List.RowSource = "SELECT FileNumber,Name,Company FROM qrySearchNames " & _
'            "WHERE " & Active & "Name LIKE """ & UserText.Text & "*"" " & _
'            "OR Company LIKE """ & UserText.Text & "*"" " & _
'            "ORDER BY Name"

        List.ColumnCount = 6
        List.ColumnWidths = ".5 in; .3 in ; .5 in; 1.25 in; 1 in"
        List.RowSource = "SELECT FileNumber,casecode,shortclientName,Name,Company FROM qrySearchNames " & _
            "WHERE " & Active & "Name LIKE ""*" & UserText.Text & "*"" " & _
            "OR Company LIKE ""*" & UserText.Text & "*"" " & _
            "ORDER BY Name"


    Case 2      ' address
        List.ColumnCount = 6
        List.ColumnWidths = ".5 in;.3 in; .5 in; 1.9 in; 1 in; 0 in"
        List.RowSource = "SELECT FileNumber,casecode,shortclientName,PropertyAddress,[Unit#] FROM qrySearchAddress " & _
            "WHERE " & Active & "(PropertyAddress LIKE ""*" & UserText.Text & "*"") " & _
            "ORDER BY PropertyAddress"
    
    Case 10      ' loan # - quick
        List.ColumnCount = 7
        List.ColumnWidths = ".5 in;.3 in; .5 in; .7 in; .7 in; .7in; .7 in"
        List.RowSource = "SELECT FileNumber,casecode,shortclientName,FCDetailsLoanNumber,COLDetailsLoanNumber,CivDetailsLoanNumber,PrimaryDefName FROM qrySearchloan " & _
            "WHERE " & Active & "(FCDetailsLoanNumber LIKE ""*" & UserText.Text & "*"" " & _
            "OR COLDetailsLoanNumber LIKE """ & UserText.Text & "*"" " & _
            "OR CivDetailsLoanNumber LIKE """ & UserText.Text & "*"") " & _
            "ORDER BY FCDetailsLoanNumber, COLDetailsLoanNumber, CIVDetailsLoanNumber"


    Case 3      ' loan # - long  'These are identical...*scratches head*
        List.ColumnCount = 7
        List.ColumnWidths = ".5 in;.3 in; .5 in; .7 in; .7 in; .7in; .7 in"
        List.RowSource = "SELECT FileNumber,casecode,shortclientName,FCDetailsLoanNumber,COLDetailsLoanNumber,CivDetailsLoanNumber,PrimaryDefName FROM qrySearchloanLike " & _
            "WHERE " & Active & "(FCDetailsLoanNumber LIKE ""*" & UserText.Text & "*"" " & _
            "OR COLDetailsLoanNumber LIKE """ & UserText.Text & "*"" " & _
            "OR CivDetailsLoanNumber LIKE """ & UserText.Text & "*"") " & _
            "ORDER BY FCDetailsLoanNumber, COLDetailsLoanNumber, CIVDetailsLoanNumber"

    Case 4      ' client #
        List.ColumnCount = 5
        List.ColumnWidths = ".5 in; .3 in; .5 in; 1.5 in; 1 in"
        List.RowSource = "SELECT FileNumber, caseCode,ShortclientName,ClientNumber,PrimaryDefName FROM qrySearchClientNumber " & _
            "WHERE " & Active & "ClientNumber LIKE """ & UserText.Text & "*"" " & _
            "ORDER BY ClientNumber"
    
    Case 5      ' court case #
        List.ColumnCount = 6
        List.ColumnWidths = ".5 in; .3 in; .5in; .7 in; .7 in; .7 in"
        List.RowSource = "SELECT FileNumber,casecode, shortclientName,CourtCaseNumber,CaseNo,CaseNumber FROM qrySearchCourtCase " & _
            "WHERE " & Active & "(CourtCaseNumber LIKE """ & UserText.Text & "*"" " & _
            "OR CaseNo LIKE """ & UserText.Text & "*"" " & _
            "OR CaseNumber LIKE """ & UserText.Text & "*"") " & _
            "ORDER BY CourtCaseNumber"

    Case 6      ' project name
        List.ColumnCount = 5
        List.ColumnWidths = ".5 in; .3 in; .5 in; 1.5 in; 1 in"
        List.RowSource = "SELECT FileNumber,casecode,shortclientName,PrimaryDefName,PropertyAddress FROM qrySearchProjects " & _
            "WHERE " & Active & "PrimaryDefName LIKE """ & UserText.Text & "*"" " & _
            "ORDER BY PrimaryDefName, PropertyAddress"
    
    Case 7      ' invoice #
        List.ColumnCount = 6
        List.ColumnWidths = ".5 in; .3 in; .5 in; .7 in; .7 in; .7 in"
        List.RowSource = "SELECT FileNumber,casecode,shortclientName,InvoiceNumber,PrimaryDefName,PropertyAddress FROM qrySearchInvoices " & _
            "WHERE " & Active & "InvoiceNumber LIKE """ & UserText.Text & "*"" " & _
            "ORDER BY InvoiceNumber, PrimaryDefName, PropertyAddress"
            
    Case 8      ' FNMA loan #
        List.ColumnCount = 5
        List.ColumnWidths = ".5 in; .3 in; .5 in; .7 in; .7 in"
        List.RowSource = "SELECT FileNumber,casecode,shortClientName,FNMALoanNumber,PrimaryDefName FROM qrySearchFNMA " & _
            "WHERE " & Active & "(FNMALoanNumber LIKE """ & UserText.Text & "*"") " & _
            "ORDER BY FNMALoanNumber"
            
    Case 9      ' FHLMC loan #
        List.ColumnCount = 5
        List.ColumnWidths = ".5 in; .3 in; .5 in; .7 in; .7 in"
        List.RowSource = "SELECT FileNumber,casecode, shortClientName,FHLMCLoanNumber,PrimaryDefName FROM qrySearchFHLMC " & _
            "WHERE " & Active & "(FHLMCLoanNumber LIKE """ & UserText.Text & "*"") " & _
            "ORDER BY FHLMCLoanNumber"
        
End Select
If List.ListCount < 1 Then
    Status = "No matches"
Else
    If List.ListCount = 1 Then
        List.Selected(0) = True
        List.Value = List.Column(0)
        Status = "1 match"
    Else
        Status = List.ListCount & " matches, select a file or type a few more letters"
    End If
End If

End Sub

Private Sub Form_Activate()
UserText.SetFocus
End Sub

Private Sub Form_Timer()
Me.TimerInterval = 0
Call UpdateList
End Sub

Private Sub List_DblClick(Cancel As Integer)
Call cmdOpen_Click
End Sub

Private Sub optActive_Click()
Call UpdateList
End Sub

Private Sub SearchBy_Click()
Call UpdateList
End Sub

Private Sub UserText_Change()
Me.TimerInterval = 500
End Sub
