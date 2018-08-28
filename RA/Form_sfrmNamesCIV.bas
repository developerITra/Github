VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmNamesCIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdDelete_Click()

On Error GoTo Err_cmdDelete_Click

DoCmd.RunCommand acCmdDeleteRecord

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click

End Sub

Private Sub cmdManageClients_Click()
On Error GoTo Err_cmdManageClients_Click

  Dim strAttorneyName As String
  
  If (IsNull([ID])) Then
    MsgBox "Please enter name details before continuing.", vbCritical, "Name Details"
    Exit Sub
  End If
  
  If (Me.Plaintiff = True Or Me.Defendant = True) Then
    MsgBox "Person cannot be a plantiff or defendant.", vbCritical, "Name Details"
    Exit Sub
  End If
  
  strAttorneyName = [First] & " " & [Last]
  DoCmd.OpenForm "sfrmNamesCIVAttorneyRep", , , "[AttorneyNameID] = " & [ID], , , strAttorneyName
  

Exit_cmdManageClients_Click:
  Exit Sub
  
Err_cmdManageClients_Click:
  MsgBox Err.Description
  Resume Exit_cmdManageClients_Click
End Sub

Private Sub cmdPrintLabel_Click()

Dim rstLabelData As Recordset, sql As String, strCriteria As String


On Error GoTo Err_cmdPrintLabel_Click

sql = "SELECT CaseList.FileNumber, CaseList.PrimaryDefName, ClientList.ShortClientName, NamesCiv.Company,  NamesCiv.Last, NamesCiv.First, NamesCiv.Address, NamesCiv.City, NamesCiv.State, NamesCiv.Zip " & _
        "FROM ClientList INNER JOIN (CaseList INNER JOIN NamesCiv ON CaseList.FileNumber = NamesCiv.FileNumber) ON ClientList.ClientID = CaseList.ClientID " & _
        "WHERE ID=" & Me!ID

Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)


Do While Not rstLabelData.EOF
  
    Call StartLabel
    Print #6, FormatName(rstLabelData!Company, rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, "", rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
    Print #6, "|FONTSIZE 8"
    Print #6, "|BOTTOM"
    Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
    Call FinishLabel
    rstLabelData.MoveNext
Loop

rstLabelData.Close
Set rstLabelData = Nothing

Exit_cmdPrintLabel_Click:
    Exit Sub

Err_cmdPrintLabel_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintLabel_Click

End Sub

Private Sub Defendant_AfterUpdate()
If Defendant Then
    Plaintiff = False
    Other = False
    DeleteAttorneyClients
    Me.AttorneyForSummary.Requery
    
   ' AttorneyForNameID = Null
End If
End Sub

Private Sub Other_AfterUpdate()
If Other Then
    Plaintiff = False
    Defendant = False
   ' AttorneyForNameID = Null
End If
End Sub

Private Sub Plaintiff_AfterUpdate()
If Plaintiff Then
    Defendant = False
    Other = False
    
    DeleteAttorneyClients
    Me.AttorneyForSummary.Requery
  '  AttorneyForNameID = Null
End If
End Sub



Public Sub DeleteAttorneyClients()

DoCmd.SetWarnings False
DoCmd.RunSQL ("DELETE from  NamesCIVAttorneyRep WHERE AttorneyNameID = " & ID)
DoCmd.SetWarnings True

End Sub
