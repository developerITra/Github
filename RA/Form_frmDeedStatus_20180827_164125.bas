VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDeedStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click


DoCmd.Close

'Forms!frmFile.Requery
'Forms!frmFile.cmdEditStatus.Enabled = False

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub


Private Sub cmdPreView_Click()
Dim FileNumber As Long
Dim rs As Recordset

FileNumber = Forms!foreclosuredetails!FileNumber

'Set rs = CurrentDb.OpenRecordset("SELECT * FROM qryInvestor WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

Select Case frmOption

Case 1
  'Investor
        Me.txtStatus = " " & Forms!foreclosuredetails!Investor & " is foreclosing in 1st position.  Title is clear for purpose of foreclosure. "
   
Case 2

        Me.txtStatus = " TITLE NOT CLEAR – DO NOT PROCEED:"

    
End Select

'rs.Close
'Set rs = Nothing

End Sub


Private Sub cmdupdate_Click()

Dim rs As Recordset
Dim FileNumber As Long

FileNumber = Forms!foreclosuredetails!FileNumber


Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

If Not rs.EOF Then

rs.Edit
rs!TitleReviewStatus = Me.txtStatus
rs.Update

rs.Close
Set rs = Nothing

MsgBox ("File updated")
Forms!foreclosuredetails.Refresh
'Forms!frmFile.Requery

End If

End Sub

