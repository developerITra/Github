VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDeedTaxes"
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

Private Sub cmdPreView_Click()
Dim FileNumber As Long

FileNumber = Forms!foreclosuredetails!FileNumber

Select Case frmOption

Case 1
    Me.txtTaxes = " ***Taxes are current*** "
   
Case 2
    Me.txtTaxes = " Year " & Me.txtTaxYear & " taxes are delinquent in the amount of $" & Me.txtTaxAmt & "."
   
End Select

End Sub



Private Sub cmdupdate_Click()

Dim rs As Recordset
Dim FileNumber As Long

FileNumber = Forms!foreclosuredetails!FileNumber


Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

If Not rs.EOF Then

rs.Edit
rs!TitleReviewTaxes = Me.txtTaxes
rs.Update

rs.Close
Set rs = Nothing

MsgBox ("File updated")
Forms!foreclosuredetails.Refresh

End If

End Sub




