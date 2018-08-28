VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ForeclosurePrintDisposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PrintDocs(PrintTo As Integer)
Dim ReportName As String, rstLabelData As Recordset, LabelData As String, sql As String, i As Integer, FeeAmount As Currency, noticecnt As Integer, rstJnl As Recordset

'On Error GoTo Err_PrintDocs
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord        ' might need to save the attorney name

If ChNoteAffidavit Then
    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If
    Call DoReport("Note Affidavit", PrintTo)
End If


If ChCollateralFileAffidavit Then
    If (IsNull([Forms]![Foreclosureprint]!Attorney)) Then
    MsgBox ("Missing attorney who will sign.")
    Exit Sub
    End If
   
    DoCmd.OpenForm "Print Affidavit Collateral file", , , , , acDialog, PrintTo
End If
 


If ChLabel Then
    If PrintTo = acViewNormal Then
        sql = "SELECT DISTINCTROW FCdetails.FileNumber, FCdetails.PrimaryLastName, FCdetails.PrimaryFirstName, FCdetails.SecondaryLastName, FCdetails.SecondaryFirstName, FCdetails.PropertyAddress, FCdetails.City, FCdetails.State, FCdetails.ZipCode, FCdetails.LoanNumber, CaseList.Active, FCdetails.Current FROM (CaseList INNER JOIN FCdetails ON CaseList.FileNumber = FCdetails.FileNumber) LEFT JOIN FCDisposition ON FCdetails.Disposition = FCDisposition.ID WHERE (((FCdetails.FileNumber)=" & Forms![Case List]!FileNumber & ") AND ((CaseList.Active)=1) AND ((FCdetails.Current)=1));"

        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rstLabelData.EOF
            Call StartLabel
            'Print #6, "|FONTSIZE 11"
            Print #6, rstLabelData!PrimaryFirstName & "  " & rstLabelData!PrimaryLastName
            Print #6, rstLabelData!SecondaryFirstName & "  " & rstLabelData!SecondaryLastName
            Print #6, rstLabelData!PropertyAddress
            Print #6, rstLabelData!City; ", " & rstLabelData!State & " " & rstLabelData!ZipCode
            Print #6, "Loan # " & rstLabelData!LoanNumber
            Print #6, "|BOTTOM"
            Print #6, "File # " & rstLabelData!FileNumber
            Call FinishLabel
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
      
        End If
End If

If chTitleReview Then Call DoReport("Title Review", PrintTo)

If chDismissCase Then Call DoReport("Dismiss Case", PrintTo)

If chCertofService Then
   DoCmd.OpenForm "Print Certificate of Service", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , , PrintTo
End If

 
Exit Sub

Err_PrintDocs:
    MsgBox Err.Description
    Exit Sub

End Sub

Private Sub Form_Current()
Me.Caption = "Print Foreclosure " & [CaseList.FileNumber] & " " & [PrimaryDefName]


End Sub

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close
Call refreshFCform
Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdWord_Click()
Call PrintDocs(-1)
End Sub

Private Sub cmdPrint_Click()
Call PrintDocs(acViewNormal)
End Sub

Private Sub cmdView_Click()
Call PrintDocs(acPreview)
End Sub

Private Sub cmdAcrobat_Click()
Call PrintDocs(-2)
End Sub


