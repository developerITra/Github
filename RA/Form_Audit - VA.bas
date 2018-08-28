VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Audit - VA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Auditor_NotInList(NewData As String, response As Integer)
MsgBox "If the auditor's name is not in this list, return to the main screen, then click Auditors. Make sure the auditor is listed properly.", vbCritical
response = 0
Auditor = Null
End Sub

Private Sub cmdAcrobat_Click()
Call PrintDoc(-2)
End Sub

Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdPrintAudit_Click()
Call PrintDoc(acNormal)
End Sub

Private Sub cmdPreviewAudit_Click()
Call PrintDoc(acPreview)
End Sub

Private Sub cmdWord_Click()
Call PrintDoc(-1)
End Sub

Private Sub Commission_AfterUpdate()
CommissionAmount = Commission * Proceeds
End Sub

Private Sub CommissionAmount_AfterUpdate()
Commission = CommissionAmount / Proceeds
End Sub

Private Sub Form_Current()

Call PropertyTaxesPOC_AfterUpdate
If Nz(Proceeds) = 0 Then Proceeds = Nz([SalePrice])
If IsNull(Abstractor) Then Abstractor = DLookup("Name", "Abstractors", "ID=" & DLookup("Abstractor", "JurisdictionList", "JurisdictionID=" & JurisdictionID))

Interest.Enabled = (Nz(Disposition) = 2)        ' enable these for 3rd party sales only
Interest3From.Enabled = (Nz(Disposition) = 2)
Interest3To.Enabled = (Nz(Disposition) = 2)
UnpaidBalance.Enabled = (Nz(Disposition) = 2)
InterestRate.Enabled = (Nz(Disposition) = 2)

If IsNull(Interest3From) Then Interest3From = Sale
If IsNull(Interest3To) Then Interest3To = Settled
If IsNull(UnpaidBalance) Then UnpaidBalance = SalePrice - Deposit
If IsNull(Interest) Then Interest = UnpaidBalance * DateDiff("d", Interest3From, Interest3To) * InterestRate / 365#

If IsNull(InterestFrom) Then InterestFrom = DateAdd("d", 1, StatementOfDebtDate)
If IsNull(InterestTo) Then InterestTo = Sale
If IsNull(InterestPerDiem) Then InterestPerDiem = StatementOfDebtPerDiem

If Nz(PropertyTaxes) = 0 And Not IsNull(Sale) Then Call CalcTaxes

Auditor.RowSource = "SELECT ID, Name FROM Auditors WHERE Jurisdiction=" & JurisdictionID & " ORDER BY Name;"

Attorney.RowSource = "SELECT Staff.ID, Staff.Name & ', ' & [CommonWealthTitle] AS CWRep " & _
                       "FROM Staff " & _
                       "WHERE (((Staff.CommonwealthTitle) Is Not Null and staff.active = true)) " & _
                       "ORDER BY Staff.CommonwealthTitle, Staff.Sort;"



End Sub

Private Sub CalcTaxes()
Dim yearbegin As Date, yearend As Date

'If PropertyTaxesPaid Then
'    If Month(Sale) <= 6 Then
'        yearend = DateSerial(Year(Sale), 6, 30)
'    Else
'        yearend = DateSerial(Year(Sale) + 1, 6, 30)
'    End If
'    PropertyTaxes = AnnualTaxes * DateDiff("d", Sale, yearend) / 365#
'Else
'    If Month(Sale) <= 6 Then
'        yearbegin = DateSerial(Year(Sale) - 1, 7, 1)
'    Else
'        yearbegin = DateSerial(Year(Sale), 7, 1)
'    End If
'    PropertyTaxes = AnnualTaxes * DateDiff("d", Sale, yearbegin) / 365#
'End If
'PropertyTaxesDesc = Format$(AnnualTaxes, "Currency") & " " & IIf(PropertyTaxesPaid, "", "un") & "paid as of " & Format$(Sale, "m/d/yyyy")

End Sub

Private Sub PropertyTaxesPOC_AfterUpdate()
If PropertyTaxesPOC Then
    PropertyTaxes = 0
    PropertyTaxes.Enabled = False
Else
    PropertyTaxes.Enabled = True
End If
End Sub

Private Sub PrintDoc(PrintTo As Integer)
Dim ErrCnt As Integer
On Error GoTo PrintDocErr

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
If MsgBox("Update Audit Filed = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
    Forms!foreclosuredetails!AuditFile = Date
    AddStatus [FileNumber], Now(), "Filed Audit"
End If

DoReport "Audit Cover Letter VA", PrintTo
DoReport "Audit VA", PrintTo
If MsgBox("Do you need the Bid Instructions?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then DoReport "Bid Instructions VA", PrintTo

DoEvents        ' seems to be needed or else report cancels itself ??????
Exit Sub

PrintDocErr:
    If Err.Number = 7878 And ErrCnt < 3 Then ' seems to be a synchronization issue when updating Forms!ForeclosureDetails!AuditFile
        ErrCnt = ErrCnt + 1
        DoEvents
        Resume
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub cmdNoteholderIsInvestor_Click()

On Error GoTo Err_cmdNoteholderIsInvestor_Click
Noteholder = Investor

Exit_cmdNoteholderIsInvestor_Click:
    Exit Sub

Err_cmdNoteholderIsInvestor_Click:
    MsgBox Err.Description
    Resume Exit_cmdNoteholderIsInvestor_Click
    
End Sub
