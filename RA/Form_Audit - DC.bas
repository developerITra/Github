VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Audit - DC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

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

Private Sub cmdWord_Click()
Call PrintDoc(-1)
End Sub

Private Sub Form_Current()

If Nz(StatementofDebt) = 0 Then StatementofDebt = StatementOfDebtAmount
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
If Nz(InterestPerDiem) = 0 Then InterestPerDiem = StatementOfDebtPerDiem

'If Nz(PropertyTaxes) = 0 Then Call CalcTaxes

If Me.State = "VA" Then
Attorney.RowSource = "SELECT Staff.ID, Staff.Name & ', ' & [CommonWealthTitle] AS CWRep " & _
                       "FROM Staff " & _
                       "WHERE  ((staff.active = true) And (Staff.Attorney =True) And(staff.PracticeVA = true )) " & _
                       "ORDER BY Staff.CommonwealthTitle, Staff.Sort;"
                   'It was  "WHERE (((Staff.CommonwealthTitle) Is Not Null)) and staff.active = true " S.A.
'staff.active=true
ElseIf Me.State = "MD" Then
  Attorney.RowSource = "SELECT Staff.ID, Staff.Name & ', Esq.' FROM Staff WHERE ((Staff.active = true ) and (Staff.Attorney = True) and (Staff.PracticeMD = True)) ORDER BY Staff.Sort;"
Else
  Attorney.RowSource = "SELECT Staff.ID, Staff.Name FROM Staff WHERE ((staff.active = true) and (Staff.Attorney = True) and (staff.PracticeDC = true ))ORDER BY Staff.Sort;"
    
End If






End Sub

'Private Sub CalcTaxes()
'Dim yearbegin As Date, yearend As Date
'
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
'
'End Sub

Private Sub cmdPrintAudit_Click()
Call PrintDoc(acNormal)
End Sub

Private Sub cmdPreviewAudit_Click()
Call PrintDoc(acPreview)
End Sub

Private Sub PrintDoc(PrintTo As Integer)

On Error GoTo PrintDocErr
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

If MsgBox("Update Audit Filed = " & Format$(Date, "m/d/yyyy") & "?", vbYesNo) = vbYes Then
    Forms!foreclosuredetails!AuditFile = Now()
    AddStatus [FileNumber], Now(), "Filed Audit"
End If

'MsgBox "These are Maryland Audit Reports, they may need to be configured for DC", vbCritical
'DoReport "Audit Cover Letter DC", PrintTo
'DoReport "Audit DC", PrintTo
'DoReport "Line Notifying of Audit Filing DC", PrintTo
DoReport "DC Accounting", PrintTo

If JurisdictionID = 7 Or JurisdictionID = 9 Or JurisdictionID = 16 Or JurisdictionID = 19 Or JurisdictionID = 22 Then   ' Caroline, Cecil, Kent, Queen Anne's, and Talbot
    
    DoCmd.OpenForm "PrintClientAudit", , , "FileNumber=" & Forms!foreclosuredetails!FileNumber, , acDialog, "Audit Affidavit Fiduciary|Audit Affidavit Fiduciary|" & PrintTo
    'DoReport "Audit Affidavit Fiduciary", PrintTo
End If
DoEvents        ' seems to be needed or else report cancels itself ??????
Exit Sub

PrintDocErr:
    MsgBox Err.Description

End Sub

Private Sub PropertyTaxesPaid_AfterUpdate()
'Call CalcTaxes
End Sub
