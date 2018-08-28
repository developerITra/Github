VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Statement of Debt with Figures Wells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
Dim args() As String
Dim InterestRateNote As String
args = Split(ReportArgs, "|")
txtName = args(0)
txtTitle = args(1)

If Forms!foreclosuredetails!LoanType <> 2 And Forms!foreclosuredetails!LoanType <> 3 Then
Select Case Forms![Print Statement of Debt]!cbxRateType
Case "Fixed"
InterestRateNote = "Per diem interest in the amount of " & Format(Forms![Print Statement of Debt]!PerDiem, "currency") & " will accrue on the principal from " & Forms![Print Statement of Debt]!txtDueDate
Case "Variable"
InterestRateNote = "Per diem interest in the amount of _____ will accrue on the principal from " & Forms![Print Statement of Debt]!txtDueDate & " to the next interest rate change date and accrue thereafter in accordance with the variable rate as set forth in the Note"
Case "HELOC"
InterestRateNote = "A daily variable per diem will accrue on the principal in accordance with the variable rate as set forth in the Note"
End Select
Else
InterestRateNote = "Interest will continue to accrue according to the terms of the note."
End If
End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
Dim X As Single
Dim y1 As Single, y2 As Single

Me.ScaleMode = 1    ' twips (1440 twips per inch)

X = Me![LeftBox].Left + Me![LeftBox].Width + 1440 / 12
y1 = Me![LeftBox].Top
y2 = Me![LeftBox].Top + Me![LeftBox].Height + 1440 / 12

Me.DrawWidth = 4    ' in pixels
Me.Line (X, y1)-(X, y2)         ' vertical line
Me.Line (Me![LeftBox].Left, y2)-(Me.Width, y2)  ' horizontal line

End Sub

Private Sub Report_Page()
Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress, , Nz([Fair Debt]))
If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 223, ""), 450, 7000, True)
End Sub
'2012.02.09 Davw: Changed:
'IIf(([RemainingPBal]>[OriginalPBal]),"Additional Interest","Paid on principal")
'=[OriginalPBal]-[RemainingPBal]
