VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Declaration Nationstar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim NextNumber As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
'chCoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And [BKdetails.Chapter] = 13)
End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
'Dim X As Single
'Dim y1 As Single, y2 As Single, y3 As Single
'
'Me.ScaleMode = 1    ' twips (1440 twips per inch)
'
'X = Me![LeftBox].Left + Me![LeftBox].Width + 1440 / 12
'y1 = Me![LeftBox].Top
'y2 = Me![LeftBox].Top + Me![LeftBox].Height + 1440 / 12
''y3 = Me![LeftBox2].Top + Me![LeftBox2].Height + 1440 / 12
'
'Me.DrawWidth = 4    ' in pixels
'Me.Line (X, y1)-(X, y3)         ' vertical line
'Me.Line (Me![LeftBox].Left, y2)-(Me.Width, y2)  ' horizontal line
'Me.Line (Me![LeftBox].Left, y3)-(Me.Width, y3)  ' horizontal line

End Sub

Private Sub Report_Close()
Forms![Print Declaration].Visible = True

End Sub

Private Sub Report_Current()
If [Forms]![BankruptcyPrint].[PosPetpayHis] = True Then
    Label136.Visible = True
    Text138.Visible = True
Else
    Label136.Visible = False
    Text138.Visible = False
End If

If [Forms]![BankruptcyPrint].[PosPetFeeCos] = True Then
    Label139.Visible = True
    Text140.Visible = True
Else
    Label139.Visible = False
    Text140.Visible = False
End If

If [Forms]![BankruptcyPrint].[PosPetTaxInsAdvAdd] = True Then
    Label141.Visible = True
    Text142.Visible = True
Else
    Label141.Visible = False
    Text142.Visible = False
End If

    
End Sub

Private Sub Report_Open(Cancel As Integer)
Forms![Print Declaration].Visible = False

If [Forms]![BankruptcyPrint].[PosPetpayHis] = True Then
    Label136.Visible = True
    Text138.Visible = True
Else
    Label136.Visible = False
    Text138.Visible = False
End If

If [Forms]![BankruptcyPrint].[PosPetFeeCos] = True Then
    Label139.Visible = True
    Text140.Visible = True
Else
    Label139.Visible = False
    Text140.Visible = False
End If

If [Forms]![BankruptcyPrint].[PosPetTaxInsAdvAdd] = True Then
    Label141.Visible = True
    Text142.Visible = True
Else
    Label141.Visible = False
    Text142.Visible = False
End If
' 2012.02.09 DaveW

'Me.lblResponseDate.Caption = Forms!BankruptcyPrint!txtResponseDate
End Sub


Private Sub Report_Page()
'If [FCdetails.State] = "VA" Then
'    Call FirmMarginVA(Me, FileNumber, , PrimaryDefName, PropertyAddress)
'Else
    'Call FirmMargin(Me, FileNumber, , PrimaryDefName, PropertyAddress, , Nz([Fair Debt]))
'End If

'If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 1487, ""), 450, 6000, True)


End Sub

Private Sub Text192_BeforeUpdate(Cancel As Integer)

End Sub
