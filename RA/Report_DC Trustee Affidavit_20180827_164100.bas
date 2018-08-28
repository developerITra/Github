VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DC Trustee Affidavit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

If IsNull(txtPub1) Then
    txtPub1 = Format(InputBox("Please enter the next Publication Date.  (2 of 4)  Format mm/dd/yyyy"), "mmmm d, yyyy")
    If Nz(txtPub1) = "" Then Cancel = 1
End If

If IsNull(txtPub2) Then
    txtPub2 = Format(InputBox("Please enter the next Publication Date.  (3 of 4)  Format mm/dd/yyyy"), "mmmm d, yyyy")
    If Nz(txtPub2) = "" Then Cancel = 1
End If

If IsNull(txtPub3) Then
    txtPub3 = Format(InputBox("Please enter the next Publication Date.  (4 of 4)  Format mm/dd/yyyy"), "mmmm d, yyyy")
    If Nz(txtPub3) = "" Then Cancel = 1
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

'Private Sub GroupFooter4_Format(Cancel As Integer, FormatCount As Integer)
'If CommissionAmount = 0 Then
'    txtCommissionLabel.Visible = False
'    CommissionAmount.Visible = False
'End If
'End Sub

Private Sub Report_Page()
Call FirmMargin(Me, FileNumber)
End Sub
