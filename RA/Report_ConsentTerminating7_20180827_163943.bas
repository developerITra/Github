VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_ConsentTerminating7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Title As String
Dim N1 As String, N2 As String

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
chCoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And [BKdetails.Chapter] = 13)

Select Case Forms![Print Consent Order Terminating 7]!optTrustee
    
    Case 0  '   not a party
        Select Case Forms![Print Consent Order Terminating 7]!optDebtor
            Case 0  ' defaulted
                Title = "Order Terminating Automatic Stay"
                N1 = "the debtor(s) having failed to respond"
                N2 = vbNewLine & vbNewLine & "          ORDERED that the Motion for Relief from Automatic Stay be and it is hereby, entered by Default as to the Debtor(s); and be it further"
            Case 1  ' agreed
                Title = "Consent Order Terminating Automatic Stay"
                N1 = "the parties having reached an agreement"
                N2 = ""
        End Select

    Case 1  ' filed report
        Select Case Forms![Print Consent Order Terminating 7]!optDebtor
            Case 0
                Title = "Order Terminating Automatic Stay"
                N1 = "the trustee having filed a report of no distribution, the debtor(s) having failed to respond"
                N2 = vbNewLine & vbNewLine & "          ORDERED that the Motion for Relief from Automatic Stay be and it is hereby, entered by Default as to the Debtor(s); and be it further"
            Case 1
                Title = "Consent Order Terminating Automatic Stay"
                N1 = "the trustee having filed a report of no distribution, the parties having reached an agreement"
                N2 = ""
        End Select

    Case 2  ' defaulted
        Select Case Forms![Print Consent Order Terminating 7]!optDebtor
            Case 0
                Title = "Order Terminating Automatic Stay"
                N1 = "the parties having failed to respond"
                N2 = vbNewLine & vbNewLine & "          ORDERED that the Motion for Relief from Automatic Stay be and it is hereby, entered by Default as to the Debtor(s) and the Chapter 7 Trustee; and be it further"
            Case 1
                Title = "Consent Order Terminating Automatic Stay as to Debtor and Default as to Chapter 7 Trustee"
                N1 = "the trustee having failed to file an answer, the debtor(s) and movant having reached an agreement"
                N2 = vbNewLine & vbNewLine & "          ORDERED that the Motion for Relief from Automatic Stay be an it is hereby, entered by Default as to the Chapter 7 Trustee; and be it further"
            End Select

    Case 3  ' agreed
        Select Case Forms![Print Consent Order Terminating 7]!optDebtor
            Case 0
                Title = "Consent Order Terminating Automatic Stay as to Trustee and Default as to Debtor(s)"
                N1 = "the trustee and movant having reached an agreement, the debtor(s) having failed to respond"
                N2 = vbNewLine & vbNewLine & "          ORDERED that the Motion for Relief from Automatic Stay be and it is hereby, entered by Default as to the Debtor(s); and be it further"
            Case 1
                Title = "Consent Order Terminating Automatic Stay"
                N1 = "the parties having reached an agreement"
                N2 = ""
        End Select

End Select
End Sub

Private Function GetTitle() As String
GetTitle = Title
End Function

Private Function GetN1() As String
GetN1 = N1
End Function

Private Function GetN2() As String
GetN2 = N2
End Function

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
Dim X As Single, x0 As Single
Dim y1 As Single, y2 As Single, y3 As Single

Me.ScaleMode = 1    ' twips (1440 twips per inch)

X = Me![LeftBox].Left + Me![LeftBox].Width + 1440 / 12
x0 = Me![LeftBox].Left
y1 = Me![LeftBox].Top
y2 = Me![LeftBox].Top + Me![LeftBox].Height + 1440 / 12
y3 = Me![LeftBox2].Top + Me![LeftBox2].Height + 1440 / 12

Me.DrawWidth = 4    ' in pixels
Me.Line (X, y1)-(X, y3)         ' vertical line
Me.Line (x0, y2)-(Me.Width, y2)  ' horizontal line
Me.Line (x0, y3)-(Me.Width, y3)  ' horizontal line

End Sub

Private Sub GroupFooter3_Format(Cancel As Integer, FormatCount As Integer)
Cancel = ([Districts.State] <> "VA")
End Sub

Private Sub Report_Page()
Call FirmMargin(Me, FileNumber, AttorneyInfo)
End Sub
