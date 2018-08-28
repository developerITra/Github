VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Motion for Relief BOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim NextNumber As Integer
Dim newNextNumber As Integer

Private Function GetNextNumber() As String
GetNextNumber = Space$(10) & NextNumber & ".  "
'Debug.Print NextNumber
NextNumber = NextNumber + 1
End Function

Private Function GetNewNextNumber() As String  'Second set of Numbers!
GetNewNextNumber = Space$(10) & newNextNumber & ".  "
'Debug.Print NextNumber
newNextNumber = newNextNumber + 1
End Function

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
chCoDebtor = (GetNames(0, 1, "BKCoDebtor=True") <> "" And [BKdetails.Chapter] = 13)
End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
Dim X As Single
Dim y1 As Single, y2 As Single, y3 As Single

Me.ScaleMode = 1    ' twips (1440 twips per inch)

X = Me![LeftBox].Left + Me![LeftBox].Width + 1440 / 12
y1 = Me![LeftBox].Top
y2 = Me![LeftBox].Top + Me![LeftBox].Height + 1440 / 12
y3 = Me![LeftBox2].Top + Me![LeftBox2].Height + 1440 / 12

Me.DrawWidth = 4    ' in pixels
Me.Line (X, y1)-(X, y3)         ' vertical line
Me.Line (Me![LeftBox].Left, y2)-(Me.Width, y2)  ' horizontal line
Me.Line (Me![LeftBox].Left, y3)-(Me.Width, y3)  ' horizontal line

End Sub

Private Sub Report_Load()
If (otherCosts = 0 Or IsNull(otherCosts)) = True Then
    Me.srptBKOtherCosts.Visible = False
Else
    Me.srptBKOtherCosts.Visible = True
End If
End Sub

Private Sub Report_Open(Cancel As Integer)
NextNumber = 1
newNextNumber = 1

If MsgBox("Is Debtor being evaluated for Loss Mit?", vbYesNo) = vbYes Then
    Me.txtLossMit.Visible = True
Else
    Me.txtLossMit.Visible = False
End If

End Sub

Private Sub Report_Page()

' It is necessary to restart the sequence after each page, because the entire
' report formats to produce each page.

NextNumber = 1
newNextNumber = 1
Call FirmMargin(Me, FileNumber)
End Sub

Private Function AssignmentInfo() As String
Select Case AssignBy
    Case 1              ' DOT
        Select Case AssignByDOT
            Case 1      ' assignment
                AssignmentInfo = OriginalBeneficiary & " assigned its interest to " & Investor
            Case 2      ' merger
                AssignmentInfo = Trim$(MergerInfo) & " is now the beneficiary of the Deed of Trust due to merger."
        End Select
    Case 2              ' note
        AssignmentInfo = "The Promissory Note has been transferred from " & OriginalBeneficiary & " to " & Investor
        
    
        


If Right$(AssignmentInfo, 1) <> "." Then AssignmentInfo = AssignmentInfo & "."


End Select

If IsNull(Right$(AssignmentInfo, 1)) Then AssignmentInfo = ""

End Function
