VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DocsWindowLitigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdComplete_Click()

If IsNull(lstDocs.Column(0)) Then
MsgBox ("Please select File")
Exit Sub
End If


Dim JurText As String

Dim i As Integer
 DoCmd.SetWarnings False
    For i = 0 To lstDocs.ListCount


        If lstDocs.Selected(i) = True Then
        JurText = "Complete the hold " & Forms!DocsWindowLitigation.lstDocs.Column(3, i)
         DoCmd.RunSQL ("UPDATE DocIndex set Hold = '' WHERE DocID = " & Forms!DocsWindowLitigation.lstDocs.Column(0, i))
         DoCmd.RunSQL ("UPDATE Accou_LitigationBillingQueue set Hold = '',Dismissed = True , MangerQ = False  WHERE DocIndexID = " & Forms!DocsWindowLitigation.lstDocs.Column(0, i))
         DoCmd.RunSQL ("Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values( '" & Forms![Case List]!FileNumber & "', #" & Now & "#,'" & GetFullName() & "','" & JurText & "',2 )")

         lstDocs.Selected(i) = False








      End If

    Next i

DoCmd.SetWarnings True
Me.lstDocs.Requery
Me.Form.Requery
Forms!Journal.Requery


   
End Sub

Private Sub Command111_Click()
DoCmd.Close

End Sub

Private Sub Form_Open(Cancel As Integer)
Forms![DocsWindowLitigation]!lstDocs.ColumnCount = 6
Forms![DocsWindowLitigation]!lstDocs.ColumnWidths = "0 in; 0.4 in; 0.75 in; 3 in; 0 in ;0.3 in "
Forms![DocsWindowLitigation]!lstDocs.RowSource = "SELECT DocID, Initials, Format(Datestamp,""mm/dd/yyyy"") AS [Date Entered], Filespec as [File Name], [doctitleid] AS DocType , Hold FROM DocIndex LEFT JOIN Staff ON DocIndex.StaffID=Staff.ID WHERE FileNumber=" & FileNumber & " AND DocGroup='B' AND Hold = 'H' AND DocTitleID = 1546 AND Filespec IS NOT NULL and DeleteDate is null"
Forms![DocsWindowLitigation]!lstDocs.Requery

End Sub




Private Sub cmdAll_Click()
Dim i As Long

On Error GoTo Err_cmdAll_Click

For i = 0 To lstDocs.ListCount - 1
     lstDocs.Selected(i) = True
Next i

Exit_cmdAll_Click:
    Exit Sub

Err_cmdAll_Click:
    MsgBox Err.Description
    Resume Exit_cmdAll_Click
    
End Sub

Private Sub cmdInvert_Click()
Dim i As Long

On Error GoTo Err_cmdInvert_Click

For i = 0 To lstDocs.ListCount - 1
    If lstDocs.Selected(i) Then
        lstDocs.Selected(i) = False
    Else
        lstDocs.Selected(i) = True
    End If
Next i

Exit_cmdInvert_Click:
    Exit Sub

Err_cmdInvert_Click:
    MsgBox Err.Description
    Resume Exit_cmdInvert_Click
    
End Sub


    

Private Sub lstDocs_DblClick(Cancel As Integer)
Call cmdView_Click

End Sub

Private Sub cmdView_Click()
Dim i As Long

On Error GoTo Err_cmdView_Click

For i = 0 To lstDocs.ListCount - 1

        Select Case lstDocs.Column(4, i)
      
        Case 1511, 1513, 1514, 1515, 1516, 1517, 1518, 1519, 1520, 1521, 1522, 1523, 1524, 1525, 1526, 1528, 1557, 1558, 1571, 1572
        'If lstDocs.Column(4, i) = (1511 Or 1513 Or 1514 Or 1515 Or 1516 Or 1517 Or 1518 Or 1519 Or 1520 Or 1521 Or 1522) Then
        If PrivSSN Then
                If lstDocs.Selected(i) Then StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\SSN\" & lstDocs.Column(3, i)
                Else
                MsgBox ("You are not authirized to open SSN")
                Exit Sub
                End If
        Case Else
        If lstDocs.Selected(i) Then StartDoc DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\" & lstDocs.Column(3, i)
       ' End If
       End Select
       
        
Next i

Exit_cmdView_Click:
    Exit Sub

Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click
    
        
    
    
End Sub
