VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EnterTitleOutReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdComplete_Click()


Dim rstdocs As Recordset
Set rstdocs = CurrentDb.OpenRecordset("Select Top 1 TitleRecieved From TitleReceivedArchive Where FileNumber =" & FileNumber & " Order by TitleRecieved DESC", dbOpenDynaset, dbSeeChanges)
If rstdocs.EOF Then
MsgBox (" Please notice that there is no old Ttile received date, Contact IT")
Exit Sub
End If
rstdocs.Close

Set rstdocs = CurrentDb.OpenRecordset("Select Top 1 TitleThrough From TitleThroughArchive Where FileNumber =" & FileNumber & " Order by TitleThrough DESC", dbOpenDynaset, dbSeeChanges)
If rstdocs.EOF Then
MsgBox (" Please notice that there is no old Ttile Through date, Contact IT")
Exit Sub
End If
rstdocs.Close

Set rstdocs = CurrentDb.OpenRecordset("Select Top 1 TitleReviewToClient From TitleReviewArchive Where FileNumber =" & FileNumber & " Order by TitleReviewToClient DESC", dbOpenDynaset, dbSeeChanges)
If rstdocs.EOF Then
MsgBox (" Please notice that there is no old Ttile Review date, Contact IT")
Exit Sub
End If
rstdocs.Close

        If IsNull(TexReason.Value) Then
        MsgBox ("Please inseart the reason")
        Exit Sub
        Else
        
        
        Forms!foreclosuredetails!cmdWizComplete.Visible = True
        Forms!foreclosuredetails!cmdWizComplete.Caption = "Title Cancelled completed"
        Forms!foreclosuredetails!cmdWizComplete.SetFocus
        Forms!foreclosuredetails!cmdWaiting.Visible = False
        Forms!foreclosuredetails!cmdcloserestart.Visible = False
        Forms!foreclosuredetails!TitleCancelledReason = Forms!EnterTitleOutReason!TexReason
        DoCmd.Close
        End If

End Sub

Private Sub cmdUpload_Click()
If IsNull(Forms!foreclosuredetails!TitleThru) Then
MsgBox ("Must have a good thrugh date")
Exit Sub
End If

BillTitle = False
BillTitleUpdate = False

DoCmd.OpenForm "DocsWindow", , , "FileNumber = " & Forms!foreclosuredetails!FileNumber
Forms!DocsWindow!Command111.Visible = True


End Sub






Private Sub Command22_Click()
TexReason.Visible = True
cmdComplete.Visible = True

End Sub
