VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SetDisposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click
If FCDis Then
FCDis = False
    If lstDisposition.Column(0) = 1 Or lstDisposition.Column(0) = 2 Then
        If IsNull(Forms!foreclosuredetails!SalePrice) Or IsNull(Forms!foreclosuredetails!Purchaser) Or IsNull(Forms!foreclosuredetails!PurchaserAddress) Then
            MsgBox ("There is missing data from one of the fields (Sale Price, Purchaser  Purchaser Address), You can not set disposition")
            CheckCancelDisposition = False
            'DoCmd.Close
            DoCmd.Close acForm, "SetDisposition"
        Else
            
            SelectedDispositionID = lstDisposition.Column(0)
            
            'DoCmd.Close
            DoCmd.Close acForm, "SetDisposition"
        End If
    Else
     SelectedDispositionID = lstDisposition.Column(0)
     'DoCmd.Close
     DoCmd.Close acForm, "SetDisposition"
    End If
Else

SelectedDispositionID = lstDisposition.Column(0)
'DoCmd.Close
DoCmd.Close acForm, "SetDisposition"
End If


Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
CheckCancelDisposition = False
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
  Dim strSQL As String
    
  Select Case getDispositionType
    Case "FC"    ' set foreclosure disposition
       Me.lstDisposition.RowSource = "SELECT FCDisposition.ID, FCDisposition.Disposition, FCDisposition.Completed, FCDisposition.StatusInfo From FCDisposition WHERE (((FCDisposition.ID) <> 33 And (FCDisposition.ID) <> 34))ORDER BY FCDisposition.Disposition;"

    Case "TR"    ' set title resolution disposition
       Me.lstDisposition.RowSource = "SELECT TRDisposition.ID, TRDisposition.Disposition, TRDisposition.Completed, TRDisposition.StatusInfo FROM TRDisposition order by Disposition;"
    Case "CD"    ' bankruptcy cramdown disposition
       Me.lstDisposition.RowSource = "SELECT BKCDDisposition.ID, BKCDDisposition.Disposition, BKCDDisposition.Completed, BKCDDisposition.StatusInfo FROM BKCDDisposition order by Disposition;"
    Case "BK-FINAL"    ' bankruptcy  disposition
       Me.lstDisposition.RowSource = "SELECT BKFinalDisposition.ID, BKFinalDisposition.Disposition FROM BKFinalDisposition order by Disposition;"
    Case "LM"          ' loss mediation
        strSQL = "SELECT LMDisposition.LMDispositionID, LMDisposition.LMDisposition"
        strSQL = strSQL + " FROM LMDisposition"
        strSQL = strSQL + " WHERE (((LMDisposition.LMDispositionID) <> 4))"
        strSQL = strSQL + " ORDER BY LMDisposition.LMDisposition;"
    'Me.lstDisposition.RowSource = "SELECT LMDisposition.LMDispositionID, LMDisposition.LMDisposition FROM LMDisposition order by LMDisposition;"
        Me.lstDisposition.RowSource = strSQL
    Case "Mo"
    Me.lstDisposition.RowSource = "SELECT FCDisposition.ID, FCDisposition.Disposition, FCDisposition.Completed, FCDisposition.StatusInfo From FCDisposition WHERE (((FCDisposition.ID) = 8 Or (FCDisposition.ID) = 6 Or (FCDisposition.ID) = 33 Or (FCDisposition.ID) = 34))ORDER BY FCDisposition.Disposition;"

  End Select
  
  

End Sub


Private Sub lstDisposition_DblClick(Cancel As Integer)
Call cmdOK_Click
End Sub

Private Function getDispositionType() As String

  getDispositionType = Me.OpenArgs

End Function
