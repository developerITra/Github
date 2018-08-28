VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmDisbursingSurplusUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ComUndo_Click()
    Me.Undo
    DoCmd.Close
End Sub

Private Sub ComUpdate_Click()
'Dont mess with me  -____-
Dim rsSurplus As Recordset
Dim strSurplus As String

strSurplus = "SELECT DisbursingSurplus.ID, DisbursingSurplus.FileNumber, DisbursingSurplus.DSurplus"
strSurplus = strSurplus + " FROM DisbursingSurplus"
strSurplus = strSurplus + " Where FileNumber =" & Forms![Case List]!FileNumber & ";"

Set rsSurplus = CurrentDb.OpenRecordset(strSurplus, dbOpenDynaset, dbSeeChanges)



If Me.NewRecord Then
    If IsNull(Me.DSurplus) Then
        Me.Undo
        MsgBox ("You must select a Disbursing Surplus Type")
        DoCmd.Close acForm, "sfrmDisbursingSurplusUpdate", acSaveNo
    Else
        Forms!Journal.cmdNewJournalEntry_Click
        Forms![Journal New Entry]!Info = "Disbursing Surplus: " & Me.DSurplus & " in the amount of " & Format$(Me.DSAmount, "Currency") + vbCrLf + vbCrLf
        Forms![Journal New Entry]!chAccounting = True
        DoCmd.Close acForm, "sfrmDisbursingSurplusUpdate", acSaveYes
        
        'Try this
        rsSurplus.MoveFirst
        If IsNull(rsSurplus!DSurplus) = True Then
            rsSurplus.Delete
        Else
            'Pay no attention to the man behind the curtain
        End If
        
        'If IsLoadedF("ForeclosureDetails") = True Then Forms!ForeclosureDetails!sfrmDisbursingSurplus.Requery
    End If
Else
    If IsNull(Me.DSurplus) Then
        Me.Undo
        MsgBox ("You can not remove all data")
        Exit Sub
    Else
        Forms!Journal.cmdNewJournalEntry_Click
        Forms![Journal New Entry]!Info = "Edited Disbursing Surplus " & Me.DSurplus & " in the amount of " & Format$(Me.DSAmount, "Currency") + vbCrLf + vbCrLf
        Forms![Journal New Entry]!chAccounting = True
        DoCmd.Close acForm, "sfrmDisbursingSurplusUpdate", acSaveYes
        'If IsLoadedF("ForeclosureDetails") = True Then Forms!ForeclosureDetails!sfrmDisbursingSurplus.Requery
    End If

End If

rsSurplus.Close
Set rsSurplus = Nothing
Forms!foreclosuredetails.sfrmDisbursingSurplusTable.Requery

End Sub


