VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CivilDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub SetAnswerDue()

Select Case CourtID
    Case 1      ' District
    Case 2      ' Circuit
        Select Case State
            Case "MD"
                AnswerDue = DateAdd("d", 30, ClientServed)
            Case "VA"
                AnswerDue = DateAdd("d", 21, ClientServed)
        End Select
        AddStatus FileNumber, Date, "Answer Due " & Format$(AnswerDue, "m/d/yyyy")

    Case 3, 5   ' Superior or Federal
        AnswerDue = DateAdd("d", 20, ClientServed)
        AddStatus FileNumber, Date, "Answer Due " & Format$(AnswerDue, "m/d/yyyy")

    Case 4      ' Bankruptcy
        AnswerDue = DateAdd("d", 30, SummonsIssued)
        AddStatus FileNumber, Date, "Answer Due " & Format$(AnswerDue, "m/d/yyyy")
End Select

End Sub

Private Sub ClientServed_AfterUpdate()
AddStatus FileNumber, ClientServed, "Client Served"
Call SetAnswerDue
End Sub

Private Sub ClientServed_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ClientServed)
End Sub

Private Sub ClientServed_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ClientServed = Date
Call ClientServed_AfterUpdate
End If

End Sub

Private Sub cmdEditPropertyDetails_Click()

On Error GoTo Err_EditPropertyDetails_Click

DoCmd.OpenForm "EditPropertyDetailsCiv", , , WhereCondition:="FileNumber= " & Forms!CivilDetails!FileNumber ' & " And Current = true"

Exit_EditPropertyDetails_Click:
    Exit Sub

Err_EditPropertyDetails_Click:
    MsgBox Err.Description
    Resume Exit_EditPropertyDetails_Click

End Sub

Private Sub ComplaintFiled_AfterUpdate()
AddStatus FileNumber, ComplaintFiled, "Complaint Filed"
End Sub

Private Sub ComplaintFiled_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(ComplaintFiled)
End Sub

Private Sub ComplaintFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
ComplaintFiled = Date
Call ComplaintFiled_AfterUpdate
End If

End Sub



Private Sub DispositionDate_AfterUpdate()
  DeleteFutureHearings (DispositionDate)
  
    
End Sub


Private Sub DispositionID_AfterUpdate()
If Not IsNull(DispositionID) Then
   
'    DispositionDate.Enabled = True
'    DispositionDate.Locked = False
'    DispositionDate.BackColor = 16777215

    Call SetObjectAttributes(DispositionDate, True)
    DispositionDate = Date
    AddStatus FileNumber, Date, "Disposition: " & DispositionID.Column(1)
 
     
'    DeleteFutureHearings (DispositionDate)
End If
End Sub

Private Sub FCTab_Change()
Select Case FCTab.Value     ' 0 based
    Case 3      ' status
        sfrmStatus.Requery
End Select
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

Private Sub cmdPrint_Click()

On Error GoTo Err_cmdPrint_Click
If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm "CivilPrint", , , "[CaseList].[FileNumber]=" & Me![FileNumber]

Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrint_Click
    
End Sub

Private Sub cmdSelectFile_Click()

On Error GoTo Err_cmdSelectFile_Click
DoCmd.Close
DoCmd.OpenForm "Select File"

Exit_cmdSelectFile_Click:
    Exit Sub

Err_cmdSelectFile_Click:
    MsgBox Err.Description
    Resume Exit_cmdSelectFile_Click
    
End Sub


Private Sub Form_Current()

Me.Caption = "Civil File " & Me![FileNumber] & " " & [PrimaryDefName]

If (IsNull(Me.DispositionID)) Then
  Call SetObjectAttributes(DispositionDate, False)
Else
  Call SetObjectAttributes(DispositionDate, True)
End If

If (IsNull(Me.ResponseID)) Then
  Call SetObjectAttributes(ResponseDate, False)
Else
  Call SetObjectAttributes(ResponseDate, True)
End If


End Sub

Private Sub Form_Open(Cancel As Integer)
If FileReadOnly Or EditDispute Then

    Dim ctl As Control
    Dim lngI As Long
    Dim bSkip As Boolean

    For Each ctl In Form.Controls
    Select Case ctl.ControlType
    Case acTextBox, acComboBox, acListBox, acOptionGroup, acCheckBox, acSubform, acOptionButton
        
            If Not (ctl.Locked) Then ctl.Locked = True
            
    Case acCommandButton
        bSkip = False
            If ctl.Name = "cmdClose" Then bSkip = True
            If ctl.Name = "cmdSelectFile" Then bSkip = True
            If Not bSkip Then ctl.Enabled = False
         
       
    End Select
    Next
End If
End Sub

Private Sub ResponseID_AfterUpdate()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
If Not IsNull(ResponseID) Then
    Call SetObjectAttributes(ResponseDate, True)
    ResponseDate = Date
    
    AddStatus FileNumber, Date, "Response: " & ResponseID.Column(1)
End If
End If

End Sub

Private Sub SummonsIssued_AfterUpdate()
AddStatus FileNumber, SummonsIssued, "Summons Issued"
Call SetAnswerDue
End Sub

Private Sub SummonsIssued_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(SummonsIssued)
End Sub

Private Sub SummonsIssued_DblClick(Cancel As Integer)
SummonsIssued = Date
Call SummonsIssued_AfterUpdate
End Sub


Private Sub DeleteFutureHearings(pDispositionDate As Date)

Dim rstHearings As Recordset
Dim i As Integer

Set rstHearings = CurrentDb.OpenRecordset("SELECT * FROM CivHearings WHERE FileNumber=" & Me!FileNumber & " AND Hearing > #" & pDispositionDate & "#;", dbOpenDynaset, dbSeeChanges)

With rstHearings

  If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
  
      If (Not IsNull(rstHearings![HearingCalendarEntryID])) Then
        DeleteCalendarEvent (rstHearings![HearingCalendarEntryID])
      End If
  
      .Delete
      .MoveNext
    Loop
  End If
  
  .Close
End With

Me.sfrmCIVHearing.Requery


End Sub
