VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmSoftHold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Current()
SetHold.Enabled = PrivSetDisposition

End Sub

Private Sub SetHold_Click()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else

    If Not PrivSetDisposition Then
    MsgBox (" You do not have permission for Soft Hold disposition")
    Exit Sub
    Else
        If MsgBox(" You are going to add Soft Hold disposition, Are you sure? ", vbYesNo) = vbNo Then
        Exit Sub
        Else
        Dim rstBillReasons As Recordset
    
        Forms![Case List]!BillCase = True
        Forms![Case List]!BillCaseUpdateUser = GetStaffID()
        Forms![Case List]!BillCaseUpdateDate = Date
        Forms![Case List]![BillCaseUpdateReasonID] = 32
        Forms![Case List]!lblBilling.Visible = True
        'Forms![Case List].SetFocus
       ' DoCmd.RunCommand acCmdSaveRecord
        'Forms![ForeclosureDetails].SetFocus
    
        Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & FileNumber, dbOpenDynaset, dbSeeChanges)
        With rstBillReasons
        .AddNew
        !FileNumber = FileNumber
        !billingreasonid = 32
        !UserID = GetStaffID
        !Date = Date
        .Update
        End With
        Set rstBillReasons = Nothing
       
        Forms!foreclosuredetails!sfrmSoftHold.Form.SoftID.Enabled = True
        Forms!foreclosuredetails!sfrmSoftHold.Form.SoftID = 1
        Forms!foreclosuredetails!sfrmSoftHold.Form.SoftDate = Now()
        Forms!foreclosuredetails!sfrmSoftHold.Form.SoftStaffInitial = GetStaffInitials(GetStaffID())
        Forms!foreclosuredetails!sfrmSoftHold.Form.SoftStaffId = GetStaffID()
         Call RemoveDates
         
        DoCmd.SetWarnings False
        Dim jnltxt As String
        jnltxt = "Adding Soft Hold disposition"
        strinfo = jnltxt
        strinfo = Replace(strinfo, "'", "''")
        strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
        DoCmd.RunSQL strSQLJournal
        DoCmd.SetWarnings True
    
    
        AddStatus FileNumber, Now(), "Adding Soft Hold "
        Forms!Journal.Requery
        
        
    
        End If
    End If

End If

End Sub

