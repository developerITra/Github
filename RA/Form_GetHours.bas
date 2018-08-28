VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetHours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click

'3_30_15*********
If Forms![Case List]!CaseType = "Bankruptcy" Then
txtProcess = "BK_NOFC"
End If
'**
If Nz(txtTotal) <= 0 Then
    MsgBox "Amount must be greater than zero", vbCritical
    Exit Sub
End If
Dim Amount As Currency, Rate As Currency, rstBillReasons As Recordset
If Forms![Case List]!CaseTypeID = 7 Then
Rate = Nz(DLookup("NonStandardevFee", "ClientList", "ClientID=" & Forms![Case List]![ClientID]), 0)
Else
Rate = Nz(DLookup("NonStandardFee", "ClientList", "ClientID=" & Forms![Case List]![ClientID]), 0)
    If Forms![Case List]!Des.Value <> "" Then
        Forms![Case List]!BillCase = True
        Forms![Case List]![BillCaseUpdateReasonID] = 31
        Forms![Case List]!BillCaseUpdateDate = Date
        Forms![Case List]!BillCaseUpdateUser = GetStaffID
        Forms![Case List]!lblBilling.Visible = True
        Forms![Case List]!lstBillingReasons.Requery
        
        Set rstBillReasons = CurrentDb.OpenRecordset("Select * FROM BillingReasonsFCarchive where filenumber=" & Forms![Case List]!FileNumber, dbOpenDynaset, dbSeeChanges)
        With rstBillReasons
        .AddNew
        !FileNumber = Forms![Case List]!FileNumber
        !billingreasonid = 31
        !UserID = GetStaffID
        !Date = Date
        .Update
        End With
    End If
End If
If Rate = 0 Then
MsgBox "Hours cannot be entered because this client does not have an hourly rate entered", vbCritical
Exit Sub
End If
Amount = Rate * txtTotal
AddInvoiceItem Forms![Case List]!FileNumber, txtProcess, txtDesc, Amount, 0, True, True, False, False
MsgBox "Hours accepted"
DoCmd.Close

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
txtProcess = Me.OpenArgs
End Sub
