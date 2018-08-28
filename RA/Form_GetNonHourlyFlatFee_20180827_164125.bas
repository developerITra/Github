VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetNonHourlyFlatFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click
'If IsNumeric(Me.txtDesc) = False Or Me.txtDesc = "" Or IsNull(Me.txtDesc) = True Then


If Forms![Case List]!CaseType = "Bankruptcy" Then
txtProcess = "BK_NOFC"
End If

If Nz(txtTotal) <= 0 Or IsNumeric(Me.txtTotal) = False Or Me.txtTotal = "" Or IsNull(Me.txtTotal) = True Then
    MsgBox "Amount must be greater than zero", vbCritical
    Exit Sub
End If
Dim Amount As Currency, rstBillReasons As Recordset
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
Amount = Format$(txtTotal, "Currency")
AddInvoiceItem Forms![Case List]!FileNumber, txtProcess, "Non-Standard Time Entry Flat Fee.", Amount, 0, True, True, False, False



DoCmd.SetWarnings False

strinfo = txtDesc
strinfo = Format$(Me.txtTotal, "Currency") & " Non-Standard Time Entry flat fee approved by client. " & strinfo

strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms![Case List].FileNumber & ",Now,GetFullName(),'" & strinfo & "',2 )"
DoCmd.RunSQL strSQLJournal
strinfo = ""

DoCmd.SetWarnings True
Forms!Journal.Requery


MsgBox "Flat Fee accepted"
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
