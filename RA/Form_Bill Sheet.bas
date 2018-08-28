VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Bill Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub

Private Sub cmdComplete_Click()
'Dim rsFee, rsCost, rsBillsheet As Recordset

'Set rsFee = CurrentDb.OpenRecordset("SELECT * FROM tb_BillSheetFees join invoiceitems on invoiceitems.invoiceid = tb_BillSheetFees.invoiceId WHERE Fee = true and Filenumber=" & txtFilenum, dbOpenDynaset, dbSeeChanges)
'Set rsCost = CurrentDb.OpenRecordset("SELECT * FROM tb_BillSheetCost join invoiceitems on invoiceitems.invoiceid = tb_BillSheetFees.invoiceId WHERE Fee = 0 and Filenumber=" & txtFilenum, dbOpenDynaset, dbSeeChanges)
'Set rsBillsheet = CurrentDb.OpenRecordset("SELECT * FROM invoiceitems WHERE Filenumber=" & txtFilenum, dbOpenDynaset, dbSeeChanges)
DoCmd.SetWarnings False

DoCmd.OpenQuery ("UpdateBillSheetFees")
DoCmd.OpenQuery ("UpdateBillSheetCost")


DoCmd.SetWarnings True
DoCmd.Close

End Sub
