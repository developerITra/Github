VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmForeclosure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
If (Forms![Case List]![CaseTypeID] = 1 Or Forms![Case List]![CaseTypeID] = 7) Then ' foreclosure or eviction
  Call SetObjectAttributes(State, False) ' cannot edit
Else
  Call SetObjectAttributes(State, True)
End If

If IsNull(LoanNumber) Then
    LoanNumber.Locked = False
    LoanNumber.BackStyle = 1
    'Call SetObjectAttributes(LoanNumber, True)
Else  ' this allows for copying/pasting
    LoanNumber.Locked = True
    LoanNumber.BackStyle = 0
    'Call SetObjectAttributes(LoanNumber, False)
End If
End Sub



Private Sub Form_Open(Cancel As Integer)
  UpdateLoanNumberVisuals
End Sub

Private Sub LoanType_AfterUpdate()

 UpdateLoanNumberVisuals

End Sub

Private Sub UpdateLoanNumberVisuals()
Dim lt As Integer

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

If IsNull(LoanType) Then
    lt = 0
Else
    lt = LoanType
End If
FHALoanNumber.Enabled = (lt = 2 Or lt = 3)    ' enable for VA or HUD
FNMALoanNumber.Enabled = (lt = 4)
FHLMCLoanNumber.Enabled = (lt = 5)

Forms![EvictionDetails].SetCashForKeysDate


  
End Sub

Private Sub Purchaser_AfterUpdate()
AddStatus FileNumber, Sale, "Property sold to " & Purchaser & " for " & Format$(SalePrice, "Currency")
End Sub

Private Sub SaleRat_AfterUpdate()
AddStatus FileNumber, SaleRat, "Sale ratified/confirmed"
End Sub

Private Sub SaleRat_BeforeUpdate(Cancel As Integer)
Cancel = CheckFutureDate(SaleRat)
End Sub
