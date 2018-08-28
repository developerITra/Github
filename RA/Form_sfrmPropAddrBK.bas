VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmPropAddrBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
   
   
   If IsNull(LoanNumber) Then
    LoanNumber.Locked = False
    LoanNumber.BackStyle = 1
    'Call SetObjectAttributes(LoanNumber, True)
  Else  ' this allows for copying
    LoanNumber.Locked = True
    LoanNumber.BackStyle = 0
    'Call SetObjectAttributes(LoanNumber, False)
  End If
  
  SetFNMAEnabled
  
End Sub

Private Sub SetFNMAEnabled()

Dim lt As Integer
  If (IsNull(Me.LoanType)) Then
    lt = 0
  Else
    lt = Me.LoanType
  End If
  
  If (lt = 4) Then
    FNMALoanNumber.Locked = False
    FNMALoanNumber.BackStyle = 1
    
    FHLMCLoanNumber.Locked = True
    FHLMCLoanNumber.BackStyle = 0
    
  ElseIf (lt = 5) Then ' FHLMC (Freddie)
    FNMALoanNumber.Locked = True
    FNMALoanNumber.BackStyle = 0
    
    FHLMCLoanNumber.Locked = False
    FHLMCLoanNumber.BackStyle = 1
    
  
  Else
    FNMALoanNumber.Locked = True
    FNMALoanNumber.BackStyle = 0
    
    FHLMCLoanNumber.Locked = True
    FHLMCLoanNumber.BackStyle = 0
    
    
  End If


End Sub

Private Sub LoanType_AfterUpdate()
  SetFNMAEnabled
  
End Sub

