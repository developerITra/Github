VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Add Check Request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboCRType_AfterUpdate()

  
  Dim ComputeAmt As Boolean
  Dim computeFunction As String
  
  
  ComputeAmt = Nz(DLookup("Computed", "Fees", GetFeeString()), False)
  If (ComputeAmt = True) Then
     computeFunction = Nz(DLookup("ComputedFunction", "Fees", GetFeeString()))
     Me.txtAmt = FetchComputedAmt(computeFunction)
  Else
    FetchAmount
  End If
  
  If (cboCRType.Column(1) = "Other") Then
    Me.lblPayableTo.Visible = True
    Me.txtPayableTo.Visible = True
    Me.lblDescription.Visible = True
    Me.txtDescription.Visible = True

  Else
    Me.lblPayableTo.Visible = False
    Me.txtPayableTo.Visible = False
    Me.lblDescription.Visible = False
    Me.txtDescription.Visible = False
  End If
End Sub

Private Sub FetchAmount()

Dim FeeAmt As Variant
Dim strFilter As String
Dim rCnt As Integer

strFilter = GetFeeString()

rCnt = DCount("ID", "Fees", strFilter)
If (rCnt = 0) Then
  MsgBox "The Check Request fee is not available.", vbCritical
  Exit Sub
End If

txtAmt = 0
FeeAmt = DLookup("Amount", "Fees", strFilter)
If (IsNull(FeeAmt)) Then
  Me.lblAmt.Visible = True
  Me.txtAmt.Visible = True
Else
  Me.lblAmt.Visible = False
  Me.txtAmt.Visible = False
  txtAmt = CCur(FeeAmt)
End If

End Sub

Private Function GetFeeString()
Dim State As String


State = DLookup("[State]", "[JurisdictionList]", "[JurisdictionID] = " & [Forms]![Case List]![JurisdictionID])
Select Case Me.cboCRType
    Case "FC-SOT", "FC-DISM", "BK-ASSIGN", "FC-ASSIGN", "FC-AUD", "FC-PROPTAX", "EV-PS30NTQ", "EV-PSCMPL", "EV-ALIASWRIT", "EV-CMPL", "EV-PPROC"
      GetFeeString = "Fees.FeeType = '" & cboCRType & "' and State = '" & State & "' "
    Case "FC-DOCK", "FC-LCERT", "FC-LCERTUPD", "EV-APPEAR", "EV-FILING", "EV-SHERIFF", "EV-WRIT", "EV-MOTIONS"
      GetFeeString = "Fees.FeeType = '" & cboCRType & "' and JurisdictionID = " & [Forms]![Case List]![JurisdictionID]
    Case "BK-MFR", "BK-OTH", "FC-OTH", "EV-OTH", "CIV-OTH", "DIL-OTH"
      GetFeeString = "Fees.FeeType = '" & cboCRType & "'"
End Select



End Function

Private Function GetDefendantCnt(strCaseType As String)

Dim FileNumber As Long

FileNumber = Forms![Case List]!FileNumber

Select Case strCaseType
    Case "FC"
      GetDefendantCnt = DCount("ID", "Names", "FileNumber = " & FileNumber & " and Mortgagor = true")
    Case "EV"
      GetDefendantCnt = DCount("ID", "Names", "FileNumber = " & FileNumber & " and (Owner = true or Tenant = true)")
'      GetDefendantCnt = GetDefendantCnt + 1   ' this will include All Occupants
    Case "BK"
      GetDefendantCnt = DCount("ID", "Names", "FileNumber = " & FileNumber & " and (BKDebtor = true or BKCoDebtor = true) and (owner = true or mortgagor = true)")
End Select



End Function


Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click
If IsNull(Me.cboCRType) Then
    MsgBox "Enter a check request type.", vbExclamation
    Exit Sub
End If

'If (txtAmt = 0) Then
'  MsgBox "Enter Fee amount.", vbExclamation
'  Exit Sub
'End If

If (Me.cboCRType.Column(1) = "Other") Then

  
  If (IsNull(Me.txtPayableTo)) Then
    MsgBox "Enter Payable To.", vbExclamation
    Exit Sub
  End If
  
  If (IsNull(Me.txtDescription)) Then
    MsgBox "Enter Description.", vbExclamation
    Exit Sub
  End If
End If

Dim s As Recordset
Dim amt As Currency
Dim Desc As String
Dim PayableTo As String
Dim State As String
Dim BulkCheck As Boolean
Dim defCount As Integer
Dim Location As String, FCType As String
Dim FileNumber As Long

FileNumber = Forms![Case List]!FileNumber

If optRequestType = 3 Then
    If Val(Nz(txtCheckNumber)) <= 0 Then
        MsgBox "Check number is required for a pre-cut check", vbCritical
        Exit Sub
    End If
End If

Set s = CurrentDb.OpenRecordset("select feetype.casetype, * from fees inner join feetype on fees.feetype = feetype.feetype where " & GetFeeString(), dbOpenDynaset, dbSeeChanges)
If Not s.EOF Then

   amt = txtAmt
   
   If (Nz(s![FeeDescription])) = "Other" Then
     Desc = Me.txtDescription
     PayableTo = Me.txtPayableTo
   Else
     Desc = Nz(s![FeeDescription])
     PayableTo = Nz(s![PayableTo])
   End If
   
   If Not IsNull(s![AmtPerDefendant]) Then  ' charge per defendant
     
     defCount = GetDefendantCnt(s![CaseType])
     
     If (defCount = 0) Then
       MsgBox "There are no defendants for this fee.  Amount may need to be updated."
     Else
     
        If Not IsNull(s![AmtPerAddlDefendant]) Then ' charge for each additional defendant
           amt = amt + (Nz(s![AmtPerDefendant])) 'first defendant
           amt = amt + ((defCount - 1) * Nz(s![AmtPerAddlDefendant]))  ' additional defendants
        Else
           amt = amt + (defCount * Nz(s![AmtPerDefendant]))  ' all defendants
        End If
     End If
     
   End If
   
   BulkCheck = Nz(s![BulkCheck])
   
   If Me.optLocation = 1 Then
        Location = "MD"

   ElseIf Me.optLocation = 2 Then
        Location = "VA"
   Else
        Location = ""
   End If
   
   
   If Me.optFCType = 1 Then
        FCType = "Pre Sale"
   ElseIf Me.optFCType = 2 Then
        FCType = "Post Sale"
   Else
       FCType = ""
   End If
   
   Call AddCheckRequest(FileNumber, amt, Desc, PayableTo, optRequestType, cboCRType, BulkCheck, Nz(txtCheckNumber), PreviouslyBilled, Location, FCType)
   
   'Call AddCheckRequest(FileNumber, amt, Desc, PayableTo, optRequestType, cboCRType, BulkCheck, Nz(txtCheckNumber), PreviouslyBilled)
   
   If (optRequestType = 1 Or optRequestType = 2) And Me.PreviouslyBilled = 0 Then ' credit card + check
        AddInvoiceItem FileNumber, Me!cboCRType, Desc, amt, 0, False, False, False, True

   End If
   
Else
  MsgBox "The Check Request details are not entered. The Check Request was not added.", vbExclamation
End If

DoCmd.Close

Forms![Case List]!sfrmCheckRequest.Requery

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
    
End Sub


Private Sub Form_Open(Cancel As Integer)
  UpdateCRList
  If Me.optDept = 1 Then Me.optFCType.Visible = True
End Sub

Private Sub optDept_AfterUpdate()
  UpdateCRList
  If Me.optDept = 1 Then
    Me.optFCType.Visible = True
  Else
    Me.optFCType.Visible = False
  End If
End Sub

Private Sub UpdateCRList()

Dim strCaseTypeCode As String

strCaseTypeCode = DLookup("[CaseCode]", "[CaseTypes]", "[CaseTypeID] = " & optDept)
cboCRType.RowSource = "SELECT FeeType.FeeType, FeeDescription " & _
                         "FROM FeeType " & _
                         "WHERE (FeeType.CheckRequest = True) AND (FeeType.CaseType = '" & strCaseTypeCode & "')" & _
                         "ORDER BY FeeDescription; "
cboCRType.Requery

End Sub

Private Function FetchComputedAmt(tfunction As String)

Dim SalePrice As Currency

Select Case tfunction
  Case "FetchVAAuditorFee"
    
    SalePrice = Nz(DLookup("[SalePrice]", "FCDetails", "FileNumber = " & Forms![Case List]!FileNumber & " and Current=true"), 0)
    FetchComputedAmt = CalculateVAAuditorFee(SalePrice)
  
End Select

End Function

Private Sub optRequestType_AfterUpdate()
txtCheckNumber.Visible = (optRequestType = 3)
lblCheckNumber.Visible = (optRequestType = 3)
End Sub
