VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print Deed Recording Cover Letter MD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cbxAbstractorTarget_AfterUpdate()

Me.cbxAbstractorFees = DLookup("RecorderPrice", "Recorders", "RecorderID= Forms![Print Deed Recording Cover Letter MD]!cbxAbstractorTarget")
'DoCmd.SetWarnings False
'DoCmd.RunSQL "UPDATE PrintInfoMD SET recorder =" & Forms![Print Deed Recording Cover Letter MD]!cbxAbstractorTarget & " WHERE FileNumber = " & Forms![Case List]!FileNumber & ";"
'DoEvents
DoCmd.SetWarnings True

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
End Sub


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

Me.Form.Requery

'Pick which Document based on Jurisdiction
If Me.OpenArgs = -1 Then 'All word Versions
    Call DoReport("Deed Recording Cover MD", Me.OpenArgs)
Else
    Select Case Forms![Case List]!JurisdictionID
    Case 13 ' Garrett County
        Call DoReport("Deed Recording Cover Garrett", Me.OpenArgs)
    Case 2 'Allegany County
        Call DoReport("Deed Recording Cover Allegany", Me.OpenArgs)
    Case 21 'Somerset County
        Call DoReport("Deed Recording Cover Somerset", Me.OpenArgs)
    Case 20 'St. Mary's
        Call DoReport("Deed Recording Cover StMary", Me.OpenArgs)
    Case 18 'PG COunty
        Call DoReport("Deed Recording Cover PG", Me.OpenArgs)
    Case 4 'Baltimore City
        Call DoReport("Deed Recording Cover Baltci", Me.OpenArgs)
    Case 5 'Baltimore County
        Call DoReport("Deed Recording Cover BaltCount", Me.OpenArgs)
    Case 3 'Anne Arundel
        Call DoReport("Deed Recording Cover Anne", Me.OpenArgs)
    Case 17
        Call DoReport("Deed Recording Cover MoCoMD", Me.OpenArgs)
    Case 10 'Charles County, MD
        Call DoReport("Deed Recording Cover Charles", Me.OpenArgs)
    Case 12 'Frederick MD
        Call DoReport("Deed Recording Cover Frederick", Me.OpenArgs)
    Case 23 'Washington County
        Call DoReport("Deed Recording Cover WashCO", Me.OpenArgs)
    Case 24 'Wicomico County
        Call DoReport("Deed Recording Cover Wicomico", Me.OpenArgs)
    Case 25 'Worcester County
        Call DoReport("Deed Recording Cover Worcester", Me.OpenArgs)
    Case 19 'Queen Anne's COunty
        Call DoReport("Deed Recording Cover QueenAnne", Me.OpenArgs)
    Case 14 'Harford County
        Call DoReport("Deed Recording Cover Harford", Me.OpenArgs)
    Case 8 ' Carroll County
        Call DoReport("Deed Recording Cover Carroll", Me.OpenArgs)
    Case 7 ' Caroline County
        Call DoReport("Deed Recording Cover Caroline", Me.OpenArgs)
    Case 6 ' Calvert County
        Call DoReport("Deed Recording Cover Calvert", Me.OpenArgs)
    Case 22 'Talbot County
        Call DoReport("Deed Recording Cover Talbot", Me.OpenArgs)
    Case 16 'Kent County
        Call DoReport("Deed Recording Cover Kent", Me.OpenArgs)
    Case 11 'Dorchester County
        Call DoReport("Deed Recording Cover Dorchester", Me.OpenArgs)
    Case 9   'Cecil
        Call DoReport("Deed Recording Cover Cecil", Me.OpenArgs)
    Case 15 'Howard
        Call DoReport("Deed Recording Cover Howard", Me.OpenArgs)
    Case Else
        Call DoReport("Deed Recording Cover MD", Me.OpenArgs)
        
    End Select
    
End If
'DoCmd.Close acForm, Me.Name
End Sub


Private Sub Form_Open(Cancel As Integer)
Dim i As Integer
If IsNull(Me.cbxAbstractorTarget) Then
    i = DLookup("RecorderID", "Recorders", "JurisdictionID = FOrms![Case List]!JurisdictionID")
    Me.cbxAbstractorTarget.SetFocus
    Me.cbxAbstractorTarget.Value = i
    Me.cbxAbstractorFees = DLookup("RecorderPrice", "Recorders", "RecorderID= Forms![Print Deed Recording Cover Letter MD]!cbxAbstractorTarget")

Else
    Me.cbxAbstractorTarget.SetFocus
    Me.cbxAbstractorFees = DLookup("RecorderPrice", "Recorders", "RecorderID= Forms![Print Deed Recording Cover Letter MD]!cbxAbstractorTarget")
End If
Me.cboMail = "US Postal"


Me.cbxAbstractorFees.Enabled = True
Me.cbxAbstractorTarget.Enabled = True
Me.cboMail.Enabled = True
Me.lblCountyTransferTax.Caption = "County Transfer Tax"

'Unlock Fields based on Jurisdictions
Select Case Forms![Case List]!JurisdictionID
Case 18 'PG County
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cboPropertyTax.Enabled = True
   
Case 4 'Baltimore City
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxWater.Enabled = True
    Me.cbxLien.Enabled = True
    'Me.Label82.Caption = "City Transfer Tax"
    
Case 5 'Baltimore County
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cboPropertyTax.Enabled = True
    'Me.cbxWater.Enabled = True
    Me.cbxLien.Enabled = True
    
Case 3 ' Anne Arundel
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxWater.Enabled = True
    Me.cbxWaterTarget.Enabled = False
Case 17
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cboPropertyTax.Enabled = True
    'Me.cbxWater.Enabled = True
    
Case 10
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxWater.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cboCertificate.Enabled = True
    Me.cboCertificate = "$20.00"
Case 12 'Frederick
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxWater.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cboTaxStatus.Enabled = True
    Me.cbxWaterTarget.Enabled = True
    Me.TaxStatus = "$20.00"
    Me.cbxWaterTarget.RowSource = "Frederick County DUSWM;City of Frederick"

Case 23 ' Washington County
    Me.RecordingFee.Enabled = True
    'Me.cbxWater.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cboAgriculture.Enabled = True
Case 24 'Wicomico County
    Me.RecordingFee.Enabled = True
    Me.cbxWater.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    'Me.cbxCountyTransfer.Enabled = True
    Me.cbxWaterTarget.Enabled = True
    Me.cbxAbstractorTarget = 52

Case 25 'Worcester County
    Me.RecordingFee.Enabled = True
    Me.cbxWater.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cbxWaterTarget.Enabled = True
    Me.cbxAbstractorTarget = 52
Case 19 'Queen Anne's
    Me.RecordingFee.Enabled = True
    Me.cbxWater.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
Case 14 'Harford
    Me.RecordingFee.Enabled = True
    Me.cbxWater.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cbxLien.Enabled = True
Case 8 'Carroll
    Me.RecordingFee.Enabled = True
    Me.cbxWater.Enabled = True
   ' Me.cboSewer.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
   ' Me.cbxCountyTransfer.Enabled = True
    Me.cbxLien.Enabled = True
Case 7 'Caroline
    Me.RecordingFee.Enabled = True
    Me.cbxWater.Enabled = True
    'Me.cboSewer.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cbxTownTax.Enabled = True
    Me.cbxWaterTarget.Enabled = True
    'Me.cbxPropertyTaxDestination.Enabled = True
    Me.cbxTownTaxTarget.Enabled = True
Case 6 'Calvert
    Me.RecordingFee.Enabled = True
    Me.cbxWater.Enabled = True

    Me.DeedCoverStateTax.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
 
Case 22 'Talbot
    Me.RecordingFee.Enabled = True
    'Me.cbxWater.Enabled = True
    'Me.cboSewer.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    'Me.cbxTownTax.Enabled = True
    Me.cbxAbstractorTarget = 52
Case 16 ' Kent
     Me.RecordingFee.Enabled = True
    Me.cbxWater.Enabled = True
    'Me.cboSewer.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    'Me.cbxTownTax.Enabled = True
Case 20 ' St Mary's
    Me.RecordingFee.Enabled = True
    Me.cbxWater.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
Case 21
    Me.RecordingFee.Enabled = True
    Me.cboCityTax.Enabled = True
    Me.cbxWater.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxCityTaxTarget.Enabled = True
    Me.cbxWaterTarget.Enabled = True
    Me.cbxAbstractorTarget = 52
Case 2 'Allegany
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxWater.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxWaterTarget.Enabled = True
   ' Me.cboCertificate.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cbxLien.Enabled = True
Case 11 'Dorchester
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxWater.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    Me.cbxWaterTarget.Enabled = True
    Me.cbxAbstractorTarget = 50
Case 13 'Garrett
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
  '  Me.cbxWater.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
Case 9 'Cecil
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.lblCountyTransferTax.Caption = "County Deed Fee"
    Me.cbxWater.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
Case 15 'Howard
    Me.RecordingFee.Enabled = True
    Me.DeedCoverStateTax.Enabled = True
    Me.cbxWater.Enabled = True
    Me.cboPropertyTax.Enabled = True
    Me.cbxCountyTransfer.Enabled = True
    'Me.cboCityTax.Enabled = True
    Me.cbxStateTransferTax.Enabled = True
    Me.cbxLien.Enabled = True
Case Else
    MsgBox (" No programmed Deed Cover Sheet for this County ")
End Select


End Sub

