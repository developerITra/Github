VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Jursidictions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Abstractor_AfterUpdate()
'update Case Abstractor when Jurisidiction abstractor is updated (only if client abstractor is not set)

  DoCmd.SetWarnings False
  DoCmd.RunSQL ("UPDATE (CaseList INNER JOIN JurisdictionList ON CaseList.JurisdictionID = JurisdictionList.JurisdictionID) INNER JOIN ClientList ON CaseList.ClientID = ClientList.ClientID SET CaseList.CaseAbstractor = " & Me.Abstractor & _
                " WHERE (ClientList.ClientAbstrator Is Null) and (CaseList.JurisdictionID = " & JurisdictionID & ")")
  DoCmd.SetWarnings True
End Sub

Private Sub Auditors_Click()
DoCmd.OpenForm "sfrmauditor", , , "jurisdictionid=" & JurisdictionID
Forms!sfrmAuditor!JurisdictionID = JurisdictionID
End Sub

Private Sub btnRecorder_Click()

On Error GoTo Err_btnRecorder_Click
DoCmd.OpenForm "Recorder Addresses"

Exit_btnRecorder_Click:
    Exit Sub

Err_btnRecorder_Click:
    MsgBox Err.Description
    Resume Exit_btnRecorder_Click

End Sub



Sub cbxSelect_AfterUpdate()
    ' Find the record that matches the control.
    Me.RecordsetClone.FindFirst "[JurisdictionID] = " & Me![cbxSelect]
    Me.Bookmark = Me.RecordsetClone.Bookmark
End Sub

Private Sub cmdAbstractorDeeds_Click()
DoCmd.OpenForm "sfrmAbstractorDeed", , , "jurisdictionid=" & JurisdictionID
Forms!sfrmAbstractorDeed!JurisdictionID = JurisdictionID
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

Private Sub cmdHUD_Click()

On Error GoTo Err_cmdHUD_Click
DoCmd.OpenForm "HUD Addresses"

Exit_cmdHUD_Click:
    Exit Sub

Err_cmdHUD_Click:
    MsgBox Err.Description
    Resume Exit_cmdHUD_Click
    
End Sub

Private Sub cmdVA_Click()

On Error GoTo Err_cmdVA_Click
DoCmd.OpenForm "VA Addresses"

Exit_cmdVA_Click:
    Exit Sub

Err_cmdVA_Click:
    MsgBox Err.Description
    Resume Exit_cmdVA_Click
    
End Sub

Private Sub cmdNew_Click()

On Error GoTo Err_cmdNew_Click
DoCmd.GoToRecord , , acNewRec

Exit_cmdNew_Click:
    Exit Sub

Err_cmdNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdNew_Click
    
End Sub

Private Sub cmdDelete_Click()

On Error GoTo Err_cmdDelete_Click
DoCmd.RunCommand acCmdDeleteRecord

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
    
End Sub

Private Sub cmdViewNewspapers_Click()
DoCmd.OpenForm "sfrmNewspapers", , , "jurisdictionid=" & JurisdictionID
Forms!sfrmNewspapers!JurisdictionID = JurisdictionID
End Sub



Private Sub Form_BeforeUpdate(Cancel As Integer)
If Not PrivJurisdic Then
    Cancel = 1
    Me.Undo
    MsgBox "You are not authorized to make changes", vbCritical
    Call cbxSelect_AfterUpdate
    'MsgBox "You are not authorized to make changes", vbCritical
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
If Not PrivJurisdic Then

    Dim ctl As Control
    Dim lngI As Long
    Dim bSkip As Boolean

    For Each ctl In Form.Controls
    Select Case ctl.ControlType
    Case acTextBox, acComboBox, acListBox, acOptionGroup, acCheckBox, acOptionButton, acToggleButton, acSubform

            bSkip = False
            If ctl.Name = "cbxSelect" Then
                    bSkip = True
                   
            End If
           
            If Not bSkip Then
            ctl.Locked = True
            End If

    End Select
    Next
End If




cmdIRSAddress.Enabled = PrivJurisdic
cmdVA.Enabled = PrivJurisdic
cmdHUD.Enabled = PrivJurisdic
cmdNew.Enabled = PrivJurisdic
'Me.btnRecorder.Enabled = PrivJurisdic   'Removed

Me.sfrmDistrictCourt.Enabled = PrivJurisdic
Me.sfrmCircuitCourt.Enabled = PrivJurisdic
Me.sfrmRecorders.Enabled = PrivJurisdic
Me.sfrmLienors.Enabled = PrivJurisdic
Me.sfrmHUDAddress.Enabled = PrivJurisdic
Me.sfrmIRSAddress.Enabled = PrivJurisdic
Me.sfrmVAAddress.Enabled = PrivJurisdic

'Me.AllowAdditions = PrivAdmin
'Me.AllowDeletions = PrivAdmin
'cmdDelete.Enabled = PrivAdmin

End Sub

Private Sub State_Change()
If Nz(LongState) = "" Then
    Select Case State.Text
        Case "DC"
            LongState = "District of Columbia"
        Case "MD"
            LongState = "Maryland"
        Case "VA"
            LongState = "Virginia"
    End Select
End If
End Sub

Private Sub cmdIRSAddress_Click()

On Error GoTo Err_cmdIRSAddress_Click
DoCmd.OpenForm "IRS Addresses"

Exit_cmdIRSAddress_Click:
    Exit Sub

Err_cmdIRSAddress_Click:
    MsgBox Err.Description
    Resume Exit_cmdIRSAddress_Click
    
End Sub

Private Sub ViewAllVendors_Click()
DoCmd.OpenForm ("VendorsInfo_1")
End Sub
