VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EVContactsByClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub LockControls()
    txt_Search.SetFocus
    Me.FormHeader.BackColor = 8421631
    Me.Detail.BackColor = 8421631
    lbl_RO.Visible = True
    FirstName.Locked = True
    LastName.Locked = True
    EMail.Locked = True
    PhoneNumber.Locked = True
    Extension.Locked = True
    Address.Locked = True
    Address2.Locked = True
    City.Locked = True
    State.Locked = True
    ZipCode.Locked = True
    Me.btn_Cancel.Enabled = False
    Me.btn_Save.Enabled = False
    Me.btn_Edit.Enabled = True
    Me.btn_New.Enabled = True
    
End Sub

Private Sub UnlockControls()
    Me.FormHeader.BackColor = -2147483633
    Me.Detail.BackColor = -2147483633
    lbl_RO.Visible = False
    FirstName.Locked = False
    LastName.Locked = False
    EMail.Locked = False
    PhoneNumber.Locked = False
    Extension.Locked = False
    Address.Locked = False
    Address2.Locked = False
    City.Locked = False
    State.Locked = False
    ZipCode.Locked = False
    Me.btn_Cancel.Enabled = True
    Me.btn_Save.Enabled = True

End Sub

Private Sub btn_Cancel_Click()
        btn_New.Enabled = False
        Call LockControls
        If Me.Dirty Then
            DoCmd.RunCommand acCmdUndo
            Cancel = True
        Else
            Cancel = True
        End If
End Sub

Private Sub btn_Edit_Click()
    Call UnlockControls
    btn_New.Enabled = Not (btn_Edit.Enabled)
End Sub

Private Sub btn_New_Click()
    btn_Edit.Enabled = Not (btn_New.Enabled)
    Call UnlockControls
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub btn_Save_Click()
    DoCmd.RunCommand acCmdSaveRecord
    Call LockControls
    btn_New.Enabled = Not (btn_New.Enabled)
End Sub



Private Sub cmb_Filter_Change()

  ' If the combo box is cleared, clear the form filter.
  If Nz(Me.cmb_Filter.Text) = "" Then
    Me.Form.Filter = ""
    Me.FilterOn = False
    
  ' If a combo box item is selected, filter for an exact match.
  ' Use the ListIndex property to check if the value is an item in the list.
  ElseIf Me.cmb_Filter.ListIndex <> -1 Then
  
    Me.Form.Filter = "[ClientID] = " & Nz(Me.cmb_Filter.Column(0))
    Me.FilterOn = True

  End If
  
  Me.cmb_Filter.SetFocus
  Me.cmb_Filter.SelStart = Len(Me.cmb_Filter.Text)


End Sub

Private Sub txt_Search_Change()
  If Nz(Me.txt_Search.Text) = "" Then
    Me.Form.Filter = ""
    Me.FilterOn = False
    
  ' If a combo box item is selected, filter for an exact match.
  ' Use the ListIndex property to check if the value is an item in the list.
  Else
  
    Me.Form.Filter = "[LastName] LIKE '*" & Nz(Me.txt_Search.Text) & "*' OR [FirstName] LIKE '*" & Nz(Me.txt_Search.Text) & "*' OR [Email] LIKE '*" & Nz(Me.txt_Search.Text) & "*'"
    Me.FilterOn = True

  End If
    Me.txt_Search.SetFocus
    Me.txt_Search.SelStart = Len(Me.txt_Search.Text)
End Sub
