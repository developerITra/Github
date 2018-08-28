VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmCertOfPubFiled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CertOfPubField_AfterUpdate()
AddStatus FileNumber, CertOfPubField, "Cert Of Pub uploaded "
End Sub

Private Sub Form_Current()
If IsLoaded("Case List") = True Then
    If Forms![Case List]!CaseTypeID = 8 Then
    Call SetObjectAttributes(CertOfPubField, False)
'    Me.CertOfPubField.Enabled = False
'    Me.CertOfPubField.Locked = True
    End If
End If

If DCTabView = False Then
    CertOfPubField.Enabled = False
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
'If DCTabView = False Then
'    CertOfPubField.Enabled = False
'End If
End Sub
