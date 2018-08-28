VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmFNMAMissingDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private Sub FNMAMissingDocID_BeforeUpdate(Cancel As Integer)
  If (Not IsNull(FNMAMissingDocID)) Then
    Me.FNMAMissingDocDate = Date
  Else
    Me.FNMAMissingDocDate = Null
  End If
End Sub

Private Sub Form_Current()
  If IsNull(FileNumber) Then FileNumber = [Forms]![fcdetails]![FileNumber]
End Sub


