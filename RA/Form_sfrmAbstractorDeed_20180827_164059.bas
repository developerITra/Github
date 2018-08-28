VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmAbstractorDeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Form_BeforeInsert(Cancel As Integer)
JurisdictionID = [Forms]![Jursidictions]![cbxSelect]
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
AllowDeletions = False
 AllowAdditions = False

End If

End Sub
