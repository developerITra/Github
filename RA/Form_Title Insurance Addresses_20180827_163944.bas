VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Title Insurance Addresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
Me.AllowAdditions = PrivAdmin
Me.AllowDeletions = PrivAdmin
Me.AllowEdits = PrivAdmin
End Sub
