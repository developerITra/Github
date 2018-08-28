VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Lost Note Affidavit Wells VA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Page()
If [FCdetails.State] = "VA" Then
Call FirmMarginVA(Me, FileNumber, , PrimaryDefName, PropertyAddress)
Else
Call FirmMarginVA(Me, FileNumber, , PrimaryDefName, PropertyAddress)
End If

'If Pages > 1 Then
'If page = 2 Then
'    If FunYesNo("Copy Available") = "vbYes" Then
'    Line6text = "6. A Copy of S"
'    Else
'    Line6text = " 6. If Wells Farg"
'    End If
'End If
'Else
'If FunYesNo("Copy Available") = "vbYes" Then
'    Line6text = "6. A Copy of S"
'
'    Else
'    Line6text = " 6. If Wells Farg"
'    End If
'End If
'


End Sub

