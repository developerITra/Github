VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CivilPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub chMotion_Click()
'opt1.Enabled = chMotion
   'If InPossession = 4 Then
       'chNoticeToTenant = 1
   'End If
    
End Sub


Private Sub cmdAcrobat_Click()
    Call PrintDocs(-2)
End Sub

Private Sub cmdClear_Click()
On Error GoTo Err_cmdClear_Click

chCreateLabel = 0

Exit_cmdClear_Click:
    Exit Sub

Err_cmdClear_Click:
    MsgBox Err.Description
    Resume Exit_cmdClear_Click
    
End Sub

Private Sub PrintDocs(PrintTo As Integer)
Dim ReportName As String

On Error GoTo Err_PrintDocs

If chCreateLabel = -1 Then
    DoCmd.OpenForm "Getlabel"
Else
End If

Exit_Err_PrintDocs:
    Exit Sub

Err_PrintDocs:
    MsgBox Err.Description
    Exit Sub

End Sub

Private Sub Form_Current()
'If IsNull(Attorney) Then
'    MsgBox "You must select an attorney before you can print.", vbCritical
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
'    cmdWord.Enabled = False
'End If
Me.Caption = "Print Civil " & [CaseList.FileNumber] & " " & [PrimaryDefName]
Me.cboAttorney = LeadAttorney

If ([Forms]![Case List]!Active = False) Then
    chCreateLabel.Enabled = False
    Label209.Visible = True
End If

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

Private Sub cmdWord_Click()
Call PrintDocs(-1)
End Sub

Private Sub cmdPrint_Click()
Call PrintDocs(acViewNormal)
End Sub

Private Sub cmdView_Click()
Call PrintDocs(acPreview)
End Sub

Private Sub Form_Open(Cancel As Integer)
If Me.State = "VA" Then

cboAttorney.RowSource = "SELECT Staff.ID, Staff.Name & ' ' " & _
                       "FROM Staff " & _
                       "WHERE  ((staff.active = True) And (Staff.Attorney =True) And(staff.PracticeVA = True )) " & _
                       "ORDER BY  Staff.Sort;"

ElseIf Me.State = "MD" Then
    cboAttorney.RowSource = "SELECT Staff.ID, Staff.Name & ', Esq.' FROM Staff WHERE ((Staff.active = true ) and (Staff.Attorney = True) and (Staff.PracticeMD = True)) ORDER BY Staff.Sort;"
Else
    cboAttorney.RowSource = "SELECT Staff.ID, Staff.Name FROM Staff WHERE ((staff.active = true) and (Staff.Attorney = True) and (staff.PracticeDC = true ))ORDER BY Staff.Sort;"
End If

Call cmdClear_Click
End Sub

