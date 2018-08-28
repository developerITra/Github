VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SetTrusteeVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
   Dim TempID As Double

   
On Error GoTo Err_cmdOK_Click

    If lstTrustee.Column(2) = "CoCounsel" Then
            TempID = lstTrustee.Column(0)
            TempID = (TempID + 0.5)
           
         
            Forms!foreclosuredetails!SaleConductedTrusteeID = TempID
          
            
    Else
    
        Forms!foreclosuredetails!SaleConductedTrusteeID = lstTrustee.Column(0)
          
    End If




DoCmd.Close acForm, "SetTrusteeVA"


Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click

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



Private Sub lstDisposition_DblClick(Cancel As Integer)
Call cmdOK_Click
End Sub

Private Sub lstTrustee_Click()
' Dim TempID As Double
'
'
'On Error GoTo Err_cmdOK_Click
'
'    If lstTrustee.Column(2) = "CoCounsel" Then
'            TempID = lstTrustee.Column(0)
'            TempID = (TempID + 0.5)
'
'            Me.Form.Requery
'
'            Forms!ForeclosureDetails!SaleConductedTrusteeID = TempID
'            'SaleConductedTrusteeID = TempID
'
'    Else
'
'        Forms!ForeclosureDetails!SaleConductedTrusteeID = lstTrustee.Column(0)
'          '  SaleConductedTrusteeID = lstTrustee.Column(0)
'    End If
'
'
''Forms!ForeclosureDetails!txtTrusteeConductedSale.SetFocus
'
'DoCmd.Close acForm, "SetTrusteeVA"
''DoCmd.Close "SetTrusteeVA"
'
'Exit_cmdOK_Click:
'    Exit Sub
'
'Err_cmdOK_Click:
'    MsgBox Err.Description
'    Resume Exit_cmdOK_Click
End Sub
