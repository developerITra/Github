VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QueueAccounting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub CmdAcconPSAdvanced_Click()
Dim LDay As Integer
LDay = Day(Date)
    If LDay = 18 Then
        If Not IsNull(DLookup("ID", "Accou_PSAdvancedInsertHoldDate", "Insertdate = date()")) Then
            DoCmd.OpenForm "QueueAccounPSAdvancedCosts"
        Else
            DoCmd.SetWarnings False
            
            Dim strInsert As String
            strInsert = "Insert into Accou_PSAdvancedInsertHoldDate (InsertDate,InsertBy) Values( #" & Date & "#,GetFullName())"
            DoCmd.RunSQL strInsert
            
            
            strInsert = "UPDATE Accou_PSAdvancedCostsPackageQueue SET " & " DateInsert = #" & Date & "# WHERE Hold = " & "'H'" & " AND MangerQ = False AND IsNull(DateInsert)"
                DoCmd.RunSQL strInsert
            strInsert = "UPDATE Accou_PSAdvancedCostsPackageQueue SET " & " Dismissed = False WHERE Hold = " & "'H'"
                DoCmd.RunSQL strInsert
                strInsert = ""
            
            
            
            DoCmd.SetWarnings True
            DoCmd.OpenForm "QueueAccounPSAdvancedCosts"
        End If
    Else
        DoCmd.OpenForm "QueueAccounPSAdvancedCosts"
    End If



End Sub

Private Sub CmdAccouLitigBill_Click()
Dim LDay As Integer
LDay = Day(Date)
    If LDay = 10 Then
        If Not IsNull(DLookup("ID", "Accou_LitigationInsertHoldDate", "Insertdate = date()")) Then
            DoCmd.OpenForm "QueueAccounLitigationBill"
        Else
            DoCmd.SetWarnings False
            
            Dim strInsert As String
            strInsert = "Insert into Accou_LitigationInsertHoldDate (InsertDate,InsertBy) Values( #" & Date & "#,GetFullName())"
            DoCmd.RunSQL strInsert
            
            
            strInsert = "UPDATE Accou_LitigationBillingQueue SET " & " DateInsert = #" & Date & "# WHERE Hold = " & "'H'" & " AND MangerQ = False AND IsNull(DateInsert)"
                DoCmd.RunSQL strInsert
            strInsert = "UPDATE Accou_LitigationBillingQueue SET " & " Dismissed = False WHERE Hold = " & "'H'"
                DoCmd.RunSQL strInsert
                strInsert = ""
            
            
            
            DoCmd.SetWarnings True
            DoCmd.OpenForm "QueueAccounLitigationBill"
        End If
    Else
        DoCmd.OpenForm "QueueAccounLitigationBill"
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

Private Sub cmdSCRA1_Click()

    Dim stDocName As String
  
    stDocName = "queSCRA1"
    DoCmd.OpenForm stDocName

End Sub



Private Sub cmdSCRA2_Click()
Dim stDocName As String

    stDocName = "queSCRA2"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA3_Click()
Dim stDocName As String

    stDocName = "queSCRA3"
    DoCmd.OpenForm stDocName
End Sub


Private Sub cmdSCRA4a_Click()
Dim stDocName As String

    stDocName = "queSCRA4a"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA4b_Click()
Dim stDocName As String

    stDocName = "queSCRA4b"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA5_Click()
Dim stDocName As String

    stDocName = "queSCRA5"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA6_Click()
Dim stDocName As String

    stDocName = "queSCRA6"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA7_Click()
Dim stDocName As String

    stDocName = "queSCRA7"
    DoCmd.OpenForm stDocName
End Sub
Private Sub cmdSCRA8_Click()
Dim stDocName As String

    stDocName = "queSCRA8"
    DoCmd.OpenForm stDocName
End Sub

Private Sub cmdSCRA9_Click()
Dim stDocName As String

    stDocName = "queSCRA9"
    DoCmd.OpenForm stDocName
End Sub

Private Sub cmdSCRA9Waiting_Click()
Dim stDocName As String

    stDocName = "queSCRA9Waiting"
    DoCmd.OpenForm stDocName
End Sub


Private Sub cmdSCRAUnionNew_Click()
 stDocName = "queSCRAFCNew"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Command75_Click()
stDocName = "queSCRABK"
    DoCmd.OpenForm stDocName
End Sub

Private Sub CmdESCQueue_Click()
'SortTable = "QueueESCSort"
Dim LDay As Integer
LDay = Day(Date)
    If LDay = 18 Then
        If Not IsNull(DLookup("ID", "Accou_ESCInsertHoldDate", "Insertdate = date()")) Then
            DoCmd.OpenForm "QueueAccounESC"
        Else
            DoCmd.SetWarnings False
            
            Dim strInsert As String
            strInsert = "Insert into Accou_ESCInsertHoldDate (InsertDate,InsertBy) Values( #" & Date & "#,GetFullName())"
            DoCmd.RunSQL strInsert
            
            
            strInsert = "UPDATE Accou_EscQueue SET " & " DateInsert = #" & Date & "# WHERE Hold = " & "'H'" & " AND MangerQ = False AND IsNull(DateInsert)"
                DoCmd.RunSQL strInsert
            strInsert = "UPDATE Accou_EscQueue SET " & " Dismissed = False WHERE Hold = " & "'H'"
                DoCmd.RunSQL strInsert
                strInsert = ""
            
            
            
            DoCmd.SetWarnings True
            DoCmd.OpenForm "QueueAccounESC"
        End If
    Else
        DoCmd.OpenForm "QueueAccounESC"
    End If

End Sub

Private Sub Command79_Click()
DoCmd.OpenForm "QueueAccountLitigationBillManager"

End Sub

Private Sub ComMgESc_Click()
DoCmd.OpenForm "QueueESCtManager"

End Sub

Private Sub ComMgPS_Click()
DoCmd.OpenForm "QueueAccountPSAdvancedCostManager"

End Sub
