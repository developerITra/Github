VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetRestarApprovalFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
 Dim strinfo As String

On Error GoTo Err_cmdOK_Click

If txtTotal = 0 And txtflat = 0 Then
        MsgBox "you need to enter percentage number of fee or flat fee Amount", vbCritical
Exit Sub
End If

If txtTotal = 0 And (txtflat) <= 0 Then
    MsgBox "FeeAmount must be greater than zero", vbCritical
    Exit Sub
End If

If txtTotal = 0 And txtflat > 0 Then

AddInvoiceItem Forms![Case List].FileNumber, txtProcess, txtDesc, Format$(Me.txtflat, "Currency"), 0, True, True, False, False

DoCmd.SetWarnings False

strinfo = txtDesc
strinfo = Format$(Me.txtflat, "Currency") & " restart fee approved by client. " & strinfo

strinfo = Replace(strinfo, "'", "''")
strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms![Case List].FileNumber & ",Now,GetFullName(),'" & strinfo & "',2 )"
DoCmd.RunSQL strSQLJournal
strinfo = ""

DoCmd.SetWarnings True
Forms!Journal.Requery

MsgBox "Fee accepted"
DoCmd.Close acForm, "GetRestarApprovalFee"
'DoCmd.Close
Exit Sub
End If

If txtTotal > 0 Then
    Dim FeeAmount As Currency
    Dim Rate As Currency
    Dim strClient As Integer

    strClient = Nz(DLookup("ClientID", "Caselist", "FileNumber=" & Forms![Case List].FileNumber))
    
    Select Case Nz(DLookup("State", "FCDetails", "FileNumber=" & Forms![Case List].FileNumber & " AND Current=1"))
    
    Case "VA"
        Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & Forms![Case List].FileNumber & " AND Current=1"))
            Case 1 'Conventional
                FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & strClient))
            Case 2 'VA or Veteran's Affairs
                FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
            Case 3 'FHA or HUD
                FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=570")) 'HUD/FHA
            Case 4
                FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=177")) 'Fannie Mae
            Case 5
                FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=263")) 'Freddie Mac
            Case Else
                FeeAmount = Nz(DLookup("FeeVAReferral", "ClientList", "ClientID=" & strClient))
        End Select
            AddInvoiceItem Forms![Case List].FileNumber, txtProcess, txtDesc, Format$(FeeAmount * txtTotal / 100, "Currency"), 0, True, True, False, False
    
    
    Case "MD"
        Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & Forms![Case List].FileNumber & " AND Current=1"))
            Case 1 'Conventional
                FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & strClient))
            Case 2 'VA or Veteran's Affairs
                FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
            Case 3 'FHA or HUD
                FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=570")) 'HUD/FHA
            Case 4
                FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=177")) 'Fannie Mae
            Case 5
                FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=263")) 'Freddie Mac
            Case Else
                FeeAmount = Nz(DLookup("FeeMDReferral", "ClientList", "ClientID=" & strClient))
        End Select
            AddInvoiceItem Forms![Case List].FileNumber, txtProcess, txtDesc, Format$(FeeAmount * txtTotal / 100, "Currency"), 0, True, True, False, False
    
    
    Case "DC"
        Select Case Nz(DLookup("LoanType", "FCDetails", "FileNumber=" & Forms![Case List].FileNumber & " AND Current=1"))
            Case 1 'Conventional
                FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=" & strClient))
            Case 2 'VA or Veteran's Affairs
                FeeAmount = Nz(DLookup("FeeDcReferral", "ClientList", "ClientID=573")) 'Veteran's Affairs
            Case 3 'FHA or HUD
                FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=570")) 'HUD/FHA
            Case 4
                FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=177")) 'Fannie Mae
            Case 5
                FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=263")) 'Freddie Mac
            Case Else
                FeeAmount = Nz(DLookup("FeeDCReferral", "ClientList", "ClientID=" & strClient))
        End Select
            AddInvoiceItem Forms![Case List].FileNumber, txtProcess, txtDesc, Format$(FeeAmount * txtTotal / 100, "Currency"), 0, True, True, False, False

    End Select

    DoCmd.SetWarnings False

    strinfo = txtDesc
    strinfo = Format$(FeeAmount * txtTotal / 100, "Currency") & " restart fee approved by client. " & strinfo

    strinfo = Replace(strinfo, "'", "''")
    strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & Forms![Case List].FileNumber & ",Now,GetFullName(),'" & strinfo & "',2 )"
    DoCmd.RunSQL strSQLJournal
    strinfo = ""

    DoCmd.SetWarnings True
    Forms!Journal.Requery

End If
MsgBox "Fee accepted"

'DoCmd.Close
DoCmd.Close acForm, "GetRestarApprovalFee"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
txtProcess = Me.OpenArgs
End Sub
