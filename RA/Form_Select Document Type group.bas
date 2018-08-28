VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Select Document Type group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub lstDocType_DblClick(Cancel As Integer)
Call cmdOK_Click
End Sub

Private Sub cmdOK_Click()

On Error GoTo Err_cmdOK_Click

 Dim costamt As Currency


 If IsNull(lstDocType) Then
    MsgBox "Select a document type, or click one of the quick-pick buttons.", vbCritical
    Exit Sub
End If

selecteddoctype = lstDocType

'8/26/2014 - Stoped by Sarab on 4/25 and moved to DocMissing moduel

'If selecteddoctype = 1371 Or selecteddoctype = 1450 Then
'
'    If MsgBox("Is the Original Note included ?", vbQuestion + vbYesNo) = vbYes Then
'
'        Dim rs As Recordset
'        Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCDetails WHERE FileNumber = " & fileno & " AND Current = True", dbOpenDynaset, dbSeeChanges)
'
'            If rs!DocBackOrigNote = False Then
'
'                rs.Edit
'                rs!DocBackOrigNote = True
'                rs.Update
'       'Add to Status line
'                AddStatus fileno, Date, "Received Original Note"
'        'added to Journal
'                DoCmd.SetWarnings False
'
'                strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & fileno & ",Now,GetFullName(),'" & "Original Note is received " & "',1 )"
'                DoCmd.RunSQL strSQLJournal
'                strSQLJournal = ""
'                Forms!Journal.Requery
'
'                DoCmd.SetWarnings True
'            End If
'    Set rs = Nothing
'    'rs.Close
'    Else
'                AddStatus fileno, Date, "Removed Original Note"
'
'    End If
'
'End If

'added Skip trace cost on 2/3/15
If selecteddoctype = 1525 Then
  DoCmd.OpenForm "GetSkipTraceCost", , , , , acDialog, "FC-SKP"
  
    'costamt = InputBox("Please enter cost, then rememeber to note the journal")
    'AddInvoiceItem Forms![Case List]!FileNumber, "FC-SKP", "Skip Trace - SSN", costamt, 0, False, True, False, False
              
End If

'added client postage actual cost
If selecteddoctype = 887 Then

  If IsLoadedF("Case List") = True And IsLoadedF("wizFairDebt") = False Then

    costamt = InputBox("Please enter Fair Debt postage cost")
    AddInvoiceItem Forms![Case List]!FileNumber, "FC-FairDebt", "Fair Debt postage", costamt, 76, False, True, False, False

  End If

End If

'added 2/24/15

If selecteddoctype = 124 Then
        
  If IsLoadedF("Case List") = True And IsLoadedF("WizDemand") = False Then
    
    costamt = InputBox("Please enter Demand postage")
    AddInvoiceItem Forms![Case List]!FileNumber, "FC-Demand", "Demand postage", costamt, 76, False, True, False, False

  End If
  
End If

'added on 3/6/15, Scra check not from Scra Queue
If selecteddoctype = 1516 Then
    If IsLoadedF("Case List") = True And IsLoadedF("queSCRAFCNew") = False And IsLoadedF("wizRestartFCdetails1") = False And IsLoadedF("wizIntake1") = False And IsLoadedF("wizIntakeRestart") = False And IsLoadedF("wizReferralII") = False And IsLoadedF("wizRestartCaseList1") = False Then
        If Forms![Case List]![ClientID] = 97 Then
            Dim cost As Currency
            Dim i As Integer
            Dim rst As Recordset
            Dim FileNum As Long
            
            FileNum = Forms![Case List].FileNumber
            Set rst = CurrentDb.OpenRecordset("SELECT FileNumber,SSN FROM Names where ((FileNumber =" & Forms![Case List].FileNumber & ") AND (SSN Is Not Null)) OR (((FileNumber)=" & Forms![Case List].FileNumber & ") AND (SSN <>""999999999""))", dbOpenDynaset, dbSeeChanges)
            
            If Not rst.EOF Then
                rst.MoveLast
                i = rst.RecordCount
            Else
                i = 0
            End If
            
            cost = i * DLookup("ivalue", "db", "ID=" & 32)

            AddInvoiceItem FileNum, "FC-DOD", "DOD Search - From File Search", cost, 0, True, True, False, False
            
            rst.Close
            Set rst = Nothing
        End If
 
    End If
 
End If

DoCmd.Close

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
selecteddoctype = 0
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
If IsLoaded("Foreclosuredetails") Then
    If Forms!foreclosuredetails!WizardSource = "Intake" Then
    Me.lstDocType.RowSource = "SELECT DocumentTitles.ID, DocumentTitles.Title FROM DocumentTitles WHERE (((DocumentTitles.ID)=1579) AND ((DocumentTitles.Status)=1)) ORDER BY DocumentTitles.Title;"
    End If
End If

selecteddoctype = 0

End Sub

Private Sub optQuickPick_AfterUpdate()
selecteddoctype = optQuickPick
DoCmd.Close
End Sub
