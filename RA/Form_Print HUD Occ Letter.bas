VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print HUD Occ Letter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close acForm, "Print HUD Occ Letter"

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub


Private Sub cmdPrint_Click()

Dim statusMsg As String, FeeAmount As Currency, JnlNote As String, sql As String, matter As String

On Error GoTo Err_cmdOK_Click

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

DoReport "HUD Occupancy Letter All Occ", acViewNormal
DoReport "HUD Occupancy Letter HH", acViewNormal
Forms!foreclosuredetails!HUDOccLetter = Now()
AddStatus [CaseList.FileNumber], Now(), "HUD Occupancy Letter sent"
AddInvoiceItem [CaseList.FileNumber], "FC-HUDOCC", "HUD Occ Letter Postage", 10.62, 76, False, False, False, True
Call StartDoc(TemplatePath & "\HUDOCC attachments.pdf")

    JnlNote = "HUD Occupancy Letter sent"
Dim lrs As Recordset, rstLabelData As Recordset

'            Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'            lrs.AddNew
'            lrs![FileNumber] = FileNumber
'            lrs![JournalDate] = Now
'            lrs![Who] = GetFullName()
'            lrs![Info] = JnlNote & vbCrLf
'            lrs![Color] = 1
'            lrs.Update
'            lrs.Close
            
            DoCmd.SetWarnings False
            strinfo = JnlNote & vbCrLf
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True

        sql = "SELECT fcdetails.PropertyAddress, fcdetails.City, fcdetails.State, fcdetails.ZipCode, CaseList.FileNumber, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (fcdetails INNER JOIN CaseList ON fcdetails.FileNumber = CaseList.FileNumber) INNER JOIN ClientList ON CaseList.ClientID = ClientList.ClientID WHERE (((CaseList.FileNumber)=" & FileNumber & "));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)                                                                                                         'FROM ClientList INNER JOIN qry45Days ON ClientList.ClientID = qry45Days.ClientID where qry45days.filenumber=" & Forms![wizNOI]!FileNumber
  
                Call StartLabel
                Print #6, FormatName("", "", "All Occupants", "", rstLabelData!PropertyAddress, "", rstLabelData!City, rstLabelData!State, rstLabelData!ZipCode)
                Print #6, "|FONTSIZE 8"
                Print #6, "|BOTTOM"
                Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
                matter = rstLabelData!PrimaryDefName
                Call FinishLabel
                Call StartLabel
                Print #6, FormatName("", "", "Head Of Household", "", rstLabelData!PropertyAddress, "", rstLabelData!City, rstLabelData!State, rstLabelData!ZipCode)
                Print #6, "|FONTSIZE 8"
                Print #6, "|BOTTOM"
                Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
                matter = rstLabelData!PrimaryDefName
                Call FinishLabel
 
                rstLabelData.MoveNext
                rstLabelData.Close

cmdCancel.Caption = "Close"

'Forms!foreclosuredetails.Requery

Exit_cmdOK_Click:
DoCmd.Close acForm, "Print HUD Occ Letter"
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    

End Sub

Private Sub cmdView_Click()

Dim statusMsg As String

On Error GoTo Err_cmdOK_Click

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

DoReport "HUD Occupancy Letter All Occ", acPreview
DoReport "HUD Occupancy Letter HH", acPreview
DoReport "HUD Occupancy Letter Notice", acViewNormal
cmdCancel.Caption = "Close"



Exit_cmdOK_Click:
DoCmd.Close acForm, "Print HUD Occ Letter"
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    

End Sub

Private Sub cmdAcrobat_Click()

Dim statusMsg As String

On Error GoTo Err_cmdOK_Click

If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
DoReport "HUD Occupancy Letter All Occ", -2
DoReport "HUD Occupancy Letter HH", -2
DoReport "HUD Occupancy Letter Notice", -2
'Call StartDoc(TemplatePath & "\HUDOCC attachments.pdf")
cmdCancel.Caption = "Close"

Forms!foreclosuredetails!cmdWizComplete.Enabled = True

Exit_cmdOK_Click:
DoCmd.Close acForm, "Print HUD Occ Letter"
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub
