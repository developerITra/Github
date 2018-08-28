VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Print News Ad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim PrintTo As Integer


Private Sub cmdCancel_Click()

DoCmd.Close
    
End Sub

Private Sub cmdOK_Click()
Dim statusMsg As String, rptType As String, Line1 As String, Line2 As String, Line3 As String, Line4 As String, Line5 As String, Line6 As String, FileName As String

If IsNull(Combo4) Then
MsgBox "Please select a newspaper"
Exit Sub
End If

''check for news ad, Word 03
'If Dir(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\Newspaper AD " & Format$(Now(), "yyyymmdd hhnn") & ".doc") <> "" Then
'FileName = DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\Newspaper AD " & Format$(Now(), "yyyymmdd hhnn") & ".doc"
''check for news ad, Word 07
'ElseIf Dir(DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\Newspaper AD " & Format$(Now(), "yyyymmdd hhnn") & ".docx") Then
'FileName = DocLocation & DocBucket(FileNumber) & "\" & FileNumber & "\Newspaper AD " & Format$(Now(), "yyyymmdd hhnn") & ".docx"
'Else
'MsgBox "No News Ad for this file is uploaded into Documents for today", vbCritical
'Exit Sub
'End If


If optDocType = 4 Then
MsgBox "Please enter the frequency the ads will run in the email"

'Dim lrs As Recordset
'  Set lrs = CurrentDb.OpenRecordset("journal", dbOpenDynaset, dbSeeChanges)
'  lrs.AddNew
'
'  lrs![FileNumber] = FileNumber
'  lrs![JournalDate] = Now
'  lrs![Who] = GetFullName()
'
'  lrs![Info] = "Frequency of Newspaper as is Other:  " & Other & vbCrLf
'  lrs![Color] = 1
'  lrs.Update
  
              DoCmd.SetWarnings False
            strinfo = "Frequency of Newspaper as is Other:  " & Other & vbCrLf
            strinfo = Replace(strinfo, "'", "''")
            strSQLJournal = "Insert into Journal (FileNumber,JournalDate,Who,Info,Color) Values(" & FileNumber & ",Now,GetFullName(),'" & strinfo & "',1 )"
            DoCmd.RunSQL strSQLJournal
            DoCmd.SetWarnings True

'lrs.Close

End If
Call DoReport("Ad VA", PrintTo)

Select Case optDocType
Case 1
Line1 = "Attached is a legal advertisement that needs to be run once a week for two weeks, not less than 8 days prior to 2nd ad, not more than 30 days from sale." & vbNewLine
Case 2
Line1 = "Attached is a legal advertisement that needs to be run every other day for 5 times." & vbNewLine
Case 3
Line1 = "Attached is a legal advertisement that needs to be run once a week for four weeks." & vbNewLine
Case 4
Line1 = "Attached is a legal advertisement that needs to be run _____." & vbNewLine
End Select

Line2 = "Run these advertisements on " & FirstPub & " and ___. " & vbNewLine

Line3 = "Please, fax or email a draft copy to my attention for review along with fees and cost prior to the date ad will run. " & vbNewLine
 
Line4 = "**IMPORTANT** Please attach a separate invoice indicating amounts owed for the ad and include invoice with the ad proof.** " & vbNewLine
       
Line5 = "If possible, also email (Cc) a draft copy to pubs@rosenberg-assoc.com. " & vbNewLine
 
Line6 = "(See attached file)  Please contact me if you have any questions or require additional information. " & vbNewLine
    
    
Dim olApp As Object
            Dim olMail As Object
            Set olApp = CreateObject("Outlook.Application")
            Set olMail = olApp.CreateItem(olMailItem)
            
            With olMail
            If Not IsNull(DLookup("email", "vendors", "id=" & Int(Combo4))) Then
                .To = DLookup("email", "vendors", "id=" & Int(Combo4))
            End If
                .Subject = "Foreclosure Pubs- " & Forms!foreclosuredetails!PropertyAddress & " " & Forms!foreclosuredetails!City
                '.Attachments.Add = FileName
                .Body = Line1 & vbNewLine & Line2 & vbNewLine & Line3 & vbNewLine & Line4 & vbNewLine & Line5 & vbNewLine & Line6
                .Display
            End With
    
cmdCancel.Caption = "Close"

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.Description
    Resume Exit_cmdOK_Click
    
End Sub


Private Sub Combo4_AfterUpdate()
Forms!foreclosuredetails!NewspaperVendor = Int(Combo4)
End Sub


Private Sub FirstPub_AfterUpdate()
If (Forms!foreclosuredetails!Sale - FirstPub) < 9 Or (Forms!foreclosuredetails!Sale - FirstPub) > 37 Then
FirstPub = Null
MsgBox "First Pub date must be between 9 and 37 days prior to the sale date", vbCritical
End If
Forms!foreclosuredetails!FirstPub = FirstPub
End Sub

Private Sub Form_Current()
PrintTo = Int(Me.OpenArgs)

End Sub


