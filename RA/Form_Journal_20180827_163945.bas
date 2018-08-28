VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Journal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()

If FileReadOnly Then
    cmdNewJournalEntry.Enabled = False
    Detail.BackColor = ReadOnlyColor
Else
    cmdNewJournalEntry.Enabled = True
    Detail.BackColor = -2147483633
End If

Call ViewJournal
End Sub

Public Sub ViewJournal()
Dim J As Recordset, TmpFile As String, BGColor As String, WhereClause As String

On Error GoTo Err_cmdViewJournal_Click

TmpFile = Environ("TEMP") & "\RA_Journal.html"
Open TmpFile For Output As #1
Print #1, "<html><head><style type=""text/css""><!--"
Print #1, "body,td,th {font-family: Verdana, Arial, Helvetica, sans-serif; color: #000000;}"
'Font Black
Print #1, ".style1 {color: #000000; font-size: 11px}"
'Font Red
Print #1, ".style2 {color: #" & DLookup("SValue", "DB", "Name='AccJournalColor'") & "; font-weight: bold; font-size: 11px}"
'Font White
Print #1, ".style3 {color: #ffffff; font-size: 11px}"
'Font Blue
Print #1, ".style4 {color: #0000FF; font-size: 11px}"
'Font Green
Print #1, ".style5 {color: #00FF00; font-size: 11px}"
'Font Purple
Print #1, ".style6 {color: #FF00FF; font-size: 11px}"
'Font Orange
Print #1, ".style6 {color: #FF9900; font-size: 11px}"
Print #1, "--></style></head><body><table width=""100%"" border=0>"

If CaseTypeID = 5 Then      ' Civil
    Print #1, "<tr><td><img src=""file://FileServer/Applications/Database/caution.bmp""></td>";
    Print #1, "<td bgcolor=""#ff0000"" class=""style3"" align=""center""><strong>ATTORNEY-CLIENT PRIVILEGED COMMUNICATION</strong><br>Do not discuss without an attorney present!";
    Print #1, "<br><br></td></tr>"
End If

WhereClause = "WHERE FileNumber = " & FileNumber
If optView = 2 Then WhereClause = WhereClause & " AND (Color=2 OR Warning>0)"

Set J = CurrentDb.OpenRecordset("SELECT * FROM Journal " & WhereClause & " ORDER BY JournalDate DESC, ID DESC", dbOpenSnapshot)
Do While Not J.EOF
    Select Case Nz(J!Warning)
        Case 0
            Print #1, "<tr><td>&nbsp;</td>";
            BGColor = "#ffffff"
        Case 50   ' Waiting for Bill
            Print #1, "<tr><td><img src=""file://FileServer/Applications/Database/dollar.jpg""></td>";
            BGColor = "#aaffaa"
        Case 100  ' Waiting for Document
            Print #1, "<tr><td><img src=""file://FileServer/Applications/Database/papertray.jpg""></td>";
            BGColor = "#bbeeff"
        Case 200  ' Title
            Print #1, "<tr><td><img src=""file://FileServer/Applications/Database/house.jpg""></td>";
            BGColor = "#ffaaff"
        Case 300  ' Caution
            Print #1, "<tr><td><img src=""file://FileServer/Applications/Database/caution.bmp""></td>";
            BGColor = "#ffffaa"
        Case 400  ' Stop
            Print #1, "<tr><td><img src=""file://FileServer/Applications/Database/stop.jpg""></td>";
            BGColor = "#ffdddd"
    End Select
    If Nz(J!Color) = 2 Then
        Print #1, "<td bgcolor=""" & BGColor & """ class=""style2"">" & Format$(J("JournalDate"), "m/d/yyyy h:mm am/pm") & "&nbsp;&nbsp;&nbsp;" & J("Who") & "<br>" & J("Info");
    ElseIf Nz(J!Color) = 4 Then
        Print #1, "<td bgcolor=""" & BGColor & """ class=""style4"">" & Format$(J("JournalDate"), "m/d/yyyy h:mm am/pm") & "&nbsp;&nbsp;&nbsp;" & J("Who") & "<br>" & J("Info");
    Else
        Print #1, "<td bgcolor=""" & BGColor & """ class=""style1"">" & Format$(J("JournalDate"), "m/d/yyyy h:mm am/pm") & "&nbsp;&nbsp;&nbsp;" & J("Who") & "<br>" & J("Info");
    End If
    Print #1, "<br><br></td></tr>"
    J.MoveNext
Loop
J.Close

Print #1, "</body></html>"
Close #1
DoEvents
webJournal.Navigate2 TmpFile

Exit_cmdViewJournal_Click:
    Close #1
    Exit Sub

Err_cmdViewJournal_Click:
    MsgBox Err.Description
    Resume Exit_cmdViewJournal_Click
    
End Sub

Public Sub cmdNewJournalEntry_Click()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else

    On Error GoTo Err_cmdNewJournalEntry_Click
    DoCmd.OpenForm "Journal New Entry", , , , , , FileNumber
    Call ViewJournal
    
Exit_cmdNewJournalEntry_Click:
        Exit Sub
    
Err_cmdNewJournalEntry_Click:
        MsgBox Err.Description
        Resume Exit_cmdNewJournalEntry_Click
End If

    
End Sub

Private Sub cmdPrintJournal_Click()

On Error GoTo Err_cmdPrintJournal_Click
DoCmd.OpenReport "Journal", , , "FileNumber=" & FileNumber

Exit_cmdPrintJournal_Click:
    Exit Sub

Err_cmdPrintJournal_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintJournal_Click
    
End Sub



Private Sub Form_Open(Cancel As Integer)
Dim l As Long, t As Long, w As Long, h As Long

l = MaximumValue(0, CLng(GetSetting("RosenbergDatabase", "Forms", Me.Name & ".Left", 8340)))
t = MaximumValue(0, CLng(GetSetting("RosenbergDatabase", "Forms", Me.Name & ".Top", 540)))
w = GetSetting("RosenbergDatabase", "Forms", Me.Name & ".Width", 5220)
h = GetSetting("RosenbergDatabase", "Forms", Me.Name & ".Height", 5595)

DoCmd.MoveSize l, t, w, h

cmdAttributes.Enabled = PrivJournalFlags

End Sub

Private Sub Form_Close()
DoCmd.Restore
SaveSetting "RosenbergDatabase", "Forms", Me.Name & ".Left", IIf(Me.WindowLeft = 0, 1, Me.WindowLeft)
SaveSetting "RosenbergDatabase", "Forms", Me.Name & ".Top", IIf(Me.WindowTop = 0, 1, Me.WindowTop)
SaveSetting "RosenbergDatabase", "Forms", Me.Name & ".Height", Me.WindowHeight
SaveSetting "RosenbergDatabase", "Forms", Me.Name & ".Width", Me.WindowWidth
End Sub

Private Sub Form_Resize()
Const MinWidth = 3800, MinHeight = 3800
On Error GoTo ResizeErr

If Me.WindowWidth >= MinWidth Then
    webJournal.Width = Me.WindowWidth - 340
End If
If Me.WindowHeight >= MinHeight Then
    webJournal.Height = Me.WindowHeight - 1500
    cmdNewJournalEntry.Top = Me.WindowHeight - 900
    cmdPrintJournal.Top = Me.WindowHeight - 900
    cmdAttributes.Top = Me.WindowHeight - 900
End If

Exit Sub

ResizeErr:
    Resume Next
End Sub

Private Sub cmdAttributes_Click()
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else

    On Error GoTo Err_cmdAttributes_Click
    
    DoCmd.OpenForm "Edit Journal", , , "FileNumber=" & FileNumber, , acDialog
    Call ViewJournal
    
Exit_cmdAttributes_Click:
        Exit Sub
    
Err_cmdAttributes_Click:
        MsgBox Err.Description
        Resume Exit_cmdAttributes_Click
End If

    
End Sub

Private Sub optView_AfterUpdate()
Call ViewJournal
End Sub
