VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJudgmentsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim strDesc As String
Option Explicit



Private Sub cmdClose_Click()

On Error GoTo Err_cmdClose_Click
DoCmd.Close


Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub


Private Sub cmdPreView_Click()
Dim FileNumber As Long
Dim strResult, strResult1 As String

FileNumber = Forms!foreclosuredetails!FileNumber


Select Case frmOption

Case 1

strResult1 = "Judgment in favor of " & Me.txtNameJud & " filed in the Circuit" & _
        " Court of " & Me.txtCourt & " at Case no. " & Me.txtcaseJud & " on " & Me.txtdateJud & " in the amount of $ " & Me.txtamrjud & "."
    
    If Not IsNull(Me.txtJudgments) Then
        strResult = Me.txtJudgments
     
       Me.txtJudgments = strResult & vbNewLine & vbNewLine & strResult1
    Else
       Me.txtJudgments = strResult1
    End If

Case 2

strResult1 = "Notice of Water/Sewer lien filed on " & Me.txtDateWater & " in the amount of $ " & Me.txtAmtWater & "."
    
    If Not IsNull(Me.txtJudgments) Then
        strResult = Me.txtJudgments
        Me.txtJudgments = strResult & vbNewLine & vbNewLine & strResult1
    Else
        Me.txtJudgments = strResult1
    End If

Case 3
       strResult1 = "Notice of HOA/Condo lien filed on " & Me.txtDateHOA & " in the amount of $ " & Me.txtAmtHOA & "."

    If Not IsNull(Me.txtJudgments) Then
       strResult = Me.txtJudgments

       Me.txtJudgments = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtJudgments = strResult1
End If

Case 4

       strResult1 = "IRS/federal lien filed on " & Me.txtDateIRS & " in case # " & Me.txtCaseIRS & " in the amount of $ " & Me.txtAmtIRS & "."

    If Not IsNull(Me.txtJudgments) Then
       strResult = Me.txtJudgments
       
       Me.txtJudgments = strResult & vbNewLine & vbNewLine & strResult1

    Else
       Me.txtJudgments = strResult1
    End If

Case 5
       strResult1 = "State Tax lien filed on " & Me.txtDateState & " in case # " & Me.txtCaseState & " in the amount of $ " & Me.txtAmtstate & "."

    If Not IsNull(Me.txtJudgments) Then
       strResult = Me.txtJudgments
        Me.txtJudgments = strResult & vbNewLine & vbNewLine & strResult1
    Else
       Me.txtJudgments = strResult1
    End If

Case 6
        strResult1 = "Foreclosure case docketed in favor of  " & Me.txtNamesFC & " on " & Me.txtDateFC & " at Case no. " & Me.txtCaseFC & "."

    If Not IsNull(Me.txtJudgments) Then
        strResult = Me.txtJudgments

        Me.txtJudgments = strResult & vbNewLine & vbNewLine & strResult1

    Else
        Me.txtJudgments = strResult1
    End If

Case 7
    If Not IsNull(Me.txtJudgments) Then
        strResult = Me.txtJudgments

        Me.txtJudgments = strResult & vbNewLine & vbNewLine
    Else
        Me.txtJudgments = Me.txtJudgments

    End If
End Select

End Sub


Private Sub cmdupdate_Click()

Dim rs As Recordset
Dim FileNumber As Long

FileNumber = Forms!foreclosuredetails!FileNumber

If Not IsNull(Forms!foreclosuredetails!sfrmFCtitle!TitleReviewJudgments) Then
    strDesc = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewJudgments
Else
    strDesc = ""
End If

'strDesc = Forms!ForeclosureDetails!sfrmFCtitle!TitleReviewJudgments

Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

'If Not rs.EOF And strDesc <> Me.txtJudgments Then

If Not rs.EOF Then

If (IsNull(txtJudgments) And Not IsNull(strDesc)) Or (Not IsNull(txtJudgments) And IsNull(strDesc)) Then GoTo tracking:
If (Not IsNull(strDesc) Or strDesc <> "") And Not IsNull(txtJudgments) And strDesc <> txtJudgments Then GoTo tracking:
If strDesc = txtJudgments Then GoTo msg:

tracking:

rs.Edit
rs!TitleReviewJudgments = Me.txtJudgments
rs.Update

rs.Close
Set rs = Nothing

'tracking:

Set rs = CurrentDb.OpenRecordset("SELECT * FROM Audit_4", dbOpenDynaset, dbSeeChanges)

rs.AddNew
rs!FileNumber = FileNumber
rs!TableName = "FCTitle"
rs!FieldName = "TitleReviewJudgments"
rs!Username = GetFullName
rs!ChangeDate = Now()
rs!ChangeType = "UPDATE"
rs!OldValue = strDesc
rs!NewValue = Me.txtJudgments
rs.Update

rs.Close
Set rs = Nothing

MsgBox ("File updated")
Forms!foreclosuredetails.Refresh

Else
msg:  MsgBox ("No changes for updating")
End If

End Sub

Private Sub Form_AfterUpdate()
'lstSelect.Requery
End Sub

Private Sub cmdNew_Click()

On Error GoTo Err_cmdNew_Click
DoCmd.GoToRecord , , acNewRec

Exit_cmdNew_Click:
    Exit Sub

Err_cmdNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdNew_Click
    
End Sub

