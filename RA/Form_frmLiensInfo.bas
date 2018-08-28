VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLiensInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
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


Private Sub cmdPre3_Click()
Dim FileNumber As Long
Dim strResult, strResult1 As String

FileNumber = Forms!foreclosuredetails!FileNumber

Select Case frmOption


Case 1
'DOT-Ours

strResult1 = "Deed of Trust dated " & Format$(Forms!foreclosuredetails!DOTdate, "mmmm d, yyyy") & " securing " & _
        " " & Forms!foreclosuredetails!OriginalBeneficiary & " in the original amount of " & Format$(Forms!foreclosuredetails!OriginalPBal, "Currency") & " and recorded on" & _
        " " & Format$(Forms!foreclosuredetails!DOTrecorded, "mmmm d, yyyy") & " " & LiberFolio(Forms!foreclosuredetails!Liber, Forms!foreclosuredetails!Folio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & Forms![Case List]!JurisdictionID.Column(1) & "."

If Not IsNull(Me.txtLiens3) Then
        strResult = Me.txtLiens3
     
       Me.txtLiens3 = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtLiens3 = strResult1
 End If
Case 2
'DOT-others

strResult1 = "Deed of Trust dated " & Format$(Me!txtDotOtherDated, "mmmm d, yyyy") & " securing " & _
        " " & Me!txtDotOtherBeneficiary & " in the original amount of " & Format$(Me!txtDotOtherAmt, "Currency") & " and recorded on" & _
        " " & Format$(Me!txtDotOtherRecordDate, "mmmm d, yyyy") & " " & LiberFolio(Me!txtDotOtherLiber, Me!txtDotOtherFolio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & Forms![Case List]!JurisdictionID.Column(1) & "."

If Not IsNull(Me.txtLiens3) Then
        strResult = Me.txtLiens3
     
       Me.txtLiens3 = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtLiens3 = strResult1
 End If

Case 3

'Mortgage
strResult1 = "Mortgage dated " & Format$(Me!txtmtgdate, "mmmm d, yyyy") & " securing " & _
        " " & Me!txtmtgBeneficiary & " in the original amount of " & Format$(Me!txtmtgAmt, "Currency") & " and recorded on" & _
        " " & Format$(Me!txtmtgRecorddate, "mmmm d, yyyy") & " " & LiberFolio(Me!txtmtgLiber, Me!txtmtgFolio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & Forms![Case List]!JurisdictionID.Column(1) & "."

If Not IsNull(Me.txtLiens3) Then
        strResult = Me.txtLiens3
     
       Me.txtLiens3 = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtLiens3 = strResult1
 End If

Case 4
'Beneficiary
strResult1 = " *** Beneficiary in Deed of Trust is Mortgage Electronic Registration Systems, Inc., acting solely as nominee for lender.***  "


If Not IsNull(Me.txtLiens3) Then
        strResult = Me.txtLiens3
     
       Me.txtLiens3 = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtLiens3 = strResult1
 End If

Case 5
'Assignment
strResult1 = "*** Assignment of DOT from " & Me!txtAssimentAssignor & " to " & _
        "" & Me!txtAssimentAssignee & " filed " & Format$(Me!txtAssimentDate, "mmmm d, yyyy") & " at " & _
        "" & LiberFolio(Me!txtAssimentLiber, Me!txtAssimentFolio, Forms!foreclosuredetails!State) & " .*** "
  
  
 If Not IsNull(Me.txtLiens3) Then
        strResult = Me.txtLiens3
     
       Me.txtLiens3 = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtLiens3 = strResult1
 End If
 
Case 6
'SOT

strResult1 = "*** Substitution of Trustee appointing " & Me!txtSOTNames & " recorded on " & _
        "" & Format$(Me!txtSOTDate, "mmmm d, yyyy") & " at " & LiberFolio(Me!txtSOTLiber, Me!txtSOTFolio, Forms!foreclosuredetails!State) & " .*** "
       
  If Not IsNull(Me.txtLiens3) Then
        strResult = Me.txtLiens3
     
       Me.txtLiens3 = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtLiens3 = strResult1
 End If
 
Case 7
'Release
   
 strResult1 = "*** Certificate of Satisfaction filed " & _
        "" & Format$(Me!txtReleaseDate, "mmmm d, yyyy") & " at " & LiberFolio(Me!txtReleaseLiber, Me!txtReleaseFolio, Forms!foreclosuredetails!State) & " .*** "


If Not IsNull(Me.txtLiens3) Then
        strResult = Me.txtLiens3
     
       Me.txtLiens3 = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtLiens3 = strResult1
 End If
 

Case 8
If Not IsNull(Me.txtLiens3) Then
        strResult = Me.txtLiens3
     
       Me.txtLiens3 = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtLiens3 = ""
 End If
   
    
End Select
End Sub

Private Sub cmdPreBlank_Click()
Dim FileNumber As Long
Dim strResult, strResult1 As String

FileNumber = Forms!foreclosuredetails!FileNumber

Select Case frmOption

Case 1, 2, 3, 4, 5, 6, 7, 8

If Not IsNull(Me.txtblank) Then
        strResult = Me.txtblank
     
       Me.txtblank = strResult & vbNewLine & vbNewLine
 Else
       Me.txtblank = ""
 End If

End Select
End Sub

Private Sub cmdPreJunior_Click()

Dim strResult, strResult1 As String
Dim FileNumber As Long

FileNumber = Forms!foreclosuredetails!FileNumber

Select Case frmOption

Case 1
'DOT-Ours

strResult1 = "Deed of Trust dated " & Format$(Forms!foreclosuredetails!DOTdate, "mmmm d, yyyy") & " securing " & _
        " " & Forms!foreclosuredetails!OriginalBeneficiary & " in the original amount of " & Format$(Forms!foreclosuredetails!OriginalPBal, "Currency") & " and recorded on" & _
        " " & Format$(Forms!foreclosuredetails!DOTrecorded, "mmmm d, yyyy") & " " & LiberFolio(Forms!foreclosuredetails!Liber, Forms!foreclosuredetails!Folio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & Forms![Case List]!JurisdictionID.Column(1) & "."

If Not IsNull(Me.txtJunior) Then
        strResult = Me.txtJunior
     
       Me.txtJunior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtJunior = strResult1
End If

Case 2
'DOT-others

strResult1 = "Deed of Trust dated " & Format$(Me!txtDotOtherDated, "mmmm d, yyyy") & " securing " & _
        " " & Me!txtDotOtherBeneficiary & " in the original amount of " & Format$(Me!txtDotOtherAmt, "Currency") & " and recorded on" & _
        " " & Format$(Me!txtDotOtherRecordDate, "mmmm d, yyyy") & " " & LiberFolio(Me!txtDotOtherLiber, Me!txtDotOtherFolio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & Forms![Case List]!JurisdictionID.Column(1) & "."

If Not IsNull(Me.txtJunior) Then
        strResult = Me.txtJunior
     
       Me.txtJunior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtJunior = strResult1
End If


Case 3
'Mortgage
strResult1 = "Mortgage dated " & Format$(Me!txtmtgdate, "mmmm d, yyyy") & " securing " & _
        " " & Me!txtmtgBeneficiary & " in the original amount of " & Format$(Me!txtmtgAmt, "Currency") & " and recorded on" & _
        " " & Format$(Me!txtmtgRecorddate, "mmmm d, yyyy") & " " & LiberFolio(Me!txtmtgLiber, Me!txtmtgFolio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & Forms![Case List]!JurisdictionID.Column(1) & "."


If Not IsNull(Me.txtJunior) Then
        strResult = Me.txtJunior
     
       Me.txtJunior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtJunior = strResult1
End If

Case 4
'Beneficiary
strResult1 = " *** Beneficiary in Deed of Trust is Mortgage Electronic Registration Systems, Inc., acting solely as nominee for lender.***  "

If Not IsNull(Me.txtJunior) Then
        strResult = Me.txtJunior
     
       Me.txtJunior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtJunior = strResult1
End If

Case 5
'Assignment
 strResult1 = "*** Assignment of DOT from " & Me!txtAssimentAssignor & " to " & _
        "" & Me!txtAssimentAssignee & " filed " & Format$(Me!txtAssimentDate, "mmmm d, yyyy") & " at " & _
        "" & LiberFolio(Me!txtAssimentLiber, Me!txtAssimentFolio, Forms!foreclosuredetails!State) & " .*** "
  
 If Not IsNull(Me.txtJunior) Then
        strResult = Me.txtJunior
     
       Me.txtJunior = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtJunior = strResult1
 End If

Case 6
'SOT

strResult1 = "*** Substitution of Trustee appointing " & Me!txtSOTNames & " recorded on " & _
        "" & Format$(Me!txtSOTDate, "mmmm d, yyyy") & " at " & LiberFolio(Me!txtSOTLiber, Me!txtSOTFolio, Forms!foreclosuredetails!State) & " .*** "
       
       
 If Not IsNull(Me.txtJunior) Then
        strResult = Me.txtJunior
     
       Me.txtJunior = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtJunior = strResult1
 End If
 
Case 7
'Release
   
 strResult1 = "*** Certificate of Satisfaction filed " & _
        "" & Format$(Me!txtReleaseDate, "mmmm d, yyyy") & " at " & LiberFolio(Me!txtReleaseLiber, Me!txtReleaseFolio, Forms!foreclosuredetails!State) & " .*** "

If Not IsNull(Me.txtJunior) Then
        strResult = Me.txtJunior
     
       Me.txtJunior = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtJunior = strResult1
 End If

Case 8
If Not IsNull(Me.txtJunior) Then
        strResult = Me.txtJunior
     
       Me.txtJunior = strResult & vbNewLine & vbNewLine & strResult1
 Else
       Me.txtJunior = ""
 End If
   
    
End Select
End Sub

Private Sub cmdPreSenior_Click()

Dim FileNumber As Long
Dim strResult, strResult1 As String

FileNumber = Forms!foreclosuredetails!FileNumber

Select Case frmOption

Case 1
'DOT-Ours

strResult1 = "Deed of Trust dated " & Format$(Forms!foreclosuredetails!DOTdate, "mmmm d, yyyy") & " securing " & _
        " " & Forms!foreclosuredetails!OriginalBeneficiary & " in the original amount of " & Format$(Forms!foreclosuredetails!OriginalPBal, "Currency") & " and recorded on" & _
        " " & Format$(Forms!foreclosuredetails!DOTrecorded, "mmmm d, yyyy") & " " & LiberFolio(Forms!foreclosuredetails!Liber, Forms!foreclosuredetails!Folio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & Forms![Case List]!JurisdictionID.Column(1) & "."

If Not IsNull(Me.txtSenior) Then
        strResult = Me.txtSenior
     
       Me.txtSenior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtSenior = strResult1
End If


Case 2
'DOT-others

strResult1 = "Deed of Trust dated " & Format$(Me!txtDotOtherDated, "mmmm d, yyyy") & " securing " & _
        " " & Me!txtDotOtherBeneficiary & " in the original amount of " & Format$(Me!txtDotOtherAmt, "Currency") & " and recorded on" & _
        " " & Format$(Me!txtDotOtherRecordDate, "mmmm d, yyyy") & " " & LiberFolio(Me!txtDotOtherLiber, Me!txtDotOtherFolio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & Forms![Case List]!JurisdictionID.Column(1) & "."

If Not IsNull(Me.txtSenior) Then
        strResult = Me.txtSenior
     
       Me.txtSenior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtSenior = strResult1
End If

Case 3
'Mortgage
strResult1 = "Mortgage dated " & Format$(Me!txtmtgdate, "mmmm d, yyyy") & " securing " & _
        " " & Me!txtmtgBeneficiary & " in the original amount of " & Format$(Me!txtmtgAmt, "Currency") & " and recorded on" & _
        " " & Format$(Me!txtmtgRecorddate, "mmmm d, yyyy") & " " & LiberFolio(Me!txtmtgLiber, Me!txtmtgFolio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & Forms![Case List]!JurisdictionID.Column(1) & "."

If Not IsNull(Me.txtSenior) Then
        strResult = Me.txtSenior
     
       Me.txtSenior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtSenior = strResult1
End If

Case 4
'Beneficiary
strResult1 = " *** Beneficiary in Deed of Trust is Mortgage Electronic Registration Systems, Inc., acting solely as nominee for lender.***  "


If Not IsNull(Me.txtSenior) Then
        strResult = Me.txtSenior
     
       Me.txtSenior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtSenior = strResult1
End If

Case 5
'Assignment
  strResult1 = "*** Assignment of DOT from " & Me!txtAssimentAssignor & " to " & _
        "" & Me!txtAssimentAssignee & " filed " & Format$(Me!txtAssimentDate, "mmmm d, yyyy") & " at " & _
        "" & LiberFolio(Me!txtAssimentLiber, Me!txtAssimentFolio, Forms!foreclosuredetails!State) & " .*** "
  
 If Not IsNull(Me.txtSenior) Then
        strResult = Me.txtSenior
     
       Me.txtSenior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtSenior = strResult1
End If
 

Case 6
'SOT

strResult1 = "*** Substitution of Trustee appointing " & Me!txtSOTNames & " recorded on " & _
        "" & Format$(Me!txtSOTDate, "mmmm d, yyyy") & " at " & LiberFolio(Me!txtSOTLiber, Me!txtSOTFolio, Forms!foreclosuredetails!State) & " .*** "
       
              
If Not IsNull(Me.txtSenior) Then
        strResult = Me.txtSenior
     
       Me.txtSenior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtSenior = strResult1
End If
       
Case 7
'Release
   
 strResult1 = "*** Certificate of Satisfaction filed " & _
        "" & Format$(Me!txtReleaseDate, "mmmm d, yyyy") & " at " & LiberFolio(Me!txtReleaseLiber, Me!txtReleaseFolio, Forms!foreclosuredetails!State) & " .*** "


If Not IsNull(Me.txtSenior) Then
        strResult = Me.txtSenior
     
       Me.txtSenior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtSenior = strResult1
End If
       
Case 8

If Not IsNull(Me.txtSenior) Then
        strResult = Me.txtSenior
     
       Me.txtSenior = strResult & vbNewLine & vbNewLine & strResult1
Else
       Me.txtSenior = ""
End If
   
    
End Select

End Sub



Private Sub cmdupdate3_Click()
'Dim rs As Recordset
'Dim FileNumber As Long

'FileNumber = Forms!ForeclosureDetails!FileNumber


'Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

'If Not rs.EOF Then

'rs.Edit
'rs!TitleReview3 = Me.txtLiens3
'rs.Update

'rs.Close
'Set rs = Nothing

'MsgBox ("File updated")
'Forms!ForeclosureDetails.Refresh
'End If

'--------------------------------

Dim rs As Recordset
Dim FileNumber As Long
Dim strDesc As String

FileNumber = Forms!foreclosuredetails!FileNumber

If Not IsNull(Forms!foreclosuredetails!sfrmFCtitle!TitleReview3) Then
    strDesc = Forms!foreclosuredetails!sfrmFCtitle!TitleReview3
Else
    strDesc = ""
End If

'strDesc = Forms!ForeclosureDetails!sfrmFCtitle!TitleReviewJudgments

Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

'If Not rs.EOF And strDesc <> Me.txtJudgments Then

If Not rs.EOF Then

If (IsNull(txtLiens3) And Not IsNull(strDesc)) Or (Not IsNull(txtLiens3) And IsNull(strDesc)) Then GoTo tracking:
If (Not IsNull(strDesc) Or strDesc <> "") And Not IsNull(txtLiens3) And strDesc <> txtLiens3 Then GoTo tracking:
If strDesc = txtLiens3 Then GoTo msg:

tracking:

rs.Edit
'rs!TitleReviewLiens = Me.txtSenior
rs!TitleReview3 = Me.txtLiens3
rs.Update

rs.Close
Set rs = Nothing

'tracking:

Set rs = CurrentDb.OpenRecordset("SELECT * FROM Audit_4", dbOpenDynaset, dbSeeChanges)

rs.AddNew
rs!FileNumber = FileNumber
rs!TableName = "FCTitle"
rs!FieldName = "TitleReview3"
rs!Username = GetFullName
rs!ChangeDate = Now()
rs!ChangeType = "UPDATE"
rs!OldValue = strDesc
rs!NewValue = Me.txtLiens3
rs.Update

rs.Close
Set rs = Nothing

MsgBox ("File updated")
Forms!foreclosuredetails.Refresh

Else
msg:  MsgBox ("No changes for updating")
End If
End Sub

Private Sub cmdupdateBlank_Click()
'Dim rs As Recordset
'Dim FileNumber As Long

'FileNumber = Forms!ForeclosureDetails!FileNumber


'Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

'If Not rs.EOF Then

'rs.Edit
'rs!TitleReviewBlank = Me.txtblank
'rs.Update

'rs.Close
'Set rs = Nothing

'MsgBox ("File updated")
'Forms!ForeclosureDetails.Refresh
'End If

Dim rs As Recordset
Dim FileNumber As Long
Dim strDesc As String

FileNumber = Forms!foreclosuredetails!FileNumber

If Not IsNull(Forms!foreclosuredetails!sfrmFCtitle!TitleReviewBlank) Then
    strDesc = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewBlank
Else
    strDesc = ""
End If

'strDesc = Forms!ForeclosureDetails!sfrmFCtitle!TitleReviewJudgments

Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

'If Not rs.EOF And strDesc <> Me.txtJudgments Then

If Not rs.EOF Then

If (IsNull(txtblank) And Not IsNull(strDesc)) Or (Not IsNull(txtblank) And IsNull(strDesc)) Then GoTo tracking:
If (Not IsNull(strDesc) Or strDesc <> "") And Not IsNull(txtblank) And strDesc <> txtblank Then GoTo tracking:
If strDesc = txtblank Then GoTo msg:

tracking:

rs.Edit
'rs!TitleReviewLiens = Me.txtSenior
rs!TitleReviewBlank = Me.txtblank
rs.Update

rs.Close
Set rs = Nothing

'tracking:

Set rs = CurrentDb.OpenRecordset("SELECT * FROM Audit_4", dbOpenDynaset, dbSeeChanges)

rs.AddNew
rs!FileNumber = FileNumber
rs!TableName = "FCTitle"
rs!FieldName = "TitleReviewBlank"
rs!Username = GetFullName
rs!ChangeDate = Now()
rs!ChangeType = "UPDATE"
rs!OldValue = strDesc
rs!NewValue = Me.txtblank
rs.Update

rs.Close
Set rs = Nothing

MsgBox ("File updated")
Forms!foreclosuredetails.Refresh

Else
msg:  MsgBox ("No changes for updating")
End If
End Sub

Private Sub cmdupdateJunior_Click()
'Dim rs As Recordset
'Dim FileNumber As Long

'FileNumber = Forms!ForeclosureDetails!FileNumber


'Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

'If Not rs.EOF Then

'rs.Edit
'rs!TitleReviewJunior = Me.txtJunior
'rs.Update

'rs.Close
'Set rs = Nothing

'MsgBox ("File updated")
'Forms!ForeclosureDetails.Refresh
'End If


'------------------------
Dim rs As Recordset
Dim FileNumber As Long
Dim strDesc As String

FileNumber = Forms!foreclosuredetails!FileNumber

If Not IsNull(Forms!foreclosuredetails!sfrmFCtitle!TitleReviewJunior) Then
    strDesc = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewJunior
Else
    strDesc = ""
End If

'strDesc = Forms!ForeclosureDetails!sfrmFCtitle!TitleReviewJudgments

Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

'If Not rs.EOF And strDesc <> Me.txtJudgments Then

If Not rs.EOF Then

If (IsNull(txtJunior) And Not IsNull(strDesc)) Or (Not IsNull(txtJunior) And IsNull(strDesc)) Then GoTo tracking:
If (Not IsNull(strDesc) Or strDesc <> "") And Not IsNull(txtJunior) And strDesc <> txtJunior Then GoTo tracking:
If strDesc = txtJunior Then GoTo msg:

tracking:

rs.Edit
'rs!TitleReviewLiens = Me.txtSenior
rs!TitleReviewJunior = Me.txtJunior
rs.Update

rs.Close
Set rs = Nothing

'tracking:

Set rs = CurrentDb.OpenRecordset("SELECT * FROM Audit_4", dbOpenDynaset, dbSeeChanges)

rs.AddNew
rs!FileNumber = FileNumber
rs!TableName = "FCTitle"
rs!FieldName = "TitleReviewJunior"
rs!Username = GetFullName
rs!ChangeDate = Now()
rs!ChangeType = "UPDATE"
rs!OldValue = strDesc
rs!NewValue = Me.txtJunior
rs.Update

rs.Close
Set rs = Nothing

MsgBox ("File updated")
Forms!foreclosuredetails.Refresh

Else
msg:  MsgBox ("No changes for updating")
End If

End Sub

Private Sub cmdupdateSenior_Click()
Dim rs As Recordset
Dim FileNumber As Long
Dim strDesc As String

FileNumber = Forms!foreclosuredetails!FileNumber

If Not IsNull(Forms!foreclosuredetails!sfrmFCtitle!TitleReviewLiens) Then
    strDesc = Forms!foreclosuredetails!sfrmFCtitle!TitleReviewLiens
Else
    strDesc = ""
End If

'strDesc = Forms!ForeclosureDetails!sfrmFCtitle!TitleReviewJudgments

Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

'If Not rs.EOF And strDesc <> Me.txtJudgments Then

If Not rs.EOF Then

If (IsNull(txtSenior) And Not IsNull(strDesc)) Or (Not IsNull(txtSenior) And IsNull(strDesc)) Then GoTo tracking:
If (Not IsNull(strDesc) Or strDesc <> "") And Not IsNull(txtSenior) And strDesc <> txtSenior Then GoTo tracking:
If strDesc = txtSenior Then GoTo msg:

tracking:

rs.Edit
rs!TitleReviewLiens = Me.txtSenior
rs.Update

rs.Close
Set rs = Nothing

'tracking:

Set rs = CurrentDb.OpenRecordset("SELECT * FROM Audit_4", dbOpenDynaset, dbSeeChanges)

rs.AddNew
rs!FileNumber = FileNumber
rs!TableName = "FCTitle"
rs!FieldName = "TitleReviewLiens"
rs!Username = GetFullName
rs!ChangeDate = Now()
rs!ChangeType = "UPDATE"
rs!OldValue = strDesc
rs!NewValue = Me.txtSenior
rs.Update

rs.Close
Set rs = Nothing

MsgBox ("File updated")
Forms!foreclosuredetails.Refresh

Else
msg:  MsgBox ("No changes for updating")
End If

End Sub

