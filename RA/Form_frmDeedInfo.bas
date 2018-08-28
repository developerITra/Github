VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDeedInfo"
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


Private Sub cmdPreView_Click()
Dim rs1, rs2, rs3, rs4 As Recordset
Dim strSoleOwner As String
Dim cnt As Integer
Dim FileNumber As Long
Dim strResult As String

FileNumber = Forms!foreclosuredetails!FileNumber

cnt = CountNames(FileNumber, "Owner=True")

'Set rs1 = CurrentDb.OpenRecordset("SELECT * FROM Files WHERE RAFileNum =" & FileNumber, dbOpenDynaset, dbSeeChanges)
Set rs4 = CurrentDb.OpenRecordset("SELECT * FROM vw_leasehold WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

Select Case frmOption

Case 1

If cnt = 1 Then
    If Forms!foreclosuredetails!State = "VA" Then
        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        "," & " sole owner of a " & rs4!Leasehold & " property by " & _
        "virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , " & _
        " " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."
    ElseIf Forms!foreclosuredetails!State = "DC" Then
        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        "," & " sole owner of a " & rs4!Leasehold & " property by " & _
        "virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , " & _
        "in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & rs4!Jurisdiction & "."
    Else
        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        "," & " sole owner of a " & rs4!Leasehold & " property by " & _
        "virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , " & _
        "in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " " & _
        "among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."
    End If
    
Else
MsgBox ("This is not sole Owner")
End If

Case 2

If cnt = 2 Then

    If Forms!foreclosuredetails!State = "VA" Then
        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", tenants by the entirety  of a " & rs4!Leasehold & " property by virtue " & _
        " of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."
    ElseIf Forms!foreclosuredetails!State = "DC" Then
        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", tenants by the entirety  of a " & rs4!Leasehold & " property by virtue " & _
        " of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & "."
    Else
    
            Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", tenants by the entirety  of a " & rs4!Leasehold & " property by virtue " & _
        " of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."

    End If

Else

MsgBox ("This is not T/E")
End If

Case 3
If cnt >= 2 Then
    If Forms!foreclosuredetails!State = "VA" Then

        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", joint tenants of a " & rs4!Leasehold & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."
    ElseIf Forms!foreclosuredetails!State = "DC" Then
                 Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", joint tenants of a " & rs4!Leasehold & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & "."

    
    Else
             Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", joint tenants of a " & rs4!Leasehold & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."

    End If
Else
    MsgBox ("Not J/T Tenant")
End If

Case 4

If cnt >= 2 Then
    If Forms!foreclosuredetails!State = "VA" Then

        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", tenants in common of a " & rs4!Leasehold & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."
    ElseIf Forms!foreclosuredetails!State = "DC" Then
        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", tenants in common of a " & rs4!Leasehold & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & "."

    Else
    
        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", tenants in common of a " & rs4!Leasehold & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."

    End If
 Else
    MsgBox ("Not TIC")
 End If

Case 5
    'If Forms!ForeclosureDetails!State = "VA" Then

        If Not IsNull(Me.txtnameof) Then
            strResult = Me.txtnameof
            Me.txtnameof = strResult & vbNewLine & vbNewLine & GetNames(FileNumber, 2, "Owner=True") & _
        ", Tenants in Common pursuant to a divorce decree dated " & Format$(Forms!frmDeedInfo!txtDateOrder, "mmmm d, yyyy") & " in the Circuit Court of " & Forms!frmDeedInfo!txtCir_court & " County, " & Forms!frmDeedInfo!txtState & " in case number  " & Forms!frmDeedInfo!txtCaseNum & "."
        Else
        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
            ", Tenants in Common pursuant to a divorce decree dated " & Format$(Forms!frmDeedInfo!txtDateOrder, "mmmm d, yyyy") & " in the Circuit Court of " & Forms!frmDeedInfo!txtCir_court & " County, " & Forms!frmDeedInfo!txtState & " in case number  " & Forms!frmDeedInfo!txtCaseNum & "."
        End If
    'Else
     'If Not IsNull(Me.txtnameof) Then
            'strResult = Me.txtnameof
            'Me.txtnameof = strResult & vbNewLine & vbNewLine & GetNames(FileNumber, 2, "Owner=True") & _
        '", Tenants in Common pursuant to a divorce decree dated " & Format$(Forms!frmDeedInfo!txtDateOrder, "mmmm d, yyyy") & " in the Circuit Court of " & Forms!frmDeedInfo!txtCir_court & " County, " & Forms!frmDeedInfo!txtState & " in case number  " & Forms!frmDeedInfo!txtCaseNum & "."
            'Else
        'Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
           ' ", Tenants in Common pursuant to a divorce decree dated " & Format$(Forms!frmDeedInfo!txtDateOrder, "mmmm d, yyyy") & " in the Circuit Court of " & Forms!frmDeedInfo!txtCir_court & " County, " & Forms!frmDeedInfo!txtState & " in case number  " & Forms!frmDeedInfo!txtCaseNum & "."
        'End If
    'End If
    
Case 6
        If Not IsNull(Me.txtnameof) Then
        strResult = Me.txtnameof
            If Forms!foreclosuredetails!State = "VA" Then
 
                Me.txtnameof = strResult & vbNewLine & vbNewLine & " " & Forms!frmDeedInfo!txtNamesHeirs & ", heirs of " & GetNames(FileNumber, 2, "Owner=True And Deceased = True") & _
                ", who died on " & Format$(Forms!frmDeedInfo!txtDOD, "mmmm d, yyyy") & " who, Owned a " & rs4!Leasehold & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."
            ElseIf Forms!foreclosuredetails!State = "DC" Then
                Me.txtnameof = strResult & vbNewLine & vbNewLine & " " & Forms!frmDeedInfo!txtNamesHeirs & ", heirs of " & GetNames(FileNumber, 2, "Owner=True And Deceased = True") & _
                ", who died on " & Format$(Forms!frmDeedInfo!txtDOD, "mmmm d, yyyy") & " who, Owned a " & rs4!Leasehold & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & "."

            Else
                Me.txtnameof = strResult & vbNewLine & vbNewLine & " " & Forms!frmDeedInfo!txtNamesHeirs & ", heirs of " & GetNames(FileNumber, 2, "Owner=True And Deceased = True") & _
                ", who died on " & Format$(Forms!frmDeedInfo!txtDOD, "mmmm d, yyyy") & " who, Owned a " & rs4!Leasehold & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."

            End If
    
    
    Else
    
        If Forms!foreclosuredetails!State = "VA" Then

            Me.txtnameof = " " & Forms!frmDeedInfo!txtNamesHeirs & ", heirs of " & GetNames(FileNumber, 2, "Owner=True") & _
            ", who died on " & Format$(Forms!frmDeedInfo!txtDOD, "mmmm d, yyyy") & ", of a " & rs4!Leasehold & " in the subject property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " ," & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."
            
        ElseIf Forms!foreclosuredetails!State = "DC" Then
        
            Me.txtnameof = " " & Forms!frmDeedInfo!txtNamesHeirs & ", heirs of " & GetNames(FileNumber, 2, "Owner=True") & _
            ", who died on " & Format$(Forms!frmDeedInfo!txtDOD, "mmmm d, yyyy") & ", of a " & rs4!Leasehold & " in the subject property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & "."

        Else
            Me.txtnameof = " " & Forms!frmDeedInfo!txtNamesHeirs & ", heirs of " & GetNames(FileNumber, 2, "Owner=True") & _
            ", who died on " & Format$(Forms!frmDeedInfo!txtDOD, "mmmm d, yyyy") & ", of a " & rs4!Leasehold & " in the subject property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & " , in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."
        End If
    End If

Case 7

    If Forms!foreclosuredetails!State = "VA" Then

        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", the             of a " & rs4!Leasehold & " with a remaining life estate interest to " & Forms!frmDeedInfo!txtNamesRem & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & ", " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."
    ElseIf Forms!foreclosuredetails!State = "DC" Then
        Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", the             of a " & rs4!Leasehold & " with a remaining life estate interest to " & Forms!frmDeedInfo!txtNamesRem & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & ", in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & "."

    Else
    
    Me.txtnameof = GetNames(FileNumber, 2, "Owner=True") & _
        ", the             of a " & rs4!Leasehold & " with a remaining life estate interest to " & Forms!frmDeedInfo!txtNamesRem & " property by virtue of a Deed dated " & Format$(Forms!frmDeedInfo!txtDeedDate, "mmmm d, yyyy") & " and Recorded " & Format$(Forms!frmDeedInfo!txtDeedRecordDate, "mmmm d, yyyy") & ", in " & LiberFolio(Forms!frmDeedInfo!txtLiber, Forms!frmDeedInfo!txtFolio, Forms!foreclosuredetails!State) & " among the Land Records of " & rs4!Jurisdiction & ", " & rs4!County & "."
    End If
   
Case 8

    If Not IsNull(Me.txtnameof) Then
        strResult = Me.txtnameof

        Me.txtnameof = strResult & vbNewLine & vbNewLine
    Else
        Me.txtnameof = Me.txtnameof

    End If
    
End Select

rs4.Close
Set rs4 = Nothing

End Sub

Private Sub lstSelect_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[BrokerID] = " & str(Nz(Me![lstSelect], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
End Sub

Private Sub cmdupdate_Click()

Dim rs As Recordset
Dim FileNumber As Long

FileNumber = Forms!foreclosuredetails!FileNumber


Set rs = CurrentDb.OpenRecordset("SELECT * FROM FCTitle WHERE FileNumber =" & FileNumber, dbOpenDynaset, dbSeeChanges)

If Not rs.EOF Then

rs.Edit
rs!TitleReviewNameOf = Me.txtnameof
rs.Update

rs.Close
Set rs = Nothing

MsgBox ("File updated")
Forms!foreclosuredetails.Refresh
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
