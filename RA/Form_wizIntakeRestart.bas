VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_wizIntakeRestart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close


If IsLoaded("wizIntake1") = True Then
DoCmd.Close acForm, "wizIntake1"
End If

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdupdate_Click()
Dim rstFCdetails As Recordset, rstNames As Recordset, rstCase As Recordset, FileNum As Integer
FileNum = txtFileNumber
Set rstCase = CurrentDb.OpenRecordset("SELECT * FROM CaseList WHERE FileNumber=" & FileNum, dbOpenDynaset, dbSeeChanges)
Set rstFCdetails = CurrentDb.OpenRecordset("SELECT * FROM FCDetails WHERE FileNumber=" & FileNum, dbOpenDynaset, dbSeeChanges)
Set rstNames = CurrentDb.OpenRecordset("SELECT * FROM Names WHERE (FileNumber=" & FileNum & " AND Mortgagor = True) OR (FileNumber=" & FileNum & " AND Owner = True)OR (FileNumber=" & FileNum & " AND Noteholder = True)", dbOpenDynaset, dbSeeChanges)



With rstCase
    .Edit
    !ReferralDate = txtReferralDate
    !PrimaryDefName = txtProjectName
    !ClientID = cbxClient
    !JurisdictionID = cbxJurisdictionID
    .Update
    .Close
End With


With rstFCdetails
    .Edit
    !LoanNumber = txtLoanNumber
    !PrimaryFirstName = txtFirstName1
    !PrimaryLastName = txtLastName1
    !SecondaryFirstName = txtFirstName2
    !SecondaryLastName = txtLastName2
    !PropertyAddress = txtPropertyAddress
    !City = txtCity
    !State = txtState
    !ZipCode = txtZipCode
    .Update
    .Close
End With

With rstNames

Do Until .EOF
.Delete
.MoveNext
Loop


.AddNew
    !FileNumber = FileNum
    !First = txtFirstName1
    !Last = txtLastName1
    !ProjName = txtLastName1 & ", " & txtFirstName1
    !SSN = txtSSN1
    !Address = txtPropertyAddress
    !City = txtCity
    !State = txtState
    !Zip = txtZipCode
    .Update
    
End With
If Not IsNull(txtLastName2) Then
    With rstNames
        .AddNew
        !FileNumber = FileNum
        !First = txtFirstName2
        !Last = txtLastName2
        !SSN = txtSSN2
        !Address = txtPropertyAddress
        !City = txtCity
        !State = txtState
        !Zip = txtZipCode
        .Update
    End With
End If
If Not IsNull(txtLastName3) Then
    With rstNames
        .AddNew
        !FileNumber = FileNum
        !First = txtFirstName3
        !Last = txtLastName3
        !SSN = txtSSN3
        !Address = txtPropertyAddress
        !City = txtCity
        !State = txtState
        !Zip = txtZipCode
        .Update
    End With
End If
If Not IsNull(txtLastName4) Then
    With rstNames
        .AddNew
        !FileNumber = FileNum
        !First = txtFirstName4
        !Last = txtLastName4
        !SSN = txtSSN2
        !Address = txtPropertyAddress
        !City = txtCity
        !State = txtState
        !Zip = txtZipCode
        .Update
    End With
End If
With rstNames
' Add an All Occupants entry for each FC
        .AddNew
        !FileNumber = FileNum
        !First = "All"
        !Last = "Occupants"
        !Address = txtPropertyAddress
        !City = txtCity
        !State = txtState
        !Zip = txtZipCode
        .Update
End With

MsgBox "Values have been updated"
DoCmd.Close acForm, Me.Name

End Sub



Private Sub txtZipCode_AfterUpdate()
Dim rstZip As Recordset, JurisdictionID As Variant

If Not IsNull(txtZipCode) Then
    Set rstZip = CurrentDb.OpenRecordset("SELECT * FROM ZipCodes WHERE ZipCode = '" & Left$(txtZipCode, 5) & "' and Preferred = 'Yes'", dbOpenSnapshot)
    If Not rstZip.EOF Then
        txtCity = StrConv(rstZip!City, vbProperCase)
        txtState = rstZip!State
        cbxJurisdictionID = Null
        JurisdictionID = DLookup("JurisdictionID", "JurisdictionList", "Jurisdiction Like '" & rstZip!County & "*'")
        If Not IsNull(JurisdictionID) Then
            If txtState = DLookup("State", "JurisdictionList", "JurisdictionID=" & JurisdictionID) Then
                cbxJurisdictionID = JurisdictionID
            End If
        End If
    Else
        txtCity = Null
        txtState = Null
        cbxJurisdictionID = Null
        MsgBox "CAUTION: unknown Zip Code", vbExclamation
    End If
    rstZip.Close
End If
End Sub


