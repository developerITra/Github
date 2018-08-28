VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GetLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnBankruptcy_Click()
Me.txtPropertyAddress2 = ""
Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""

Dim strCriteria As String
Dim rsBKCourts As Recordset
Dim sql As String

sql = "SELECT Districts.BKdistrictName, Districts.ID, Districts.BKDistrictAddress, Districts.BKDistrictAddress2, Districts.BKDistrictState, Districts.BKDistrictState, Districts.BKDistrictZip, Districts.BKDistrictAttn, Districts.BKDistrictCity"
sql = sql + " FROM BKdetails INNER JOIN Districts ON BKdetails.District = Districts.ID;"
Set rsBKCourts = CurrentDb.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)

strCriteria = "[ID] =" & Forms![BankruptcyDetails]!District.Column(0) & ""

rsBKCourts.FindFirst (strCriteria)

If rsBKCourts.NoMatch = True Then
    MsgBox "No matching label data for this Jurisdiction"
    rsBKCourts.Close
    Set rsBKCourts = Nothing
    Exit Sub
Else

Me.txtPrimaryFirstName = rsBKCourts!BKdistrictName
Me.txtSecondaryFirstName = "Attn: " & rsBKCourts!BKDistrictAttn
Me.txtPropertyAddress = rsBKCourts!BKDistrictAddress
Me.txtPropertyAddress2 = rsBKCourts!BKDistrictAddress2
Me.txtState = rsBKCourts!BKDistrictState
Me.txtCity = rsBKCourts!BKDistrictCity
Me.txtZipCode = rsBKCourts!BKDistrictZip

End If

rsBKCourts.Close
Set rsBKCourts = Nothing
End Sub

Private Sub btnBethesda_Click()
Me.txtPropertyAddress2 = ""
Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""


Me.txtPrimaryFirstName = "Rosenberg and Associates"
Me.txtSecondaryFirstName = "Attn: "
Me.txtPropertyAddress = "7910 Woodmont Avenue"
Me.txtPropertyAddress2 = "Suite 750"
Me.txtCity = "Bethesda"
Me.txtState = "MD"
Me.txtZipCode = "20814"

End Sub

Private Sub btnClear_Click()

Me.txtPropertyAddress2 = ""
Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""

End Sub

Private Sub btnClerk_Click()

Me.txtPropertyAddress2 = ""
Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""

Dim strCriteria As String
Dim rsCircuitCourts As Recordset
Dim sql As String

sql = "SELECT CircuitCourts.CCID, CircuitCourts.CCName, CircuitCourts.CCAddress, CircuitCourts.CCAddress2, CircuitCourts.CCCity, CircuitCourts.CCState, CircuitCourts.CCZip, CircuitCourts.CCAttn, JurisdictionList.JurisdictionID"
sql = sql + " FROM CircuitCourts INNER JOIN JurisdictionList ON CircuitCourts.JurisdictionID = JurisdictionList.JurisdictionID;"
Set rsCircuitCourts = CurrentDb.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)

strCriteria = "[JurisdictionID] =" & Forms![Case List]!JurisdictionID & ""

rsCircuitCourts.FindFirst (strCriteria)

If rsCircuitCourts.NoMatch = True Then
    MsgBox "No matching label data for this Jurisdiction"
    rsCircuitCourts.Close
    Set rsCircuitCourts = Nothing
    Exit Sub
Else

Me.txtPrimaryFirstName = rsCircuitCourts!CCName
Me.txtSecondaryFirstName = rsCircuitCourts!CCAttn
Me.txtPropertyAddress = rsCircuitCourts!CCAddress
Me.txtPropertyAddress2 = rsCircuitCourts!CCAddress2
Me.txtState = rsCircuitCourts!CCState
Me.txtCity = rsCircuitCourts!CCCity
Me.txtZipCode = rsCircuitCourts!CCZip

End If

rsCircuitCourts.Close
Set rsCircuitCourts = Nothing

End Sub

Private Sub btnDeedRec_Click()

Me.txtPropertyAddress2 = ""
Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""

Dim sql As String
Dim rsDeed As Recordset
Dim strCriteria As String

sql = "SELECT JurisdictionList.JurisdictionID, Recorders.RecorderName, Recorders.RecorderAddress, Recorders.RecorderAddress2, Recorders.RecorderCity, Recorders.RecorderState, Recorders.RecorderZip, Recorders.RecorderEmail, Recorders.RecorderPhone, Recorders.RecorderATTN, Recorders.RecorderPrice, Recorders.JurisdictionID"
sql = sql + " FROM Recorders INNER JOIN JurisdictionList ON Recorders.JurisdictionID = JurisdictionList.JurisdictionID;"
'sql = sql + " WHERE (((JurisdictionList.JurisdictionID)=[Forms]![Case List]![JurisdictionID]));"

Set rsDeed = CurrentDb.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
strCriteria = "JurisdictionList![JurisdictionID] =" & Forms![Case List]!JurisdictionID & ""

rsDeed.FindFirst (strCriteria)
If rsDeed.NoMatch = True Then
    MsgBox "No matching label data for this Jurisdiction"
    rsDeed.Close
    Set rsDeed = Nothing
    Exit Sub
Else
Me.txtPrimaryFirstName = rsDeed!RecorderName
Me.txtSecondaryFirstName = rsDeed!RecorderATTN
Me.txtPropertyAddress = rsDeed!RecorderAddress
Me.txtPropertyAddress2 = rsDeed!RecorderAddress2
Me.txtState = rsDeed!RecorderState
Me.txtCity = rsDeed!RecorderCity
Me.txtZipCode = rsDeed!RecorderZip
End If

rsDeed.Close
Set rsDeed = Nothing


End Sub

Private Sub btnDistrict_Click()

Me.txtPropertyAddress2 = ""
Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""

Dim strCriteria As String
Dim rsDistrictCourts As Recordset
Dim sql As String

sql = "SELECT DistrictCourts.DCID, DistrictCourts.DCName, DistrictCourts.DCAddress, DistrictCourts.DCAddress2, DistrictCourts.DCCity, DistrictCourts.DCState, DistrictCourts.DCZip, DistrictCourts.DCAttn, JurisdictionList.JurisdictionID"
sql = sql + " FROM DistrictCourts INNER JOIN JurisdictionList ON DistrictCourts.JurisdictionID = JurisdictionList.JurisdictionID;"
Set rsDistrictCourts = CurrentDb.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)

strCriteria = "[JurisdictionID] =" & Forms![Case List]!JurisdictionID & ""

rsDistrictCourts.FindFirst (strCriteria)

If rsDistrictCourts.NoMatch = True Then
    MsgBox "No matching label data for this Jurisdiction"
    rsDistrictCourts.Close
    Set rsDistrictCourts = Nothing
    Exit Sub
Else

Me.txtPrimaryFirstName = rsDistrictCourts!DCName
Me.txtSecondaryFirstName = rsDistrictCourts!DCAttn
Me.txtPropertyAddress = rsDistrictCourts!DCAddress
Me.txtPropertyAddress2 = rsDistrictCourts!DCAddress2
Me.txtState = rsDistrictCourts!DCState
Me.txtCity = rsDistrictCourts!DCCity
Me.txtZipCode = rsDistrictCourts!DCZip

End If

rsDistrictCourts.Close
Set rsDistrictCourts = Nothing

End Sub

Private Sub btnEV_Click()
Dim sql As String
Dim rstLabelData As Recordset
Dim i As Integer
   
        sql = "SELECT Names.Company, Names.First, Names.Last, Names.AKA, Names.Address, Names.Address2, Names.City, Names.State, Names.Zip, CaseList.FileNumber, ClientList.FairDebt, ClientList.ShortClientName, CaseList.PrimaryDefName FROM (ClientList RIGHT JOIN (CaseList RIGHT JOIN [Names] ON CaseList.FileNumber=Names.FileNumber) ON ClientList.ClientID=CaseList.ClientID) LEFT JOIN EVdetails ON CaseList.FileNumber=EVdetails.FileNumber WHERE (((CaseList.FileNumber)=" & Forms![Case List]!FileNumber & ") And ((Names.EV)=True) And ((EVdetails.Current)=True));"
        Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        If rstLabelData.EOF Then
            MsgBox "Cannot print file label because no EV names checked.", vbCritical
            Exit Sub
            rstLabelData.Close
        Else
        Do While Not rstLabelData.EOF
            For i = 1 To 4
                Call StartLabel
                Print #6, FormatName(rstLabelData!Company, rstLabelData!First, rstLabelData!Last, "", rstLabelData!Address, rstLabelData!Address2, rstLabelData!City, rstLabelData!State, rstLabelData!Zip)
                Print #6, "|FONTSIZE 8"
                Print #6, "|BOTTOM"
                Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
                Call FinishLabel
            Next i
            rstLabelData.MoveNext
        Loop
        rstLabelData.Close
        MsgBox "EV labels have been printed", vbInformation
    End If


End Sub

Private Sub btnHUD_Click()

Me.txtPropertyAddress2 = ""
Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""

Dim sql As String
Dim rsHud As Recordset
Dim strCriteria As String

sql = "SELECT JurisdictionList.HUDAddress,HUDAddress.HUDNewAddress, HUDAddress.HUDNewAddress2, HUDAddress.HUDCity, HUDAddress.HUDState, HUDAddress.HUDZip, HUDAddress.HUDAttn, JurisdictionList.JurisdictionID"
sql = sql + " FROM JurisdictionList INNER JOIN HUDAddress ON JurisdictionList.HUDAddress = HUDAddress.ID;"

Set rsHud = CurrentDb.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
strCriteria = "[JurisdictionID] =" & Forms![Case List]!JurisdictionID & ""

rsHud.FindFirst (strCriteria)
If rsHud.NoMatch = True Then
    MsgBox "No matching label data for this Jurisdiction"
    rsHud.Close
    Set rsHud = Nothing
    Exit Sub
Else
Me.txtPrimaryFirstName = rsHud!HUDAttn
'Me.txtSecondaryFirstName = rsHud!LienorAttn
Me.txtPropertyAddress = rsHud!HUDNewAddress
Me.txtPropertyAddress2 = rsHud!HUDNewAddress2
Me.txtState = rsHud!HUDState
Me.txtCity = rsHud!HUDCity
Me.txtZipCode = rsHud!HUDZip

End If

rsHud.Close
Set rsHud = Nothing

End Sub

Private Sub btnIRS_Click() 'Needs ATTN

Me.txtPropertyAddress2 = ""
Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""

Dim sql As String
Dim rsIRS As Recordset
Dim strCriteria As String

sql = "SELECT JurisdictionList.IRSAddress, IRSAddress.IRSAddress AS Address, JurisdictionList.JurisdictionID, IRSAddress.dear, IRSAddress.IRSNewAddress, irsaddress.irsnewaddress2, irsaddress.irsCity, irsaddress.IRSstate, irsaddress.IRSzip, irsaddress.irsName "
sql = sql + " FROM JurisdictionList INNER JOIN IRSAddress ON JurisdictionList.IRSAddress = IRSAddress.ID;"

Set rsIRS = CurrentDb.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
strCriteria = "[JurisdictionID] =" & Forms![Case List]!JurisdictionID & ""

rsIRS.FindFirst (strCriteria)
If rsIRS.NoMatch = True Then
    MsgBox "No matching label data for this Jurisdiction"
    rsIRS.Close
    Set rsIRS = Nothing
    Exit Sub
Else
Me.txtPrimaryFirstName = rsIRS!IRSName
Me.txtSecondaryFirstName = "Attn: " & rsIRS!Dear
Me.txtPropertyAddress = rsIRS!IRSNewAddress
Me.txtPropertyAddress2 = rsIRS!IRSNewAddress2
Me.txtState = rsIRS!IRSState
Me.txtCity = rsIRS!IRSCity
Me.txtZipCode = rsIRS!IRSZip
End If

rsIRS.Close
Set rsIRS = Nothing

End Sub

Private Sub btnLienCerts_Click()

Me.txtPropertyAddress2 = ""
Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""

Dim sql As String
Dim rsLienors As Recordset
Dim strCriteria As String

sql = "SELECT JurisdictionList.JurisdictionID, Lienors.LienorName, Lienors.LienorAddress, Lienors.LienorAddress2, Lienors.LienorCity, Lienors.LienorState, Lienors.LienorZip, Lienors.LienorAttn"
sql = sql + " FROM JurisdictionList INNER JOIN Lienors ON JurisdictionList.JurisdictionID = Lienors.JurisdictionID;"
'sql = sql + " WHERE (((JurisdictionList.JurisdictionID)=[Forms]![Case List]![JurisdictionID]));"

Set rsLienors = CurrentDb.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
strCriteria = "[JurisdictionID] =" & Forms![Case List]!JurisdictionID & ""

rsLienors.FindFirst (strCriteria)
If rsLienors.NoMatch = True Then
    MsgBox "No Matching Label Data for this Jurisdiction"
    rsLienors.Close
    Set rsLienors = Nothing
    Exit Sub
Else
Me.txtPrimaryFirstName = rsLienors!LienorName
Me.txtSecondaryFirstName = rsLienors!LienorAttn
Me.txtPropertyAddress = rsLienors!LienorAddress
Me.txtPropertyAddress2 = rsLienors!LienorAddress2
Me.txtState = rsLienors!LienorState
Me.txtCity = rsLienors!LienorCity
Me.txtZipCode = rsLienors!LienorZip
End If

rsLienors.Close
Set rsLienors = Nothing

End Sub

Private Sub btnVA_Click()

Me.txtPropertyAddress2 = ""
Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""

Dim sql As String
Dim rsVA As Recordset
Dim strCriteria As String

sql = "SELECT JurisdictionList.VAAddress, VAAddress.VAAddress AS Address, JurisdictionList.JurisdictionID,VAAddress.vaNewAddress, VAAddress.vanewaddress2, VAAddress.Vacity, VAAddress.vastate, VAAddress.vaZip, VAAddress.vaAttn, VAAddress.vaName"
sql = sql + " FROM JurisdictionList INNER JOIN VAAddress ON JurisdictionList.VAAddress = VAAddress.ID;"

Set rsVA = CurrentDb.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
strCriteria = "[JurisdictionID] =" & Forms![Case List]!JurisdictionID & ""

rsVA.FindFirst (strCriteria)
If rsVA.NoMatch = True Then
    MsgBox "No matching label data for this Jurisdiction"
    rsVA.Close
    Set rsVA = Nothing
    Exit Sub
Else
Me.txtPrimaryFirstName = rsVA!VAName
Me.txtSecondaryFirstName = "Attn: " & rsVA!VAAttn
Me.txtPropertyAddress = rsVA!VANewAddress
Me.txtPropertyAddress2 = rsVA!VANewAddress2
Me.txtState = rsVA!VAState
Me.txtCity = rsVA!VACity
Me.txtZipCode = rsVA!VAZip
End If
rsVA.Close
Set rsVA = Nothing

End Sub

Private Sub btnVienna_Click()

Me.txtPrimaryFirstName = ""
Me.txtSecondaryFirstName = ""
Me.txtPropertyAddress = ""
Me.txtPropertyAddress2 = ""
Me.txtCity = ""
Me.txtState = ""
Me.txtZipCode = ""

Me.txtPrimaryFirstName = "Rosenberg and Associates"
Me.txtSecondaryFirstName = "Attn:"
Me.txtPropertyAddress = "8601 Westwood Center Dr"
Me.txtPropertyAddress2 = "Suite 255"
Me.txtCity = "Vienna"
Me.txtState = "VA"
Me.txtZipCode = "22182"

End Sub

Private Sub cmdCancel_Click()

On Error GoTo Err_cmdCancel_Click
DoCmd.Close

Exit_cmdCancel_Click:
    Exit Sub

Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Private Sub cmdupdate_Click() 'Print Function goes here

    Dim sql As String
    Dim rstLabelData As Recordset
              
    sql = "SELECT fcdetails.PropertyAddress, fcdetails.City, fcdetails.State, fcdetails.ZipCode, CaseList.FileNumber, ClientList.ShortClientName, CaseList.PrimaryDefName, CaseList.CaseTypeID, CaseList.Active, fcdetails.Current"
    sql = sql + " FROM (fcdetails INNER JOIN CaseList ON fcdetails.FileNumber = CaseList.FileNumber) INNER JOIN ClientList ON CaseList.ClientID = ClientList.ClientID"
    sql = sql + " WHERE CaseList.Active=True AND fcdetails.Current=True AND  CaseList.FileNumber =" & Forms![Case List]!FileNumber & ";"
      
          
          
    Set rstLabelData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        Call StartLabel
        'Make a seperate one that can tell if ATTN is in or not, what a pain
        'Print #6, FormatName("", Me.txtPrimaryFirstName & " " & Me.txtPrimaryLastName, "", "", Me.txtSecondaryFirstName & " " & Me.txtSecondaryLastName, Me.txtPropertyAddress, Me.txtCity, Me.txtState, Me.txtZipCode)
        Print #6, FormatName2(Me.txtPrimaryFirstName, Me.txtSecondaryFirstName, "", "", Me.txtPropertyAddress, Me.txtPropertyAddress2, Me.txtCity, Me.txtState, Me.txtZipCode)
        Print #6, "|FONTSIZE 8"
        Print #6, "|BOTTOM"
        Print #6, rstLabelData!FileNumber & " / " & rstLabelData!ShortClientName & " / " & rstLabelData!PrimaryDefName
        Call FinishLabel
        'MsgBox ("Label Created")
 
        
    AddInvoiceItem Forms![Case List]!FileNumber, "LabelPrint", "Postage Stamp", Nz(DLookup("Value", "StandardCharges", "ID=" & 1)), 76, False, False, False, True
   
    rstLabelData.Close
    Set rstLabelData = Nothing
    
End Sub


Private Sub Form_Current()

'Dim sql As String
'Dim rstFormData As Recordset

'sql = "SELECT CaseList.FileNumber, CaseList.CaseTypeID, CaseList.Active, fcdetails.Current"
'sql = sql + " FROM (fcdetails INNER JOIN CaseList ON fcdetails.FileNumber = CaseList.FileNumber) INNER JOIN ClientList ON CaseList.ClientID = ClientList.ClientID"
'sql = sql + " WHERE CaseList.Active=True AND fcdetails.Current=True AND  CaseList.FileNumber =" & Forms![Case List]!FileNumber & ";"

'Set rstFormData = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
'    If rstFormData!CaseTypeID = 1 Then Me.btnClerk.Enabled = True
'    If rstFormData!CaseTypeID = 2 Then Me.btnBankruptcy.Enabled = True
'    If rstFormData!CaseTypeID = 7 Then Me.btnDistrict.Enabled = True


'rstFormData.Close
'Set rstFormData = Nothing

End Sub

