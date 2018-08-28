VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Workflow FNMA FC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim NoData As Boolean

Private Sub Report_NoData(Cancel As Integer)
NoData = True
End Sub

Private Function GetRecordCount() As String
If NoData Then
    GetRecordCount = "No files"
Else
    If txtRC = 1 Then
        GetRecordCount = "1 file"
    Else
        GetRecordCount = txtRC & " files"
    End If
End If
End Function

Private Sub Report_Open(Cancel As Integer)

'Dim rsMyRs As Recordset
'Dim strCriteria As String

'Set rsMyRs = CurrentDb.OpenRecordset("rqryFNMAFCInventory New", dbOpenDynaset)
'strCriteria = rsMyRs![Servicer Name]

'If Not (rsMyRs.BOF And rsMyRs.EOF) Then
'    rsMyRs.MoveFirst
'Else
'    Exit Sub
'End If
'
'Do While Not rsMyRs.EOF
'    If rsMyRs.EOF = True Then
'        Exit Do
'    Else
'    End If
       ' rsMyRs.Edit
'        If strCriteria = "Bank  of America, N.A." Then
'            rsMyRs![Serviced Name] = "Bank of America Home Loans"
        '    rsMyRs.Update
'            rsMyRs.MoveNext
'        ElseIf strCriteria = "Branch Banking and Trust Company " Then
'            rsMyRs![Serviced Name] = "Branch Banking & Trust Company"
        '    rsMyRs.Update
 '           rsMyRs.MoveNext
  '      ElseIf strCriteria = "Carrington Mortgage Services, LLC" Then
   '         rsMyRs![Serviced Name] = "Carrington Mortgage Services"
         '   rsMyRs.Update
    '        rsMyRs.MoveNext
     '   ElseIf strCriteria = "Colonial Savings, F.A." Then
      '      rsMyRs![Serviced Name] = "Colonial Savings"
          '  rsMyRs.Update
       '     rsMyRs.MoveNext
       ' ElseIf strCriteria = "Dovenmuehle Mortgage, Inc." Then
        '    rsMyRs![Serviced Name] = "Dovenmuehle Mortgage Co."
           ' rsMyRs.Update
         '   rsMyRs.MoveNext
        'ElseIf strCriteria = "Green Tree Servicing  LLC" Then
         '   rsMyRs![Serviced Name] = "Green Tree Servicing, LLC."
            'rsMyRs.Update
         '   rsMyRs.MoveNext
        'ElseIf strCriteria = "JPMorgan Chase Bank, National Association" Then
         '   rsMyRs![Serviced Name] = "JPMorgan Chase Bank, N.A."
            'rsMyRs.Update
          '  rsMyRs.MoveNext
        'ElseIf strCriteria = "LoanCare, A Division of FNF Servicing, Inc." Then
         '   rsMyRs![Serviced Name] = "Loan Care Servicing Center, Inc."
           ' rsMyRs.Update
       '     rsMyRs.MoveNext
      '  ElseIf strCriteria = "PHH Mortgage Corporation" Then
      '      rsMyRs![Serviced Name] = "PHH Mortgage"
           ' rsMyRs.Update
      '      rsMyRs.MoveNext
      '  ElseIf strCriteria = "PNC Bank, NA" Then
      '      rsMyRs![Serviced Name] = "PNC Mortgage"
           ' rsMyRs.Update
       '     rsMyRs.MoveNext
       ' ElseIf strCriteria = "Provident Bank" Then
       '     rsMyRs![Serviced Name] = "Provident Funding Associates, L.P."
           ' rsMyRs.Update
       '     rsMyRs.MoveNext
      '  ElseIf strCriteria = "SanA-Wells Fargo Bank, N.A." Then
       '     rsMyRs![Serviced Name] = "Wells Fargo Bank, N.A."
           ' rsMyRs.Update
       '     rsMyRs.MoveNext
      '  ElseIf strCriteria = "Saxon Special Servicing - MSP" Then
      '      rsMyRs![Serviced Name] = "Saxon Mortgage Services, Inc."
          '  rsMyRs.Update
      '      rsMyRs.MoveNext
      '  ElseIf strCriteria = "SC-Wells Fargo Home Mortgage" Then
      '      rsMyRs![Serviced Name] = "Wells Fargo Bank, N.A."
          '  rsMyRs.Update
       '     rsMyRs.MoveNext
       ' ElseIf strCriteria = "Sovereign Bank, N.A." Then
       '     rsMyRs![Serviced Name] = "Sovereign Bank, FSB."
          '  rsMyRs.Update
      '      rsMyRs.MoveNext
     '   ElseIf strCriteria = "TX-Green Tree Servicing LLC" Then
      '      rsMyRs![Serviced Name] = "Green Tree Servicing, LLC."
           ' rsMyRs.Update
      '      rsMyRs.MoveNext
      '  Else
      '      rsMyRs![Serviced Name] = strCriteria
           ' rsMyRs.Update
      '      rsMyRs.MoveNext
      '  End If
'Loop

'Set rsMyRs = Nothing
'rsMyRs.Close

End Sub
