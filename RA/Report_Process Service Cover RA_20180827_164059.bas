VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Process Service Cover RA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_NoData(Cancel As Integer)
MsgBox "Process Server cover letter cannot be completed.  Make sure you have selected a Jurisdiction.  And that a Process Server has been selected for that Jurisdiction.", vbCritical
Cancel = 1
End Sub
