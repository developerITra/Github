VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_IRS Notice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_NoData(Cancel As Integer)
MsgBox "Unable to print IRS Notice.  Make sure you have specified an IRS address for the Jurisdiction.", vbCritical
Cancel = 1
End Sub
