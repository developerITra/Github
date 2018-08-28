VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Military Affidavit Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Page()
Call FirmMargin(Me, FileNumber)
If page = 1 Then Call DrawBarcode(Me, AddDocPreIndex([FileNumber], 152, ""), 450, 7000, True)
End Sub
