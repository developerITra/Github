VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_OptionalValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdComplete_Click()
Dim FileNum As Long

FileNum = FileNumber

If btn1 = True Then
NOI = Null
chkClientSentNOI = False
End If
If btn2 = True Then
FairDebt = Null
End If
If btn3 = True Then
AccelerationLetter = Null
ClientSentAcceleration = False
End If

DoCmd.Close

Call ReleaseFile(FileNum)
Call RestartCallFromQueue(FileNum)

End Sub
