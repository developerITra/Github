VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmLisPendens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub LisPendensFiled_AfterUpdate()

If Not IsNull(LisPendensFiled) Then
 AddStatus FileNumber, LisPendensFiled, "Lis Pendens recorded on " & Format(LisPendensFiled, "mm/dd/yyyy")
Else
 AddStatus FileNumber, Now(), "Removed Lis Pendens filed date"
End If

End Sub

Private Sub LisPendensFiled_DblClick(Cancel As Integer)
If FileReadOnly Or EditDispute Then
   DoCmd.CancelEvent
Else
    LisPendensFiled = Now()
    Call LisPendensFiled_AfterUpdate
End If
End Sub
