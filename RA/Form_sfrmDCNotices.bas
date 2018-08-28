VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmDCNotices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub NoticePosted_AfterUpdate()
If Not IsNull(NoticePosted) Then
 AddStatus FileNumber, NoticePosted, "Notice Posted"
Else
 AddStatus FileNumber, Now(), "Removed Notice Posted date"
End If
End Sub

Private Sub NoticePosted_DblClick(Cancel As Integer)
NoticePosted = Now()
Call NoticePosted_AfterUpdate
End Sub
