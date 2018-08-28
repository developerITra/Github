VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DateChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Open(Cancel As Integer)
    Dim s As String
    Me.Calendar0.SetFocus
    Me.Calendar0.Value = Utility.dtDateChooser
    s = Utility.strDateChooserCpt
    Me.Caption = s
    If 0 < Len(Utility.strDateChooserTxt) Then s = Utility.strDateChooserTxt
    Me.txtDateChooser.Visible = (s <> " ")
    Me.txtDateChooser.Value = s
End Sub

Private Sub cmdBack_Click()
    Utility.dtDateChooser = 0
    Me.Visible = False
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdNext_Click()
    Dim dt As Date
    dt = Me.Calendar0.Value
    Utility.dtDateChooser = dt
    Me.Visible = False
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Public Sub dtDateChooserSet(dt As Date)
    Utility.dtDateChooser = dt
    'Me.Calendar0.SetFocus
    'Me.Calendar0.Value = dt
    'On Error Resume Next
    'Me.Calendar0.Refresh
    'Resume Next
End Sub

Public Function dtDateChooserGet() As Date
    dtDateChooserGet = Utility.dtDateChooser
End Function

Public Sub strDateChooserCaptionSet(str As String)
    Utility.strDateChooserCpt = str
    'Me.SetFocus
    'Me.Caption = str
    'On Error Resume Next
    'Me.Refresh
    'Resume Next
End Sub

Public Sub strDateChooserTextSet(str As String)
    Utility.strDateChooserTxt = str
    'Me.txtDateChooser.SetFocus
    'Me.txtDateChooser.Value = str
    'On Error Resume Next
    'Me.Refresh
    'Resume Next
End Sub

