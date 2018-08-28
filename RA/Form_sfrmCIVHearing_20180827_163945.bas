VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmCIVHearing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
  If (Nz(Hearing) <> Nz(Hearing.OldValue)) Then
    HearingCalendarEntryID = UpdateCalendar(Hearing.OldValue, Hearing, Nz(HearingCalendarEntryID))
  End If

End Sub


Private Sub Form_Current()
If IsNull(FileNumber) Then FileNumber = [Forms]![CivilDetails]![FileNumber]
End Sub

Private Sub Hearing_AfterUpdate()
 AddStatus FileNumber, Date, "Civil Hearing scheduled for " & Hearing
End Sub

Private Function UpdateCalendar(calendarDateOldValue As Variant, calendarDate As Variant, calendarID As String) As Variant

'UpdateCalendar = Null

'Exit Function

Dim emailGroup As String

UpdateCalendar = calendarID
' If existing date changed but we don't know the EntryID then user must update calendar manually
If (Not IsNull(calendarDateOldValue) And calendarID = "") Then
    MsgBox "Please update the Shared Calendar", vbExclamation
    Exit Function
End If

If (IsNull(calendarDate) And calendarID <> "") Then
    Call DeleteCalendarEvent(calendarID)
    UpdateCalendar = Null
    Exit Function
End If

Dim Location As String
If (IsNull(Forms![CivilDetails]![CourtID])) Then
  Location = "Unknown Civil Court Location"
Else
  Location = DLookup("[CourtName]", "CIV_Court", "[CourtID] = " & Forms![CivilDetails]![CourtID])
End If


emailGroup = "SharedCalRecipCivil"
If (calendarID = "") Then     ' new event on calendar
    UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, Forms![Case List]!PrimaryDefName, Location, 8, emailGroup)
Else                                    ' change existing event on calendar

   If (IsNull(calendarDateOldValue)) Then   ' new date
    UpdateCalendar = AddCalendarEvent(CDate(calendarDate), False, Forms![Case List]!PrimaryDefName, Location, 8, emailGroup)
      
   Else ' otherwise update calendar event
    Call UpdateCalendarEvent(calendarID, CDate(calendarDate), False, Forms![Case List]!PrimaryDefName, Location, 8)
   End If
End If
    
End Function

Private Sub Hearing_BeforeUpdate(Cancel As Integer)
If Not IsNull(Hearing) Then

    If HearingCheking(Hearing, 1) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(Hearing, 2) = 1 Then
    Cancel = 1
    End If
    If HearingCheking(Hearing, 3) = 1 Then
    Cancel = 1
    End If
End If
End Sub
