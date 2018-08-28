VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LIMBO_DC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, "LIMBO_DC"
End Sub

Private Sub cmdRefresh_Click()

End Sub

Private Sub ComArchive_Click()


End Sub



Private Sub ComRefreshAll_Click()

End Sub

Private Sub ComSCRACancel_Click()



End Sub

Private Sub ComRule_Click()
DoCmd.OpenForm "QueueRules", , , "ruleid = " & 3
End Sub

Private Sub ComSendExcel_Click()
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputQuery, "LimboDCExcel", acFormatXLS, TemplatePath & "Limbo DC All colores.xlt", True
    DoCmd.SetWarnings True
End Sub

Private Sub Fnumber_AfterUpdate()
If Not IsNull(Fnumber) Then
Dim Value As String
Dim blnFound As Boolean
blnFound = False
Dim J As Integer
Dim A As Integer
Dim i As Integer



        For J = 0 To lstFiles.ListCount - 1
           Value = lstFiles.Column(0, J)
           If InStr(Value, Fnumber.Value) Then
                blnFound = True
                 A = J
                Me.lstFiles.Selected(A) = True
               
                
                For i = 0 To lstFilesY.ListCount - 1
                If lstFilesY.Selected(i) Then
                lstFilesY.Selected(i) = False
                End If
                Next i
               
                For i = 0 To lstFilesR.ListCount - 1
                If lstFilesR.Selected(i) Then
                lstFilesR.Selected(i) = False
                End If
                Next i
               
               
               
            Exit For
            End If
        Next J
        
            If Not blnFound Then
                J = 0
                A = 0
                
                For J = 0 To lstFilesY.ListCount - 1
                   Value = lstFilesY.Column(0, J)
                   If InStr(Value, Fnumber.Value) Then
                        blnFound = True
                         A = J
                        Me.lstFilesY.Selected(A) = True
                       
                         For i = 0 To lstFiles.ListCount - 1
                         If lstFiles.Selected(i) Then
                         lstFiles.Selected(i) = False
                         End If
                         Next i
                        
                         For i = 0 To lstFilesR.ListCount - 1
                         If lstFilesR.Selected(i) Then
                         lstFilesR.Selected(i) = False
                         End If
                         Next i
                       
                    Exit For
                    End If
                Next J
                
                If Not blnFound Then
         
         
                    J = 0
                    A = 0
                    
                    For J = 0 To lstFilesR.ListCount - 1
                       Value = lstFilesR.Column(0, J)
                       If InStr(Value, Fnumber.Value) Then
                            blnFound = True
                             A = J
                            Me.lstFilesR.Selected(A) = True
                            
                             For i = 0 To lstFiles.ListCount - 1
                             If lstFiles.Selected(i) Then
                             lstFiles.Selected(i) = False
                             End If
                             Next i
                            
                             For i = 0 To lstFilesY.ListCount - 1
                             If lstFilesY.Selected(i) Then
                             lstFilesY.Selected(i) = False
                             End If
                             Next i
                           
                           
                        Exit For
                        End If
                    Next J
                    
                    If Not blnFound Then
                    MsgBox ("File not in the queue.")
                    Fnumber.SetFocus
                    End If
                End If
            End If
            
                    
        
        End If


End Sub

Private Sub Fnumber_DblClick(Cancel As Integer)
Fnumber.Value = Null

End Sub



Private Sub Form_Current()

Dim rstqueue As Integer
Dim rstqueueY As Integer
Dim rstqueueR As Integer

rstqueue = DCount("filenumber", "LimboDC")
QueueCount = rstqueue

rstqueueY = DCount("filenumber", "LimboDC_Yellow")
QueueCountY = rstqueueY

rstqueueR = DCount("filenumber", "LimboDC_Red")
QueueCountR = rstqueueR



End Sub



Private Sub lstFiles_Click()
lstFiles.SetFocus
lstFilesR = Null
lstFilesY = Null
End Sub

Private Sub lstFiles_DblClick(Cancel As Integer)
AddToList (lstFiles)
Call Limbo_OpenWizard(Me.lstFiles, "Limbo_DC", "White")
End Sub

Private Sub lstFilesR_Click()
lstFilesR.SetFocus
lstFiles = Null
lstFilesY = Null
End Sub

Private Sub lstFilesR_DblClick(Cancel As Integer)
AddToList (lstFilesR)
Call Limbo_OpenWizard(Me.lstFilesR, "Limbo_DC", "Red")
End Sub

Private Sub lstFilesY_Click()
lstFilesY.SetFocus
lstFiles = Null
lstFilesR = Null
End Sub

Private Sub lstFilesY_DblClick(Cancel As Integer)
AddToList (lstFilesY)
Call Limbo_OpenWizard(Me.lstFilesY, "Limbo_DC", "Yellow")
End Sub
