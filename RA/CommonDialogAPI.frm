VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonDialogAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private mstrFileName As String
Private mblnStatus As Boolean

Public Property Let GetName(strName As String)
    mstrFileName = strName
End Property

Public Property Get GetName() As String
    GetName = mstrFileName
End Property

Public Property Let GetStatus(blnStatus As Boolean)
    mblnStatus = blnStatus
End Property

Public Property Get GetStatus() As Boolean
    GetStatus = mblnStatus
End Property

Public Function OpenFileDialog(lngFormHwnd As Long, lngAppInstance As Long, strInitDir As String, strFileFilter As String) As Long

Dim OpenFile As OPENFILENAME
Dim X As Long

With OpenFile
    .lStructSize = Len(OpenFile)
    .hwndOwner = lngFormHwnd
    .hInstance = lngAppInstance
    .lpstrFilter = strFileFilter
    .nFilterIndex = 1
    .lpstrFile = String(257, 0)
    .nMaxFile = Len(OpenFile.lpstrFile) - 1
    .lpstrFileTitle = OpenFile.lpstrFile
    .nMaxFileTitle = OpenFile.nMaxFile
    .lpstrInitialDir = strInitDir
    .lpstrTitle = "Open File"
    .Flags = 0
End With

X = GetOpenFileName(OpenFile)
If X = 0 Then
    mstrFileName = "none"
    mblnStatus = False
Else
    mstrFileName = Trim(OpenFile.lpstrFile)
    mblnStatus = True
End If
End Function

Public Function SaveFileDialog(lngFormHwnd As Long, lngAppInstance As Long, strInitDir As String, strFileFilter As String) As Long

Dim SaveFile As OPENFILENAME
Dim X As Long
        
With SaveFile
    .lStructSize = Len(SaveFile)
    .hwndOwner = lngFormHwnd
    .hInstance = lngAppInstance
    .lpstrFilter = strFileFilter
    .nFilterIndex = 1
    .lpstrFile = String(257, 0)
    .nMaxFile = Len(SaveFile.lpstrFile) - 1
    .lpstrFileTitle = SaveFile.lpstrFile
    .nMaxFileTitle = SaveFile.nMaxFile
    .lpstrInitialDir = strInitDir
    .lpstrTitle = "Save File"
    .Flags = 0
End With

X = GetSaveFileName(SaveFile)
If X = 0 Then
    mstrFileName = "none"
    mblnStatus = False
Else
    mstrFileName = Trim(SaveFile.lpstrFile)
    mblnStatus = True
End If
End Function

