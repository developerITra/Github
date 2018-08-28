Attribute VB_Name = "SetPrinter"
Option Compare Database
Option Explicit

Declare Function GetProfileString Lib "kernel32" _
    Alias "GetProfileStringA" _
    (ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long) As Long

Declare Function WriteProfileString Lib "kernel32" _
    Alias "WriteProfileStringA" _
    (ByVal lpszSection As String, _
    ByVal lpszKeyName As String, _
    ByVal lpszString As String) As Long

Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lparam As String) As Long

Public Const HWND_BROADCAST = &HFFFF
Public Const WM_WININICHANGE = &H1A

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Declare Function GetVersionExA Lib "kernel32" _
    (lpVersionInformation As OSVERSIONINFO) As Integer

Public Declare Function OpenPrinter Lib "winspool.drv" _
    Alias "OpenPrinterA" _
    (ByVal pPrinterName As String, _
    phPrinter As Long, _
    pDefault As PRINTER_DEFAULTS) As Long
    
Public Declare Function SetPrinter Lib "winspool.drv" _
    Alias "SetPrinterA" _
    (ByVal hPrinter As Long, _
    ByVal Level As Long, _
    pPrinter As Any, _
    ByVal Command As Long) As Long

Public Declare Function GetPrinter Lib "winspool.drv" _
    Alias "GetPrinterA" _
    (ByVal hPrinter As Long, _
    ByVal Level As Long, _
    pPrinter As Any, _
    ByVal cbBuf As Long, _
    pcbNeeded As Long) As Long

Public Declare Function lstrcpy Lib "kernel32" _
    Alias "lstrcpyA" _
    (ByVal lpString1 As String, _
    ByVal lpString2 As Any) As Long

Public Declare Function ClosePrinter Lib "winspool.drv" _
    (ByVal hPrinter As Long) As Long

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

' constants for DesiredAccess member of PRINTER_DEFAULTS
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
    PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Public Const PRINTER_ATTRIBUTE_DEFAULT = 4

Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long        ' // Windows 95 only
    dmICMIntent As Long        ' // Windows 95 only
    dmMediaType As Long        ' // Windows 95 only
    dmDitherType As Long       ' // Windows 95 only
    dmReserved1 As Long        ' // Windows 95 only
    dmReserved2 As Long        ' // Windows 95 only
End Type

Public Type PRINTER_INFO_5
    pPrinterName As String
    pPortName As String
    Attributes As Long
    DeviceNotSelectedTimeout As Long
    TransmissionRetryTimeout As Long
End Type

Public Type PRINTER_DEFAULTS
    pDatatype As Long
    pDevMode As Long
    DesiredAccess As Long
End Type

Private Function PtrCtoVbString(Add As Long) As String
Dim sTemp As String * 512, X As Long

X = lstrcpy(sTemp, Add)
If (InStr(1, sTemp, Chr(0)) = 0) Then
     PtrCtoVbString = ""
  Else
     PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
End If
End Function
      
Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
Dim DeviceLine As String
Dim R As Long
Dim l As Long

DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
' Store the new printer information in the [WINDOWS] section of
' the WIN.INI file for the DEVICE= item
R = WriteProfileString("windows", "Device", DeviceLine)
' Cause all applications to reload the INI file:
l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub
      
Private Function Win95SetDefaultPrinter(PrinterName As String) As Boolean
Dim Handle As Long          'handle to printer
Dim pd As PRINTER_DEFAULTS
Dim X As Long
Dim need As Long            ' bytes needed
Dim pi5 As PRINTER_INFO_5   ' your PRINTER_INFO structure
Dim LastError As Long          ' determine which printer was selected

If PrinterName = "" Then
    Exit Function
End If
' set the PRINTER_DEFAULTS members
pd.pDatatype = 0&
pd.DesiredAccess = PRINTER_ALL_ACCESS
' Get a handle to the printer
X = OpenPrinter(PrinterName, Handle, pd)          ' failed the open
If X = False Then
    MsgBox "Unable to open printer " & PrinterName
    Win95SetDefaultPrinter = False
    Exit Function
End If
' Make an initial call to GetPrinter, requesting Level 5
' (PRINTER_INFO_5) information, to determine how many bytes
' you need
X = GetPrinter(Handle, 5, ByVal 0&, 0, need)
' don't want to check Err.LastDllError here - it's supposed
' to fail
' with a 122 - ERROR_INSUFFICIENT_BUFFER
' redim t as large as you need
ReDim t((need \ 4)) As Long
' and call GetPrinter for keepers this time
X = GetPrinter(Handle, 5, t(0), need, need)
' failed the GetPrinter
If X = False Then
    MsgBox "Can't get printer information for " & PrinterName
    Win95SetDefaultPrinter = False
    Exit Function
End If
' set the members of the pi5 structure for use with SetPrinter.
' PtrCtoVbString copies the memory pointed at by the two string
' pointers contained in the t() array into a Visual Basic string.
' The other three elements are just DWORDS (long integers) and
' don't require any conversion
pi5.pPrinterName = PtrCtoVbString(t(0))
pi5.pPortName = PtrCtoVbString(t(1))
pi5.Attributes = t(2)
pi5.DeviceNotSelectedTimeout = t(3)
pi5.TransmissionRetryTimeout = t(4)
' this is the critical flag that makes it the default printer
pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT
' call SetPrinter to set it
X = SetPrinter(Handle, 5, pi5, 0)
' failed the SetPrinter
If X = False Then
    MsgBox "SetPrinter failed, error code: " & Err.LastDllError
    Win95SetDefaultPrinter = False
    Exit Function
End If          ' and close the handle
ClosePrinter (Handle)
Win95SetDefaultPrinter = True
End Function
      
Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As String, _
    PrinterPort As String)
Dim iDriver As Integer
Dim iPort As Integer

DriverName = ""
PrinterPort = ""
'The driver name is first in the string terminated by a comma
iDriver = InStr(Buffer, ",")
If iDriver > 0 Then
    'Strip out the driver name
    DriverName = Left(Buffer, iDriver - 1)
    'The port name is the second entry after the driver name
    'separated by commas.
    iPort = InStr(iDriver + 1, Buffer, ",")
    If iPort > 0 Then                  'Strip out the port name
        PrinterPort = Mid(Buffer, iDriver + 1, iPort - iDriver - 1)
    End If
End If
End Sub

Private Function WinNTSetDefaultPrinter(PrinterName As String) As Boolean
Dim Buffer As String
Dim DeviceName As String
Dim DriverName As String
Dim PrinterPort As String
Dim R As Long

'Get the printer information from the WIN.INI file.
Buffer = Space(1024)
R = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))
'Parse the driver name and port name out of the buffer
GetDriverAndPort Buffer, DriverName, PrinterPort
If DriverName <> "" And PrinterPort <> "" Then
    SetDefaultPrinter PrinterName, DriverName, PrinterPort
End If
WinNTSetDefaultPrinter = True
End Function
      
Public Function SelectPrinter(PrinterName As String) As Boolean
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer

osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)
'MsgBox "MajorVersion=" & osinfo.dwMajorVersion & "   MinorVersion=" & osinfo.dwMinorVersion & "   PlatformID=" & osinfo.dwPlatformId, vbInformation

If osinfo.dwMajorVersion = 3 And osinfo.dwMinorVersion = 51 _
    And osinfo.dwPlatformId = 2 Then
    SelectPrinter = WinNTSetDefaultPrinter(PrinterName)
ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 _
    And osinfo.dwPlatformId = 1 _
    Then SelectPrinter = Win95SetDefaultPrinter(PrinterName)
ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 _
    And osinfo.dwPlatformId = 2 Then
    SelectPrinter = WinNTSetDefaultPrinter(PrinterName)
ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 10 Then  ' Win98
    SelectPrinter = Win95SetDefaultPrinter(PrinterName)
ElseIf osinfo.dwMajorVersion = 5 Then       ' win 2000
    SelectPrinter = WinNTSetDefaultPrinter(PrinterName)

End If
End Function

Public Function DefaultPrinter() As String
Dim Buffer As String
Dim R As Long

Buffer = Space(1024)
R = GetProfileString("windows", "device", "", Buffer, Len(Buffer))
DefaultPrinter = Left$(Buffer, InStr(1, Buffer, ",") - 1)
End Function

