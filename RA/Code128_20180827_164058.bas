Attribute VB_Name = "Code128"
Option Compare Database
Option Explicit

Private Const StartA = "11010000100"
Private Const StartB = "11010010000"
Private Const StartC = "11010011100"
Private Const StopCode = "1100011101011"

Public Sub DrawBarcode(G As Report, Info As String, xPos As Single, yPos As Single, Optional Vertical As Boolean = False)
Dim Encoding As String, EncodingLen As Integer, i As Integer, X As Single, Y As Single
Const BarWidth = 20
Const BarHeight = 600

Encoding = Encode(Info)
EncodingLen = Len(Encoding)

X = xPos
Y = yPos

If Vertical Then
    For i = 1 To EncodingLen
        If Mid(Encoding, i, 1) = "1" Then
            G.Line (X, Y + (CSng(i) * BarWidth))-Step(CSng(BarHeight), CSng(BarWidth)), RGB(0, 0, 0), BF
        End If
        X = xPos
        'y = y + BarWidth
    Next i
Else
    For i = 1 To EncodingLen
        If Mid(Encoding, i, 1) = "1" Then
            G.Line (X + (CSng(i) * BarWidth), Y)-Step(CSng(BarWidth), CSng(BarHeight)), RGB(0, 0, 0), BF
        End If
        'x = x + BarWidth
        Y = yPos
    Next i
End If
End Sub

Private Function Encode(Info As String) As String
  Dim Total As Integer, Encoding As String, BarCodeCharacter As Integer, i As Integer
  
  Encoding = StartB
  Total = 104
  For i = 1 To Len(Info)
    BarCodeCharacter = Asc(Mid(Info, i, 1)) - 32
    Total = Total + (i * BarCodeCharacter)
    Encoding = Encoding + GetCharEncoding(BarCodeCharacter)
  Next
  
  Encoding = Encoding + GetCharEncoding(Total Mod 103) + StopCode
  Encode = Encoding
  
End Function

Private Function GetCharEncoding(CharIndex As Integer) As String
  If CharIndex < 25 Then
    Select Case CharIndex
    Case 0
      GetCharEncoding = "11011001100"
    Case 1
      GetCharEncoding = "11001101100"
    Case 2
      GetCharEncoding = "11001100110"
    Case 3
      GetCharEncoding = "10010011000"
    Case 4
      GetCharEncoding = "10010001100"
    Case 5
      GetCharEncoding = "10001001100"
    Case 6
      GetCharEncoding = "10011001000"
    Case 7
      GetCharEncoding = "10011000100"
    Case 8
      GetCharEncoding = "10001100100"
    Case 9
      GetCharEncoding = "11001001000"
    Case 10
      GetCharEncoding = "11001000100"
    Case 11
      GetCharEncoding = "11000100100"
    Case 12
      GetCharEncoding = "10110011100"
    Case 13
      GetCharEncoding = "10011011100"
    Case 14
      GetCharEncoding = "10011001110"
    Case 15
      GetCharEncoding = "10111001100"
    Case 16
      GetCharEncoding = "10011101100"
    Case 17
      GetCharEncoding = "10011100110"
    Case 18
      GetCharEncoding = "11001110010"
    Case 19
      GetCharEncoding = "11001011100"
    Case 20
      GetCharEncoding = "11001001110"
    Case 21
      GetCharEncoding = "11011100100"
    Case 22
      GetCharEncoding = "11001110100"
    Case 23
      GetCharEncoding = "11101101110"
    Case 24
      GetCharEncoding = "11101001100"
    End Select
  ElseIf CharIndex < 50 Then
    Select Case CharIndex
    Case 25
      GetCharEncoding = "11100101100"
    Case 26
      GetCharEncoding = "11100100110"
    Case 27
      GetCharEncoding = "11101100100"
    Case 28
      GetCharEncoding = "11100110100"
    Case 29
      GetCharEncoding = "11100110010"
    Case 30
      GetCharEncoding = "11011011000"
    Case 31
      GetCharEncoding = "11011000110"
    Case 32
      GetCharEncoding = "11000110110"
    Case 33
      GetCharEncoding = "10100011000"
    Case 34
      GetCharEncoding = "10001011000"
    Case 35
      GetCharEncoding = "10001000110"
    Case 36
      GetCharEncoding = "10110001000"
    Case 37
      GetCharEncoding = "10001101000"
    Case 38
      GetCharEncoding = "10001100010"
    Case 39
      GetCharEncoding = "11010001000"
    Case 40
      GetCharEncoding = "11000101000"
    Case 41
      GetCharEncoding = "11000100010"
    Case 42
      GetCharEncoding = "10110111000"
    Case 43
      GetCharEncoding = "10110001110"
    Case 44
      GetCharEncoding = "10001101110"
    Case 45
      GetCharEncoding = "10111011000"
    Case 46
      GetCharEncoding = "10111000110"
    Case 47
      GetCharEncoding = "10001110110"
    Case 48
      GetCharEncoding = "11101110110"
    Case 49
      GetCharEncoding = "11010001110"
    End Select
  ElseIf CharIndex < 75 Then
    Select Case CharIndex
    Case 50
      GetCharEncoding = "11000101110"
    Case 51
      GetCharEncoding = "11011101000"
    Case 52
      GetCharEncoding = "11011100010"
    Case 53
      GetCharEncoding = "11011101110"
    Case 54
      GetCharEncoding = "11101011000"
    Case 55
      GetCharEncoding = "11101000110"
    Case 56
      GetCharEncoding = "11100010110"
    Case 57
      GetCharEncoding = "11101101000"
    Case 58
      GetCharEncoding = "11101100010"
    Case 59
      GetCharEncoding = "11100011010"
    Case 60
      GetCharEncoding = "11101111010"
    Case 61
      GetCharEncoding = "11001000010"
    Case 62
      GetCharEncoding = "11110001010"
    Case 63
      GetCharEncoding = "10100110000"
    Case 64
      GetCharEncoding = "10100001100"
    Case 65
      GetCharEncoding = "10010110000"
    Case 66
      GetCharEncoding = "10010000110"
    Case 67
      GetCharEncoding = "10000101100"
    Case 68
      GetCharEncoding = "10000100110"
    Case 69
      GetCharEncoding = "10110010000"
    Case 70
      GetCharEncoding = "10110000100"
    Case 71
      GetCharEncoding = "10011010000"
    Case 72
      GetCharEncoding = "10011000010"
    Case 73
      GetCharEncoding = "10000110100"
    Case 74
      GetCharEncoding = "10000110010"
    End Select
  Else
    Select Case CharIndex
    Case 75
      GetCharEncoding = "11000010010"
    Case 76
      GetCharEncoding = "11001010000"
    Case 77
      GetCharEncoding = "11110111010"
    Case 78
      GetCharEncoding = "11000010100"
    Case 79
      GetCharEncoding = "10001111010"
    Case 80
      GetCharEncoding = "10100111100"
    Case 81
      GetCharEncoding = "10010111100"
    Case 82
      GetCharEncoding = "10010011110"
    Case 83
      GetCharEncoding = "10111100100"
    Case 84
      GetCharEncoding = "10011110100"
    Case 85
      GetCharEncoding = "10011110010"
    Case 86
      GetCharEncoding = "11110100100"
    Case 87
      GetCharEncoding = "11110010100"
    Case 88
      GetCharEncoding = "11110010010"
    Case 89
      GetCharEncoding = "11011011110"
    Case 90
      GetCharEncoding = "11011110110"
    Case 91
      GetCharEncoding = "11110110110"
    Case 92
      GetCharEncoding = "10101111000"
    Case 93
      GetCharEncoding = "10100011110"
    Case 94
      GetCharEncoding = "10001011110"
    Case 95
      GetCharEncoding = "10111101000"
    Case 96
      GetCharEncoding = "10111100010"
    Case 97
      GetCharEncoding = "11110101000"
    Case 98
      GetCharEncoding = "11110100010"
    Case 99
      GetCharEncoding = "10111011110"
    Case 100
      GetCharEncoding = "10111101110"
    Case 101
      GetCharEncoding = "11101011110"
    Case 102
      GetCharEncoding = "11110101110"
    End Select
  End If
  
End Function

