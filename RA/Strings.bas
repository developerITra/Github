Attribute VB_Name = "Strings"
Option Compare Database
Option Explicit
Public Function FindDelims(sInp As String, sDelim As String) As Variant
    ' Return either -1, or array of posiotions (IO=1) of sDelim in sInp
    Dim n, nTot, nLth, n0, nWhere() As Integer

    ' Find the delims:
    nLth = Len(sDelim)
    n = 1 - nLth
    nTot = 0
    Do
        n = InStr(n + nLth, sInp, sDelim)
        If 0 = n Then Exit Do
        nTot = nTot + 1
        ReDim Preserve nWhere(nTot - 1)
        nWhere(nTot - 1) = n
    Loop
    If 0 = nTot Then FindDelims = -1 Else FindDelims = nWhere
    
End Function

Public Function PrePendArray(vInp As Variant, ParamArray vX() As Variant) As Variant
    Dim v, vRes() As Variant
    Dim n, nL, nU, nL1, nU1 As Integer
    If Not IsArray(vInp) _
    Then
        v = vInp
        ReDim vInp(0)
        vInp(0) = v
    End If
    
    If IsMissing(vX) _
    Then
        vRes = vInp
    Else
        nL = LBound(vInp)
        nU = UBound(vInp)
        nL1 = LBound(vX)
        nU1 = UBound(vX)
        ReDim vRes(nU + nU1 - nL - nL1 + 1)
        For n = nL1 To nU1
            vRes(n) = vX(n)
        Next n
        For n = nL To nU
            vRes(n + nU1 + 1) = vInp(n)
        Next n
    End If
    PrePendArray = vRes
End Function

Public Function AppendArray(vInp As Variant, ParamArray vX() As Variant) As Variant
    Dim v, vRes() As Variant
    Dim n, nL, nU, nL1, nU1, lth As Integer
    If Not IsArray(vInp) _
    Then
        v = vInp
        ReDim vInp(0)
        vInp(0) = v
    End If
    
    If IsMissing(vX) _
    Then
        vRes = vInp
    Else
        nL = LBound(vInp)
        nU = UBound(vInp)
        nL1 = LBound(vX)
        nU1 = UBound(vX)
        ReDim vRes(nU + nU1 - nL - nL1 + 1)
        For n = nL1 To nU1
            vRes(n + nU + 1) = vX(n)
        Next n
        For n = nL To nU
            vRes(n) = vInp(n)
        Next n
    End If
    AppendArray = vRes
End Function

Public Function RemoveEmpties(sInp As String, ParamArray empties() As Variant) As String
    ' Remove empties from vbNewLine-delimited string.
    Dim bMT As Boolean
    Dim bRmvCt, i, J, lNL, lth As Integer
    Dim Removes(), s, sResult As String
    Dim v As Variant

    If 0 = Len(sInp) _
    Then
        RemoveEmpties = sInp
        Exit Function
    End If
    
    v = FindDelims(sInp, vbNewLine)
    If Not IsArray(v) _
    Then
        RemoveEmpties = sInp
        Exit Function
    End If
    
    bRmvCt = 0

    
    If IsMissing(empties) _
    Then
        bMT = True
    Else
        For i = LBound(empties) To UBound(empties)
            s = CStr(empties(i))
            If 0 = Len(s) _
            Then
                bMT = True
            Else
                bRmvCt = 1 + bRmvCt
                ReDim Preserve Removes(bRmvCt)
                Removes(bRmvCt) = s
            End If
        Next i
    End If
    
    lNL = Len(vbNewLine)
    v = PrePendArray(v, 1 - lNL)
    v = AppendArray(v, Len(sInp) + 1)
    
    sResult = ""
    For i = LBound(v) To UBound(v) - 1
        lth = v(i + 1) - v(i) - lNL
        If 0 = lth _
        Then
            If bMT Then GoTo RemoveEmpties_Skip
        End If
        s = Mid(sInp, v(i) + lNL, lth)

        If i < UBound(v) - 1 Then s = s & vbNewLine
        sResult = sResult & s
RemoveEmpties_Skip:
    Next i
        
    RemoveEmpties = sResult
    
End Function

Public Function SplitIntoLines(sInp As String, Optional nSkip As Integer = 0, Optional sDelim As String = ", ") As String
    ' Break a string into multiple lineas at appearances of sDelim.
    ' nSkip can be negative, meaning ignore this many at end (specifically for City, State ).
    ' DaveW 2012.02.17
    '  Debug.Print SplitIntoLines("asdf,qwer,tyui")
    'asdf
    'qwer
    'tyui
    '  debug.print SplitIntoLines("asdf,qwer,tyui",-1)
    'asdf
    'qwer,tyui

    
    Dim s, sRes As String
    Dim n, nTot, nLth, n0, nWhere() As Integer

    ' Find the delims:
    nLth = Len(sDelim)
    n = 1 - nLth
    nTot = 0
    Do
        n = InStr(n + nLth, sInp, sDelim)
        If 0 = n Then Exit Do
        nTot = nTot + 1
        ReDim Preserve nWhere(nTot)
        nWhere(nTot - 1) = n
    Loop
    
    If (0 = nTot) Or (Abs(nSkip) >= nTot) _
    Then
        SplitIntoLines = sInp
        Exit Function
    ElseIf 0 > nSkip _
    Then
        nTot = nTot + nSkip ' Really subtracting Abs(nSkip)
        ReDim Preserve nWhere(nTot)
    ElseIf 0 < nSkip _
    Then
        For n = 0 To nTot - nSkip - 1
            nWhere(n) = nWhere(n + nSkip)
        Next n
        nTot = nTot - nSkip
        ReDim Preserve nWhere(nTot)
    End If
    ' Pretend that there is a delimiter just beyond the end of the string:
    ReDim Preserve nWhere(nTot + 1)
    nWhere(nTot) = Len(sInp) + 1
    
   
    sRes = ""
    ' Allow for leading delim:
    If 1 < nWhere(0) _
    Then
        n0 = 1
        sRes = Left(sInp, nWhere(0) - 1)
    Else
        n0 = 2
        sRes = ""
        'W = 0
    End If
    For n = 1 To nTot
        s = Mid(sInp, nWhere(n - 1) + nLth, nWhere(n) - nWhere(n - 1) - nLth)
        sRes = sRes & vbNewLine & s
        'Debug.Print N & " " & nW & " " & nWhere(N) & sRes
        'nW = nWhere(N)
    Next n
    SplitIntoLines = sRes
   
End Function

Public Sub DumpArray(vArr As Variant)
    Dim n, nL, nU As Integer
    Dim s As String
        
    If Not IsArray(vArr) _
    Then
        Debug.Print "Single value: !" & vArr & "!"
    Else
        nL = LBound(vArr)
        nU = UBound(vArr)
        Debug.Print "From " & nL & " to " & nU & ":"
        s = "!"
        For n = nL To nU
            s = s & vArr(n) & "!"
        Next n
        Debug.Print s
    End If
End Sub
Private Sub testme()
    Dim w(1) As Integer
    w(0) = 2
    w(1) = 4
    DumpArray (AppendArray(w, -1, -2))
    
End Sub
