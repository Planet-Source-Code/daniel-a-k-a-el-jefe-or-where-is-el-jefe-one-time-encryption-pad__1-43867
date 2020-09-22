Attribute VB_Name = "modMain"
Option Explicit

Global SplicedPad(1 To 9, 1 To 20) As String
Global LoadedMessage(1 To 2) As String
Function GeneratePad() As String
Dim i, j, a, b, c, CurrNum, CurrKeyLen, CurrTextPos As Integer 'We Need These for numbering
Dim CurrKey, CurrLine As String 'These for forming each line

For i = 1 To 20
    CurrKey = "" 'reset Key
    CurrKeyLen = Rand(4, 6)
    CurrTextPos = Rand(2, 4)
    
        For b = 1 To CurrKeyLen
            If b = CurrTextPos Then
                CurrKey = CurrKey & Chr(Rand(97, 122))
            Else
                CurrKey = CurrKey & Rand(0, 9)
            End If
        Next b
    
    CurrLine = CurrKey 'Start the current line
    
    SplicedPad(1, i) = CurrKey 'Put in then array, no use splicing text file when we can do it here
    'i-1 because the array works from 0-19 (counts 0 too)
    For j = 1 To 9
        CurrKeyLen = Rand(4, 6)
        CurrTextPos = Rand(2, 4)
        CurrKey = "" 'Reset Key
            For c = 1 To CurrKeyLen
                If c = CurrTextPos Then
                    CurrKey = CurrKey & Chr(Rand(97, 122))
                Else
                    CurrKey = CurrKey & Rand(0, 9)
                End If
            Next c
        
        CurrLine = CurrLine & "," & CurrKey
        SplicedPad(j, i) = CurrKey 'continue to add
    Next j
    
    If i <> 20 Then
        CurrLine = CurrLine & "," & vbCrLf
    Else
        CurrLine = CurrLine & ","
    End If
    GeneratePad = GeneratePad & CurrLine
Next i
End Function

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
  Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Function DisplaySplicedPad() As String
'mimic the GeneratePad, just display instead of add...
If SplicedPad(1, 1) = Null Then
    DisplaySplicedPad = "0"
    Exit Function
End If

Dim i, j, CurrNum As Integer    'We Need These for numbering
Dim CurrKey, CurrLine As String 'These for forming each line

For i = 1 To 20
    CurrKey = SplicedPad(1, i)
    CurrLine = CurrKey
    For j = 1 To 9
        CurrKey = SplicedPad(j, i)
        CurrLine = CurrLine & "," & CurrKey
    Next j
    If i <> 20 Then
        CurrLine = CurrLine & "," & vbCrLf
    Else
        CurrLine = CurrLine & ","
    End If
    DisplaySplicedPad = DisplaySplicedPad & CurrLine
Next i
End Function

Function Encrypt(ByVal message As String) As String
Dim PulledOut, StringLeft As String
Dim i, j, k, l As Integer
'Begin with the insane
message = UCase(message)
k = Len(message)
For i = 1 To k
    PulledOut = Mid(message, i, 1)
    
    Select Case PulledOut
        Case "0"
            PulledOut = SplicedPad(9, 9)
        Case "1"
            PulledOut = SplicedPad(8, 9)
        Case "2"
            PulledOut = SplicedPad(7, 9)
        Case "3"
            PulledOut = SplicedPad(6, 9)
        Case "4"
            PulledOut = SplicedPad(5, 9)
        Case "5"
            PulledOut = SplicedPad(4, 9)
        Case "6"
            PulledOut = SplicedPad(3, 9)
        Case "7"
            PulledOut = SplicedPad(2, 9)
        Case "8"
            PulledOut = SplicedPad(1, 9)
        Case "9"
            PulledOut = SplicedPad(4, 20)
        Case "A"
            PulledOut = SplicedPad(1, 10)
        Case "B"
            PulledOut = SplicedPad(8, 15)
        Case "C"
            PulledOut = SplicedPad(1, 1)
        Case "D"
            PulledOut = SplicedPad(2, 1)
        Case "E"
            PulledOut = SplicedPad(1, 2)
        Case "F"
            PulledOut = SplicedPad(2, 2)
        Case "G"
            PulledOut = SplicedPad(3, 1)
        Case "H"
            PulledOut = SplicedPad(1, 3)
        Case "I"
            PulledOut = SplicedPad(3, 3)
        Case "J"
            PulledOut = SplicedPad(4, 1)
        Case "K"
            PulledOut = SplicedPad(1, 4)
        Case "L"
            PulledOut = SplicedPad(4, 4)
        Case "M"
            PulledOut = SplicedPad(5, 1)
        Case "N"
            PulledOut = SplicedPad(1, 5)
        Case "O"
            PulledOut = SplicedPad(5, 5)
        Case "P"
            PulledOut = SplicedPad(6, 1)
        Case "Q"
            PulledOut = SplicedPad(1, 6)
        Case "R"
            PulledOut = SplicedPad(6, 6)
        Case "S"
            PulledOut = SplicedPad(7, 1)
        Case "T"
            PulledOut = SplicedPad(1, 7)
        Case "U"
            PulledOut = SplicedPad(7, 7)
        Case "V"
            PulledOut = SplicedPad(8, 1)
        Case "W"
            PulledOut = SplicedPad(1, 8)
        Case "X"
            PulledOut = SplicedPad(8, 8)
        Case "Y"
            PulledOut = SplicedPad(9, 1)
        Case "Z"
            PulledOut = SplicedPad(1, 9)
        Case "."
            PulledOut = SplicedPad(1, 12)
        Case Chr(32)
            PulledOut = SplicedPad(8, 12)
        Case ","
            PulledOut = SplicedPad(9, 13)
    End Select

    Encrypt = Encrypt & PulledOut
    PulledOut = ""
Next i
End Function

Function Extracted(ByVal message As String, ByVal ToRemove As String) As String
Dim a, i As Integer
a = InStr(1, message, ToRemove, vbTextCompare)
End Function

Function Decrypt(ByVal message As String) As String
message = Replace(message, SplicedPad(9, 9), "0")
message = Replace(message, SplicedPad(8, 9), "1")
message = Replace(message, SplicedPad(7, 9), "2")
message = Replace(message, SplicedPad(6, 9), "3")
message = Replace(message, SplicedPad(5, 9), "4")
message = Replace(message, SplicedPad(4, 9), "5")
message = Replace(message, SplicedPad(3, 9), "6")
message = Replace(message, SplicedPad(2, 9), "7")
message = Replace(message, SplicedPad(1, 9), "8")
message = Replace(message, SplicedPad(4, 20), "9")
message = Replace(message, SplicedPad(1, 10), "A")
message = Replace(message, SplicedPad(8, 15), "B")
message = Replace(message, SplicedPad(1, 1), "C")
message = Replace(message, SplicedPad(2, 1), "D")
message = Replace(message, SplicedPad(1, 2), "E")
message = Replace(message, SplicedPad(2, 2), "F")
message = Replace(message, SplicedPad(3, 1), "G")
message = Replace(message, SplicedPad(1, 3), "H")
message = Replace(message, SplicedPad(3, 3), "I")
message = Replace(message, SplicedPad(4, 1), "J")
message = Replace(message, SplicedPad(1, 4), "K")
message = Replace(message, SplicedPad(4, 4), "L")
message = Replace(message, SplicedPad(5, 1), "M")
message = Replace(message, SplicedPad(1, 5), "N")
message = Replace(message, SplicedPad(5, 5), "O")
message = Replace(message, SplicedPad(6, 1), "P")
message = Replace(message, SplicedPad(1, 6), "Q")
message = Replace(message, SplicedPad(6, 6), "R")
message = Replace(message, SplicedPad(7, 1), "S")
message = Replace(message, SplicedPad(1, 7), "T")
message = Replace(message, SplicedPad(7, 7), "U")
message = Replace(message, SplicedPad(8, 1), "V")
message = Replace(message, SplicedPad(1, 8), "W")
message = Replace(message, SplicedPad(8, 8), "X")
message = Replace(message, SplicedPad(9, 1), "Y")
message = Replace(message, SplicedPad(1, 9), "Z")
message = Replace(message, SplicedPad(1, 12), ".")
message = Replace(message, SplicedPad(8, 12), Chr(32))
message = Replace(message, SplicedPad(9, 13), ",")
Decrypt = message
End Function

Function LoadPad(ByVal theFile As String) As Boolean
Dim fnum As Integer
Dim num_lines As Long
Dim i As Long

    Screen.MousePointer = vbHourglass
    DoEvents

    fnum = FreeFile
    On Error GoTo errorhandler
    Open theFile For Input As fnum
    i = 1
    Do While Not EOF(fnum)
        Input #fnum, SplicedPad(1, i), SplicedPad(2, i), SplicedPad(3, i), SplicedPad(4, 1), SplicedPad(5, i), SplicedPad(6, i), SplicedPad(7, i), SplicedPad(8, i), SplicedPad(9, i)
        i = i + 1
    Loop
    Close #fnum
    
    Screen.MousePointer = vbDefault
    LoadPad = True
    Exit Function
errorhandler:
    LoadPad = False
    Screen.MousePointer = vbDefault
End Function
