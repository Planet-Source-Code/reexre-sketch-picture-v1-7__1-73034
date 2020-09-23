Attribute VB_Name = "modNumAlpha"
Option Explicit

'Private Const KN = 52 'AAzz

Private Const KN = 62        '09AaBbZz

Public ALPHAname       As String

Public Function Num2CharO(N, Digits_1or2) As String
'Max=2703 'kn=52


    Dim V1             As Integer
    Dim V2             As Integer

    V1 = N \ KN
    V2 = N Mod KN

    If V1 >= 26 Then V1 = V1 + 6
    If V2 >= 26 Then V2 = V2 + 6

    If Digits_1or2 = 1 Then
        Num2CharO = Chr(65 + V2)
    End If
    If Digits_1or2 = 2 Then
        Num2CharO = Chr(65 + V1) & Chr(65 + V2)
    End If

End Function

Public Function Char2NumO(Str As String) As Integer
    Dim V1             As Integer
    Dim V2             As Integer

    V1 = 65
    V2 = Asc(Right$(Str, 1))
    If V2 >= 91 Then V2 = V2 - 6

    If Len(Str) > 1 Then
        V1 = Asc(Str)
        If V1 >= 91 Then V1 = V1 - 6
    End If

    Char2NumO = (V1 - 65) * KN + (V2 - 65)

End Function

Public Function Num2Char(N, Digits_1or2) As String
'Max=3843 'kn=62

    Dim V1             As Integer
    Dim V2             As Integer

    V1 = N \ KN
    V2 = N Mod KN

    If V1 >= 10 Then V1 = V1 + 7
    If V2 >= 10 Then V2 = V2 + 7

    If V1 >= 43 Then V1 = V1 + 6
    If V2 >= 43 Then V2 = V2 + 6


    If Digits_1or2 = 1 Then
        Num2Char = Chr(48 + V2)
    End If
    If Digits_1or2 = 2 Then
        Num2Char = Chr(48 + V1) & Chr(48 + V2)
    End If

End Function

Public Function Char2Num(Str As String) As Integer
    Dim V1             As Integer
    Dim V2             As Integer
    'Stop

    V1 = 48
    V2 = Asc(Right$(Str, 1))
    If V2 >= 58 Then V2 = V2 - 7
    If V2 >= 91 Then V2 = V2 - 6


    If Len(Str) > 1 Then
        V1 = Asc(Str)
        If V1 >= 58 Then V1 = V1 - 7
        If V1 >= 91 Then V1 = V1 - 6
    End If

    Char2Num = (V1 - 48) * KN + (V2 - 48)

End Function
Public Sub ALPHAnameFromSettings()

    Dim N              As Integer

    Open App.Path & "\Stroke.txt" For Input As 99
    Input #99, N
    ALPHAname = Num2Char(N, 2)
    Input #99, N
    ALPHAname = ALPHAname & Num2Char(N, 2)

    While Not (EOF(99))
        Input #99, N
        ALPHAname = ALPHAname & Num2Char(N, 1)
    Wend
    Close 99

    ALPHAname = ALPHAname & "-"

    Open App.Path & "\Gabor.txt" For Input As 99

    While Not (EOF(99))
        Input #99, N
        ALPHAname = ALPHAname & Num2Char(N, 2)
    Wend
    Close 99

End Sub

Public Sub SettingsFromALPHAname(S As String)
    Dim L1             As Integer
    Dim L2             As Integer
    Dim s1             As String
    Dim s2             As String
    'Stop

    L1 = Len(S) - 4 - 19
    L2 = Len(S) - 4 - 7

    s1 = Mid$(S, L1, 11)
    s2 = Mid$(S, L2, 8)

    Open App.Path & "\Stroke.txt" For Output As 99
    Print #99, Char2Num(Left$(s1, 2))
    Print #99, Char2Num(Mid$(s1, 3, 2))

    Print #99, Char2Num(Mid$(s1, 5, 1))
    Print #99, Char2Num(Mid$(s1, 6, 1))
    Print #99, Char2Num(Mid$(s1, 7, 1))
    Print #99, Char2Num(Mid$(s1, 8, 1))
    Print #99, Char2Num(Mid$(s1, 9, 1))
    Print #99, Char2Num(Mid$(s1, 10, 1))
    Print #99, Char2Num(Mid$(s1, 11, 1))

    Close 99

    Open App.Path & "\Gabor.txt" For Output As 99
    Print #99, Char2Num(Left$(s2, 2))
    Print #99, Char2Num(Mid$(s2, 3, 2))

    Print #99, Char2Num(Mid$(s2, 5, 2))
    Print #99, Char2Num(Mid$(s2, 7, 2))

    Close 99

End Sub
