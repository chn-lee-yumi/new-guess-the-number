Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public ran As Variant
Public djpass As Integer
Public djkey As Integer
Public djbomb As Integer, djcy As Integer, djhhy As Integer, jf As Integer, choose As Integer
Public checkdata As Long

Public Sub play(ByVal musicpath As String)
 Call mciSendString("close OpenFile", vbNullString, 0, 0)
 Call mciSendString("open " & musicpath & " alias OpenFile type MPEGVideo", vbNullString, 0, 0)
 Call mciSendString("play OpenFile", vbNullString, 0, 0)
End Sub

Public Function dieguess(dmin As Integer, dmax As Integer) As Long
Dim judge As Integer
ran = Minute(Time): Randomize ran
judge = Int(Rnd * 9)
ran = Minute(Time): Randomize ran
If judge < 5 Then
    If dmax - dmin >= 60 Then
        dieguess = dmin + Int(Rnd * 5) + 4
    ElseIf dmax - dmin >= 40 And dmax - dmin < 60 Then
        dieguess = dmin + Int(Rnd * 3) + 3
    ElseIf dmax - dmin > 13 And dmax - dmin < 40 Then
        dieguess = dmin + Int(Rnd * 2) + 1
    ElseIf dmax - dmin <= 13 Then
        dieguess = dmin + 1
    End If
Else
    If dmax - dmin >= 60 Then
        dieguess = dmax - Int(Rnd * 5) - 4
    ElseIf dmax - dmin >= 40 And dmax - dmin < 60 Then
        dieguess = dmax - Int(Rnd * 3) - 3
    ElseIf dmax - dmin > 13 And dmax - dmin < 40 Then
        dieguess = dmax - Int(Rnd * 2) - 1
    ElseIf dmax - dmin <= 13 Then
        dieguess = dmax - 1
    End If
End If
End Function
Public Function aiguess(amin As Integer, amax As Integer) As Integer
ran = Minute(Time): Randomize ran
If amax - amin >= 50 Then
    aiguess = Int(Rnd * 11 + ((amax + amin) \ 2) - 4)
ElseIf amax - amin >= 25 And amax - amin < 50 Then
    aiguess = Int(Rnd * 7 + ((amax + amin) \ 2) - 3)
ElseIf amax - amin >= 13 And amax - amin < 25 Then
    aiguess = Int(Rnd * 5 + ((amax + amin) \ 2) - 2)
ElseIf amax - amin >= 6 And amax - amin < 13 Then
    aiguess = Int(Rnd * 3 + ((amax + amin) \ 2) - 1)
ElseIf amax - amin < 6 Then
    'If Int(Rnd * 101) > 50 Then: aiguess = amax - 1: Else aiguess = amin + 1
    aiguess = amax - 1
End If
End Function
Public Function longx() As Long
Dim judge As Integer
ran = Minute(Time): Randomize ran
judge = Int(Rnd * 9)
ran = Minute(Time): Randomize ran
If judge < 5 Then
longx = -(Int(Rnd * 2147483646 - 1))
ElseIf judge > 5 Then
longx = Int(Rnd * 2147483646 - 1)
Else
longx = Int(Rnd * 19999 - 10000)
End If
End Function

Public Sub savedj()
'On Error Resume Next
'Kill App.Path & "\data\猜数字数据.txt"
checkdata = djpass \ 2 + djkey * 2 + djbomb * 3 + djcy * djhhy + jf + 12 '数据验证，防止被恶意纂改。
Open App.Path & "\data\猜数字数据.txt" For Output As #1
Write #1, djpass, djkey, djbomb, djcy, djhhy, jf, checkdata
Close
End Sub

'以下为格斗模式AI猜数
Public Function fgguess(min As Integer, max As Integer) As Integer
If max - min > 3 Then
    fgguess = Int(Rnd * 3 + (min + max) \ 2 - 1)
Else
    fgguess = min + 1
End If
End Function

Public Function btdyfguess(min As Integer, max As Integer, t As Integer, x As Integer) As Integer
Dim judge
If t = 2 Then
    btdyfguess = x
Else
    judge = Rnd * 2
    If judge < 0.5 Then
        btdyfguess = x
    Else
        If max - min > 4 Then
            btdyfguess = Int(Rnd * 3 + (min + max) \ 2 - 1)
        Else
            btdyfguess = max - 1
        End If
    End If
End If
End Function

Public Function dyfguess(min As Integer, max As Integer, t As Integer, x As Integer) As Integer
Dim judge As Single
If t = 3 Then
    dyfguess = x
ElseIf t = 2 Then
    judge = Rnd * 2
    If judge < 1 Then
        dyfguess = x
    Else
        If max - min > 4 Then
            dyfguess = Int(Rnd * 3 + (min + max) \ 2 - 1)
        Else
            dyfguess = max - 1
        End If
    End If
Else
    dyfguess = Int(Rnd * 3 + (min + max) \ 2 - 1)
End If
End Function

Public Function Labaiguess(min As Integer, max As Integer, x As Integer, level As Integer) As Integer
Dim ABC As Single
ABC = Rnd
Select Case level
    Case -3
        If ABC > 0.5 Then
            Labaiguess = x
        Else
            Labaiguess = Round((max + min) / 2)
        End If
    Case -2
        If ABC > 0.75 Then
            Labaiguess = x
        Else
            Labaiguess = Round((max + min) / 2)
        End If
    Case -1
        If ABC > 0.875 Then
            Labaiguess = x
        Else
            Labaiguess = Round((max + min) / 2)
        End If
    Case 0
        Labaiguess = Round((max + min) / 2)
    Case 1
        Labaiguess = Round((max + min) / 2)
        If Labaiguess = x And ABC > 0.875 Then
            Labaiguess = Round((max + min) / 2) + 5 - Int(Rnd * 10 + 1)
        End If
    Case 2
        Labaiguess = Round((max + min) / 2)
        If Labaiguess = x And ABC > 0.75 And max - x > 6 And x - min > 6 Then
            Labaiguess = Round((max + min) / 2) + 5 - Int(Rnd * 10 + 1)
        End If
    Case 3
        Labaiguess = Round((max + min) / 2)
        If Labaiguess = x And ABC > 0.5 And max - x > 6 And x - min > 6 Then
            Labaiguess = Round((max + min) / 2) + 5 - Int(Rnd * 10 + 1)
        End If
    Case 4
        If max - x > 299 Then
            Labaiguess = x + Int(Rnd * 50 + 1)
        Else
            Labaiguess = Round((max + min) / 2)
        End If
        If Labaiguess = x And ABC > 0.25 And max - x > 6 And x - min > 6 Then
            Labaiguess = Round((max + min) / 2) + 5 - Int(Rnd * 10 + 1)
        End If
    Case 5
        If max - x > 99 Then
            Labaiguess = x + Int(Rnd * 30 + 1)
        Else
            Labaiguess = Round((max + min) / 2)
        End If
        If Labaiguess = x And max - x > 6 And x - min > 6 Then
            Labaiguess = Round((max + min) / 2) + 5 - Int(Rnd * 10 + 1)
        End If
End Select
End Function


