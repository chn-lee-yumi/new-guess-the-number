VERSION 5.00
Begin VB.Form die 
   BackColor       =   &H00FFFF00&
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   Picture         =   "a3.frx":0000
   ScaleHeight     =   3705
   ScaleWidth      =   6135
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "道具"
      Height          =   1215
      Left            =   1440
      TabIndex        =   9
      Top             =   2400
      Width           =   2655
      Begin VB.CommandButton Command5 
         Caption         =   "道具说明"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Key"
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pass"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label labkey 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label labpass 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox t1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "确认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label l4 
      Alignment       =   2  'Center
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label l3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "　　死亡模式：猜一个数字，谁猜中谁就死掉(出局)，你要让其他三个电脑一个个地出局。最后只剩下你自己。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "在此输入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3240
      Picture         =   "a3.frx":0342
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "范围：          ~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   6135
   End
   Begin VB.Label lab3 
      Caption         =   "正在玩的：                        你、电脑1、电脑2、电脑3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   6135
   End
End
Attribute VB_Name = "die"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim min As Integer
Dim max As Integer
Dim t As Long
Dim x As Integer
Dim guess As Integer
Dim p As Integer
Private Sub Command1_Click()
    On Error Resume Next
    Cls '清除开G痕迹
    If t1 = "ke" Then t1 = "": GoTo key
    guess = Val(t1)
    If t1 = "pa" Then t1 = "": GoTo pass
    t1 = ""
    '判断自己猜的数
    If guess <= min Or guess >= max Then
        MsgBox "你所选的数不在范围内！", 16, "错误"
        MsgBox "请重新输入", 64, "消息"
        Exit Sub
    End If
    If guess > min And guess < max Then
    If guess > x Then
        MsgBox "太大了", 64, ""
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", 64, ""
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "你猜中了~~", 64, "恭喜你！"
        MsgBox "你输了！！！！hahahaha..." & vbCrLf & "点击确定退出。", 64, "消息"
        Unload Me
        Exit Sub
    End If
    End If
pass:
Select Case p
Case 1 '自己+AI1+AI2+AI3
    'AI1猜数
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能1猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能1猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能1死掉了！！！！出局！", 64, "很遗憾~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "正在玩的：                        你、电脑2、电脑3"
        MsgBox "请准备下一场比赛。", 64, "消息"
        p = 5
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
    'AI2猜数
    guess = aiguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能2猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能2猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能2死掉了！！！！出局！", 64, "很遗憾~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "正在玩的：                        你、电脑1、电脑3"
        MsgBox "请准备下一场比赛。", 64, "消息"
        p = 7
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
    'AI3猜数
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能3猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能3猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能3死掉了！！！！出局！", 64, "很遗憾~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "正在玩的：                        你、电脑1、电脑2"
        MsgBox "请准备下一场比赛。", 64, "消息"
        p = 3
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
Case 2 '自己+AI1
    'AI1猜数
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能1猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能1猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能1死掉了！！！！出局！", 64, "很遗憾~~"
        lab3 = "正在玩的：                        你"
        MsgBox "很遗憾，你赢了。QAQ" & vbCrLf & "点击确定退出。", 48, "消息"
        Unload Me
    End If
Case 3 '自己+AI1+AI2
    'AI1猜数
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能1猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能1猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能1死掉了！！！！出局！", 64, "很遗憾~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "正在玩的：                        你、电脑2"
        MsgBox "请准备下一场比赛。", 64, "消息"
        p = 4
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
    'AI2猜数
    guess = aiguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能2猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能2猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能2死掉了！！！！出局！", 64, "很遗憾~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "正在玩的：                        你、电脑1"
        MsgBox "请准备下一场比赛。", 64, "消息"
        p = 2
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
Case 4 '自己+AI2
    'AI2猜数
    guess = aiguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能2猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能2猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能2死掉了！！！！出局！", 64, "很遗憾~~"
        lab3 = "正在玩的：                        你"
        MsgBox "很遗憾，你赢了。QAQ" & vbCrLf & "点击确定退出。", 48, "消息"
        Unload Me
    End If
Case 5 '自己+AI2+AI3
    'AI2猜数
    guess = aiguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能2猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能2猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能2死掉了！！！！出局！", 64, "很遗憾~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "正在玩的：                        你、电脑3"
        MsgBox "请准备下一场比赛。", 64, "消息"
        p = 6
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
    'AI3猜数
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能3猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能3猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能3死掉了！！！！出局！", 64, "很遗憾~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "正在玩的：                        你、电脑2"
        MsgBox "请准备下一场比赛。", 64, "消息"
        p = 4
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
Case 6 '自己+AI3
    'AI3猜数
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能3猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能3猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能3死掉了！！！！出局！", 64, "很遗憾~~"
        lab3 = "正在玩的：                        你"
        MsgBox "很遗憾，你赢了。QAQ" & vbCrLf & "点击确定退出。", 48, "消息"
        Unload Me
    End If
Case 7 '自己+AI1+AI3
    'AI1猜数
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能1猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能1猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能1死掉了！！！！出局！", 64, "很遗憾~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "正在玩的：                        你、电脑3"
        MsgBox "请准备下一场比赛。", 64, "消息"
        p = 6
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
    'AI3猜数
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
        MsgBox "           " & guess, , "人工智能3猜的数字为："
    If guess > x Then
        MsgBox "太大了", , "结果"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "太小了", , "结果"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能3猜中了！", 48, "很遗憾~~"
        MsgBox "人工智能3死掉了！！！！出局！", 64, "很遗憾~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "正在玩的：                        你、电脑1"
        MsgBox "请准备下一场比赛。", 64, "消息"
        p = 2
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
End Select
t1 = ""
t1.SetFocus
Exit Sub
key:
MsgBox "电脑产生的随机数是：" & x
End Sub

Private Sub Command2_Click()
Unload Me
b1.Show
End Sub

Private Sub Command3_Click()
If djpass = 0 Then MsgBox "你没有该道具，获取方法请看说明。", , "消息": Exit Sub
djpass = djpass - 1
labpass = Val(labpass) - 1 & "个"
t1 = "pa"
Command1_Click
End Sub

Private Sub Command4_Click()
If djkey = 0 Then MsgBox "你没有该道具，获取方法请看说明。", , "消息": Exit Sub
djkey = djkey - 1
labkey = Val(labkey) - 1 & "个"
t1 = "ke"
Command1_Click
End Sub

Private Sub Form_DblClick()
Print x '开G
End Sub
Private Sub Form_Load()
min = 0
max = 100
ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
x = Int(Rnd * 99 + 1)
p = 1
labpass = djpass & "个"
labkey = djkey & "个"
End Sub

Private Sub Form_Unload(Cancel As Integer)
savedj
b.Show
End Sub

Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click
End Sub
Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
    MsgBox "请输入数字"
  End If
End Sub
