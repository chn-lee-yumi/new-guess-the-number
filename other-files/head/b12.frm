VERSION 5.00
Begin VB.Form b12 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "创新玩法-死亡模式-死亡大乱斗-"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6105
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "确认"
      Default         =   -1  'True
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
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
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
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "道具"
      Height          =   1215
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   3135
      Begin VB.CommandButton Command6 
         Caption         =   "Bomb"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pass"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Key"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "道具说明"
         Height          =   855
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "春哥药"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label labcy 
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
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.Label labbomb 
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
         TabIndex        =   16
         Top             =   720
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
         TabIndex        =   6
         Top             =   240
         Width           =   495
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
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label lab3 
      Caption         =   "正在玩的：                                             你、AI1、AI2、AI3、AI4、AI5、AI6、AI7、AI8、AI9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   14
      Top             =   480
      Width           =   6135
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3240
      Picture         =   "b12.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   735
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
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
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
      TabIndex        =   11
      Top             =   0
      Width           =   6135
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
      TabIndex        =   10
      Top             =   1200
      Width           =   615
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
      TabIndex        =   9
      Top             =   1200
      Width           =   615
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
      TabIndex        =   13
      Top             =   1200
      Width           =   6135
   End
End
Attribute VB_Name = "b12"
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
Dim p(1 To 9) As Boolean
Dim ABC As Integer
Dim judge As Integer
Dim def As Integer
Private Sub Command1_Click()
'正在玩的：                                             你、AI1、AI2、AI3、AI4、AI5、AI6、AI7、AI8、AI9
    On Error Resume Next
    Cls '清除开G痕迹
    If t1 = "ke" Then t1 = "": GoTo key
    guess = Val(t1)
    If t1 = "pa" Then t1 = "": GoTo pass
    t1 = ""
    '判断自己猜的数
    If guess <= min Or guess >= max Then
    MsgBox "你所选的数不在范围内！请重新输入。", 16, "错误"
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
            If djcy <> 0 Then
                judge = MsgBox("你是否使用春哥药？", 64 + vbYesNo)
                If judge = vbYes Then
                    djcy = djcy - 1
                    savedj
                    labcy = djcy & "个"
                    MsgBox "你复活了！", , "真垃圾~"
                    Exit Sub
                Else
                    GoTo lost
                End If
            End If
            GoTo lost
        Else
lost:       play "data\die.mp3": MsgBox "你输了！！！！hahahaha..." & vbCrLf & "点击确定退出。", 64, "消息"
            Unload Me
            b11.Show
            Exit Sub
        
        End If
    End If
pass:
For ABC = 1 To choose
If p(ABC) = True Then
    'AI猜数
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" & vbCrLf & "min max guess:" & min & max & guess '错误检测
    If guess > x Then
        MsgBox "        " & guess & ",太大了", , "人工智能" & ABC & "猜的数字为："
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "        " & guess & ",太小了", , "人工智能" & ABC & "猜的数字为："
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "人工智能" & ABC & "猜的数字是" & guess & "，猜中了！", 48, "很遗憾~~"
        p(ABC) = False
        play "data\goodbye.mp3"
        MsgBox "人工智能" & ABC & "死掉了！！！！出局！", 64, "很遗憾~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        judge = 0
        For def = 1 To 9
            If p(def) = False Then
                judge = judge + 1
            End If
        Next def
        If judge = 9 Then GoTo aidie
        MsgBox "请准备下一场比赛。", 64, "消息"
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 99 + 1)
        GoTo aidie
    End If
End If
Next ABC
t1 = ""
t1.SetFocus
Exit Sub
key:
MsgBox "电脑产生的随机数是：" & x
Exit Sub
aidie:
lab3.Caption = "正在玩的：                                             你"
For ABC = 1 To 9
If p(ABC) = True Then
lab3 = lab3 & "、AI" & ABC
End If
Next ABC
'If p(1) = False And p(2) = False And p(3) = False And p(4) = False And p(5) = False Then
If lab3 = "正在玩的：                                             你" Then
    If choose = 3 Then
    djpass = djpass + 1
    MsgBox "恭喜你获得一个Pass道具" & vbCrLf & "该道具可在死亡模式使用。"
    ElseIf choose = 5 Then
    djkey = djkey + 1
    MsgBox "恭喜你获得一个Key道具" & vbCrLf & "该道具可在死亡模式使用。"
    ElseIf choose = 9 Then
    djbomb = djbomb + 1
    MsgBox "恭喜你获得一个Bomb道具" & vbCrLf & "该道具可在死亡模式使用。"
    End If
    savedj
MsgBox "很遗憾，你赢了。QAQ" & vbCrLf & "点击确定退出。", 48, "消息"
Unload Me
b11.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
b11.Show
End Sub

Private Sub Command3_Click()
If djpass = 0 Then MsgBox "你没有该道具，获取方法请看说明。", , "消息": Exit Sub
djpass = djpass - 1
savedj
labpass = Val(labpass) - 1 & "个"
t1 = "pa"
Command1_Click
End Sub

Private Sub Command4_Click()
If djkey = 0 Then MsgBox "你没有该道具，获取方法请看说明。", , "消息": Exit Sub
djkey = djkey - 1
savedj
labkey = Val(labkey) - 1 & "个"
t1 = "ke"
Command1_Click
End Sub

Private Sub Command5_Click()
Shell "notepad.exe " & App.Path & "\data\道具说明(帮助)", 1
End Sub

Private Sub Command6_Click()
If djbomb > 0 Then
    djbomb = djbomb - 1
    savedj
    labbomb = djbomb
    For ABC = 1 To 9
        If p(ABC) = True Then
            p(ABC) = False
            '=======================================处理AI被炸死
            play "data\goodbye.mp3"
            MsgBox "AI" & ABC & "被玩家炸死！", , "真贱- -"
            lab3.Caption = "正在玩的：                                             你"
            For judge = 1 To choose
                If p(judge) = True Then
                    lab3 = lab3 & "、AI" & judge
                End If
            Next judge
            Exit For
            '=======================================处理完毕
        End If
    Next ABC
    If lab3 = "正在玩的：                                             你" Then
        If choose = 3 Then
            djpass = djpass + 1
            MsgBox "恭喜你获得一个Pass道具" & vbCrLf & "该道具可在死亡模式使用。"
        ElseIf choose = 5 Then
            djkey = djkey + 1
            MsgBox "恭喜你获得一个Key道具" & vbCrLf & "该道具可在死亡模式使用。"
        ElseIf choose = 9 Then
            djbomb = djbomb + 1
            MsgBox "恭喜你获得一个Bomb道具" & vbCrLf & "该道具可在死亡模式使用。"
        End If
    savedj
    play "data\laugh1.mp3"
    MsgBox "很遗憾，你赢了。QAQ" & vbCrLf & "点击确定退出。", 48, "消息"
    Unload Me
    b1.Show
    End If
Else
    MsgBox "你没有该道具，获取方法请看说明。", , "消息"
End If
End Sub

Private Sub Form_DblClick()
Print x '开G
End Sub
Private Sub Form_Load()
min = 0
max = 100
ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
x = Int(Rnd * 99 + 1)
labpass = djpass & "个"
labkey = djkey & "个"
labbomb = djbomb & "个"
labcy = djcy & "个"
For ABC = 1 To choose
    p(ABC) = True
Next ABC
If choose = 3 Then
    Me.Caption = Me.Caption & "简单"
    lab3 = "正在玩的：                                             你、AI1、AI2、AI3"
ElseIf choose = 5 Then
    Me.Caption = Me.Caption & "普通"
    lab3 = "正在玩的：                                             你、AI1、AI2、AI3、AI4、AI5"
ElseIf choose = 9 Then
    Me.Caption = Me.Caption & "困难"
    lab3 = "正在玩的：                                             你、AI1、AI2、AI3、AI4、AI5、AI6、AI7、AI8、AI9"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
savedj
For ABC = 1 To 9
p(ABC) = False
Next ABC
End Sub

Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click
End Sub
Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
  End If
End Sub

