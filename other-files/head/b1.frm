VERSION 5.00
Begin VB.Form d3 
   BackColor       =   &H00FFFF00&
   Caption         =   "单挑死亡模式的AI，玩普通模式"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
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
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label l1 
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
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3120
      Picture         =   "b1.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
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
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "   猜一个数字，和一个人工智能一起玩。看谁最先猜中。"
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
      TabIndex        =   5
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label l2 
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
      Left            =   2520
      TabIndex        =   4
      Top             =   480
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
      Left            =   1680
      TabIndex        =   3
      Top             =   840
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
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "范围：       ~"
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
      TabIndex        =   9
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "比分： 你    ：   电脑"
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
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "d3"
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
Private Sub Command1_Click()
Cls '清除开G痕迹
guess = Val(t1)
t1 = ""
If guess <= min Or guess >= max Then
    MsgBox "你所选的数不在范围内！", 16, "错误"
    MsgBox "请重新输入", 64, "消息"
    Exit Sub
End If
If guess > min And guess < max Then
    If guess > x Then
        MsgBox "恭喜你——猜错了,太大了", 64, "提示"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "恭喜你——猜错了,太小了", 64, "提示"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "很遗憾,你猜对了!!!(你妹的RP真好!)", 64, "消息"
        l1 = Val(l1) + 1
        MsgBox "你得了一分！", 64, "很遗憾~"
        MsgBox "如果要在比一次，请按确定。" & vbCrLf & "如果要退出，请关闭窗体。", 64, "消息"
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        x = Int(Rnd * 99)
        Exit Sub
    End If
End If
guess = dieguess(min, max)
If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" '错误检测
MsgBox "           " & guess, , "人工智能猜的数字为："
If guess > x Then
    MsgBox "太大了", , "结果"
    max = guess
    l4 = max
ElseIf guess < x Then
    MsgBox "太小了", , "结果"
    min = guess
    l3 = min
ElseIf guess = x Then
    MsgBox "人工智能猜对了！", 48, "恭喜你"
    l2 = Val(l2) + 1
    MsgBox "人工智能得了一分！", 64, "很遗憾~"
    min = 0
    max = 100
    l3 = "0"
    l4 = "100"
    MsgBox "如果要在比一次，请按确定。" & vbCrLf & "如果要退出，请关闭窗体。", 64, "消息"
    ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
    x = Int(Rnd * 99)
End If
t1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
b.Show
End Sub

Private Sub Form_DblClick()
Print x '开G
End Sub
Private Sub Form_Load()
max = 100
ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
x = Int(Rnd * 99)
End Sub
Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click
End Sub
Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
    MsgBox "请输入数字"
  End If
End Sub
