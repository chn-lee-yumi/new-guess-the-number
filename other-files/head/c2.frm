VERSION 5.00
Begin VB.Form d2 
   Caption         =   "AI 测试"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4575
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "开始"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "为什么每次都是后面的AI猜中- -换了AI也一样？？"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label l1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "   猜一个数字，两个人工智能在比赛。看那个人工智能最先猜中。"
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
      Width           =   4575
   End
   Begin VB.Label l2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   840
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
      Top             =   1560
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
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "比分：normalAI:dieAI"
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
      TabIndex        =   7
      Top             =   480
      Width           =   4575
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
      TabIndex        =   8
      Top             =   1560
      Width           =   4575
   End
End
Attribute VB_Name = "d2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim min As Integer
Dim max As Integer
Dim x As Integer
Dim guess As Integer
Private Sub Command1_Click()
Cls '清除开G痕迹
'AI1=======================
guess = aiguess(min, max)
If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" '错误检测
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
    MsgBox "人工智能1猜对了！"
    l1 = Val(l1) + 1
    MsgBox "人工智能1得了一分！"
    min = 0
    max = 100
    l3 = "0"
    l4 = "100"
    MsgBox "如果要在比一次，请按开始。" & vbCrLf & "如果要退出，请关闭窗体。", 64, "消息"
    ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
    x = Int(Rnd * 99)
    Exit Sub
End If
'AI1 end===================
'AI2=======================
guess = dieguess(min, max)
If guess <= min Or guess >= max Then MsgBox "测试期间出现错误：" & vbCrLf & "人工智能猜的数字在范围外！" '错误检测
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
    MsgBox "人工智能2猜对了！"
    l2 = Val(l2) + 1
    MsgBox "人工智能2得了一分！"
    min = 0
    max = 100
    l3 = "0"
    l4 = "100"
    MsgBox "如果要在比一次，请按开始。" & vbCrLf & "如果要退出，请关闭窗体。", 64, "消息"
    ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
    x = Int(Rnd * 99)
    Exit Sub
End If
End Sub
Private Sub Command2_Click()
Unload Me
main.Show
End Sub
Private Sub Form_DblClick()
Print x '开G
End Sub
Private Sub Form_Load()
max = 100
ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
x = Int(Rnd * 99)
End Sub

