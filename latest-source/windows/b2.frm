VERSION 5.00
Begin VB.Form a24 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "传统玩法-大范围模式"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5085
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2280
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
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1560
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label l4 
      Alignment       =   2  'Center
      Caption         =   "1000000"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   600
      Width           =   1335
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "最高纪录："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   1335
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
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3360
      Picture         =   "b2.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "  猜一个数字，范围0-1000000"
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
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "范围：         ~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "a24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim min As Long
Dim max As Long
Dim t As Long
Dim x As Long
Dim guess As Long
Private Sub Command1_Click()
Dim jilu As String * 2
Cls '清除开G痕迹
If t1 = "-1" Then t1 = "": MsgBox "恭喜你打爆机！！！": Label4 = "-100": Exit Sub
guess = Val(t1)
t1 = ""
If guess <= min Or guess >= max Then
    MsgBox "你所选的数不在范围内！请重新输入。", 16, "错误"
End If
    t = t + 1
If guess > min And guess < max Then
    If guess > x Then
        MsgBox "恭喜你——猜错了,太大了", 64, "提示"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "恭喜你——猜错了,太小了", 64, "提示"
        min = guess
        l3 = guess
    ElseIf guess = x Then
        MsgBox "很遗憾,你猜对了!!!(你妹的RP真好!)" & vbCrLf & "你共猜了" & t & "次" & vbCrLf & "你获得了5积分", 64, "消息"
        jf = jf + 5
        savedj
        If Label4 = "\" Then Label4 = t: GoTo jiluline
        If t < Val(Label4) Then
            play "data\happy.mp3"
            MsgBox "你刷新了记录！！！", , "Good job!"
            Label4 = t
jiluline:   jilu = Label4
            Open App.Path & "\data\猜数字记录.txt" For Random As #2
                Put #2, 2, jilu
            Close
        End If
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
        x = Int(Rnd * 999999)
        t = 0
        l3 = 0
        l4 = 1000000
        MsgBox "点击确定重新开始玩。", 64, "消息"
    End If
End If
t1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
a2.Show
End Sub

Private Sub Form_DblClick()
Print x '开G
End Sub
Private Sub Form_Load()
Dim jilu1 As String * 2
Dim jilu2 As Integer
max = 1000000
ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran换种子
x = Int(Rnd * 999999 + 1)
Open App.Path & "\data\猜数字记录.txt" For Random As #2
    Get #2, 2, jilu1
    jilu2 = Val(Trim(jilu1))
    If jilu2 = 0 Then Label4 = "\" Else Label4 = jilu2
Close
End Sub
Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click '13是回车键
End Sub
Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
  End If
End Sub
