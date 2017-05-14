VERSION 5.00
Begin VB.Form b8 
   BackColor       =   &H00C0E0FF&
   Caption         =   "创新玩法-欺骗模式"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   4440
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "规则说明"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始游戏"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   1695
      Begin VB.OptionButton o4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "1000"
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
         Left            =   720
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton o3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "100"
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.OptionButton o2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "30%"
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
      Left            =   1680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton o1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "15%"
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
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
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
      TabIndex        =   12
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3000
      Picture         =   "b8.frx":0000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
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
      Left            =   2760
      TabIndex        =   10
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "你猜的次数"
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
      Left            =   840
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "难度："
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
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "b8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim min As Integer, max As Integer, p As Single '最小值，最大值，误报概率
Dim x As Integer, guess As Integer '你要猜的数，你猜的数
Dim ptemp As Single '用来与p比较


Private Sub Command1_Click()
If o1.Value = True Then
    p = 0.15
Else
    p = 0.3
End If
If o3.Value = True Then
    max = 100
    x = Int(Rnd * 99 + 1)
    Text1.MaxLength = 2
Else
    max = 1000
    x = Int(Rnd * 999 + 1)
    Text1.MaxLength = 3
End If
min = 0
Command1.Enabled = False
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Shell "notepad.exe " & App.Path & "\data\欺骗模式(帮助)", 1
End Sub

Private Sub Command3_Click()
Unload Me
b.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
  End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then '13是回车
    guess = Val(Text1)
    Text1 = ""
    If guess >= max Or guess <= min Then
        MsgBox "输入数字超出范围！", vbInformation, "提示"
        Exit Sub
    Else
        Label3 = Val(Label3) + 1
    End If
    ptemp = Rnd
    If guess = x Then
        MsgBox "！！！！！你猜中了！！！！！" & vbCrLf & "共猜了" & Label3.Caption & "次……", , "我们很抱歉的告诉您"
        If o1.Value = True Then
            If o3.Value = True Then
                jf = jf + 1: MsgBox "你获得了1积分"
            Else
                jf = jf + 3: MsgBox "你获得了3积分"
            End If
        Else
            If o3.Value = True Then
                jf = jf + 2: MsgBox "你获得了2积分"
            Else
                jf = jf + 6: MsgBox "你获得了6积分"
            End If
        End If
        savedj
        Command1.Enabled = True
        Text1.Enabled = False
        Label3 = 0
    ElseIf guess > x Then
        If ptemp < p Then
            MsgBox "太小了！", , "提示"
        Else
            MsgBox "太大了！", , "提示"
        End If
    Else
        If ptemp < p Then
            MsgBox "太大了！", , "提示"
        Else
            MsgBox "太小了！", , "提示"
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Label1 = Label1 & vbCrLf & "范围："
End Sub


