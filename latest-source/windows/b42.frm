VERSION 5.00
Begin VB.Form b62 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "创新玩法-极限模式-极限二"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5730
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "点击开始"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "注意：这只是一个之前写的VBS脚本！想赢请看脚本代码！（放在data文件夹）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "  单挑终极人工智能，死亡模式。"
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
      Width           =   5775
   End
End
Attribute VB_Name = "b62"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iguess123 As Integer
Dim aiguess123 As Integer
Dim min As Integer
Dim max As Integer
Dim x As Integer

Function bossdie(x, min, max, iguess)
bossdie = x + x - iguess123
If bossdie >= max Then
bossdie = (max - 1)
ElseIf bossdie <= min Then
bossdie = (min + 1)
End If
End Function

Sub aa()
On Error GoTo errline
    Do
iguess123 = InputBox("请输入数字（范围" & min & "~" & max & "）", "请猜数")
    Do While iguess123 >= max Or iguess123 <= min
MsgBox "你猜的数在范围外!" & vbcelf & "请重新输入。", 48, "(error0x000015" & x & ")"
iguess123 = InputBox("请输入数字（范围" & min & "~" & max & "）", "请猜数")
iguess123 = iguess123 + 0
    Loop

If iguess123 = x Then
MsgBox "你猜中了!", 48, "你输了!"
Exit Sub
ElseIf iguess123 < x Then
min = iguess123
MsgBox "太小了!"
ElseIf iguess123 > x Then
max = iguess123
MsgBox "太大了!"
End If

aiguess123 = bossdie(x, min, max, iguess)
MsgBox "终极人工智能猜的是" & aiguess123
If aiguess123 = x Then
MsgBox "终极人工智能猜中了!", 48, "你赢了!"
Exit Sub
ElseIf aiguess123 < x Then
min = aiguess123
MsgBox "太小了!"
ElseIf aiguess123 > x Then
max = aiguess123
MsgBox "太大了!"
End If
    Loop

errline:
MsgBox "程序因用户错误操作而停止。（抱歉。。。）"
End Sub

Private Sub Command1_Click()
Unload Me
b6.Show
MsgBox "欢迎挑战终极人工智能（死亡模式1vs1）!! ", , "欢迎使用本程序（vbs脚本= =）! "
Randomize (Minute(Time))
x = Int(Rnd * 99 + 1)
min = 0
max = 100
aa
End Sub

Private Sub Command2_Click()
Unload Me
b.Show
End Sub
