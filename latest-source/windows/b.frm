VERSION 5.00
Begin VB.Form b 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "创新玩法"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7815
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command6 
      Caption         =   "  勇者之塔    (即将到来)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "欺骗模式"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "街机模式"
      Height          =   615
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "极限模式"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "一些道具的说明"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton c5 
      Caption         =   "格斗模式"
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton c2 
      Caption         =   "  挑战模式    (共4关)"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton c1 
      Caption         =   "死亡模式"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton c3 
      Caption         =   "RP模式"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "创意提供："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   3840
      TabIndex        =   11
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3840
      TabIndex        =   10
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub c1_Click()
b1.Show
Unload Me
End Sub
Private Sub c2_Click()
b2.Show
Unload Me
End Sub
Private Sub c3_Click()
Dim yesno As Integer
yesno = MsgBox("玩法：你要一步步的缩小范围，直到猜中数字。(数字包括范围两端)" & vbCrLf & "如果赢了，你会获得一个珍贵的道具；" & vbCrLf & "如果输了，你将会被扣10积分。" & vbcelf & "玩之前请确定你有10积分。", 48 + vbOKCancel, "你确定要玩吗？")
If yesno = vbOK Then
    If jf < 10 Then MsgBox "你没有10积分，所以你不能玩。": Exit Sub
    'Shell "shutdown -s -t 120"
    b3.Show
    Unload Me
End If
End Sub


Private Sub c5_Click()
b51.Show
Unload Me
End Sub

Private Sub Command1_Click()
Shell "notepad.exe " & App.Path & "\data\道具说明(帮助)", 1
End Sub

Private Sub Command2_Click()
Unload Me
main.Show
End Sub

Private Sub Command3_Click()
Unload Me
b6.Show
End Sub

Private Sub Command4_Click()
Unload Me
b7.Show
End Sub

Private Sub Command5_Click()
Unload Me
b8.Show
End Sub

Private Sub Form_Load()
Set b1 = Nothing
Set b2 = Nothing
Set b3 = Nothing
Set b4 = Nothing
Set b51 = Nothing
Set b6 = Nothing
Set b7 = Nothing
Set b8 = Nothing
Label1.Caption = "死亡实验室(AI实验室)：CHN-Lee-玉米" & vbCrLf & _
"挑战模式：CHN-Lee-玉米" & vbCrLf & _
"RP模式：CHN-Lee-玉米" & vbCrLf & _
"格斗模式：审判者Crj." & vbCrLf & _
"极限模式：人人都能想到吧？(划掉)" & vbCrLf & _
"街机模式：CHN-Lee-玉米" & vbCrLf & _
"欺骗模式：CHN-Lee-玉米" & vbCrLf & _
"勇者之塔：审判者Crj."
End Sub

