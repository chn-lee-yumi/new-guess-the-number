VERSION 5.00
Begin VB.Form b3 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "创新玩法-RP模式（创意by:CHN-Lee-玉米）"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4935
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "确认选择"
      Height          =   375
      Left            =   3885
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1575
      Width           =   945
   End
   Begin VB.CommandButton c5 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton c3 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton c4 
      Caption         =   "46~90"
      Height          =   375
      Left            =   2730
      TabIndex        =   5
      Top             =   1575
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton c2 
      Caption         =   "1~45"
      Height          =   375
      Left            =   945
      TabIndex        =   4
      Top             =   1575
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton c1 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton o2 
      Caption         =   "Bomb"
      Height          =   180
      Left            =   3360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton o1 
      Caption         =   "春哥药"
      Height          =   180
      Left            =   2400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.Label labhhy 
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
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
      Left            =   4440
      TabIndex        =   12
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label labjh 
      BackColor       =   &H00C0FFFF&
      Caption         =   "5"
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
      Left            =   3000
      TabIndex        =   10
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "    现在，你要一步步的缩小范围，直到猜中数字。(数字包括范围两端)，如果你在机会用完前猜中数字，那么你会获得你想要的那个道具。"
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "请选择你想要的道具："
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "你剩余的机会(次数)： 后悔药："
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
      Left            =   -15
      TabIndex        =   9
      Top             =   2040
      Width           =   4935
   End
End
Attribute VB_Name = "b3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim floor As Integer
Dim x As Integer
Dim jh As Integer
Dim daoju As String
Private Sub cmdclick(cc)
Cls
Dim cmdmin As Integer
Dim cmdmax As Integer
Dim cmdlen As Integer
Dim yesno As Integer
cmdlen = Len(cc)
cmdmin = Val(cc)
If cmdlen = 3 Then
    cmdmax = Val(Right(cc, 1))
Else 'If cmdlen = 4 Or cmdlen = 5 Then
    cmdmax = Val(Right(cc, 2))
End If
If x < cmdmin Or x > cmdmax Then
    MsgBox "你选错了！", 48, "警告！"
    jh = jh - 1
    labjh = jh
    If jh = 0 Then
        yesno = MsgBox("你输了，是否使用后悔药？", 64 + 4, "提示")
        If yesno = 6 Then
            If djhhy = 0 Then MsgBox "你没有后悔药！", 48, "我擦！": GoTo lost
            djhhy = djhhy - 1
            MsgBox "你使用了后悔药。该游戏可悔，人生不可悔！", , "珍惜现在"
            jf = jf + 10
            savedj
            'Shell "shutdown -a"
            Unload Me
            b.Show
        Else
lost:
            Unload Me
            a21.Show
            a21.Command2.Visible = False
            a21.Label1.Visible = True
            Exit Sub
        End If
    End If
Else
    Select Case floor
    Case 1, 2
        c1.Visible = True
        c2.Visible = False
        c3.Visible = True
        c4.Visible = False
        c5.Visible = True
        c1.Caption = cmdmin & "~" & (cmdmin - 1 + (cmdmax - cmdmin + 1) / 3)
        c3.Caption = (cmdmin + (cmdmax - cmdmin + 1) / 3) & "~" & (cmdmin - 1 + (cmdmax - cmdmin + 1) / 3 * 2)
        c5.Caption = ((cmdmin + (cmdmax - cmdmin + 1) / 3 * 2)) & "~" & cmdmax
        floor = floor + 1
    Case 3
        c2.Visible = True
        c4.Visible = True
        c1.Caption = cmdmin
        c2.Caption = (cmdmin + 1)
        c3.Caption = (cmdmin + 2)
        c4.Caption = (cmdmin + 3)
        c5.Caption = cmdmax
        floor = floor + 1
    Case 4
        If o1.Value Then
        djcy = djcy + 1
        Else
        djbomb = djbomb + 1
        End If
        floor = floor + 1
        play "data\win.mp3"
        MsgBox "很遗憾，你赢了！" & vbCrLf & "你获得了一个“" & daoju & "”", vbInformation, "天哪！"
        savedj
        Unload Me
        b.Show
    End Select
End If
End Sub

Private Sub c1_Click()
cmdclick c1.Caption
End Sub

Private Sub c2_Click()
cmdclick c2.Caption
End Sub

Private Sub c3_Click()
cmdclick c3.Caption
End Sub

Private Sub c4_Click()
cmdclick c4.Caption
End Sub

Private Sub c5_Click()
cmdclick c5.Caption
End Sub

Private Sub Command2_Click()
If o1.Value = False And o2 = False Then
MsgBox "请选择你想要的道具！", 48
Exit Sub
End If
If o1.Value = True Then
daoju = "春哥药"
Else
daoju = "Bomb"
End If
Command2.Visible = False
c2.Visible = True
c4.Visible = True
o1.Enabled = False
o2.Enabled = False
End Sub

Private Sub Form_DblClick()
Print x
End Sub

Private Sub Form_Load()
ran = Minute(Time): Randomize ran
x = Int(Rnd * 90 + 1)
floor = 1
jh = 5
labhhy = djhhy
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> 1 Then
    jf = jf - 10
End If
End Sub
