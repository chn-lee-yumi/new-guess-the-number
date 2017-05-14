VERSION 5.00
Begin VB.Form dela1 
   Caption         =   "原始版"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "原始版代码"
      Height          =   855
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "原始版(初二信息技术书上的“猜数字”游戏)[1~100]"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "dela1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "按Esc返回。"
a11.Show
Unload Me
End Sub

Private Sub Command2_Click()
a12.Show
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
a.Show
End Sub
