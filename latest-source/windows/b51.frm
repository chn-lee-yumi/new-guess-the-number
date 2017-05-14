VERSION 5.00
Begin VB.Form b51 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "创新玩法-格斗模式"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4470
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "猪"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "地狱坟(作者)"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "  地狱坟    暴走状态"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "请选择人工智能："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "b51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
choose = 1
Unload Me
b52.Show
End Sub

Private Sub Command2_Click()
choose = 2
Unload Me
b52.Show
End Sub

Private Sub Command3_Click()
choose = 3
Unload Me
b52.Show
End Sub

Private Sub Command5_Click()
Unload Me
b.Show
End Sub

Private Sub Form_Load()
Set b52 = Nothing
End Sub
