VERSION 5.00
Begin VB.Form a2 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "传统玩法"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4950
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton c4 
      Caption         =   "推理模式"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "大范围模式"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "点此返回"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "单人+2AI"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "单人+1AI"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "单人"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "a2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer

Private Sub c4_Click()
b4.Show
Unload Me
End Sub

Private Sub Command1_Click()
a21.Show
Unload Me
End Sub

Private Sub Command2_Click()
a22.Show
Unload Me
End Sub

Private Sub Command3_Click()
a23.Show
Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
main.Show
End Sub

Private Sub Command5_Click()
Unload Me
a24.Show
End Sub

Private Sub Form_Load()
Set a21 = Nothing
Set a22 = Nothing
Set a23 = Nothing
Set a24 = Nothing
End Sub
