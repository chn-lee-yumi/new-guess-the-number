VERSION 5.00
Begin VB.Form b11 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����淨-����ģʽ-�������Ҷ�"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4785
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command4 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "(����)"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(��ͨ)"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "(��)"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "b11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
choose = 3
b12.Show
End Sub

Private Sub Command2_Click()
Unload Me
choose = 5
b12.Show
End Sub

Private Sub Command3_Click()
Unload Me
choose = 9
b12.Show
End Sub

Private Sub Command4_Click()
Unload Me
b1.Show
End Sub

Private Sub Form_Load()
Set b12 = Nothing
End Sub
