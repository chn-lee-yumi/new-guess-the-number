VERSION 5.00
Begin VB.Form b6 
   BackColor       =   &H00FFFF00&
   Caption         =   "�����淨-����ģʽ"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4545
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command3 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���޶������ռ��˹�����������ģʽ"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����һ����������"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   $"�����淨-����ģʽ.frx":0000
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   4575
   End
End
Attribute VB_Name = "b6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
b61.Show
End Sub

Private Sub Command2_Click()
b62.Show
End Sub

Private Sub Command3_Click()
Unload Me
b.Show
End Sub

Private Sub Form_Load()
Set b61 = Nothing
End Sub
