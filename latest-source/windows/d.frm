VERSION 5.00
Begin VB.Form d 
   BackColor       =   &H00FFFF80&
   Caption         =   "���ش���"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "��������ģʽ��AI������ͨģʽ"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ȡ�����ֵĴ���"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ַ���ת��"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
d1.Show
End Sub

Private Sub Command2_Click()
d2.Show
End Sub

Private Sub Command3_Click()
d3.Show
End Sub
