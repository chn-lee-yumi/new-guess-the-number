VERSION 5.00
Begin VB.Form d3 
   BackColor       =   &H00FFFF00&
   Caption         =   "��������ģʽ��AI������ͨģʽ"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox t1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label l1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3120
      Picture         =   "b1.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "�ڴ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "   ��һ�����֣���һ���˹�����һ���档��˭���Ȳ��С�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label l2 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Label l3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Label l4 
      Alignment       =   2  'Center
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "��Χ��       ~"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "�ȷ֣� ��    ��   ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "d3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim min As Integer
Dim max As Integer
Dim t As Long
Dim x As Integer
Dim guess As Integer
Private Sub Command1_Click()
Cls '�����G�ۼ�
guess = Val(t1)
t1 = ""
If guess <= min Or guess >= max Then
    MsgBox "����ѡ�������ڷ�Χ�ڣ�", 16, "����"
    MsgBox "����������", 64, "��Ϣ"
    Exit Sub
End If
If guess > min And guess < max Then
    If guess > x Then
        MsgBox "��ϲ�㡪���´���,̫����", 64, "��ʾ"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "��ϲ�㡪���´���,̫С��", 64, "��ʾ"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "���ź�,��¶���!!!(���õ�RP���!)", 64, "��Ϣ"
        l1 = Val(l1) + 1
        MsgBox "�����һ�֣�", 64, "���ź�~"
        MsgBox "���Ҫ�ڱ�һ�Σ��밴ȷ����" & vbCrLf & "���Ҫ�˳�����رմ��塣", 64, "��Ϣ"
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        x = Int(Rnd * 99)
        Exit Sub
    End If
End If
guess = dieguess(min, max)
If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" '������
MsgBox "           " & guess, , "�˹����ܲµ�����Ϊ��"
If guess > x Then
    MsgBox "̫����", , "���"
    max = guess
    l4 = max
ElseIf guess < x Then
    MsgBox "̫С��", , "���"
    min = guess
    l3 = min
ElseIf guess = x Then
    MsgBox "�˹����ܲ¶��ˣ�", 48, "��ϲ��"
    l2 = Val(l2) + 1
    MsgBox "�˹����ܵ���һ�֣�", 64, "���ź�~"
    min = 0
    max = 100
    l3 = "0"
    l4 = "100"
    MsgBox "���Ҫ�ڱ�һ�Σ��밴ȷ����" & vbCrLf & "���Ҫ�˳�����رմ��塣", 64, "��Ϣ"
    ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
    x = Int(Rnd * 99)
End If
t1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
b.Show
End Sub

Private Sub Form_DblClick()
Print x '��G
End Sub
Private Sub Form_Load()
max = 100
ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
x = Int(Rnd * 99)
End Sub
Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click
End Sub
Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
    MsgBox "����������"
  End If
End Sub
