VERSION 5.00
Begin VB.Form a22 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ͳ�淨-����+1AI"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
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
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
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
      TabIndex        =   9
      Top             =   840
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
      TabIndex        =   8
      Top             =   840
      Width           =   615
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
      TabIndex        =   6
      Top             =   480
      Width           =   615
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
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3240
      Picture         =   "a22.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   735
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
      TabIndex        =   2
      Top             =   480
      Width           =   615
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
      TabIndex        =   3
      Top             =   480
      Width           =   4695
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
      TabIndex        =   7
      Top             =   840
      Width           =   4695
   End
End
Attribute VB_Name = "a22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim min As Integer
Dim max As Integer
Dim x As Integer
Dim guess As Integer
Dim liansheng As Integer '��ʤ��¼

Private Sub Command1_Click()
Cls '�����G�ۼ�
guess = Val(t1)
t1 = ""
If guess <= min Or guess >= max Then
    MsgBox "����ѡ�������ڷ�Χ�ڣ����������롣", 16, "����"
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
        liansheng = liansheng + 1
        If liansheng >= 2 Then
            '2��ʤ1���֣�3��ʤ2���֡���
            jf = jf + liansheng - 1
            savedj
            MsgBox "���ź�,��¶���!!!(���õ�RP���!)" & vbCrLf & liansheng & "��ʤ��������" & (liansheng - 1) & "���֣�", 64, "���ź�OAO"
        Else
            MsgBox "���ź�,��¶���!!!(���õ�RP���!)", 64, "���ź�OAO"
        End If
        l1 = Val(l1) + 1
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
End If
guess = aiguess(min, max)
If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" '������
If guess > x Then
    MsgBox guess & vbCrLf & "̫����", , "�˹����ܲµ�����Ϊ��"
    max = guess
    l4 = max
ElseIf guess < x Then
    MsgBox guess & vbCrLf & "̫С��", , "�˹����ܲµ�����Ϊ��"
    min = guess
    l3 = min
ElseIf guess = x Then
    l2 = Val(l2) + 1
    MsgBox "�˹����ܲµ�����Ϊ��" & guess & vbCrLf & "�˹����ܲ¶��ˣ�" & vbCrLf & "�˹����ܵ���һ�֣�", 48, "��ϲ��OwO"
    liansheng = 0
    min = 0
    max = 100
    l3 = "0"
    l4 = "100"
    ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
    x = Int(Rnd * 99 + 1)
End If
t1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
a2.Show
End Sub

Private Sub Form_DblClick()
Print x '��G
End Sub
Private Sub Form_Load()
max = 100
ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
x = Int(Rnd * 99 + 1)
End Sub
Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click
End Sub
Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
  End If
End Sub
