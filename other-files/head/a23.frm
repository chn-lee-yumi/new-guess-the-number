VERSION 5.00
Begin VB.Form a23 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ͳ�淨-����+2AI"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2520
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
      Top             =   2400
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
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label l22 
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
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l33 
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
      Left            =   960
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l11 
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
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "�ȷ֣� �� ������1 ������2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   4695
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3240
      Picture         =   "a23.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
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
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "   ��һ�����֣��������˹�����һ���档��˭���Ȳ��С�"
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
      TabIndex        =   4
      Top             =   0
      Width           =   4695
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
      Top             =   1200
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
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
      TabIndex        =   7
      Top             =   1200
      Width           =   4695
   End
End
Attribute VB_Name = "a23"
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
        If liansheng = 1 Then
            jf = jf + 1
            savedj
            MsgBox "���ź�,��¶���!!!(���õ�RP���!)" & vbCrLf & "������1����", 64, "��Ϣ"
        Else
            jf = jf + liansheng
            savedj
            MsgBox "���ź�,��¶���!!!(���õ�RP���!)" & vbCrLf & liansheng & "��ʤ��������" & liansheng & "���֣�", 64, "���ź�OAO"
        End If
        l33 = Val(l33) + 1
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
End If
'�˹�����1��ʼ������
guess = aiguess(min, max)
If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" '������
If guess > x Then
    MsgBox guess & vbCrLf & "̫����", , "�˹�����1�µ�����Ϊ��"
    max = guess
    l4 = max
ElseIf guess < x Then
    MsgBox guess & vbCrLf & "̫С��", , "�˹�����1�µ�����Ϊ��"
    min = guess
    l3 = min
ElseIf guess = x Then
    MsgBox "�˹�����1�µ�����Ϊ��" & guess & vbCrLf & "�˹�����1�¶��ˣ�" & vbCrLf & "�˹�����1����һ�֣�", 48, "��ϲ��OwO"
    liansheng = 0
    l11 = Val(l11) + 1
    min = 0
    max = 100
    l3 = "0"
    l4 = "100"
    ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
    x = Int(Rnd * 99 + 1)
    Exit Sub
End If
'�˹�����1��������
'========================================================
'�˹�����2������
guess = aiguess(min, max)
If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" '������
If guess > x Then
    MsgBox guess & vbCrLf & "̫����", , "�˹�����2�µ�����Ϊ��"
    max = guess
    l4 = max
ElseIf guess < x Then
    MsgBox guess & vbCrLf & "̫С��", , "�˹�����w�µ�����Ϊ��"
    min = guess
    l3 = min
ElseIf guess = x Then
    MsgBox "�˹�����2�µ�����Ϊ��" & guess & vbCrLf & "�˹�����2�¶��ˣ�" & vbCrLf & "�˹�����2����һ�֣�", 48, "��ϲ��OwO"
    liansheng = 9
    l22 = Val(l22) + 1
    min = 0
    max = 100
    l3 = "0"
    l4 = "100"
    ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
    x = Int(Rnd * 99 + 1)
End If
'�˹�����2��������
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
