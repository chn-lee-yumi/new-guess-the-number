VERSION 5.00
Begin VB.Form b61 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����淨-����ģʽ-����һ"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5730
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   240
      TabIndex        =   9
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
      Left            =   4320
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
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "��߼�¼��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "          ��һ������,              ��Χ-2147483647~2147483647"
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
      TabIndex        =   7
      Top             =   0
      Width           =   5775
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   4080
      Picture         =   "b4.frx":0000
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
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label l3 
      Alignment       =   2  'Center
      Caption         =   "-2147483647"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label l4 
      Alignment       =   2  'Center
      Caption         =   "2147483647"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "��Χ��             ~"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   5775
   End
End
Attribute VB_Name = "b61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim min As Long
Dim max As Long
Dim t As Long
Dim x As Long
Dim guess As Single
Private Sub Command1_Click()
Dim jilu As String * 2
Cls '�����G�ۼ�
guess = Val(t1)
t1 = ""
If guess <= min Or guess >= max Then
    MsgBox "����ѡ�������ڷ�Χ�ڣ����������롣", 16, "����"
End If
    t = t + 1
If guess > min And guess < max Then
    If guess > x Then
        MsgBox "��ϲ�㡪���´���,̫����", 64, "��ʾ"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "��ϲ�㡪���´���,̫С��", 64, "��ʾ"
        min = guess
        l3 = guess
    ElseIf guess = x Then
        MsgBox "���ź�,��¶���!!!(���õ�RP���!)" & vbCrLf & "�㹲����" & t & "��", 64, "��Ϣ"
        If Label4 = "\" Then Label4 = t: GoTo jiluline
        If t < Val(Label4) Then
            play "data\happy.mp3"
            MsgBox "��ˢ���˼�¼������", , "Good job!"
            Label4 = t
jiluline:   jilu = Label4
            Open App.Path & "\data\�����ּ�¼.txt" For Random As #2
                Put #2, 3, jilu
            Close
        End If
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = longx
        t = 0
        l3 = -2147483647
        l4 = 2147483647
        MsgBox "���ȷ�����¿�ʼ�档", 64, "��Ϣ"
    End If
End If
t1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
b6.Show
End Sub

Private Sub Form_DblClick()
Print x '��G
End Sub
Private Sub Form_Load()
Dim jilu1 As String * 2
Dim jilu2 As Integer
min = -2147483647
max = 2147483647
ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
x = longx
Open App.Path & "\data\�����ּ�¼.txt" For Random As #2
    Get #2, 3, jilu1
    jilu2 = Val(Trim(jilu1))
    If jilu2 = 0 Then Label4 = "\" Else Label4 = jilu2
Close
End Sub
Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click '13�ǻس���
End Sub
Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
  End If
End Sub
