VERSION 5.00
Begin VB.Form b2 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����淨-��սģʽ"
   ClientHeight    =   3096
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3096
   ScaleWidth      =   5820
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox t1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
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
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lab3 
      Alignment       =   1  'Right Justify
      Caption         =   "14.3%"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "�Ѷȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label l3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��Χ��                  ~"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Label lab0 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "����(��4��)��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lab2 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "��µĴ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lab1 
      Alignment       =   1  'Right Justify
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "���ƴ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "  ��һ�����֣����ܳ������ƵĴ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5895
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
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3480
      Picture         =   "b3.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "b2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim min As Long
Dim max As Long
Dim t As Long '��ҲµĴ���
Dim t0 As Long '�ؿ�
Dim x As Long
Dim guess As Long

Private Sub Command1_Click()
    On Error Resume Next
    Cls '�����G�ۼ�
    guess = Val(t1)
    t1 = ""
Select Case t0
Case 1
If guess <= min Or guess >= max Then
    MsgBox "����ѡ�������ڷ�Χ�ڣ����������롣", 16, "����"
    Exit Sub
End If
    t = t + 1
    lab2 = t
    If guess > x Then
        MsgBox "��ϲ�㡪���´���,̫����", 64, "��ʾ"
        l4 = guess
        max = guess
        If t = (Val(lab1)) Then MsgBox "�����ˣ�", 48, "��ϲ�㣡": Unload Me: b.Show: Set b2 = Nothing
    ElseIf guess < x Then
        MsgBox "��ϲ�㡪���´���,̫С��", 64, "��ʾ"
        l3 = guess
        min = guess
        If t = (Val(lab1)) Then MsgBox "�����ˣ�", 48, "��ϲ�㣡": Unload Me: b.Show: Set b2 = Nothing
    ElseIf guess = x Then
        MsgBox "���ź�,����Խ�����һ���ˡ�" & vbCrLf & "����ȷ��", 64, "���ģ�"
        t0 = 2
        min = 0
        max = 1000
        l3 = min
        l4 = max
        lab0 = 2
        lab1 = 9
        lab2 = 0
        lab3 = "18.2%"
        t = 0
        x = Int(Rnd * 999 + 1)
    End If
Case 2
If guess <= min Or guess >= max Then
    MsgBox "����ѡ�������ڷ�Χ�ڣ�", 16, "����"
    MsgBox "����������", 64, "��Ϣ"
    Exit Sub
End If
    t = t + 1
    lab2 = t
    If guess > x Then
        MsgBox "��ϲ�㡪���´���,̫����", 64, "��ʾ"
        l4 = guess
        max = guess
        If t = (Val(lab1)) Then MsgBox "�����ˣ�", 48, "��ϲ�㣡": Unload Me: b.Show: Set b2 = Nothing
    ElseIf guess < x Then
        MsgBox "��ϲ�㡪���´���,̫С��", 64, "��ʾ"
        l3 = guess
        min = guess
        If t = (Val(lab1)) Then MsgBox "�����ˣ�", 48, "��ϲ�㣡": Unload Me: b.Show: Set b2 = Nothing
    ElseIf guess = x Then
        MsgBox "���ź�,����Խ�����һ���ˡ�" & vbCrLf & "����ȷ��", 64, "���ģ�"
        t0 = 3
        min = 0
        max = 10000
        l3 = min
        l4 = max
        lab0 = t0
        lab1 = 11
        lab2 = 0
        lab3 = "21.4%"
        t = 0
        x = Int(Rnd * 9999 + 1)
    End If
Case 3
If guess <= min Or guess >= max Then
    MsgBox "����ѡ�������ڷ�Χ�ڣ�", 16, "����"
    MsgBox "����������", 64, "��Ϣ"
    Exit Sub
End If
    t = t + 1
    lab2 = t
    If guess > x Then
        MsgBox "��ϲ�㡪���´���,̫����", 64, "��ʾ"
        l4 = guess
        max = guess
        If t = (Val(lab1)) Then MsgBox "�����ˣ�", 48, "��ϲ�㣡": Unload Me: b.Show: Set b2 = Nothing
    ElseIf guess < x Then
        MsgBox "��ϲ�㡪���´���,̫С��", 64, "��ʾ"
        l3 = guess
        min = guess
        If t = (Val(lab1)) Then MsgBox "�����ˣ�", 48, "��ϲ�㣡": Unload Me: b.Show: Set b2 = Nothing
    ElseIf guess = x Then
        MsgBox "���ź�,����Խ�����һ���ˡ�" & vbCrLf & "����ȷ��", 64, "���ģ�"
        t0 = 4
        min = 0
        max = 100000
        l3 = min
        l4 = max
        lab0 = t0
        lab1 = 13
        lab2 = 0
        lab3 = "23.5%"
        t = 0
        x = Int(Rnd * 99999 + 1)
    End If
Case 4
If guess <= min Or guess >= max Then
    MsgBox "����ѡ�������ڷ�Χ�ڣ�", 16, "����"
    MsgBox "����������", 64, "��Ϣ"
    Exit Sub
End If
    t = t + 1
    lab2 = t
    If guess > x Then
        MsgBox "��ϲ�㡪���´���,̫����", 64, "��ʾ"
        l4 = guess
        max = guess
        If t = (Val(lab1)) Then MsgBox "�����ˣ�", 48, "��ϲ�㣡": Unload Me: b.Show: Set b2 = Nothing
    ElseIf guess < x Then
        MsgBox "��ϲ�㡪���´���,̫С��", 64, "��ʾ"
        l3 = guess
        min = guess
        If t = (Val(lab1)) Then MsgBox "�����ˣ�", 48, "��ϲ�㣡": Unload Me: b.Show: Set b2 = Nothing
    ElseIf guess = x Then
        jf = jf + 10
        MsgBox "���ź�,��򱬻���QAQ" & vbCrLf & "��ϲ������10���֣�" & vbCrLf & "���ȷ���ص�ģʽѡ����档", 64, "���ģ�"
        savedj
        Unload Me
        b.Show
        Exit Sub '�������Ļ����򱬻���ر����д���Ҳ�ᷢ�ֲ����ֳ����ں�̨����
    End If
End Select
t1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
b.Show
End Sub

Private Sub Form_DblClick()
Print x
End Sub

Private Sub Form_Load()
Debug.Print t
'If t <> 0 Then t = 0 '����û����set b2 = nothing ʱ����û����䣬��֪��Ϊʲôt��ֵ��ж���������󻹻ᱣ�档   ����set b2 =- nothing ��t��ֵ�����ˡ���û�취�����𣿣���
t0 = 1
min = 0
max = 100
ran = Minute(Time): Randomize ran
x = Int(Rnd * 99 + 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
savedj
b.Show
End Sub

Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click '13�ǻس���,32�ǿո��
End Sub
Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
  End If
End Sub
