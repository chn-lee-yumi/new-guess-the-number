VERSION 5.00
Begin VB.Form die 
   BackColor       =   &H00FFFF00&
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   Picture         =   "a3.frx":0000
   ScaleHeight     =   3705
   ScaleWidth      =   6135
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����"
      Height          =   1215
      Left            =   1440
      TabIndex        =   9
      Top             =   2400
      Width           =   2655
      Begin VB.CommandButton Command5 
         Caption         =   "����˵��"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Key"
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pass"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label labkey 
         BackColor       =   &H0080FFFF&
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
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label labpass 
         BackColor       =   &H0080FFFF&
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
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
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
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1920
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
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
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
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
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "��������ģʽ����һ�����֣�˭����˭������(����)����Ҫ��������������һ�����س��֡����ֻʣ�����Լ���"
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
      Width           =   6135
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
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3240
      Picture         =   "a3.frx":0342
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "��Χ��          ~"
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
      Width           =   6135
   End
   Begin VB.Label lab3 
      Caption         =   "������ģ�                        �㡢����1������2������3"
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
      TabIndex        =   2
      Top             =   480
      Width           =   6135
   End
End
Attribute VB_Name = "die"
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
Dim p As Integer
Private Sub Command1_Click()
    On Error Resume Next
    Cls '�����G�ۼ�
    If t1 = "ke" Then t1 = "": GoTo key
    guess = Val(t1)
    If t1 = "pa" Then t1 = "": GoTo pass
    t1 = ""
    '�ж��Լ��µ���
    If guess <= min Or guess >= max Then
        MsgBox "����ѡ�������ڷ�Χ�ڣ�", 16, "����"
        MsgBox "����������", 64, "��Ϣ"
        Exit Sub
    End If
    If guess > min And guess < max Then
    If guess > x Then
        MsgBox "̫����", 64, ""
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", 64, ""
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�������~~", 64, "��ϲ�㣡"
        MsgBox "�����ˣ�������hahahaha..." & vbCrLf & "���ȷ���˳���", 64, "��Ϣ"
        Unload Me
        Exit Sub
    End If
    End If
pass:
Select Case p
Case 1 '�Լ�+AI1+AI2+AI3
    'AI1����
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����1�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����1�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����1�����ˣ����������֣�", 64, "���ź�~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "������ģ�                        �㡢����2������3"
        MsgBox "��׼����һ��������", 64, "��Ϣ"
        p = 5
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
    'AI2����
    guess = aiguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����2�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����2�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����2�����ˣ����������֣�", 64, "���ź�~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "������ģ�                        �㡢����1������3"
        MsgBox "��׼����һ��������", 64, "��Ϣ"
        p = 7
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
    'AI3����
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����3�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����3�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����3�����ˣ����������֣�", 64, "���ź�~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "������ģ�                        �㡢����1������2"
        MsgBox "��׼����һ��������", 64, "��Ϣ"
        p = 3
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
Case 2 '�Լ�+AI1
    'AI1����
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����1�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����1�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����1�����ˣ����������֣�", 64, "���ź�~~"
        lab3 = "������ģ�                        ��"
        MsgBox "���ź�����Ӯ�ˡ�QAQ" & vbCrLf & "���ȷ���˳���", 48, "��Ϣ"
        Unload Me
    End If
Case 3 '�Լ�+AI1+AI2
    'AI1����
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����1�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����1�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����1�����ˣ����������֣�", 64, "���ź�~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "������ģ�                        �㡢����2"
        MsgBox "��׼����һ��������", 64, "��Ϣ"
        p = 4
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
    'AI2����
    guess = aiguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����2�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����2�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����2�����ˣ����������֣�", 64, "���ź�~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "������ģ�                        �㡢����1"
        MsgBox "��׼����һ��������", 64, "��Ϣ"
        p = 2
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
Case 4 '�Լ�+AI2
    'AI2����
    guess = aiguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����2�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����2�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����2�����ˣ����������֣�", 64, "���ź�~~"
        lab3 = "������ģ�                        ��"
        MsgBox "���ź�����Ӯ�ˡ�QAQ" & vbCrLf & "���ȷ���˳���", 48, "��Ϣ"
        Unload Me
    End If
Case 5 '�Լ�+AI2+AI3
    'AI2����
    guess = aiguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����2�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����2�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����2�����ˣ����������֣�", 64, "���ź�~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "������ģ�                        �㡢����3"
        MsgBox "��׼����һ��������", 64, "��Ϣ"
        p = 6
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
    'AI3����
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����3�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����3�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����3�����ˣ����������֣�", 64, "���ź�~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "������ģ�                        �㡢����2"
        MsgBox "��׼����һ��������", 64, "��Ϣ"
        p = 4
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
Case 6 '�Լ�+AI3
    'AI3����
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����3�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����3�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����3�����ˣ����������֣�", 64, "���ź�~~"
        lab3 = "������ģ�                        ��"
        MsgBox "���ź�����Ӯ�ˡ�QAQ" & vbCrLf & "���ȷ���˳���", 48, "��Ϣ"
        Unload Me
    End If
Case 7 '�Լ�+AI1+AI3
    'AI1����
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����1�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����1�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����1�����ˣ����������֣�", 64, "���ź�~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "������ģ�                        �㡢����3"
        MsgBox "��׼����һ��������", 64, "��Ϣ"
        p = 6
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
    'AI3����
    guess = dieguess(min, max)
    If guess <= min Or guess >= max Then MsgBox "�����ڼ���ִ���" & vbCrLf & "�˹����ܲµ������ڷ�Χ�⣡" & vbCrLf & "min max guess:" & min & max & guess '������
        MsgBox "           " & guess, , "�˹�����3�µ�����Ϊ��"
    If guess > x Then
        MsgBox "̫����", , "���"
        max = guess
        l4 = max
    ElseIf guess < x Then
        MsgBox "̫С��", , "���"
        min = guess
        l3 = min
    ElseIf guess = x Then
        MsgBox "�˹�����3�����ˣ�", 48, "���ź�~~"
        MsgBox "�˹�����3�����ˣ����������֣�", 64, "���ź�~~"
        min = 0
        max = 100
        l3 = "0"
        l4 = "100"
        lab3 = "������ģ�                        �㡢����1"
        MsgBox "��׼����һ��������", 64, "��Ϣ"
        p = 2
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        x = Int(Rnd * 99 + 1)
        Exit Sub
    End If
End Select
t1 = ""
t1.SetFocus
Exit Sub
key:
MsgBox "���Բ�����������ǣ�" & x
End Sub

Private Sub Command2_Click()
Unload Me
b1.Show
End Sub

Private Sub Command3_Click()
If djpass = 0 Then MsgBox "��û�иõ��ߣ���ȡ�����뿴˵����", , "��Ϣ": Exit Sub
djpass = djpass - 1
labpass = Val(labpass) - 1 & "��"
t1 = "pa"
Command1_Click
End Sub

Private Sub Command4_Click()
If djkey = 0 Then MsgBox "��û�иõ��ߣ���ȡ�����뿴˵����", , "��Ϣ": Exit Sub
djkey = djkey - 1
labkey = Val(labkey) - 1 & "��"
t1 = "ke"
Command1_Click
End Sub

Private Sub Form_DblClick()
Print x '��G
End Sub
Private Sub Form_Load()
min = 0
max = 100
ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
x = Int(Rnd * 99 + 1)
p = 1
labpass = djpass & "��"
labkey = djkey & "��"
End Sub

Private Sub Form_Unload(Cancel As Integer)
savedj
b.Show
End Sub

Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click
End Sub
Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
    MsgBox "����������"
  End If
End Sub
