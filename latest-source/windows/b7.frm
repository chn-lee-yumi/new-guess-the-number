VERSION 5.00
Begin VB.Form b7 
   BackColor       =   &H00FFFF00&
   Caption         =   "�����淨-�ֻ�ģʽ(ver1.0)"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   8388
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8388
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton c3 
      Caption         =   "����"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   5880
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   8040
      Top             =   4200
   End
   Begin VB.Timer t 
      Enabled         =   0   'False
      Index           =   3
      Left            =   8040
      Top             =   3840
   End
   Begin VB.Timer t 
      Enabled         =   0   'False
      Index           =   2
      Left            =   8040
      Top             =   3480
   End
   Begin VB.Timer t 
      Enabled         =   0   'False
      Index           =   1
      Left            =   8040
      Top             =   3120
   End
   Begin VB.Timer t 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   3000
      Left            =   8040
      Top             =   2760
   End
   Begin VB.CommandButton c2 
      Caption         =   "����"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton c1 
      Caption         =   "��ʼ"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   4
      X1              =   5520
      X2              =   5520
      Y1              =   720
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   3
      X1              =   4080
      X2              =   4080
      Y1              =   720
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   2
      X1              =   2520
      X2              =   2520
      Y1              =   720
      Y2              =   6600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      Index           =   5
      X1              =   1080
      X2              =   7080
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      Index           =   4
      X1              =   1080
      X2              =   7080
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      Index           =   3
      X1              =   1080
      X2              =   7080
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      Index           =   2
      X1              =   1080
      X2              =   7080
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      Index           =   1
      X1              =   1080
      X2              =   7080
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      Index           =   0
      X1              =   1080
      X2              =   7080
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   1080
      X2              =   1080
      Y1              =   -360
      Y2              =   6360
   End
   Begin VB.Label lp 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lp 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lp 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lp 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "POINTS"
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
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Image s 
      Height          =   600
      Index           =   3
      Left            =   6000
      Picture         =   "b7.frx":0000
      Stretch         =   -1  'True
      Top             =   5180
      Width           =   600
   End
   Begin VB.Image s 
      Height          =   600
      Index           =   2
      Left            =   4560
      Picture         =   "b7.frx":531A
      Stretch         =   -1  'True
      Top             =   5125
      Width           =   600
   End
   Begin VB.Image s 
      Height          =   600
      Index           =   1
      Left            =   3000
      Picture         =   "b7.frx":A972
      Stretch         =   -1  'True
      Top             =   5130
      Width           =   600
   End
   Begin VB.Image s 
      Height          =   600
      Index           =   0
      Left            =   1560
      Picture         =   "b7.frx":10114
      Stretch         =   -1  'True
      Top             =   5130
      Width           =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   1
      X1              =   7080
      X2              =   7080
      Y1              =   -360
      Y2              =   6360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6372
      Left            =   7080
      TabIndex        =   12
      Top             =   -360
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "WIN!"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   11
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "0-10"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   3
      Left            =   5520
      TabIndex        =   10
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "0-10"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   2
      Left            =   4080
      TabIndex        =   9
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "0-10"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "0-10"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   0
      Left            =   1080
      TabIndex        =   7
      Top             =   5760
      Width           =   1455
   End
End
Attribute VB_Name = "b7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim points(3) As Single
Dim max(3) As Integer, min(3) As Integer, x(3) As Integer
Dim level(3) As Integer
Dim ABC As Integer 'ѭ����
Dim Bonus(3) As Single '��������
Dim t0 As Integer, t1 As Integer
Dim slow_times As Integer '�������Ӵ���

Private Sub c2_Click()
MsgBox _
"ʹ��С���ּ��̺ͷ�������ơ�Enter��Ϊȷ�ϡ�" & vbCrLf & _
"��������㣬ÿ��һ�����û�¶Զ��ܻ��һЩ�������¶�����ǰ��һ���ȵ�WIN����Ӯ��" & vbCrLf & _
"������÷�����ͷ����ּ��ܣ�(��������ȴʱ��)" & vbCrLf & _
"�ϣ�(���)�Լ�ǰ��һ����15��(���Լ������һ�񣬸ü�����Ч)" & vbCrLf & _
"�£����Լ�ǰ���˺���һ����20��" & vbCrLf & _
"��(����)AI�����ٶȼ��룬�������룬��7��" & vbCrLf & _
"�ң�˲����ַ��������Σ���10��" & vbCrLf & vbCrLf & _
"ע��" & vbCrLf & _
"AI��ʱֻ���ü��ܡ�����" & vbCrLf & _
"AI�ڲ�����ͬʱ�ͷż���" & vbCrLf & _
"����������3�����ظ�ʹ�����ٴμ���AI�����ٶȣ��Ұѳ���ʱ������Ϊ���룬������3��" & vbCrLf & _
"ǿ�ҽ���ʹ��С���ּ��̺ͷ����������Ϸ"
End Sub

Private Sub c3_Click()
Unload Me
b.Show
End Sub

Private Sub Form_DblClick()
MsgBox "", , "Information"
End Sub

Private Sub Form_Load()
ran = Minute(Time): Randomize ran
t0 = 800: t1 = 200 '����AI�ѶȵĲ���
For ABC = 0 To 3
    min(ABC) = 0: max(ABC) = 10: x(ABC) = Int(Rnd * 9 + 1)
    level(ABC) = 1
    points(ABC) = 0
    Bonus(ABC) = 1
Next
'points(0) = 999
End Sub

Private Sub c1_Click()
Dim guess As Integer
If c1.Caption = "��ʼ" Then
    c1.Caption = "ȷ��"
    For ABC = 1 To 3
        t(ABC).Interval = Int(t0 + Rnd * t1)
        t(ABC).Enabled = True
    Next
End If
Text1.SetFocus
guess = Val(Text1.Text)
Text1.Text = ""
        If guess <= min(0) Or guess >= max(0) Then Exit Sub
        If guess > x(0) Then
            points(0) = points(0) + (max(0) - guess) * Bonus(0) '��ȡ����
            max(0) = guess
            l(0) = min(0) & "-" & max(0)
            lp(0).Caption = Round(points(0), 2)
        ElseIf guess < x(0) Then
            points(0) = points(0) + (guess - min(0)) * Bonus(0)
            min(0) = guess
            l(0) = min(0) & "-" & max(0)
            lp(0).Caption = Round(points(0), 2)
        ElseIf guess = x(0) Then
            If level(0) = 7 Then
                s(0).Move s(0).Left, s(0).Top - 730
                For ABC = 1 To 3
                    t(ABC).Enabled = False
                Next
                jf = jf + 7
                savedj
                MsgBox "YOU WIN!!!!" & vbCrLf & "������7����", , "��Ϸ����"
                Unload b7
                b.Show
                Set b7 = Nothing
            Else
                level(0) = level(0) + 1
                s(0).Move s(0).Left, s(0).Top - 730
                remm (0)
                l(0) = min(0) & "-" & max(0)
            End If
        End If
End Sub

Private Sub Command1_Click() '�Լ���һ��
If points(0) < 15 Or level(0) = 7 Then

Else
    points(0) = points(0) - 15
    lp(0).Caption = Round(points(0), 2)
'    If level(0) = 7 Then
'        For ABC = 1 To 3
'            t(ABC).Enabled = False
'        Next
'        jf = jf + 7
'        savedj
'        MsgBox "YOU WIN!!!!" & vbCrLf & "������7���֣�", , "��Ϸ����"
'        Unload b7
'        b.Show
'        Set b7 = Nothing
'    Else
        level(0) = level(0) + 1
        s(0).Move s(0).Left, s(0).Top - 730
        remm (0)
        l(0) = min(0) & "-" & max(0)
'    End If
End If
End Sub

Private Sub Command2_Click() '���Լ��ߵĽ�һ��
If points(0) < 20 Or level(0) = 1 Then

Else
    points(0) = points(0) - 20
    lp(0).Caption = Round(points(0), 2)
    For ABC = 1 To 3
        If level(ABC) > level(0) Then
            level(ABC) = level(ABC) - 1
            s(ABC).Move s(ABC).Left, s(ABC).Top + 730
            remm (ABC)
        End If
    Next
End If
End Sub

Private Sub Command3_Click() 'AI�����ٶȼ��룬��������
If points(0) < 7 Or slow_times = 3 Then

Else
    slow_times = slow_times + 1
    points(0) = points(0) - 7
    lp(0).Caption = Round(points(0), 2)
    Timer1.Enabled = False
    t(0).Enabled = True
    For ABC = 1 To 3
        t(ABC).Interval = t(ABC).Interval * 2
    Next
End If
End Sub

Private Sub Command4_Click() '���ַ���������
Dim guess As Integer
If points(0) < 10 Then

Else
    points(0) = points(0) - 10
    lp(0).Caption = Round(points(0), 2)
    guess = (min(0) + max(0)) \ 2
    Text1.Text = guess
    If t(1).Enabled = True Then '�����д��䣬Ӯ��֮����ܻ���c1_Click()��Text1.SetFocus����
        c1_Click
        guess = (min(0) + max(0)) \ 2
        Text1.Text = guess
        c1_Click
    End If
End If
End Sub




Private Sub t_Timer(Index As Integer)
Dim guess As Integer
Select Case Index
    Case 0
        Timer1.Enabled = True
        t(0).Enabled = False
        Timer1_Timer
    Case 1, 2, 3
        guess = aiguess(min(Index), max(Index))
        If guess > x(Index) Then
            points(Index) = points(Index) + (max(Index) - guess) * Bonus(Index)
            max(Index) = guess
            l(Index) = min(Index) & "-" & max(Index)
            lp(Index).Caption = Round(points(Index), 2)
            AIusePoint (Index)
        ElseIf guess < x(Index) Then
            points(Index) = points(Index) + (guess - min(Index)) * Bonus(Index)
            min(Index) = guess
            l(Index) = min(Index) & "-" & max(Index)
            lp(Index).Caption = Round(points(Index), 2)
            AIusePoint (Index)
        ElseIf guess = x(Index) Then
            If level(Index) = 7 Then
                s(Index).Move s(Index).Left, s(Index).Top - 730
                For ABC = 1 To 3
                    t(ABC).Enabled = False
                Next
                MsgBox "AI" & Index & " WIN!!!!", , "��Ϸ����"
                Unload b7
                b.Show
                Set b7 = Nothing
            Else
                level(Index) = level(Index) + 1
                s(Index).Move s(Index).Left, s(Index).Top - 730
                remm (Index)
                l(Index) = min(Index) & "-" & max(Index)
            End If
        End If
End Select

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 32 Then c1_Click '�س���ո�
If KeyCode = 38 Then Command1_Click
If KeyCode = 40 Then Command2_Click
If KeyCode = 37 Then Command3_Click
If KeyCode = 39 Then Command4_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer) '���ֹ���
If KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 32 Then Exit Sub '8��Backspace
If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
    For ABC = 1 To 3
        t(ABC).Interval = Int(t0 + Rnd * t1)
    Next
    slow_times = 0
End Sub

Private Sub AIusePoint(Index As Integer)  'AIʹ�÷���
Dim A '������
If points(Index) >= 15 Then
    For ABC = 0 To 3
        If level(ABC) > level(Index) Then A = A + 1
    Next
If A >= 1 Then 'ԭ��û��=
        points(Index) = points(Index) - 15
        lp(Index).Caption = Round(points(Index), 2)
        For ABC = 0 To 3
            If level(ABC) > level(Index) Then
                level(ABC) = level(ABC) - 1
                s(ABC).Move s(ABC).Left, s(ABC).Top + 730
                remm (ABC)
                l(ABC) = min(ABC) & "-" & max(ABC)
            End If
        Next
End If
End If
End Sub

Private Sub remm(Index As Integer) 'reflash min max
min(Index) = 0
Select Case level(Index)
    Case 3
        max(Index) = 20: x(Index) = Int(Rnd * 19 + 1): Bonus(Index) = 0.5
    Case 4
        max(Index) = 30: x(Index) = Int(Rnd * 29 + 1): Bonus(Index) = 0.33
    Case 5
        max(Index) = 40: x(Index) = Int(Rnd * 39 + 1): Bonus(Index) = 0.25
    Case 6
        max(Index) = 50: x(Index) = Int(Rnd * 49 + 1): Bonus(Index) = 0.2
    Case 7
        max(Index) = 99: x(Index) = Int(Rnd * 98 + 1): Bonus(Index) = 0.1
    Case 2
        max(Index) = 10: x(Index) = Int(Rnd * 9 + 1): Bonus(Index) = 1
End Select
End Sub
