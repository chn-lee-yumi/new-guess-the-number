VERSION 5.00
Begin VB.Form b8 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�����淨-��ƭģʽ"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   4440
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����˵��"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ��Ϸ"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   1695
      Begin VB.OptionButton o4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "1000"
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
         Left            =   720
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton o3 
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
         Height          =   255
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.OptionButton o2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "30%"
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
      Left            =   1680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton o1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "15%"
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
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
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
      Left            =   360
      TabIndex        =   12
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   3000
      Picture         =   "b8.frx":0000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
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
      Left            =   2760
      TabIndex        =   10
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "��µĴ���"
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
      Left            =   840
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�Ѷȣ�"
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
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "b8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim min As Integer, max As Integer, p As Single '��Сֵ�����ֵ���󱨸���
Dim x As Integer, guess As Integer '��Ҫ�µ�������µ���
Dim ptemp As Single '������p�Ƚ�


Private Sub Command1_Click()
If o1.Value = True Then
    p = 0.15
Else
    p = 0.3
End If
If o3.Value = True Then
    max = 100
    x = Int(Rnd * 99 + 1)
    Text1.MaxLength = 2
Else
    max = 1000
    x = Int(Rnd * 999 + 1)
    Text1.MaxLength = 3
End If
min = 0
Command1.Enabled = False
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Shell "notepad.exe " & App.Path & "\data\��ƭģʽ(����)", 1
End Sub

Private Sub Command3_Click()
Unload Me
b.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
  End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then '13�ǻس�
    guess = Val(Text1)
    Text1 = ""
    If guess >= max Or guess <= min Then
        MsgBox "�������ֳ�����Χ��", vbInformation, "��ʾ"
        Exit Sub
    Else
        Label3 = Val(Label3) + 1
    End If
    ptemp = Rnd
    If guess = x Then
        MsgBox "����������������ˣ���������" & vbCrLf & "������" & Label3.Caption & "�Ρ���", , "���Ǻܱ�Ǹ�ĸ�����"
        If o1.Value = True Then
            If o3.Value = True Then
                jf = jf + 1: MsgBox "������1����"
            Else
                jf = jf + 3: MsgBox "������3����"
            End If
        Else
            If o3.Value = True Then
                jf = jf + 2: MsgBox "������2����"
            Else
                jf = jf + 6: MsgBox "������6����"
            End If
        End If
        savedj
        Command1.Enabled = True
        Text1.Enabled = False
        Label3 = 0
    ElseIf guess > x Then
        If ptemp < p Then
            MsgBox "̫С�ˣ�", , "��ʾ"
        Else
            MsgBox "̫���ˣ�", , "��ʾ"
        End If
    Else
        If ptemp < p Then
            MsgBox "̫���ˣ�", , "��ʾ"
        Else
            MsgBox "̫С�ˣ�", , "��ʾ"
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Label1 = Label1 & vbCrLf & "��Χ��"
End Sub


