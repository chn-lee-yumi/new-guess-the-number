VERSION 5.00
Begin VB.Form b13 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����淨-����ģʽ-����ʵ����"
   ClientHeight    =   3585
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7320
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox p1 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   7275
      TabIndex        =   46
      Top             =   360
      Width           =   7335
   End
   Begin VB.PictureBox p2 
      Height          =   855
      Index           =   5
      Left            =   5040
      ScaleHeight     =   795
      ScaleWidth      =   2235
      TabIndex        =   52
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox p2 
      Height          =   855
      Index           =   4
      Left            =   2760
      ScaleHeight     =   795
      ScaleWidth      =   2235
      TabIndex        =   51
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox p2 
      Height          =   855
      Index           =   3
      Left            =   2760
      ScaleHeight     =   795
      ScaleWidth      =   2235
      TabIndex        =   50
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox p2 
      Height          =   855
      Index           =   2
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   2235
      TabIndex        =   49
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox p2 
      Height          =   855
      Index           =   1
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   2235
      TabIndex        =   48
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox p2 
      Height          =   855
      Index           =   6
      Left            =   5040
      ScaleHeight     =   795
      ScaleWidth      =   2235
      TabIndex        =   47
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ȷ��"
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
      Left            =   2880
      TabIndex        =   30
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Left            =   1440
      TabIndex        =   29
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "��һ��"
      Height          =   375
      Index           =   5
      Left            =   6600
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "��һ��"
      Height          =   375
      Index           =   4
      Left            =   4320
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "��һ��"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "��һ��"
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "��һ��"
      Height          =   375
      Index           =   6
      Left            =   6600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȷ���淨"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.OptionButton o2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "������ʤ��"
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
      Left            =   2040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton o1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "������"
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��Ϸ����"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Labjs 
      BackColor       =   &H00FF80FF&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   45
      Top             =   3240
      Width           =   7335
   End
   Begin VB.Label Labdb 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "1"
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
      Left            =   6720
      TabIndex        =   43
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Labwb 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "1"
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
      Left            =   6720
      TabIndex        =   44
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�� ��"
      Height          =   375
      Left            =   6480
      TabIndex        =   42
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�� ��"
      Height          =   375
      Left            =   6480
      TabIndex        =   41
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Labwpoints 
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   5520
      TabIndex        =   40
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Labdpoints 
      BackColor       =   &H00FFFFC0&
      Caption         =   "334"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   5520
      TabIndex        =   39
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�ҷ�������"
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
      Left            =   3960
      TabIndex        =   38
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�з�������"
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
      Left            =   3960
      TabIndex        =   37
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "�ȼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   5880
      TabIndex        =   36
      Top             =   840
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   5
      Left            =   5040
      Picture         =   "b13.frx":0000
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   4
      Left            =   2760
      Picture         =   "b13.frx":0DAD
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   3
      Left            =   480
      Picture         =   "b13.frx":1C4A
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "���Լ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   18
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   34
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label num2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   33
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label num1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   32
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      Caption         =   "��Χ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "�ȼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   3600
      TabIndex        =   27
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "�ȼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3600
      TabIndex        =   26
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "�ȼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   25
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "�ȼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   5880
      TabIndex        =   24
      Top             =   2760
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   2
      Left            =   5040
      Picture         =   "b13.frx":2C56
      Stretch         =   -1  'True
      Top             =   480
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   1
      Left            =   2760
      Picture         =   "b13.frx":3A7C
      Stretch         =   -1  'True
      Top             =   480
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   0
      Left            =   480
      Picture         =   "b13.frx":472B
      Stretch         =   -1  'True
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AI5"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5880
      TabIndex        =   23
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AI4"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5880
      TabIndex        =   22
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AI3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3600
      TabIndex        =   21
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AI2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   20
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AI1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   19
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Height          =   135
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Width           =   7455
   End
   Begin VB.Label lablevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   12
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lablevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   4
      Left            =   4440
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lablevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   10
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lablevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lablevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   6
      Left            =   6720
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�淨��"
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "��                                ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "��                                ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   35
      Top             =   2400
      Width           =   7455
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   1320
      Y2              =   2400
   End
End
Attribute VB_Name = "b13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wpoints As Long, dpoints As Long '����
Dim level(2 To 6) As Integer '�ȼ�
Dim pname(1 To 6) As String  '����
Dim A As Integer '�����ֵ���������
Dim guess As Integer
Dim palive(1 To 6) As Boolean
Dim wanfa As Integer
Dim DB As Single, WB As Single '�ӳɣ����˵ģ��Լ���
Dim walive As Integer, dalive As Integer '�����������Լ��ģ����˵�
Dim cha As Integer '�ҷ�������з������Ĳ�
Dim d As Integer, w As Integer '������ʤ�Ƶıȷ֣����˵ģ��Լ���
Dim min As Integer, max As Integer
Dim x As Integer
Dim Ifrist As Boolean '�Ƿ��Լ��Ȳ�
Dim ABC As Integer, def As Integer
Dim aitop As Integer


Private Sub cmd_Click(Index As Integer)
'�����ȼ��Ĵ��롣
Dim need As Integer
Select Case Index
    Case 3, 5
        If level(Index) = 5 Then MsgBox "���ĵȼ��Ѿ����˶�����", vbInformation, "��Ϣ": Exit Sub
        need = Abs(Val(lablevel(Index).Caption)) * 40 + 20
        If wpoints < need Then
           MsgBox "�㲻���������������Ϊ" & need, vbInformation, "��Ϣ"
        Else
            wpoints = wpoints - need
            Labwpoints = wpoints
            level(Index) = level(Index) + 1
            lablevel(Index) = level(Index)
        End If
    Case 2, 4, 6
        If level(Index) = -3 Then MsgBox "���ĵȼ��Ѿ����˵׼���", vbInformation, "��Ϣ": Exit Sub
        need = Abs(Val(lablevel(Index).Caption) - 1) * 40 + 20
        If wpoints < need Then
           MsgBox "�㲻���������������Ϊ" & need, vbInformation, "��Ϣ"
        Else
            wpoints = wpoints - need
            Labwpoints = wpoints
            level(Index) = level(Index) - 1
            lablevel(Index) = level(Index)
        End If
End Select
End Sub

Private Sub Command1_Click()
Shell "notepad.exe " & App.Path & "\data\����ʵ����(����)", 1
End Sub

Private Sub Command2_Click()
p1.Visible = False
If o1.Value = True Then wanfa = 1
If o2.Value = True Then wanfa = 2
Command2.Enabled = False
o1.Enabled = False
o2.Enabled = False
End Sub

Private Sub Command3_Click()
Unload Me
b1.Show
End Sub

Private Sub Command4_Click()
Cls '�����G�ۼ�
'�����ֲ���
If A = 1 And palive(1) = True Then
    guess = Val(Text1)
    If guess <= min Or guess >= max Then
        MsgBox "��µ����ֳ�����Χ��", vbExclamation, "����"
        GoTo exits
    End If
    Text1.Enabled = False
    GoTo result
Else
aig:
    If A >= aitop + 1 Then '���ڴ���������±�Խ������
        A = 1
        If palive(1) = True Then Text1.Enabled = True
        GoTo exits
    End If
    If palive(A) = True Then 'ERROR:palive(A)�±�Խ�� A=7 aitop=3/5/6
        guess = Labaiguess(min, max, x, level(A)) 'AI����
    Else
        A = A + 1
        If A = aitop + 1 Then
            A = 1
            If palive(1) = True Then Text1.Enabled = True
            GoTo exits
        End If
        GoTo aig
    End If
End If

result:    If guess = x Then
        Labjs = pname(A) & "�µ���" & x & "�������ˣ�"
        MsgBox pname(A) & "�µ���" & x & "�������ˣ�", vbInformation, "��Ϣ"
        Labjs = "���������ã������²�����"
        If wanfa = 1 Then GoTo aidie1
        If wanfa = 2 Then GoTo aidie2
    ElseIf guess < x Then
        Labjs = pname(A) & "�µ���" & guess & "��̫С�ˡ�"
        num1 = guess
        If A = 1 Or A = 3 Or A = 5 Then
            wpoints = Int((wpoints - min + guess) * WB) '��ȡ����
            Labwpoints = wpoints
        Else
            dpoints = Int((dpoints - min + guess) * DB) '��ȡ����
            Labdpoints = dpoints
        End If
        min = guess
    ElseIf guess > x Then
        num2 = guess
        Labjs = pname(A) & "�µ���" & guess & "��̫���ˡ�"
        If A = 1 Or A = 3 Or A = 5 Then
            wpoints = Int((wpoints + max - guess) * WB) '��ȡ����
            Labwpoints = wpoints
        Else
            dpoints = Int((dpoints + max - guess) * DB) '��ȡ����
            Labdpoints = dpoints
        End If
        max = guess
    End If
    
A = A + 1
AIusepoints '�ڴ˴��ӶԷ��÷�����Sub
If A = 7 And palive(1) = True Then A = 1: Text1.Enabled = True
If A = 7 And palive(1) = False Then A = 2
Exit Sub '=============================error������aidie1��AI4���ˣ�û�о���aitop���ж�������A������aitop���±�Խ��
aidie1:
palive(A) = False
p2(A).Visible = True
walive = 0: dalive = 0 '�����������������������������������
If palive(1) = True Then walive = walive + 1
If palive(2) = True Then dalive = dalive + 1
If palive(3) = True Then walive = walive + 1
If palive(4) = True Then dalive = dalive + 1
If palive(5) = True Then walive = walive + 1
If palive(6) = True Then dalive = dalive + 1
If walive = 0 Then 'ʤ���ж�
    play "data\aaaaah.mp3"
    MsgBox "��Ȼ�����3��С���ӡ���", vbExclamation, "�������ˡ���"
    Unload Me
    b1.Show
    GoTo exits
ElseIf dalive = 0 Then
    play "data\win.mp3"
    MsgBox "���ź�����Ӯ��QAQ", vbExclamation, "No������"
    '�ڴ˿��Լӻ��ֻ����
    jf = jf + 6
    savedj
    MsgBox "���ź���������6���֣�", vbInformation, "��Ϣ"
    Unload Me
    b1.Show
    GoTo exits
Else
    min = 0
    max = 1000
    x = Int(Rnd * 1000 + 1)
    num1 = "0"
    num2 = "1000"
    A = A + 1
    Getaitop
    If A = aitop + 1 Then A = 1: Text1.Enabled = True
    If palive(1) = False Then A = 2: Text1.Enabled = False
End If
cha = walive - dalive
Select Case cha '���ݲ�¼ӳ�
    Case 1
        DB = 1.4: WB = 1: Labdb = "1.4": Labwb = "1"
    Case 2
        DB = 2: WB = 1: Labdb = "2": Labwb = "1"
    Case 0
        DB = 1: WB = 1: Labdb = "1": Labwb = "1"
    Case -1
        DB = 1: WB = 1.4: Labdb = "1": Labwb = "1.4"
    Case -2
        DB = 1: WB = 2: Labdb = "1": Labwb = "2"
End Select
Labjs = "�÷ֱ���Ϊ���ѷ�" & WB & "������" & DB & "��"
Getaitop
If A > 1 Then Command4_Click
Exit Sub '================================
aidie2:
If A = 1 Or A = 3 Or A = 5 Then
    d = d + 1
    MsgBox "�����ҷ���з��ıȷ���" & vbCrLf & w & ":" & d
    If d = 3 Then
        play "data\aaaaah.mp3"
        MsgBox "��ϲ�㣬�����ˣ�����", vbExclamation, "������"
        Unload Me
        b1.Show
        GoTo exits
    End If
Else
    w = w + 1
    MsgBox "�����ҷ���з��ıȷ���" & vbCrLf & w & ":" & d
    If w = 3 Then
        play "data\win.mp3"
        MsgBox "���ź�����Ӯ��QAQ" & vbCrLf & "������6���֣�", vbExclamation, "No������"
        '�ڴ˿��Լӻ��ֻ����
        jf = jf + 6
        savedj
        Unload Me
        b1.Show
        GoTo exits
    End If
End If
Resetdata
If Ifrist = False Then
    Ifrist = True
    dpoints = 334
    Labdpoints = "334"
    A = 1
    Text1.Enabled = True
    Labjs = "���Ȳ¡�"
Else
    Ifrist = False
    wpoints = 334
    Labwpoints = "334"
    A = 2
    Text1.Enabled = False
    Labjs = "AI1�Ȳ¡�"
End If
exits:
End Sub

Private Sub Form_DblClick()
Print x
End Sub

Private Sub Form_Load()
Randomize Time
'wpoints = 100000 '������
pname(1) = "��"
For ran = 2 To 6
    pname(ran) = "AI" & (ran - 1)
Next
Ifrist = True
Labjs = "���Ȳ¡�"
Resetdata
dpoints = 334
Labdpoints = "334"
End Sub
'============================================================================
Public Sub Resetdata()
min = 0: max = 1000
num1 = "0"
num2 = "1000"
x = Int(Rnd * 1000 + 1)
For A = 1 To 6
    palive(A) = True
Next
For A = 2 To 6
    level(A) = 0
    lablevel(A) = "0"
Next
A = 1
wpoints = 0
dpoints = 0
DB = 1
WB = 1
Labwpoints = "0"
Labdpoints = "0"
Text1.Enabled = True
aitop = 6
End Sub


Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command4_Click '13�ǻس���
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
  End If
End Sub

Public Sub AIusepoints() '��ַ����points
Dim need As Integer
If wanfa = 1 Then
    For ABC = 2 To 6 Step 2
        If level(ABC) < 5 And palive(ABC) = True Then
            need = Abs(level(ABC)) * 40 + 20
            If dpoints < need Then
                GoTo exitsub
            Else
                dpoints = dpoints - need
                Labdpoints = dpoints
                level(ABC) = level(ABC) + 1
                lablevel(ABC) = level(ABC)
                Debug.Print ABC '�淨һ����
                AIusepoints '�������ܲ�����������ͬ
            End If
        End If
        
    Next
    For ABC = 3 To 5 Step 2
        If level(ABC) > -3 And palive(ABC) = True Then
            need = Abs(level(ABC)) * 40 + 20
            If dpoints < need Then
                GoTo exitsub
            Else
                dpoints = dpoints - need
                Labdpoints = dpoints
                level(ABC) = level(ABC) - 1
                lablevel(ABC) = level(ABC)
                AIusepoints
            End If
        End If
    Next
Else
    For ABC = 5 To 3 Step -2
        If level(ABC) > -3 And palive(ABC) = True Then
            need = Abs(level(ABC)) * 40 + 20
            If dpoints < need Then
                GoTo exitsub
            Else
                dpoints = dpoints - need
                Labdpoints = dpoints
                level(ABC) = level(ABC) - 1
                lablevel(ABC) = level(ABC)
                AIusepoints
            End If
        End If
    Next
    For ABC = 2 To 6 Step 2
        If level(ABC) < 3 And palive(ABC) = True Then
            need = Abs(level(ABC)) * 40 + 20
            If dpoints < need Then
                GoTo exitsub
            Else
                dpoints = dpoints - need
                Labdpoints = dpoints
                level(ABC) = level(ABC) + 1
                lablevel(ABC) = level(ABC)
                AIusepoints
            End If
        End If
    Next
    For ABC = 2 To 6 Step 2
        If level(ABC) < 5 And palive(ABC) = True Then
            need = Abs(level(ABC)) * 40 + 20
            If dpoints < need Then
                GoTo exitsub
            Else
                dpoints = dpoints - need
                Labdpoints = dpoints
                level(ABC) = level(ABC) + 1
                lablevel(ABC) = level(ABC)
                AIusepoints
            End If
        End If
    Next
End If
exitsub:
End Sub

Private Sub Getaitop()
    For def = 6 To 2 Step -1
        If palive(def) = True Then
            aitop = def + 1
            GoTo exits
        End If
    Next
exits:
End Sub
