VERSION 5.00
Begin VB.Form b52 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����淨-��ģʽ"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6105
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command6 
      Caption         =   "��˷���"
      Height          =   420
      Left            =   4920
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����˵��"
      Height          =   420
      Left            =   3480
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox t2 
      Height          =   2775
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "��ģʽ.frx":0000
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   975
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
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Enabled         =   0   'False
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   3375
      Begin VB.CommandButton Command4 
         Caption         =   "��ɱ��"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "��ͨ����"
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "����"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   720
      Picture         =   "��ģʽ.frx":0009
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
   Begin VB.Shape hpme 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   4
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   135
   End
   Begin VB.Shape hpme 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   3
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   135
   End
   Begin VB.Shape hpme 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   2
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   135
   End
   Begin VB.Shape hpme 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   1
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   135
   End
   Begin VB.Shape mpme 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   5
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape mpme 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   4
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape mpme 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   3
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape mpme 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   2
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape mpme 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   1
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape mphe 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   5
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape mphe 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   4
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape mphe 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   3
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape mphe 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   2
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape mphe 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   1
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape hphe 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   4
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape hphe 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   3
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape hphe 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   2
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape hphe 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   1
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape mpme 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   0
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape hpme 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   0
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "��"
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
      Index           =   4
      Left            =   1440
      TabIndex        =   16
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ѫ"
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
      Index           =   3
      Left            =   1440
      TabIndex        =   15
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   720
      Picture         =   "��ģʽ.frx":09B7
      Top             =   2760
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   0
      TabIndex        =   14
      Top             =   2760
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   3360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   3360
      Y1              =   840
      Y2              =   1800
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
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label l4 
      Alignment       =   2  'Center
      Caption         =   "20"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   960
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
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3360
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape mphe 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape hphe 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   135
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ѫ"
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
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "��"
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
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label l5 
      Caption         =   "��Χ��   ~"
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
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "b52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dhp As Integer, wmp As Integer, whp As Integer, dmp As Integer
Dim x As Integer, judge As Single, t As Integer, tai As Integer
Dim min As Integer, max As Integer
Dim ran As Double
Dim ainame As String
'choose��ֵ��1-��,2-������,3-�����ر���״̬

Private Sub Command1_Click()
Cls
guess = Val(t1)
t1 = ""
If guess <= min Or guess >= max Then
    MsgBox "����ѡ�������ڷ�Χ�ڣ����������롣", 16, "����"
    Exit Sub
End If
    t = t + 1
If guess > min And guess < max Then
    If guess > x Then
        MsgBox "��ϲ�㡪���´���,̫����", 64, "��ʾ"
        max = guess
        l4 = max
        guess = getaiguess(choose)
        t2 = t2 & vbCrLf & ainame & "�µ�������" & guess
        If guess > x Then
            max = guess
            l4 = max

        ElseIf guess < x Then
            min = guess
            l3 = min

        ElseIf guess = x Then
            t2 = t2 & vbCrLf & ainame & "�����ˡ�"
            ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
            min = 0
            max = 20
            l3 = "0"
            l4 = "20"
            x = Int(Rnd * 19 + 1)
            t2 = t2 & vbCrLf & ainame & "�����һ���غϡ�"
            getai choose, False
        End If
    ElseIf guess < x Then
        MsgBox "��ϲ�㡪���´���,̫С��", 64, "��ʾ"
        min = guess
        l3 = min
        guess = getaiguess(choose)
        t2 = t2 & vbCrLf & ainame & "�µ�������" & guess
        t2.SetFocus: t2.SelStart = Len(t2)
        If guess > x Then
            max = guess
            l4 = max

        ElseIf guess < x Then
            min = guess
            l3 = min

        ElseIf guess = x Then
            t2 = t2 & vbCrLf & ainame & "�����ˡ�"
            t2.SetFocus: t2.SelStart = Len(t2)
            ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
            min = 0
            max = 20
            l3 = "0"
            l4 = "20"
            x = Int(Rnd * 19 + 1)
            t2 = t2 & vbCrLf & ainame & "�����һ���غϡ�"
            t2.SetFocus: t2.SelStart = Len(t2)
            getai choose, False
        End If
    ElseIf guess = x Then
        MsgBox "���ź�,��¶���!!!(���õ�RP���!)" & vbCrLf & "������һ���غ�", 64, "��Ϣ"
        ran = Minute(Time): Randomize ran 'ran = Minute(Time):Randomize ran������
        min = 0
        max = 20
        l3 = "0"
        l4 = "20"
        x = Int(Rnd * 19 + 1)
        t2 = t2 & vbCrLf & "������һ���غϡ�"
        t2.SetFocus: t2.SelStart = Len(t2)
        Command1.Enabled = False
        Frame1.Enabled = True
    End If
End If
t1.SetFocus
End Sub

Private Sub Command2_Click()
If wmp = 6 Then
    MsgBox "��������ˣ�"
Else
    wmp = wmp + 1
    t2 = t2 & vbCrLf & "�����һ������"
    mpme(wmp - 1).Visible = True
    Command1.Enabled = True
    Frame1.Enabled = False
    t2.SetFocus: t2.SelStart = Len(t2)
    t1.SetFocus
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If wmp > 0 Then
    t2 = t2 & vbCrLf & "���" & ainame & "�����˹�����"
    t2.SetFocus: t2.SelStart = Len(t2)
    wmp = wmp - 1
    mpme(wmp).Visible = False
    getai choose, True
    Command1.Enabled = True
    Frame1.Enabled = False
    t1.SetFocus
Else
    MsgBox "��û������"
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If wmp > 2 Then
    wmp = wmp - 3
    t2 = t2 & vbCrLf & "���" & ainame & "������ɱ��"
    t2.SetFocus: t2.SelStart = Len(t2)
    dhp = dhp - 2
        If dhp < 1 Then
        play "data\laugh1.mp3"
        MsgBox "��Ӯ�ˣ�"
        calcjf
        Unload Me
        b51.Show
        Exit Sub
    End If
    hphe(dhp).Visible = False
    hphe(dhp + 1).Visible = False
    mpme(wmp).Visible = False
    mpme(wmp + 1).Visible = False
    mpme(wmp + 2).Visible = False
    Command1.Enabled = True
    Frame1.Enabled = False
    t1.SetFocus
Else
    MsgBox "�㲻������"
End If
End Sub

Private Sub Command5_Click()
Shell "notepad.exe " & App.Path & "\data\��ģʽ(����)", 1
End Sub

Private Sub Command6_Click()
Unload Me
b51.Show
End Sub

Private Sub Form_DblClick()
Print x
End Sub

Private Sub Form_Load()
ran = Minute(Time): Randomize ran
x = Int(Rnd * 19 + 1)
min = 0
max = 20
dhp = 5
whp = 5

If choose = 2 Then
Label1.Caption = "Dyf"
Label1.Font.Size = 28
Image1.Picture = LoadPicture(App.Path & "\data\dyf.bmp")
Image1.Move 840
Image1.Width = 735
Label1.Width = 855
ainame = "������"
ElseIf choose = 3 Then
Label1.Caption = "Dyf"
Label1.Font.Size = 28
Image1.Picture = LoadPicture(App.Path & "\data\btdyf.bmp")
Image1.Move 840
Image1.Width = 735
Label1.Width = 855
ainame = "������"
Else
ainame = "��"
End If
End Sub

Private Sub Form_Resize()
t2 = "��¼��"
End Sub

Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
If Command1.Enabled = True Then
    If KeyCode = 13 Then Command1_Click '13�ǻس���
End If
End Sub





'����Ϊ��ģʽAI����===========================================================================================
Public Sub dyf(att As Boolean)
If att = True Then
    If dhp < 4 And dhp > 2 Then
        t2 = t2 & vbCrLf & "��������2�����ֵ��˹�����"
        t2.SetFocus: t2.SelStart = Len(t2)
        dmp = dmp - 2
        mphe(dmp).Visible = False
        mphe(dmp + 1).Visible = False
    Else
        t2 = t2 & vbCrLf & "�����ص���һ��Ѫ��"
        t2.SetFocus: t2.SelStart = Len(t2)
        dhp = dhp - 1
        If dhp < 1 Then
            play "data\laugh1.mp3"
            MsgBox "��Ӯ�ˣ�"
            calcjf
            Unload Me
            b51.Show
            Exit Sub
        End If
        hphe(dhp).Visible = False
    End If
Else
    If dmp = 0 Then
        dmp = dmp + 1
        mphe(dmp - 1).Visible = True
        t2 = t2 & vbCrLf & "�����س���һ������"
        t2.SetFocus: t2.SelStart = Len(t2)
    ElseIf whp = 5 Then
        dmp = dmp - 1
        mphe(dmp).Visible = False
        t2 = t2 & vbCrLf & "�����ؽ�����ͨ������"
        t2.SetFocus: t2.SelStart = Len(t2)
        If wmp > 1 Then
            defend = MsgBox("�����ؽ�������ͨ���������Ƿ�ֵ���", vbYesNo + vbInformation)
            If defend = vbYes Then
                t2 = t2 & vbcelf & "��ֵ��˹�����"
                t2.SetFocus: t2.SelStart = Len(t2)
                wmp = wmp - 2
                mpme(wmp).Visible = False
                mpme(wmp + 1).Visible = False
            Else
                t2 = t2 & vbCrLf & "�����һ��Ѫ��"
                t2.SetFocus: t2.SelStart = Len(t2)
                whp = whp - 1
                hpme(whp).Visible = False
                If whp = 0 Then play "data\laugh2.mp3": MsgBox "�����ˣ�": Unload Me: b51.Show: Exit Sub
            End If
        Else
            t2 = t2 & vbCrLf & "�����һ��Ѫ��"
            t2.SetFocus: t2.SelStart = Len(t2)
            whp = whp - 1
            hpme(whp).Visible = False
            If whp = 0 Then play "data\laugh2.mp3": MsgBox "�����ˣ�": Unload Me: b51.Show: Exit Sub
        End If
    ElseIf dmp < 3 Then
        dmp = dmp + 1
        mphe(dmp - 1).Visible = True
        t2 = t2 & vbCrLf & "�����س���һ������"
        t2.SetFocus: t2.SelStart = Len(t2)
    ElseIf dmp > 2 Then
        dmp = dmp - 3
        mphe(dmp).Visible = False
        mphe(dmp + 1).Visible = False
        mphe(dmp + 2).Visible = False
        t2 = t2 & vbCrLf & "���������˱�ɱ����" & vbCrLf & "���������Ѫ��"
        t2.SetFocus: t2.SelStart = Len(t2)
        whp = whp - 2
        If whp = 0 Then play "data\laugh2.mp3": MsgBox "�����ˣ�": Unload Me: b51.Show: Exit Sub
        hpme(whp).Visible = False
        hpme(whp + 1).Visible = False
    Else
        MsgBox "error"
    End If
End If
End Sub
Public Sub btdyf(att As Boolean)
If att = True Then
    t2 = t2 & vbCrLf & "�����ص���һ��Ѫ��"
    t2.SetFocus: t2.SelStart = Len(t2)
    dhp = dhp - 1
    If dhp < 1 Then
        play "data\laugh1.mp3"
        MsgBox "��Ӯ�ˣ�"
        calcjf
        Unload Me
        b51.Show
        Exit Sub
    End If
    hphe(dhp).Visible = False
Else
    If dmp = 0 Then
        dmp = dmp + 1
        mphe(dmp - 1).Visible = True
        t2 = t2 & vbCrLf & "�����س���һ������"
        t2.SetFocus: t2.SelStart = Len(t2)
    ElseIf whp = 5 Then
        dmp = dmp - 1
        mphe(dmp).Visible = False
        t2 = t2 & vbCrLf & "�����ؽ�����ͨ������"
        t2.SetFocus: t2.SelStart = Len(t2)
        If wmp > 1 Then
            defend = MsgBox("�����ؽ�������ͨ���������Ƿ�ֵ���", vbYesNo + vbInformation)
            If defend = vbYes Then
                t2 = t2 & vbcelf & "��ֵ��˹�����"
                t2.SetFocus: t2.SelStart = Len(t2)
                wmp = wmp - 2
                mpme(wmp).Visible = False
                mpme(wmp + 1).Visible = False
            Else
                t2 = t2 & vbCrLf & "�����һ��Ѫ��"
                t2.SetFocus: t2.SelStart = Len(t2)
                whp = whp - 1
                hpme(whp).Visible = False
                If whp = 0 Then play "data\laugh2.mp3": MsgBox "�����ˣ�": Unload Me: b51.Show: Exit Sub
            End If
        Else
            t2 = t2 & vbCrLf & "�����һ��Ѫ��"
            t2.SetFocus: t2.SelStart = Len(t2)
            whp = whp - 1
            hpme(whp).Visible = False
            If whp = 0 Then play "data\laugh2.mp3": MsgBox "�����ˣ�": Unload Me: b51.Show: Exit Sub
        End If
    ElseIf dmp < 3 Then
        dmp = dmp + 1
        mphe(dmp - 1).Visible = True
        t2 = t2 & vbCrLf & "�����س���һ������"
        t2.SetFocus: t2.SelStart = Len(t2)
    ElseIf dmp = 3 Then
        dmp = dmp - 3
        mphe(dmp).Visible = False
        mphe(dmp + 1).Visible = False
        mphe(dmp + 2).Visible = False
        t2 = t2 & vbCrLf & "���������˱�ɱ����" & vbCrLf & "���������Ѫ��"
        t2.SetFocus: t2.SelStart = Len(t2)
        whp = whp - 2
        If whp = 0 Then play "data\laugh2.mp3": MsgBox "�����ˣ�": Unload Me: b51.Show: Exit Sub
        hpme(whp).Visible = False
        hpme(whp + 1).Visible = False
    Else
        MsgBox "error"
    End If
End If
End Sub
Public Sub fg(att As Boolean) 'hphe As Integer, mphe As Integer, hpme As Integer, mpme As Integer)
Dim defend As Integer
judge = Rnd
If att = True Then
    If dmp >= 2 Then
        t2 = t2 & vbCrLf & "����2�����ֵ��˹�����"
        t2.SetFocus: t2.SelStart = Len(t2)
        dmp = dmp - 2
        mphe(dmp).Visible = False
        mphe(dmp + 1).Visible = False
    Else
    t2 = t2 & vbCrLf & "�����һ��Ѫ��"
    t2.SetFocus: t2.SelStart = Len(t2)
    dhp = dhp - 1
    If dhp < 1 Then
        play "data\laugh1.mp3"
        MsgBox "��Ӯ�ˣ�"
        calcjf
        Unload Me
        b51.Show
        Exit Sub
    End If
    hphe(dhp).Visible = False
    End If
Else
    If dmp = 0 Or judge > 0.75 And dmp < 6 Then
        dmp = dmp + 1
        mphe(dmp - 1).Visible = True
        t2 = t2 & vbCrLf & "�����һ������"
        t2.SetFocus: t2.SelStart = Len(t2)
    ElseIf judge > 0.35 And dmp >= 3 And dhp >= 2 Then
        dmp = dmp - 3
        mphe(dmp).Visible = False
        mphe(dmp + 1).Visible = False
        mphe(dmp + 2).Visible = False
        t2 = t2 & vbCrLf & "�����˱�ɱ����" & vbCrLf & "���������Ѫ��"
        t2.SetFocus: t2.SelStart = Len(t2)
        whp = whp - 2
        If whp = 0 Then play "data\laugh2.mp3": MsgBox "�����ˣ�": Unload Me: b51.Show: Exit Sub
        hpme(whp).Visible = False
        hpme(whp + 1).Visible = False
    Else
        dmp = dmp - 1
        mphe(dmp).Visible = False
        t2 = t2 & vbCrLf & "�������ͨ������"
        t2.SetFocus: t2.SelStart = Len(t2)
        If wmp > 1 Then
            defend = MsgBox("���������ͨ���������Ƿ�ֵ���", vbYesNo + vbInformation)
            If defend = vbYes Then
                t2 = t2 & vbcelf & "��ֵ��˹�����"
                t2.SetFocus: t2.SelStart = Len(t2)
                wmp = wmp - 2
                mpme(wmp).Visible = False
                mpme(wmp + 1).Visible = False
            Else
                t2 = t2 & vbCrLf & "�����һ��Ѫ��"
                t2.SetFocus: t2.SelStart = Len(t2)
                whp = whp - 1
                hpme(whp).Visible = False
                If whp = 0 Then play "data\laugh2.mp3": MsgBox "�����ˣ�": Unload Me: b51.Show: Exit Sub
            End If
        Else
            t2 = t2 & vbCrLf & "�����һ��Ѫ��"
            t2.SetFocus: t2.SelStart = Len(t2)
            whp = whp - 1
            hpme(whp).Visible = False
            If whp = 0 Then play "data\laugh2.mp3": MsgBox "�����ˣ�": Unload Me: b51.Show: Exit Sub
        End If
    End If
End If
End Sub


Public Function getaiguess(choose As Integer) As Integer
If choose = 1 Then
getaiguess = fgguess(min, max)
ElseIf choose = 2 Then
tai = tai + 1
getaiguess = dyfguess(min, max, tai, x)
ElseIf choose = 3 Then
tai = tai + 1
getaiguess = btdyfguess(min, max, tai, x)
End If
End Function

Public Sub getai(choose As Integer, att As Boolean)
If choose = 1 Then
fg (att)
ElseIf choose = 2 Then
dyf (att)
ElseIf choose = 3 Then
btdyf (att)
End If
End Sub

Public Sub calcjf()
Select Case choose '�ӻ���
    Case 1
        jf = jf + 2
        MsgBox "������2���֣�"
    Case 2
        jf = jf + 4
        MsgBox "������4���֣�"
    Case 3
        jf = jf + 6
        MsgBox "������6���֣�"
End Select
savedj
End Sub
