VERSION 5.00
Begin VB.Form b 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����淨"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7815
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command6 
      Caption         =   "  ����֮��    (��������)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��ƭģʽ"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�ֻ�ģʽ"
      Height          =   615
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����ģʽ"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "һЩ���ߵ�˵��"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton c5 
      Caption         =   "��ģʽ"
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton c2 
      Caption         =   "  ��սģʽ    (��4��)"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton c1 
      Caption         =   "����ģʽ"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton c3 
      Caption         =   "RPģʽ"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "�����ṩ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   3840
      TabIndex        =   11
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3840
      TabIndex        =   10
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub c1_Click()
b1.Show
Unload Me
End Sub
Private Sub c2_Click()
b2.Show
Unload Me
End Sub
Private Sub c3_Click()
Dim yesno As Integer
yesno = MsgBox("�淨����Ҫһ��������С��Χ��ֱ���������֡�(���ְ�����Χ����)" & vbCrLf & "���Ӯ�ˣ������һ�����ĵ��ߣ�" & vbCrLf & "������ˣ��㽫�ᱻ��10���֡�" & vbcelf & "��֮ǰ��ȷ������10���֡�", 48 + vbOKCancel, "��ȷ��Ҫ����")
If yesno = vbOK Then
    If jf < 10 Then MsgBox "��û��10���֣������㲻���档": Exit Sub
    'Shell "shutdown -s -t 120"
    b3.Show
    Unload Me
End If
End Sub


Private Sub c5_Click()
b51.Show
Unload Me
End Sub

Private Sub Command1_Click()
Shell "notepad.exe " & App.Path & "\data\����˵��(����)", 1
End Sub

Private Sub Command2_Click()
Unload Me
main.Show
End Sub

Private Sub Command3_Click()
Unload Me
b6.Show
End Sub

Private Sub Command4_Click()
Unload Me
b7.Show
End Sub

Private Sub Command5_Click()
Unload Me
b8.Show
End Sub

Private Sub Form_Load()
Set b1 = Nothing
Set b2 = Nothing
Set b3 = Nothing
Set b4 = Nothing
Set b51 = Nothing
Set b6 = Nothing
Set b7 = Nothing
Set b8 = Nothing
Label1.Caption = "����ʵ����(AIʵ����)��CHN-Lee-����" & vbCrLf & _
"��սģʽ��CHN-Lee-����" & vbCrLf & _
"RPģʽ��CHN-Lee-����" & vbCrLf & _
"��ģʽ��������Crj." & vbCrLf & _
"����ģʽ�����˶����뵽�ɣ�(����)" & vbCrLf & _
"�ֻ�ģʽ��CHN-Lee-����" & vbCrLf & _
"��ƭģʽ��CHN-Lee-����" & vbCrLf & _
"����֮����������Crj."
End Sub

