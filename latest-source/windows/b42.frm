VERSION 5.00
Begin VB.Form b62 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����淨-����ģʽ-���޶�"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5730
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "��˷���"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����ʼ"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "ע�⣺��ֻ��һ��֮ǰд��VBS�ű�����Ӯ�뿴�ű����룡������data�ļ��У�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "  �����ռ��˹����ܣ�����ģʽ��"
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
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "b62"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iguess123 As Integer
Dim aiguess123 As Integer
Dim min As Integer
Dim max As Integer
Dim x As Integer

Function bossdie(x, min, max, iguess)
bossdie = x + x - iguess123
If bossdie >= max Then
bossdie = (max - 1)
ElseIf bossdie <= min Then
bossdie = (min + 1)
End If
End Function

Sub aa()
On Error GoTo errline
    Do
iguess123 = InputBox("���������֣���Χ" & min & "~" & max & "��", "�����")
    Do While iguess123 >= max Or iguess123 <= min
MsgBox "��µ����ڷ�Χ��!" & vbcelf & "���������롣", 48, "(error0x000015" & x & ")"
iguess123 = InputBox("���������֣���Χ" & min & "~" & max & "��", "�����")
iguess123 = iguess123 + 0
    Loop

If iguess123 = x Then
MsgBox "�������!", 48, "������!"
Exit Sub
ElseIf iguess123 < x Then
min = iguess123
MsgBox "̫С��!"
ElseIf iguess123 > x Then
max = iguess123
MsgBox "̫����!"
End If

aiguess123 = bossdie(x, min, max, iguess)
MsgBox "�ռ��˹����ܲµ���" & aiguess123
If aiguess123 = x Then
MsgBox "�ռ��˹����ܲ�����!", 48, "��Ӯ��!"
Exit Sub
ElseIf aiguess123 < x Then
min = aiguess123
MsgBox "̫С��!"
ElseIf aiguess123 > x Then
max = aiguess123
MsgBox "̫����!"
End If
    Loop

errline:
MsgBox "�������û����������ֹͣ������Ǹ��������"
End Sub

Private Sub Command1_Click()
Unload Me
b6.Show
MsgBox "��ӭ��ս�ռ��˹����ܣ�����ģʽ1vs1��!! ", , "��ӭʹ�ñ�����vbs�ű�= =��! "
Randomize (Minute(Time))
x = Int(Rnd * 99 + 1)
min = 0
max = 100
aa
End Sub

Private Sub Command2_Click()
Unload Me
b.Show
End Sub
