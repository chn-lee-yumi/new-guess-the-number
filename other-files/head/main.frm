VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������(��ӭʹ�ã�)"
   ClientHeight    =   5340
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4920
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command6 
      Caption         =   "License"
      Height          =   345
      Left            =   150
      TabIndex        =   8
      Top             =   1800
      Width           =   4530
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��λ�û���?"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "�����������"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "9���ֻ�һ�����ҩ"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton c1 
      Caption         =   "5���ֻ�һ��Pass"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "    �����淨    (ǿ���Ƽ�)"
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "��ͳ�淨"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�������ݣ�"
      Height          =   3390
      Left            =   90
      TabIndex        =   9
      Top             =   2235
      Width           =   4680
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "jf"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "���֣�"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2011-9999 CHN-Lee-����
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.

Dim t As Integer

Private Sub c1_Click()
If jf > 4 Then
jf = jf - 5
djpass = djpass + 1
Label2 = jf
savedj
Else
MsgBox "�㲻�����֣�"
End If
End Sub

Private Sub Command1_Click()
Unload Me
a2.Show
End Sub

Private Sub Command2_Click()
Unload Me
b.Show
End Sub

Private Sub Command3_Click()
If jf > 8 Then
jf = jf - 9
djhhy = djhhy + 1
Label2 = jf
savedj
Else
MsgBox "�㲻�����֣�"
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
ran = MsgBox("��ȷ��Ҫ�������?", vbExclamation + vbYesNo, "���ȷ��")
If ran = vbYes Then Kill (App.Path & "\data\����������.txt"): MsgBox "�����������", vbInformation, "��Ϣ"
End Sub

Private Sub Command5_Click()
Shell "notepad.exe " & App.Path & "\data\���ֻ�ȡϸ��.txt", 1
End Sub

Private Sub Command6_Click()
frmlicense.Show
End Sub

Private Sub Form_DblClick()
d.Show
'If timer = False Then timer = True
't = t + 1
'If t > 5 Then d1.Show: d2.Show
't = 0
'timer = False
End Sub

Private Sub Form_Load()

Label3 = _
"�������ڣ�2016.4.16" & vbCrLf & vbCrLf & _
"���θ��£�2.11.130 �� 2.11.132" & vbCrLf & _
"�������ݣ�" & vbCrLf & _
"1.��΢�޸���һЩ˵���ı�" & vbCrLf & _
"2.�޸�����սģʽ�������˳�����������ں�̨���е�BUG" & vbCrLf & _
"3.�ֻ�ģʽʹ�ü��ܵķ��������˵���" & vbCrLf & vbCrLf & _
"�ϴθ��£�2.11.127 �� 2.11.130  ��2015.12.6��" & vbCrLf & _
"�������ݣ�" & vbCrLf & _
"1.������Ч���ڵ����Ի���ʱ���š���ǰ���ǵ���Ի����Ų��š�" & vbCrLf & _
"2.��������ƭģʽ����ʾ���õĻ�����" & vbCrLf & _
"3.���ֻ�ȡ;�������ˣ�(������������λ�ȡ���֡�)" & vbCrLf & _
"4.�޸��˽ֻ�ģʽʤ��ʱ���ܳ��ֵ����д�������Ӯ�˽ֻ�ģʽ����ȷ�ؼ�7����"

On Error GoTo Err
Open App.Path & "\data\����������.txt" For Input As #1
Input #1, djpass, djkey, djbomb, djcy, djhhy, jf, checkdata
Close

If checkdata <> djpass \ 2 + djkey * 2 + djbomb * 3 + djcy * djhhy + jf + 12 Then
    MsgBox "�������𻵣���ɾ��������������.txt���ٴ򿪳���", 48, "����"
    End
    '������֤����ֹ��������ġ�
End If

Label2 = jf
Exit Sub

Err:
Close
savedj
Label2 = jf
End Sub

