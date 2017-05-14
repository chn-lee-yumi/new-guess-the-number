VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "猜数字(欢迎使用！)"
   ClientHeight    =   5340
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4920
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command6 
      Caption         =   "License"
      Height          =   345
      Left            =   150
      TabIndex        =   8
      Top             =   1800
      Width           =   4530
   End
   Begin VB.CommandButton Command5 
      Caption         =   "如何获得积分?"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "清除所有数据"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "9积分换一个后悔药"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton c1 
      Caption         =   "5积分换一个Pass"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "    创新玩法    (强烈推荐)"
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "传统玩法"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "更新内容："
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
         Name            =   "宋体"
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
      Caption         =   "积分："
      BeginProperty Font 
         Name            =   "宋体"
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
'Copyright 2011-9999 CHN-Lee-玉米
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
MsgBox "你不够积分！"
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
MsgBox "你不够积分！"
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
ran = MsgBox("你确定要清除数据?", vbExclamation + vbYesNo, "最后确认")
If ran = vbYes Then Kill (App.Path & "\data\猜数字数据.txt"): MsgBox "数据已清除。", vbInformation, "消息"
End Sub

Private Sub Command5_Click()
Shell "notepad.exe " & App.Path & "\data\积分获取细则.txt", 1
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
"发布日期：2016.4.16" & vbCrLf & vbCrLf & _
"本次更新：2.11.130 → 2.11.132" & vbCrLf & _
"更新内容：" & vbCrLf & _
"1.稍微修改了一些说明文本" & vbCrLf & _
"2.修复了挑战模式爆机后退出程序程序仍在后台运行的BUG" & vbCrLf & _
"3.街机模式使用技能的分数进行了调整" & vbCrLf & vbCrLf & _
"上次更新：2.11.127 → 2.11.130  （2015.12.6）" & vbCrLf & _
"更新内容：" & vbCrLf & _
"1.现在音效会在弹出对话框时播放。以前则是点击对话框后才播放。" & vbCrLf & _
"2.现在玩欺骗模式会提示你获得的积分数" & vbCrLf & _
"3.积分获取途径更多了！(详情请点击“如何获取积分”)" & vbCrLf & _
"4.修复了街机模式胜利时可能出现的运行错误。现在赢了街机模式能正确地加7积分"

On Error GoTo Err
Open App.Path & "\data\猜数字数据.txt" For Input As #1
Input #1, djpass, djkey, djbomb, djcy, djhhy, jf, checkdata
Close

If checkdata <> djpass \ 2 + djkey * 2 + djbomb * 3 + djcy * djhhy + jf + 12 Then
    MsgBox "数据已损坏，请删除“猜数字数据.txt”再打开程序！", 48, "错误！"
    End
    '数据验证，防止被恶意纂改。
End If

Label2 = jf
Exit Sub

Err:
Close
savedj
Label2 = jf
End Sub

