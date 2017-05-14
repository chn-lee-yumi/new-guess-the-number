VERSION 5.00
Begin VB.Form b4 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "传统玩法-推理模式"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6345
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   48
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "玩法"
      Height          =   2415
      Left            =   120
      TabIndex        =   46
      Top             =   720
      Width           =   4455
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   $"b53.frx":0000
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   43
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame f1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "结果"
      Height          =   3015
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1575
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "结果"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   45
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "数字"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   0
         Width           =   735
      End
      Begin VB.Label result 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   42
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label result 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   41
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label result 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   40
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label result 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   39
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label result 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   38
         Top             =   960
         Width           =   735
      End
      Begin VB.Label result 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   37
         Top             =   600
         Width           =   735
      End
      Begin VB.Label result 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin VB.Label result 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   35
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   30
         Left            =   240
         TabIndex        =   34
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   29
         Left            =   120
         TabIndex        =   33
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   31
         Left            =   360
         TabIndex        =   32
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   28
         Left            =   480
         TabIndex        =   31
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   27
         Left            =   360
         TabIndex        =   30
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   29
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   28
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   24
         Left            =   480
         TabIndex        =   27
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   23
         Left            =   360
         TabIndex        =   26
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   20
         Left            =   480
         TabIndex        =   23
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   19
         Left            =   360
         TabIndex        =   22
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   16
         Left            =   480
         TabIndex        =   19
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   15
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   15
         Top             =   960
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   32
         Left            =   480
         TabIndex        =   5
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   135
      End
      Begin VB.Label numberlabel 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.TextBox TextInput 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "在此输入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "b4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RndNumber As String
Dim Counter As Integer
Private Sub CmdOK_Click()
Dim InputString As String
Dim InputArray(1 To 4) As String
Dim x As Integer, A As Integer, b As Integer
Dim getjf As Integer

Cls '清除开G痕迹

InputString = TextInput.Text

ReadText InputString, InputArray()

For x = 1 To 4
  If Asc(InputArray(x)) < 48 Or Asc(InputArray(x)) > 57 Then
  MsgBox "输入的必须是数字!", vbExclamation
  GoTo endl
  Else
  End If
Next

If CheckRepetition(InputString) = True Then
  MsgBox "数字各位不能重复!", vbExclamation
  GoTo endl
Else
End If

Counter = Counter + 1

For x = 1 To 4
  numberlabel((Counter - 1) * 4 + x).Caption = InputArray(x)
Next

ReturnAandB RndNumber, InputString, A, b
result(Counter).Caption = A & "A" & b & "B"
 
If result(Counter).Caption = "4A0B" Then
  MsgBox "你成功猜出了该数字！", vbInformation, "很遗憾！"
  getjf = 9 - Counter
  jf = jf + getjf
  savedj
  MsgBox "你获得了" & getjf & "积分。", vbInformation, "消息"
  Call Reset
ElseIf Counter = 8 Then
  MsgBox "你未能猜出正确的数字，哈哈哈！" & vbCrLf & "正确的数字是:" & RndNumber & "。", vbInformation, "恭喜你！"
  Call Reset
End If

endl:
TextInput.SetFocus
TextInput.Text = ""
End Sub

Private Sub Command1_Click()
Dim xx As Integer
xx = MsgBox("你确定要返回？", vbInformation + vbYesNo, "消息")
If xx = vbYes Then
    Unload Me
    a2.Show
End If
End Sub

Private Sub Form_Activate()
TextInput.SetFocus
End Sub

Private Sub Form_DblClick()
Print RndNumber
End Sub

Private Sub Form_Load()
Randomize Time
Call Reset
End Sub
Private Sub ReadText(TextString As String, TextArray() As String)
Dim x As Integer
For x = 1 To 4
  TextArray(x) = Mid(TextString, x, 1)
Next
End Sub
Private Function CheckRepetition(TextString As String) As Boolean
Dim j As Integer, k As Integer, x As Integer
Dim TextArray(1 To 4) As String
ReadText TextString, TextArray()

CheckRepetition = False

For j = 1 To 3
  k = j + 1
  Do Until k > 4
  If TextArray(k) = TextArray(j) Then CheckRepetition = True
  k = k + 1
  Loop
Next

End Function
Private Sub ReturnAandB(RightNumber As String, InputNumber As String, A As Integer, b As Integer)
Dim j As Integer, k As Integer
Dim RightNumberArray(1 To 4) As String, InputNumberArray(1 To 4) As String

ReadText RightNumber, RightNumberArray()
ReadText InputNumber, InputNumberArray()

For j = 1 To 4
  For k = 1 To 4
  If InputNumberArray(k) = RightNumberArray(j) And k = j Then
  A = A + 1
  ElseIf InputNumberArray(k) = RightNumberArray(j) And k <> j Then
  b = b + 1
  End If
  Next
Next

End Sub
Private Sub Reset()
Dim m As Integer, x As Boolean

Counter = 0

For m = 1 To 32
  numberlabel(m).Caption = ""
Next

For m = 1 To 8
  result(m).Caption = ""
Next

x = True

Do While x = True
    
  RndNumber = Format(Int(Rnd * 10000), "0000")
  x = CheckRepetition(RndNumber)
    
Loop

End Sub

Private Sub TextInput_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdOK_Click
    Exit Sub
End If
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
    KeyAscii = 0
End If
End Sub
