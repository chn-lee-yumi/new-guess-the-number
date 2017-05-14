VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4236
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4236
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   3650
         Left            =   6360
         Top             =   720
      End
      Begin VB.Label Label2 
         Caption         =   "使用 Visual Basic 6.0 编写"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4440
         TabIndex        =   6
         Top             =   3120
         Width           =   2565
      End
      Begin VB.Label Label1 
         Caption         =   "2011.12.21开始制作  最近更新：2016.04.16"
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image Image5 
         Height          =   600
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Top             =   1560
         Width           =   600
      End
      Begin VB.Image Image4 
         Height          =   600
         Left            =   1080
         Picture         =   "frmSplash.frx":2FDE
         Top             =   1560
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   600
         Left            =   360
         Picture         =   "frmSplash.frx":5FB0
         Top             =   2280
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   600
         Left            =   1080
         Picture         =   "frmSplash.frx":8F82
         Top             =   2280
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   600
         Left            =   1080
         Picture         =   "frmSplash.frx":BF54
         Top             =   840
         Width           =   600
      End
      Begin VB.Image imgLogo 
         Height          =   600
         Left            =   360
         Picture         =   "frmSplash.frx":EF26
         Top             =   840
         Width           =   600
      End
      Begin VB.Label lblCopyright 
         Caption         =   "使用 Apache License V2"
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "又有新想法了~~！本版本新增【街机模式】和【欺骗模式】  （CHN-Lee-玉米 制作）"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   3600
         Width           =   6885
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "版本"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6240
         TabIndex        =   3
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "产品"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   32.4
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2400
         TabIndex        =   4
         Top             =   1200
         Width           =   2430
      End
   End
End
Attribute VB_Name = "frmSplash"
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
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    main.Show
    Unload Me
End Sub

Private Sub Form_Load()
    play "data\start.mp3"
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    main.Show
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Unload Me
    main.Show
End Sub
