VERSION 5.00
Begin VB.Form d2 
   BackColor       =   &H00FFFF80&
   Caption         =   "��ȡ�����ֵĴ���"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   3180
   StartUpPosition =   2  '��Ļ����
End
Attribute VB_Name = "d2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim y As Long, x As Long, t As Long

Private Sub Form_Click()
x = InputBox("���������ޣ�������ΧΪ0~x", "����������")
y = x
Do
    y = Round(y / 2)
    t = t + 1
Loop While y <> 1
Print t
t = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Print KeyAscii
If KeyAscii = 13 Then Form_Click
End Sub
