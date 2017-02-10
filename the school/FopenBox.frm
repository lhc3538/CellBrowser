VERSION 5.00
Begin VB.Form FopenBox 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FopenBox.frx":0000
   ScaleHeight     =   1905
   ScaleMode       =   0  'User
   ScaleWidth      =   615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox ComOpen 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   0
      Width           =   615
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   120
         Top             =   120
      End
   End
End
Attribute VB_Name = "FopenBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fy As Integer
Dim M_downed As Boolean
Dim MyTop As Integer

Private Sub ComOpen_Click()
If Me.Top = MyTop Then '窗体位置未改变时，展开功能窗体
 If Fcbox.GoLeft = False Then
   Fcbox.GoLeft = True
 Else
   Fcbox.GoLeft = False
 End If
End If
Fcbox.Timermove.Enabled = True

End Sub

Private Sub ComOpen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ComOpen.BackColor = vbBlack
Me.ComOpen.Cls
Call PaintPng(App.Path & "\ui\openbox.png", Me.ComOpen.hdc, 0, 0)
Fy = Y
M_downed = True '鼠标按下
MyTop = Me.Top
End Sub

Private Sub ComOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If M_downed = True Then
  Me.Top = Me.Top + (Y - Fy) '移动窗体
 End If
End Sub

Private Sub ComOpen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ComOpen.BackColor = &H404040
Me.ComOpen.Cls
Call PaintPng(App.Path & "\ui\openbox.png", Me.ComOpen.hdc, 0, 0)
M_downed = False '鼠标松开

End Sub

Private Sub Form_Load()
Me.Top = WebPage.Height / 2 - Me.Height / 2
Me.Left = -32
M_downed = False '鼠标松开
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为在所有窗口前端
'SetWindowLong hwnd, (-20), &H80000
'SetLayeredWindowAttributes Me.hwnd, vbRed, 5, &H1
'-删除 红色背景
End Sub

Private Sub Timer1_Timer()
Me.ComOpen.Cls
Call PaintPng(App.Path & "\ui\openbox.png", Me.ComOpen.hdc, 0, 0)
Timer1.Enabled = False
End Sub

Public Sub ShowOpenpic()
Me.ComOpen.Cls
Call PaintPng(App.Path & "\ui\openbox.png", Me.ComOpen.hdc, 0, 0)
End Sub
