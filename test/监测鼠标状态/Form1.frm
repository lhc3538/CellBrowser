VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public LMouseState As Long, RMouseState As Long, LastL As Long, LastR As Long
Private Sub Form_Load()



Timer1.Enabled = True
Timer1.Interval = 1000
Me.AutoRedraw = True
End Sub
Private Sub Timer1_Timer()
LMouseState = GetAsyncKeyState(vbKeyLButton)
If LMouseState < 0 Then Print "鼠标 左键 处于 按住 状态"
If LMouseState = 1 Then Print "鼠标 左键 按过一次"
If LastL < 0 And LMouseState = 0 Then
Print "鼠标 左键 已 抬起 （按过一次）"
LastL = LMouseState
Else
LastL = LMouseState
End If
RMouseState = GetAsyncKeyState(vbKeyRButton)
If RMouseState < 0 Then Print "鼠标 右键 处于 按住 状态"
If RMouseState = 1 Then Print "鼠标 右键 按过一次"
If LastR < 0 And RMouseState = 0 Then
Print "鼠标 右键 已 抬起 （按过一次）"
LastR = RMouseState
Else
LastR = RMouseState
End If
End Sub
