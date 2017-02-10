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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const PM_REMOVE = &H1

Private Type Msg
 hWnd As Long
 Message As Long
 wParam As Long

 time As Long
End Type

Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Private Declare Function WaitMessage Lib "user32" () As Long
Private bCancel As Boolean
Private Const WM_MOUSEWHEEL = 522

Private Sub ProcessMessages()
 Dim Message As Msg

 Do While Not bCancel
 WaitMessage '等待消息

 If PeekMessage(Message, Me.hWnd, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then '...when the mousewheel is used...

 If Message.wParam < 0 Then '向上滚动
 Me.Top = Me.Top + 240
 Else '向下滚动
 Me.Top = Me.Top - 240
 End If
 End If
  
 DoEvents
 Loop
 End Sub

Private Sub Form_Load()
 Me.AutoRedraw = True
 Me.Print "请使用鼠标滚轮改变本窗体的位置。"


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ProcessMessages
End Sub

Private Sub Form_Unload(Cancel As Integer)
 bCancel = True
End Sub

