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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If hookId = 0 Then
        hookId = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, App.hInstance, 0)
    ElseIf hookId <> 0 Then
        UnhookWindowsHookEx hookId
        hookId = 0
    End If
End Sub
    
Private Sub Form_Load()
    Timer1.Interval = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hookId <> 0 Then
        UnhookWindowsHookEx hookId
        hookId = 0
    End If
End Sub

Private Sub Timer1_Timer()
    If EventRaised = True Then
        Debug.Print Direction
        'If Direction = True Then Me.Top = Me.Top - 10
        'If Direction = False Then Me.Top = Me.Top + 10
        EventRaised = False
    End If
End Sub
