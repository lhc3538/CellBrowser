VERSION 5.00
Begin VB.Form SearchBox 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timermove 
      Enabled         =   0   'False
      Interval        =   64
      Left            =   1800
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体-PUA"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton Commandsearch 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Comsearch 
      Height          =   615
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   0
      Picture         =   "SearchBox.frx":0000
      Top             =   0
      Width           =   5160
   End
End
Attribute VB_Name = "SearchBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SearchShow As Boolean

Private Sub Commandsearch_Click()
Comsearch_Click
End Sub

Private Sub Comsearch_Click()

    FrmWeb(PageNum).ComUrl.Text = "http://www.google.com.hk/search?hl=zh-CN&newwindow=1&safe=strict&tbo=d&site=webhp&source=hp&q=" & Text1.Text
    FrmWeb(PageNum).Show
    FrmWeb(PageNum).WebID.Text = Str(PageNum)
    PageNum = PageNum + 1
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()

SetWindowLong hwnd, (-20), &H80000
SetLayeredWindowAttributes Me.hwnd, vbRed, 5, &H1
'-删除 红色背景
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为在所有窗口前端
Me.Width = 0
End Sub



Private Sub Text1_LostFocus()
SearchShow = False
Timermove.Enabled = True
End Sub

Private Sub Timermove_Timer()
 If SearchShow = True Then
  Me.Width = Me.Width + (Image1.Width - Me.Width) / 2
  If Me.Width >= Image1.Width Then
   Me.Width = Image1.Width + Image1.Top
   Timermove.Enabled = False
  End If
 Else
  Me.Width = Me.Width / 2
  If Me.Width <= 1 Then
   Me.Width = 0
   Timermove.Enabled = False
  End If
 End If
End Sub
