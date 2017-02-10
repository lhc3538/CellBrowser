VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FormTranslate 
   BackColor       =   &H00E0E0E0&
   Caption         =   "细胞工具-有道翻译"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11730
   Icon            =   "FormTranslate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   11730
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4920
      Top             =   4200
   End
   Begin VB.PictureBox PBaiDu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   3495
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
      Begin SHDocVwCtl.WebBrowser WebBaiDu 
         Height          =   3375
         Left            =   -30
         TabIndex        =   5
         Top             =   -960
         Width           =   3135
         ExtentX         =   5530
         ExtentY         =   5953
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.PictureBox PYouDAo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3255
      ScaleWidth      =   3735
      TabIndex        =   1
      Top             =   480
      Width           =   3735
      Begin SHDocVwCtl.WebBrowser WebYouDao 
         Height          =   3495
         Left            =   -32
         TabIndex        =   4
         Top             =   -1440
         Width           =   3135
         ExtentX         =   5530
         ExtentY         =   6165
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.PictureBox PGoogle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3255
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
      Begin SHDocVwCtl.WebBrowser WebGoogle 
         Height          =   3375
         Left            =   -32
         TabIndex        =   3
         Top             =   -1360
         Width           =   3135
         ExtentX         =   5530
         ExtentY         =   5953
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Image BaiDu 
      Height          =   495
      Left            =   7800
      MouseIcon       =   "FormTranslate.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "FormTranslate.frx":0E1C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image YouDao 
      Height          =   495
      Left            =   5400
      MouseIcon       =   "FormTranslate.frx":113DA
      MousePointer    =   99  'Custom
      Picture         =   "FormTranslate.frx":1152C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Google 
      Height          =   495
      Left            =   2880
      MouseIcon       =   "FormTranslate.frx":1A2CF
      MousePointer    =   99  'Custom
      Picture         =   "FormTranslate.frx":1A421
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   5280
      Picture         =   "FormTranslate.frx":1C7A1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   5865
      Left            =   -240
      Picture         =   "FormTranslate.frx":1CC50
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   12000
   End
End
Attribute VB_Name = "FormTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Dim MeHeight As Integer
Dim MeWidth As Integer
Dim ClickOne As String
Dim TouMing As Integer



Private Sub Form_Load()
WebYouDao.Navigate "http://fanyi.youdao.com/"
WebGoogle.Navigate "http://translate.google.cn/"
WebBaiDu.Navigate "http://fanyi.baidu.com/"

MeHeight = Me.Height
MeWidth = Me.Width
PGoogle.Width = Me.Width
PGoogle.Height = Me.Height - Google.Height
WebGoogle.Width = Me.Width
WebGoogle.Height = 10000
PYouDAo.Width = Me.Width
PYouDAo.Height = Me.Height - Google.Height
WebYouDao.Width = WebGoogle.Width
WebYouDao.Height = 50000
PBaiDu.Width = Me.Width
PBaiDu.Height = Me.Height - Google.Height
WebBaiDu.Width = Me.Width
WebBaiDu.Height = 10000

TouMing = 255

End Sub


Private Sub Form_Resize()
On Error Resume Next
Me.Height = MeHeight
Me.Width = MeWidth
End Sub

Private Sub Google_Click()
ClickOne = "google"
Timer1.Enabled = True
Image2.Left = Google.Left - 120
Me.Caption = "细胞工具-谷歌翻译"
End Sub

Private Sub WebYouDao_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
Shell App.Path & "\cell.exe " & WebYouDao.Document.activeelement.href


End Sub

Private Sub YouDao_Click()
ClickOne = "youdao"
Timer1.Enabled = True
Image2.Left = YouDao.Left - 120
Me.Caption = "细胞工具-有道翻译"
End Sub

Private Sub BaiDu_Click()
ClickOne = "baidu"
Timer1.Enabled = True
Image2.Left = BaiDu.Left - 120
Me.Caption = "细胞工具-百度翻译"
End Sub

Private Sub WebBaiDu_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
WebBaiDu.Silent = True
pDisp.Document.parentWindow.execScript "window.alert=null;"
pDisp.Document.parentWindow.execScript "window.confirm=null;"
End Sub




Private Sub WebGoogle_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
WebGoogle.Silent = True
pDisp.Document.parentWindow.execScript "window.alert=null;"
pDisp.Document.parentWindow.execScript "window.confirm=null;"
End Sub



Private Sub WebYouDao_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
WebYouDao.Silent = True
pDisp.Document.parentWindow.execScript "window.alert=null;"
pDisp.Document.parentWindow.execScript "window.confirm=null;"
End Sub



Private Sub Timer1_Timer()
If TouMing > 0 Then
TouMing = TouMing - 15
Else
Timer1.Enabled = False
Timer2.Enabled = True
End If
  Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, TouMing, LWA_ALPHA

End Sub

Private Sub Timer2_Timer()
If ClickOne = "google" Then
PYouDAo.Visible = False
PBaiDu.Visible = False
PGoogle.Visible = True
End If
If ClickOne = "baidu" Then
PGoogle.Visible = False
PYouDAo.Visible = False
PBaiDu.Visible = True
End If
If ClickOne = "youdao" Then
PBaiDu.Visible = False
PGoogle.Visible = False
PYouDAo.Visible = True
End If
If TouMing < 255 Then
TouMing = TouMing + 15
Else
Timer2.Enabled = False
End If

  Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, TouMing, LWA_ALPHA
End Sub

