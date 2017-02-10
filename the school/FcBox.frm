VERSION 5.00
Begin VB.Form Fcbox 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9945
   ClientLeft      =   1065
   ClientTop       =   1695
   ClientWidth     =   930
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   930
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timerLoadpic 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   8880
   End
   Begin VB.TextBox ConBack 
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox ComSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1024
      Left            =   0
      Picture         =   "FcBox.frx":0000
      ScaleHeight     =   1020
      ScaleWidth      =   1020
      TabIndex        =   7
      Top             =   4320
      Width           =   1024
   End
   Begin VB.PictureBox ComFavorite 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1024
      Left            =   0
      Picture         =   "FcBox.frx":00F0
      ScaleHeight     =   1020
      ScaleWidth      =   1020
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   1024
   End
   Begin VB.PictureBox ComFrmList 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   0
      ScaleHeight     =   1500
      ScaleWidth      =   1020
      TabIndex        =   5
      ToolTipText     =   "Right"
      Top             =   6480
      Width           =   1024
      Begin VB.Image ImageRight 
         Height          =   1020
         Left            =   120
         Picture         =   "FcBox.frx":06D6
         Top             =   120
         Width           =   1020
      End
      Begin VB.Image ImageLeft 
         Height          =   1020
         Left            =   -195
         Picture         =   "FcBox.frx":07C0
         Top             =   0
         Width           =   1020
      End
   End
   Begin VB.PictureBox ComHome 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1024
      Left            =   0
      Picture         =   "FcBox.frx":08AB
      ScaleHeight     =   1020
      ScaleWidth      =   1020
      TabIndex        =   4
      Top             =   3240
      Width           =   1024
   End
   Begin VB.PictureBox ComRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1024
      Left            =   0
      Picture         =   "FcBox.frx":0BE0
      ScaleHeight     =   1020
      ScaleWidth      =   1020
      TabIndex        =   3
      Top             =   2160
      Width           =   1024
   End
   Begin VB.PictureBox ComDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1024
      Left            =   0
      Picture         =   "FcBox.frx":0CE9
      ScaleHeight     =   1020
      ScaleWidth      =   1020
      TabIndex        =   2
      Top             =   1080
      Width           =   1024
   End
   Begin VB.PictureBox ComUp 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1024
      Left            =   0
      Picture         =   "FcBox.frx":0DE2
      ScaleHeight     =   1020
      ScaleWidth      =   1020
      TabIndex        =   1
      Top             =   0
      Width           =   1024
   End
   Begin VB.Timer Timermove 
      Enabled         =   0   'False
      Interval        =   32
      Left            =   240
      Top             =   9240
   End
   Begin VB.CommandButton ComEnd 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Top             =   6720
      Width           =   375
   End
   Begin VB.Image ImageEnd 
      Height          =   735
      Left            =   120
      MouseIcon       =   "FcBox.frx":0EDB
      MousePointer    =   99  'Custom
      Picture         =   "FcBox.frx":102D
      Stretch         =   -1  'True
      ToolTipText     =   "关闭程序"
      Top             =   7800
      Width           =   735
   End
End
Attribute VB_Name = "Fcbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'各种定义
Public GoLeft As Boolean
Dim Reconback As String '被激活的前一个按钮

Private Sub ComDown_Click()
On Error Resume Next
FrmWeb(ActivePage).Web1.GoForward
End Sub

Private Sub ComDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConBack.Text = "ComDown"
End Sub

Private Sub ComEnd_Click()
bCancel = True '结束鼠标滚轮监测
Unload Fcbox
Unload FrmList
Unload SearchBox
Unload WebPage
 Dim i As Integer
 For i = 0 To PageNum Step 1
  If FrmWeb(i).Caption <> "" Then
     Unload FrmWeb(i)
  End If
 Next i
End
End Sub

Private Sub ComFavorite_Click()
FrmList.Show
End Sub

Private Sub ComFavorite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConBack.Text = "ComFavorite"

End Sub

Private Sub ComFrmList_Click()
FrmList.Show
FrmList.TimerGoRight.Enabled = True
End Sub

Private Sub ComHome_Click()
    FrmWeb(PageNum).Web1.navigate ("http://www.1616.net")
    FrmWeb(PageNum).Show
    FrmWeb(PageNum).WebID.Text = Str(PageNum)
    FrmList.PageBack(PageNum).Visible = True
    PageNum = PageNum + 1
End Sub

Private Sub ComHome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConBack.Text = "ComHome"
End Sub


Private Sub ComFrmlist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConBack.Text = "ComFrmlist"
End Sub

Private Sub ComRefresh_Click()
FrmWeb(ActivePage).Web1.Refresh

End Sub

Private Sub ComRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConBack.Text = "ComRefresh"
End Sub

Private Sub Comsearch_Click()
With SearchBox
 .Show
 .Left = Me.Left + Me.Width
 .Top = Me.Top + ComSearch.Top
 .SearchShow = True
 .Timermove.Enabled = True
 .ZOrder
End With
End Sub



Private Sub ComSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConBack.Text = "ComSearch"
End Sub

Private Sub ComUp_Click()
On Error Resume Next
FrmWeb(ActivePage).Web1.GoBack
End Sub

Private Sub ComUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConBack.Text = "ComUp"
End Sub

Private Sub ConBack_Change()
'Exit Sub '暂时不用此过程
Reset_comBackPic

 With ConBack
   If .Text = "ComRefresh" Then
     ComRefresh.BackColor = RGB(0, 0, 0)
     ''Call PaintPng(App.Path & "\ui\comrefresh.png", Me.ComRefresh.hdc, 0, 0)
   End If
   If .Text = "ComUp" Then
     ComUp.BackColor = RGB(0, 0, 0)
     ''Call PaintPng(App.Path & "\ui\ComUp.png", Me.ComUp.hdc, 0, 0)
   End If
   If .Text = "ComDown" Then
     ComDown.BackColor = RGB(0, 0, 0)
     ''Call PaintPng(App.Path & "\ui\ComDown.png", Me.ComDown.hdc, 0, 0)
   End If
   If .Text = "ComHome" Then
     ComHome.BackColor = RGB(0, 0, 0)
     ''Call PaintPng(App.Path & "\ui\ComHome.png", Me.ComHome.hdc, 0, 0)
   End If
   If .Text = "ComFrmlist" Then
     ComFrmList.BackColor = RGB(0, 0, 0)
     ''Call PaintPng(App.Path & "\ui\ComFrmlist.png", Me.ComFrmList.hdc, 0, 0)
   End If
   If .Text = "ComFavorite" Then
     ComFavorite.BackColor = RGB(0, 0, 0)
     ''Call PaintPng(App.Path & "\ui\ComFavorite.png", Me.ComFavorite.hdc, 0, 0)
   End If
   If .Text = "ComSearch" Then
     ComSearch.BackColor = RGB(0, 0, 0)
     ''Call PaintPng(App.Path & "\ui\ComSearch.png", Me.Comsearch.hdc, 0, 0)
   End If
 End With
Reconback = ConBack.Text
End Sub
Public Sub Reset_comBackPic() '按钮背景颜色恢复
   If Reconback = "ComRefresh" Then
     ComRefresh.BackColor = &H404040
     'Call PaintPng(App.Path & "\ui\comrefresh.png", Me.ComRefresh.hdc, 0, 0)
   End If
   If Reconback = "ComUp" Then
     ComUp.BackColor = &H404040
     'Call PaintPng(App.Path & "\ui\ComUp.png", Me.ComUp.hdc, 0, 0)
   End If
   If Reconback = "ComDown" Then
     ComDown.BackColor = &H404040
     'Call PaintPng(App.Path & "\ui\ComDown.png", Me.ComDown.hdc, 0, 0)
   End If
   If Reconback = "ComHome" Then
     ComHome.BackColor = &H404040
     'Call PaintPng(App.Path & "\ui\ComHome.png", Me.ComHome.hdc, 0, 0)
   End If
   If Reconback = "ComFrmlist" Then
     ComFrmList.BackColor = &H404040
     'Call PaintPng(App.Path & "\ui\ComFrmlist.png", Me.ComFrmList.hdc, 0, 0)
   End If
   If Reconback = "ComFavorite" Then
     ComFavorite.BackColor = &H404040
     'Call PaintPng(App.Path & "\ui\ComFavorite.png", Me.ComFavorite.hdc, 0, 0)
   End If
   If Reconback = "ComSearch" Then
     ComSearch.BackColor = &H404040
     'Call PaintPng(App.Path & "\ui\ComSearch.png", Me.ComSearch.hdc, 0, 0)
   End If
End Sub

'-------------------------------------------------------------------以上为删除背景色API
Private Sub Form_Load()


PageNum = 0 '网页初始数量赋值
If Command() = "" Then
    ComHome_Click
Else
    FrmWeb(PageNum).Web1.navigate Command()
    FrmWeb(PageNum).Show
    FrmWeb(PageNum).WebID.Text = Str(PageNum)
    FrmList.PageBack(PageNum).Visible = True
    PageNum = PageNum + 1
End If

SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为在所有窗口前端
'SetWindowLong hwnd, (-20), &H80000
'SetLayeredWindowAttributes Me.hwnd, vbRed, 5, &H1
'-删除 红色背景
GoLeft = True
'Timermove.Enabled = True '隐藏
'排列窗体
Me.Top = 0
Me.Left = 0
Me.Height = WebPage.Height
ImageEnd.Top = Me.Height - ImageEnd.Height
ComFrmList.Top = ComSearch.Top + ComSearch.Height
ComFrmList.Height = ImageEnd.Top - ComFrmList.Top
ImageLeft.Top = ComFrmList.Height / 2 - ImageLeft.Height / 2
ImageRight.Top = ImageLeft.Top + 120

FrmList.Show

End Sub

Private Sub Imageend_Click()
ComEnd_Click
End Sub

Private Sub TimerGoRight_Timer()
Me.Left = FrmList.Width
End Sub

Private Sub TimerGoRight_Change()
Me.Left = Val(TimerGoRight.Text)
End Sub



Private Sub ImageLeft_Click()
ComFrmList_Click
ConBack.Text = "ComFrmlist"
End Sub

Private Sub ImageRight_Click()
ComFrmList_Click
ConBack.Text = "ComFrmlist"
End Sub

Private Sub timerLoadpic_Timer()
     'Call PaintPng(App.Path & "\ui\comrefresh.png", Me.ComRefresh.hdc, 0, 0)
     'Call PaintPng(App.Path & "\ui\comrefresh.png", Me.ComRefresh.hdc, 0, 0)
     'Call PaintPng(App.Path & "\ui\ComUp.png", Me.ComUp.hdc, 0, 0)
     'Call PaintPng(App.Path & "\ui\ComDown.png", Me.ComDown.hdc, 0, 0)
     'Call PaintPng(App.Path & "\ui\ComHome.png", Me.ComHome.hdc, 0, 0)
     'Call PaintPng(App.Path & "\ui\ComFrmlist.png", Me.ComFrmList.hdc, 0, 0)
     'Call PaintPng(App.Path & "\ui\ComFavorite.png", Me.ComFavorite.hdc, 0, 0)
     'Call PaintPng(App.Path & "\ui\ComSearch.png", Me.ComSearch.hdc, 0, 0)
timerLoadpic.Enabled = False
End Sub

Private Sub Timermove_Timer()

 If GoLeft = False Then '向右展开
 
  Me.Width = Me.Width * 2
  If Me.Width >= 1024 Then
   Me.Width = 1024
   Timermove.Enabled = False
    timerLoadpic.Enabled = True '载入图片
  End If
 End If
 
 If GoLeft = True Then '向左收起
   Me.Width = Me.Width / 2
   If Me.Width <= 60 Then
    Timermove.Enabled = False
    Me.Width = 60
  End If
 End If

FopenBox.Left = Me.Width + Me.Left
End Sub

