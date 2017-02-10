VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FrmWeb 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "BlankPage"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   Icon            =   "FrmWeb.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   0
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   3015
      LargeChange     =   10
      Left            =   4680
      TabIndex        =   4
      Top             =   1000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   2
      Top             =   1000
      Visible         =   0   'False
      Width           =   4695
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   2955
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   195
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   308
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   4657
      End
   End
   Begin VB.Timer RefreshWebname 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   0
   End
   Begin SHDocVwCtl.WebBrowser WebPage 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      ExtentX         =   7646
      ExtentY         =   4683
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
   Begin VB.CommandButton WebImage 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "FrmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngLeft As Long
Private mlngTop As Long
Private mlngWidth As Long
Private mlngHeight As Long
'保存webbrowser的大小位置

Dim WithEvents Web_V1 As SHDocVwCtl.WebBrowser_V1   '保证连接在一个窗口打开相关对象
Attribute Web_V1.VB_VarHelpID = -1

Dim OpenID As Integer
Dim WebImageTF As Boolean '判断是否已截屏


Private Sub BrowserFullSize()
    Dim WidthInPixel As Double
    Dim HeightInPixel As Double
    Dim htmDoc As Object

    Set htmDoc = WebPage.Document

    WidthInPixel = htmDoc.body.scrollWidth + (htmDoc.body.offsetWidth - htmDoc.body.ClientWidth)
    HeightInPixel = htmDoc.body.scrollHeight + (htmDoc.body.offsetHeight - htmDoc.body.ClientHeight)

    WebPage.Left = 0
    WebPage.Top = 0
   ' WebPage.Width = WidthInPixel * Screen.TwipsPerPixelX
   ' WebPage.Height = HeightInPixel * Screen.TwipsPerPixelY
    Picture1.Width = WebPage.Width / 15
    Picture1.Height = WebPage.Height / 15
    Set Picture1.Picture = LoadPicture() '清除旧的图像
    Picture1.AutoRedraw = True
    VScroll1.Max = Picture1.ScaleHeight - Picture2.ScaleHeight
    If VScroll1.Max > 0 Then VScroll1.Enabled = True
    HScroll1.Max = Picture1.ScaleWidth - Picture2.ScaleWidth
    If HScroll1.Max > 0 Then HScroll1.Enabled = True
    Set htmDoc = Nothing
End Sub
Private Sub BrowserNormalSize()
    WebPage.Left = mlngLeft
    WebPage.Top = mlngTop
    WebPage.Width = mlngWidth
    WebPage.Height = mlngHeight
End Sub
Private Sub Form_Activate()
   Set Web_V1 = WebPage.Object '把webpage与webV1相联系


   Fa.ActiveUrl.Text = WebPage.LocationURL
   Fa.PageBack(Fa.PublicIndex).Picture = LoadPicture(App.path & "\skin\pageback0.gif")
   Fa.PublicIndex = Val(Me.Caption)
   Fa.PageBackClick
End Sub

Private Sub Form_Load()
Fa.openListpageback.Enabled = True
Me.Top = 0
 If SearchNum > 0 Then
  Me.Left = FormSearch.Width
  Me.Width = Fa.Width * (13 / 19)
 Else
  Me.Left = 0
  Me.Width = Fa.Width
 End If

Me.Height = Fa.Height - (Fa.P1.Height + Fa.P2.Height)
WebPage.Top = 0
WebPage.Left = 0
WebPage.Height = Me.Height

'webimage
    VScroll1.Enabled = False
    HScroll1.Enabled = False
   ' WebPage.Width = Screen.Width
    mlngLeft = WebPage.Left
    mlngTop = WebPage.Top
    mlngWidth = WebPage.Width
    mlngHeight = WebPage.Height

WebImageTF = False
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Fa.PageBack(Val(Me.Caption)).Visible = False
Fa.PageBack(Val(Me.Caption)).Cls
Fa.Page(Val(Me.Caption)).Picture = LoadPicture("")

'添加到恢复记忆中
Fa.ComboRestore.Text = WebPage.Document.Title
Fa.CloseHistory.Text = WebPage.LocationURL
End Sub

Private Sub Form_Resize()
WebPage.Width = Me.Width
WebPage.Height = Me.Height
    mlngHeight = WebPage.Height
'webimage
    Picture2.Move 0, 0, Me.ScaleWidth - VScroll1.Width, Me.ScaleHeight - HScroll1.Height
    VScroll1.Move Me.ScaleWidth - VScroll1.Width, 0, VScroll1.Width, Me.ScaleHeight - WebImage.Height
    HScroll1.Move 0, Me.ScaleHeight - HScroll1.Height, Me.ScaleWidth - WebImage.Width
    WebImage.Move Me.ScaleWidth - WebImage.Width, Me.ScaleHeight - WebImage.Height
    VScroll1.Max = Picture1.ScaleHeight - Picture2.ScaleHeight
    If VScroll1.Max > 0 Then VScroll1.Enabled = True
    HScroll1.Max = Picture1.ScaleWidth - Picture2.ScaleWidth
    If HScroll1.Max > 0 Then HScroll1.Enabled = True
End Sub



Private Sub Form_Unload(Cancel As Integer)
    ShowCaret (HScroll1.hWnd)
    ShowCaret (VScroll1.hWnd)
End Sub








Private Sub RefreshWebname_Timer()

On Error Resume Next
Fa.WebName(Val(Me.Caption)) = WebPage.Document.Title

If WebPage.Busy = False Then

  WebImage.SetFocus
 WebImage_Click '截取网页
  RefreshWebname.Enabled = False

End If

Fa.Page(Val(Me.Caption)).Picture = Picture1.Image '将网页显示在列表中
Picture1.Cls
Picture2.Cls
End Sub





Private Sub Web_V1_NavigateComplete(ByVal URL As String)
Fa.ActiveUrl.Text = WebPage.LocationURL
End Sub

Private Sub Web_V1_NewWindow(ByVal URL As String, _
    ByVal flags As Long, ByVal TargetFrameName As String, _
    PostData As Variant, ByVal Headers As String, Processed As Boolean) '打开新窗口时触发
 Processed = True
Fa.OpenNewPage.Text = URL
End Sub


Private Sub WebImage_Click()

 
 
On Error Resume Next
Dim lngHwnd As Long
    
    '完整显示整个网页
    BrowserFullSize
    
    '获得WebBrowser的hwnd
    lngHwnd = FindWindowEx(Me.hWnd, 0, "Shell Embedding", vbNullString)
    lngHwnd = FindWindowEx(lngHwnd, 0, "Shell DocObject View", vbNullString)
    '网页快照，显示在Picture1中，之后把这个图像保存成文件即可
    SendMessage lngHwnd, WM_PRINT, Picture1.hDC, PRF_CHILDREN Or PRF_CLIENT Or PRF_ERASEBKGND Or PRF_NONCLIENT Or PRF_OWNED
    DoEvents
   ' Picture1.Refresh
    Picture1.Visible = True
'    mnuSave.Enabled = True
 '   mnuClipboard.Enabled = True
    '......
    
    '恢复原来的大小
    BrowserNormalSize
   ' webpage.Visible = False
Form_Resize

End Sub

Private Sub WebPage_GotFocus()
Fa.ActiveUrl.Visible = False
End Sub
Private Sub vScroll1_GotFocus()
    HideCaret (VScroll1.hWnd)
End Sub

Private Sub vScroll1_LostFocus()
    HideCaret (VScroll1.hWnd)
End Sub

Private Sub HScroll1_GotFocus()
    HideCaret (HScroll1.hWnd)
End Sub

Private Sub HScroll1_LostFocus()
    HideCaret (HScroll1.hWnd)
End Sub

Private Sub HScroll1_Scroll()
Picture1.Left = -HScroll1.Value
 HideCaret (HScroll1.hWnd)
End Sub

Private Sub VScroll1_Scroll()
Picture1.Top = -VScroll1.Value
 HideCaret (VScroll1.hWnd)
End Sub
Private Sub VScroll1_Change()
Picture1.Top = -VScroll1.Value
 HideCaret (VScroll1.hWnd)
End Sub
Private Sub HScroll1_Change()
Picture1.Left = -HScroll1.Value
 HideCaret (HScroll1.hWnd)
End Sub



Private Sub WebPage_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
WebPage.Silent = True
pDisp.Document.parentWindow.execScript "window.alert=null;"
pDisp.Document.parentWindow.execScript "window.confirm=null;"

End Sub
