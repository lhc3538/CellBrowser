VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   10965
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2175
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   3836
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
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   3000
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
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   3015
      LargeChange     =   10
      Left            =   4680
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Private mlngLeft As Long
Private mlngTop As Long
Private mlngWidth As Long
Private mlngHeight As Long
'方法一：由WebBrowser直接抓取图片
Private Sub Command1_Click()
    Dim lngHwnd As Long
    
    '完整显示整个网页
    BrowserFullSize
    
    '获得WebBrowser的hwnd
    lngHwnd = FindWindowEx(Me.hwnd, 0, "Shell Embedding", vbNullString)
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
   ' WebBrowser1.Visible = False
End Sub





Private Sub Command2_Click()
Command1_Click
Image1.Picture = Picture1.Image
End Sub

Private Sub Form_Load()
    Command1.Enabled = False
    VScroll1.Enabled = False
    HScroll1.Enabled = False
    WebBrowser1.Width = Screen.Width
    mlngLeft = WebBrowser1.Left
    mlngTop = WebBrowser1.Top
    mlngWidth = WebBrowser1.Width
    mlngHeight = WebBrowser1.Height
    'WebBrowser1.Navigate "http://www.codefans.net" '"http://www.baidu.com";
        WebBrowser1.Navigate "http://www.1616.net"
        
End Sub

Private Sub Form_Resize()
    Picture2.Move 0, 0, Me.ScaleWidth - VScroll1.Width, Me.ScaleHeight - HScroll1.Height
    VScroll1.Move Me.ScaleWidth - VScroll1.Width, 0, VScroll1.Width, Me.ScaleHeight - Command1.Height
    HScroll1.Move 0, Me.ScaleHeight - HScroll1.Height, Me.ScaleWidth - Command1.Width
    Command1.Move Me.ScaleWidth - Command1.Width, Me.ScaleHeight - Command1.Height
    VScroll1.Max = Picture1.ScaleHeight - Picture2.ScaleHeight
    If VScroll1.Max > 0 Then VScroll1.Enabled = True
    HScroll1.Max = Picture1.ScaleWidth - Picture2.ScaleWidth
    If HScroll1.Max > 0 Then HScroll1.Enabled = True
End Sub

Private Sub BrowserFullSize()
    Dim WidthInPixel As Double
    Dim HeightInPixel As Double
    Dim htmDoc As Object

    Set htmDoc = WebBrowser1.Document

    WidthInPixel = htmDoc.body.scrollWidth + (htmDoc.body.offsetWidth - htmDoc.body.ClientWidth)
    HeightInPixel = htmDoc.body.scrollHeight + (htmDoc.body.offsetHeight - htmDoc.body.ClientHeight)

    WebBrowser1.Left = 0
    WebBrowser1.Top = 0
    WebBrowser1.Width = WidthInPixel * Screen.TwipsPerPixelX
    WebBrowser1.Height = HeightInPixel * Screen.TwipsPerPixelY
    Picture1.Width = WebBrowser1.Width / 15
    Picture1.Height = WebBrowser1.Height / 15
    Set Picture1.Picture = LoadPicture() '清除旧的图像
    Picture1.AutoRedraw = True
    VScroll1.Max = Picture1.ScaleHeight - Picture2.ScaleHeight
    If VScroll1.Max > 0 Then VScroll1.Enabled = True
    HScroll1.Max = Picture1.ScaleWidth - Picture2.ScaleWidth
    If HScroll1.Max > 0 Then HScroll1.Enabled = True
    Set htmDoc = Nothing
End Sub

Private Sub BrowserNormalSize()
    WebBrowser1.Left = mlngLeft
    WebBrowser1.Top = mlngTop
    WebBrowser1.Width = mlngWidth
    WebBrowser1.Height = mlngHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowCaret (HScroll1.hwnd)
    ShowCaret (VScroll1.hwnd)
End Sub

Private Sub mnuClipboard_Click()
    Clipboard.SetData Picture1.Image
End Sub












Private Sub vScroll1_GotFocus()
    HideCaret (VScroll1.hwnd)
End Sub

Private Sub vScroll1_LostFocus()
    HideCaret (VScroll1.hwnd)
End Sub

Private Sub HScroll1_GotFocus()
    HideCaret (HScroll1.hwnd)
End Sub

Private Sub HScroll1_LostFocus()
    HideCaret (HScroll1.hwnd)
End Sub

Private Sub HScroll1_Scroll()
Picture1.Left = -HScroll1.Value
 HideCaret (HScroll1.hwnd)
End Sub

Private Sub VScroll1_Scroll()
Picture1.Top = -VScroll1.Value
 HideCaret (VScroll1.hwnd)
End Sub
Private Sub VScroll1_Change()
Picture1.Top = -VScroll1.Value
 HideCaret (VScroll1.hwnd)
End Sub
Private Sub HScroll1_Change()
Picture1.Left = -HScroll1.Value
 HideCaret (HScroll1.hwnd)
End Sub

Private Sub WebBrowser1_DownloadComplete()
DoEvents
Command1.Enabled = True
'mnuPicture.Enabled = True
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
 '可以屏蔽部分弹出窗口
    Cancel = True
End Sub

