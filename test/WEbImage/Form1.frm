VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   11730
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   720
      ScaleHeight     =   2115
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   2880
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   720
      ScaleHeight     =   2235
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   360
      Width           =   4815
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2175
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      ExtentX         =   8493
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlngLeft As Long
Private mlngTop As Long
Private mlngWidth As Long
Private mlngHeight As Long
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
    mnuSave.Enabled = True
    mnuClipboard.Enabled = True
    '......
    
    '恢复原来的大小
    BrowserNormalSize
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "http://www.baidu.com"
End Sub
