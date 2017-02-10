VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FormAddLink 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "请给我地址"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   Icon            =   "FormAddLink.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4410
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   3000
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   8
      Top             =   840
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
         TabIndex        =   9
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
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   3000
      TabIndex        =   6
      Top             =   3795
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2818.701
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   2250
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3225
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   5
         Top             =   30
         Width           =   90
      End
   End
   Begin VB.CommandButton ComOK 
      Caption         =   "确定"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "预览"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox TextUrl 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin SHDocVwCtl.WebBrowser WebPage 
      Height          =   2655
      Left            =   3240
      TabIndex        =   10
      Top             =   5000
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
      Location        =   "http://www.soso.com/?unc=y400372_2"
   End
   Begin VB.CommandButton WebImage 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "地址："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "FormAddLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngLeft As Long
Private mlngTop As Long
Private mlngWidth As Long
Private mlngHeight As Long
Dim OpenID As Integer
'保存webbrowser的大小位置
Private Sub Command1_Click()
If TextUrl.Text <> "" Then
Page.Picture = LoadPicture("")
WebPage.Navigate TextUrl.Text
OpenID = 0
End If

End Sub


Private Sub BrowserFullSize()
    Dim WidthInPixel As Double
    Dim HeightInPixel As Double
    Dim htmDoc As Object

    Set htmDoc = WebPage.Document

    WidthInPixel = htmDoc.body.scrollWidth + (htmDoc.body.offsetWidth - htmDoc.body.ClientWidth)
    HeightInPixel = htmDoc.body.scrollHeight + (htmDoc.body.offsetHeight - htmDoc.body.ClientHeight)

    WebPage.Left = 0
    WebPage.Top = 0
    WebPage.Width = WidthInPixel * Screen.TwipsPerPixelX
    WebPage.Height = HeightInPixel * Screen.TwipsPerPixelY
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

Private Sub ComOK_Click()
Dim n As Integer
Dim neirong As String
Dim s(0 To 8) As String
Dim i As Integer
n = 0
Open App.path & "\collect.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, neirong$
s(n) = neirong
n = n + 1
Loop
Close #1

s(LinkPageNum) = WebName.Caption & "," & TextUrl.Text
SavePicture Page.Picture, App.path & "\webimage\" & WebName.Caption & ".gif"
On Error GoTo x1
Kill (App.path & "\collect.dat")
x1:

Open App.path & "\collect.dat" For Append As #1
For i = 0 To 8
Print #1, s(i)
Next i
Close #1

 MsgBox "保存成功", vbOKOnly, "成功"
Unload Me


End Sub

Private Sub Form_Load()
'webimage
    VScroll1.Enabled = False
    HScroll1.Enabled = False
    WebPage.Width = Screen.Width
    mlngLeft = WebPage.Left
    mlngTop = WebPage.Top
    mlngWidth = WebPage.Width
    mlngHeight = WebPage.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowCaret (HScroll1.hWnd)
    ShowCaret (VScroll1.hWnd)
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
   Page.Picture = Picture1.Image
   ComOK.Enabled = True
End Sub

Private Sub WebPage_DownloadComplete()
On Error Resume Next
 OpenID = OpenID + 1

If OpenID = 1 Then
DoEvents
WebImage.SetFocus
WebImage_Click '控制事件只截取一次网页
End If
WebName.Caption = WebPage.Document.Title
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

