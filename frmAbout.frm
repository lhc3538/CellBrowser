VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于我的应用程序"
   ClientHeight    =   3375
   ClientLeft      =   8550
   ClientTop       =   5865
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2329.485
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2385
      ScaleWidth      =   5745
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      Begin VB.Label SayThanks 
         BackStyle       =   0  'Transparent
         Caption         =   "鸣谢：  ?恒达 提供测试及打包             零时迷子 、?骏森 提供测试        以及各界网友提供的技术支持"
         Height          =   615
         Left            =   2520
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "版本"
         Height          =   225
         Left            =   2520
         TabIndex        =   4
         Top             =   660
         Width           =   2205
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "应用程序标题"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3000
         TabIndex        =   3
         Top             =   120
         Width           =   1605
      End
      Begin VB.Image Image1 
         Height          =   2970
         Left            =   -360
         MouseIcon       =   "frmAbout.frx":000C
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":015E
         Stretch         =   -1  'True
         Top             =   -240
         Width           =   2970
      End
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   975
      Left            =   4200
      TabIndex        =   1
      Top             =   1200
      Width           =   30
      ExtentX         =   53
      ExtentY         =   1720
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
   Begin VB.Image ComAddQQ 
      Height          =   555
      Left            =   4899
      MouseIcon       =   "frmAbout.frx":11807
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":11959
      Top             =   2640
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   3493.272
      X2              =   3493.272
      Y1              =   1656.523
      Y2              =   2319.132
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5296.251
      Y1              =   1656.523
      Y2              =   1656.523
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "问题咨询："
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "  本软件的下载和使用是完全免费的，用户拥有使用权。  感谢您对本软件的支持，希望能给您带来便利。"
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   3525
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub ComAddQQ_Click()
Web1.Navigate ("http://wpa.qq.com/msgrd?V=1&Uin=969461192")
ComAddQQ.Enabled = False

End Sub



Private Sub Form_Load()
    Me.Caption = "关于 " & App.Title
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub


Private Sub Image1_Click()
Fa.OpenNewPage.Text = "http://120343.24la.com.cn/"
Unload Me
End Sub
