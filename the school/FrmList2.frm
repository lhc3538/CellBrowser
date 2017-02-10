VERSION 5.00
Begin VB.Form FrmList 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer TimerCorrect 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   3240
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   5
      Left            =   120
      Picture         =   "FrmList.frx":0000
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   15
      Top             =   4560
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   5
         Left            =   2280
         MouseIcon       =   "FrmList.frx":02E9
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":043B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   60
         TabIndex        =   17
         Top             =   0
         Width           =   1065
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   5
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   4
      Left            =   120
      Picture         =   "FrmList.frx":0D83
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   12
      Top             =   4560
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   0
         Width           =   1065
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   4
         Left            =   2280
         MouseIcon       =   "FrmList.frx":106C
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":11BE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   4
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   3
      Left            =   120
      Picture         =   "FrmList.frx":1B06
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   9
      Top             =   4560
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   60
         TabIndex        =   11
         Top             =   0
         Width           =   1065
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   3
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1DEF
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1F41
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   3
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   2
      Left            =   120
      Picture         =   "FrmList.frx":2889
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   6
      Top             =   4560
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   0
         Width           =   1065
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   2
         Left            =   2280
         MouseIcon       =   "FrmList.frx":2B72
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":2CC4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   2
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   1
      Left            =   120
      Picture         =   "FrmList.frx":360C
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   3
      Top             =   4560
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   0
         Width           =   1065
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   1
         Left            =   2280
         MouseIcon       =   "FrmList.frx":38F5
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":3A47
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
   End
   Begin VB.Timer TimerGoRight 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   2040
      Top             =   5520
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   0
      Left            =   120
      Picture         =   "FrmList.frx":438F
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   0
         Left            =   2280
         MouseIcon       =   "FrmList.frx":4678
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":47CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'鼠标位置
Dim MX As Single
Dim MY As Single
Dim DyX, DyY As Single
Dim Tvalue As Boolean
Dim FormMove As Boolean
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
X As Long
Y As Long
End Type
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'通用变量定义区:
Dim Xleft As Integer '最左边pageback的位置
Dim MouseStep As Boolean '鼠标是否按下
Dim Xdistance As Integer '鼠标按下位置到xleft的距离
Dim FirstPageBack As Integer '纠正pageback位置（第一个）
Dim NearestPageback As Integer '纠正pggeback位置（最接近0）
Public PublicIndex As Integer '通用序号

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = Fcbox.Height
Me.Width = 1
End Sub

Private Sub TimerCorrect_Timer()
PageBack(FirstPageBack).Top = PageBack(FirstPageBack).Top - 560
CorrectPageBack


 If PageBack(NearestPageback).Top < 0 Then
  Dim DistanceNtoF As Integer
  DistanceNtoF = PageBack(NearestPageback).Top - PageBack(FirstPageBack).Top
  PageBack(NearestPageback).Top = 0
  PageBack(FirstPageBack).Top = 0 - DistanceNtoF
  CorrectPageBack
  Xleft = PageBack(FirstPageBack).Top
  TimerCorrect.Enabled = False
 End If
 

End Sub
Private Sub CorrectPageBack() ' 纠正网页列表
Dim i As Integer
Dim n As Integer
 For i = 0 To 5
  If PageBack(i).Visible = True Then
   PageBack(i).Top = PageBack(FirstPageBack).Top + PageBack(0).Width * n
   n = n + 1
  End If
 Next i
 
End Sub
Private Sub TimerGoRight_Timer()
If Fcbox.ComFrmList.Caption = "Right" Then
 If Me.Width = 1 Then Me.Width = 60
  Me.Width = Me.Width * 2
  Fcbox.Left = Me.Width
  FrmWeb(ActivePage).Left = Me.Width + Fcbox.Width - 60
  If Me.Width >= 1024 Then
   Me.Width = 6045
   FrmList.TimerGoRight.Enabled = False
   Fcbox.ComFrmList.Caption = "Left"
  End If
End If

If Fcbox.ComFrmList.Caption = "Left" Then
  Me.Width = Me.Width / 2
  If Me.Width < 400 Then
   Me.Width = 1
   FrmList.TimerGoRight.Enabled = False
   Fcbox.ComFrmList.Caption = "Right"
  End If
  Fcbox.Left = Me.Width
  FrmWeb(ActivePage).Left = Me.Width + Fcbox.Width - 60
End If
End Sub
Private Sub Page_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
MouseStep = True
Dim Point As POINTAPI
GetCursorPos Point
Dim i As Integer
 For i = 0 To 5
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
Xdistance = Point.X * 15 - PageBack(FirstPageBack).Top
End Sub

Private Sub Page_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseStep = True Then
Dim Point As POINTAPI
GetCursorPos Point
Xleft = Point.X * 15 - Xdistance
ListPageBack
End If
End Sub

Private Sub Page_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = False

Dim i As Integer
Dim n As Integer
 For i = 0 To 5
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
 
 For n = 0 To 5
  If PageBack(n).Visible = True And PageBack(n).Top < PageBack(0).Height And PageBack(n).Top >= 0 Then
   NearestPageback = n
   Exit For
  Else
   NearestPageback = FirstPageBack
  End If
 Next n

TimerCorrect.Enabled = True
End Sub






Private Sub PageBack_Click(Index As Integer)
PageBack(PublicIndex).Picture = LoadPicture(App.Path & "\skin\pageback0.gif")
PublicIndex = Index
PageBackClick
End Sub

Private Sub PageBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = True
Dim Point As POINTAPI
GetCursorPos Point
Dim i As Integer
 For i = 0 To 5
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
Xdistance = Point.X * 15 - PageBack(FirstPageBack).Top
End Sub

Private Sub PageBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseStep = True Then
Dim Point As POINTAPI
GetCursorPos Point
Xleft = Point.X * 15 - Xdistance
ListPageBack
End If
End Sub
Private Sub ListPageBack() '排列网页列表
On Error Resume Next
Dim Xnum As Integer
Dim i As Integer
Xnum = Xleft

For i = 0 To 5
 If PageBack(i).Visible = True Then
  PageBack(i).Top = Xnum
  Xnum = Xnum + PageBack(i).Height '会溢出
 End If
Next i
End Sub
Private Sub PageBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = False

Dim i As Integer
Dim n As Integer
 For i = 0 To 5
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
 
 For n = 0 To 5
  If PageBack(n).Visible = True And PageBack(n).Top < PageBack(0).Height And PageBack(n).Top >= 0 Then
   NearestPageback = n
   Exit For
  Else
   NearestPageback = FirstPageBack
  End If
 Next n

TimerCorrect.Enabled = True
End Sub
