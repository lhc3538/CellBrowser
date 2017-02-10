VERSION 5.00
Begin VB.Form LinkPage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "LinkPage"
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   Icon            =   "LinkPage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   4560
      Top             =   3240
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   8
      Left            =   0
      ScaleHeight     =   2818.701
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   24
      Top             =   0
      Width           =   3495
      Begin VB.TextBox WebURL 
         Height          =   270
         Index           =   8
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   8
         Left            =   120
         MouseIcon       =   "LinkPage.frx":000C
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3225
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   60
         TabIndex        =   26
         Top             =   63
         Width           =   630
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   8
         Left            =   3000
         Picture         =   "LinkPage.frx":015E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   7
      Left            =   0
      ScaleHeight     =   2818.701
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   21
      Top             =   0
      Width           =   3495
      Begin VB.TextBox WebURL 
         Height          =   270
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   7
         Left            =   120
         MouseIcon       =   "LinkPage.frx":0AA6
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3225
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   7
         Left            =   60
         TabIndex        =   23
         Top             =   63
         Width           =   630
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   7
         Left            =   3000
         Picture         =   "LinkPage.frx":0BF8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   6
      Left            =   0
      ScaleHeight     =   2818.701
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   18
      Top             =   0
      Width           =   3495
      Begin VB.TextBox WebURL 
         Height          =   270
         Index           =   6
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   6
         Left            =   120
         MouseIcon       =   "LinkPage.frx":1540
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3225
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   6
         Left            =   60
         TabIndex        =   20
         Top             =   63
         Width           =   630
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   6
         Left            =   3000
         Picture         =   "LinkPage.frx":1692
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   5
      Left            =   0
      ScaleHeight     =   2818.701
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   15
      Top             =   0
      Width           =   3495
      Begin VB.TextBox WebURL 
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   5
         Left            =   120
         MouseIcon       =   "LinkPage.frx":1FDA
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3225
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   17
         Top             =   63
         Width           =   630
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   5
         Left            =   3000
         Picture         =   "LinkPage.frx":212C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   4
      Left            =   0
      ScaleHeight     =   2818.701
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   12
      Top             =   0
      Width           =   3495
      Begin VB.TextBox WebURL 
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   4
         Left            =   120
         MouseIcon       =   "LinkPage.frx":2A74
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3225
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   63
         Width           =   630
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   4
         Left            =   3000
         Picture         =   "LinkPage.frx":2BC6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   3
      Left            =   0
      ScaleHeight     =   2818.701
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   9
      Top             =   0
      Width           =   3495
      Begin VB.TextBox WebURL 
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   3
         Left            =   120
         MouseIcon       =   "LinkPage.frx":350E
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3225
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   11
         Top             =   63
         Width           =   630
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   3
         Left            =   3000
         Picture         =   "LinkPage.frx":3660
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   2
      Left            =   0
      ScaleHeight     =   2818.701
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   6
      Top             =   0
      Width           =   3495
      Begin VB.TextBox WebURL 
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   2
         Left            =   120
         MouseIcon       =   "LinkPage.frx":3FA8
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3225
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   63
         Width           =   630
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   2
         Left            =   3000
         Picture         =   "LinkPage.frx":40FA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   1
      Left            =   0
      ScaleHeight     =   2818.701
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   3
      Top             =   0
      Width           =   3495
      Begin VB.TextBox WebURL 
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   1
         Left            =   120
         MouseIcon       =   "LinkPage.frx":4A42
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3225
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   63
         Width           =   630
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   1
         Left            =   3000
         Picture         =   "LinkPage.frx":4B94
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   0
      Left            =   4080
      ScaleHeight     =   2818.701
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.TextBox WebURL 
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   1
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
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   63
         Width           =   630
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   0
         Left            =   120
         MouseIcon       =   "LinkPage.frx":54DC
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3225
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   0
         Left            =   3000
         Picture         =   "LinkPage.frx":562E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
End
Attribute VB_Name = "LinkPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128 ' Maintenance string for PSS usage
osName As String ' 操作系统的名称
End Type

' 获得 Windows 操作系统的版本
' OSVERSIONINFO 结构中的 osName 返回操作系统的名称
Private Function GetWindowsVersion() As OSVERSIONINFO
Dim ver As OSVERSIONINFO
ver.dwOSVersionInfoSize = 148
GetVersionEx ver
With ver
Select Case .dwPlatformId
Case 1
Select Case .dwMinorVersion
Case 0
.osName = "Windows 95"
Case 10
.osName = "Windows 98"
Case 90
.osName = "Windows Mellinnium"
End Select
Case 2
Select Case .dwMajorVersion
Case 3
.osName = "Windows NT 3.51"
Case 4
.osName = "Windows NT 4.0"
Case 5
Select Case .dwMinorVersion
Case 0
.osName = "Windows 2000"
Case 1
.osName = "Windows XP"
Case 2
.osName = "Windows Server 2003"
End Select
Case 6
.osName = "Windows7" '新增加部分
End Select
Case Else
.osName = "Failed"
End Select
End With
GetWindowsVersion = ver
End Function

Private Sub ComEndWeb_Click(Index As Integer)
WebName(Index).Caption = "WebName"
Page(Index).Picture = LoadPicture("")
WebURL(Index).Text = ""

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

s(Index) = ""

i = 0
On Error GoTo x1
Kill (App.path & "\collect.dat")
x1:

Open App.path & "\collect.dat" For Append As #1
For i = 0 To 8
Print #1, s(i)
Next i
Close #1
End Sub

Private Sub Form_Activate()
Fa.ActiveUrl.Text = "Blank Page"
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = Fa.Height - Fa.P1.Height - Fa.P2.Height
Me.Width = Fa.Width


Dim n As Integer
Dim neirong As String
Dim s(0 To 8) As String
n = 0
Open App.path & "\collect.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, neirong$
s(n) = neirong
l = Split(s(n), ",")
If s(n) <> "" Then
WebName(n).Caption = l(0)
WebURL(n).Text = l(1)
End If
n = n + 1
Loop
Close #1

LoadBrowserSet

End Sub
Private Sub MorenBrowser()
If SetStr(0) = "true" Then
Dim w As Object
On Error Resume Next
'检测系统
Dim osnameP
Dim ver As OSVERSIONINFO
ver = GetWindowsVersion()
With ver
'在label1上显示
osnameP = .osName
End With
'win7提示
If osnameP = "Windows7" Then
Dim WshShell As Object
Dim strIP As String
    Set WshShell = CreateObject("WScript.Shell")
    strIP = WshShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice\progid")
 If strIP = "Cell.Http" Then
 Exit Sub
 Else
 DefaultSet.Show
 End If


 'XP提示
Else
Set w = CreateObject("wscript.shell")
a = w.RegRead("HKEY_CLASSES_ROOT\https\shell\")
 If a = "Cell.Http" Then
  Set w = CreateObject("wscript.shell")
  a = w.RegRead("HKEY_CLASSES_ROOT\https\shell\Cell.Http\command\")
   If Left(Right(a, 12), 8) = "cell.exe" Then
   Timer2.Enabled = False
   Exit Sub
   Else
      DefaultSet.Show
   End If
 Else
 DefaultSet.Show
End If
End If
End If

End Sub
Private Sub Form_Resize()
Dim Xdistance As Integer
Dim Ydistance As Integer
If Me.Width > PageBack(0) * 3 And Me.Height > PageBack(0) * 3 Then
Xdistance = (Me.Width - PageBack(0).Width * 3) / 4
Ydistance = (Me.Height - PageBack(0).Height * 3) / 4
Else
Xdistance = 0
Ydistance = 0
End If

Dim i As Integer
For i = 0 To 8
If i <= 2 Then
PageBack(i).Left = Xdistance * (i + 1) + PageBack(0).Width * i
PageBack(i).Top = Ydistance
End If
If i <= 5 And i > 2 Then
PageBack(i).Left = Xdistance * (i - 3 + 1) + PageBack(0).Width * (i - 3)
PageBack(i).Top = Ydistance * 2 + PageBack(0).Height
End If
If i <= 8 And i > 5 Then
PageBack(i).Left = Xdistance * (i - 6 + 1) + PageBack(0).Width * (i - 6)
PageBack(i).Top = Ydistance * 3 + PageBack(0).Height * 2
End If
Next i
End Sub

Private Sub Page_Click(Index As Integer)
If WebURL(Index).Text = "" Then
FormAddLink.Show
LinkPageNum = Index
Else
Fa.OpenNewPage.Text = WebURL(Index).Text
End If


End Sub

Private Sub Timer1_Timer() '启动那个默认浏览器的检测功能
MorenBrowser
Timer1.Enabled = False
End Sub

Private Sub WebURL_Change(Index As Integer)
On Error Resume Next
Page(Index).Picture = LoadPicture(App.path & "\webimage\" & WebName(Index).Caption & ".gif")

End Sub
