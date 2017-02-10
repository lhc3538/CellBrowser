VERSION 5.00
Begin VB.Form FormSet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "细胞浏览器设置"
   ClientHeight    =   4305
   ClientLeft      =   6000
   ClientTop       =   4065
   ClientWidth     =   8640
   ClipControls    =   0   'False
   Icon            =   "FormSet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2971.387
   ScaleMode       =   0  'User
   ScaleWidth      =   8113.409
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox PBasic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   1560
      ScaleHeight     =   4335
      ScaleWidth      =   7095
      TabIndex        =   3
      Top             =   0
      Width           =   7095
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3495
         ScaleWidth      =   7095
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   7095
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   19
            Left            =   5280
            TabIndex        =   36
            Top             =   2160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   18
            Left            =   4800
            TabIndex        =   35
            Top             =   2160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   17
            Left            =   4320
            TabIndex        =   34
            Top             =   2160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   16
            Left            =   3840
            TabIndex        =   33
            Top             =   2160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   15
            Left            =   3240
            TabIndex        =   32
            Top             =   2160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   14
            Left            =   1440
            TabIndex        =   31
            Top             =   2520
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   13
            Left            =   5160
            TabIndex        =   30
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   12
            Left            =   4440
            TabIndex        =   29
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   11
            Left            =   3720
            TabIndex        =   28
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   10
            Left            =   1080
            TabIndex        =   27
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   9
            Left            =   480
            TabIndex        =   26
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   8
            Left            =   5400
            TabIndex        =   25
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   7
            Left            =   4680
            TabIndex        =   24
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   6
            Left            =   3960
            TabIndex        =   23
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   5
            Left            =   3360
            TabIndex        =   22
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   4
            Left            =   2760
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   3
            Left            =   2160
            TabIndex        =   20
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   2
            Left            =   1560
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   1
            Left            =   1080
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TextCell 
            Height          =   270
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Labelhuanyuan 
            BackStyle       =   0  'Transparent
            Caption         =   "还原默认图标"
            Height          =   255
            Left            =   5880
            MouseIcon       =   "FormSet.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   37
            Top             =   3240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Image ComDown 
            Height          =   495
            Left            =   4680
            MouseIcon       =   "FormSet.frx":015E
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image ComUP 
            Height          =   495
            Left            =   4200
            MouseIcon       =   "FormSet.frx":02B0
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image ComExitWeb 
            Height          =   495
            Left            =   5160
            MouseIcon       =   "FormSet.frx":0402
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image ComWebList 
            Height          =   495
            Left            =   3720
            MouseIcon       =   "FormSet.frx":0554
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image ComGO 
            Height          =   615
            Left            =   3120
            MouseIcon       =   "FormSet.frx":06A6
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   615
         End
         Begin VB.Image P1 
            Height          =   1695
            Left            =   120
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Image BackDown 
            Height          =   765
            Left            =   5280
            Stretch         =   -1  'True
            Top             =   0
            Width           =   765
         End
         Begin VB.Image BackUP 
            Height          =   765
            Left            =   4440
            Stretch         =   -1  'True
            Top             =   0
            Width           =   765
         End
         Begin VB.Image ComMin 
            Height          =   735
            Left            =   3600
            Stretch         =   -1  'True
            Top             =   840
            Width           =   735
         End
         Begin VB.Image ComMax 
            Height          =   735
            Left            =   4320
            Stretch         =   -1  'True
            Top             =   840
            Width           =   735
         End
         Begin VB.Image ComExit 
            Height          =   735
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   840
            Width           =   735
         End
         Begin VB.Image ComRestore 
            Height          =   615
            Left            =   3840
            Stretch         =   -1  'True
            Top             =   120
            Width           =   615
         End
         Begin VB.Image ComFavorite 
            Height          =   615
            Left            =   3240
            Stretch         =   -1  'True
            Top             =   120
            Width           =   615
         End
         Begin VB.Image ComHome 
            Height          =   615
            Left            =   2640
            Stretch         =   -1  'True
            Top             =   120
            Width           =   615
         End
         Begin VB.Image ComRefresh 
            Height          =   615
            Left            =   2040
            Stretch         =   -1  'True
            Top             =   120
            Width           =   615
         End
         Begin VB.Image ComAdvance 
            Height          =   615
            Left            =   1440
            Stretch         =   -1  'True
            Top             =   120
            Width           =   615
         End
         Begin VB.Image ComBack 
            Height          =   615
            Left            =   840
            Stretch         =   -1  'True
            Top             =   120
            Width           =   615
         End
         Begin VB.Image ComCell 
            Height          =   765
            Left            =   0
            MouseIcon       =   "FormSet.frx":07F8
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   0
            Width           =   765
         End
         Begin VB.Image ComSearch 
            Height          =   615
            Left            =   360
            Stretch         =   -1  'True
            Top             =   960
            Width           =   615
         End
         Begin VB.Image ComTranslate 
            Height          =   615
            Left            =   960
            Stretch         =   -1  'True
            Top             =   960
            Width           =   615
         End
         Begin VB.Image Image4 
            Height          =   4785
            Left            =   -120
            Picture         =   "FormSet.frx":094A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   7335
         End
      End
      Begin VB.CheckBox OpOpenT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   4320
         TabIndex        =   13
         Top             =   2040
         Width           =   200
      End
      Begin VB.CheckBox OpExitT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   4320
         TabIndex        =   12
         Top             =   1560
         Width           =   200
      End
      Begin VB.TextBox TxtHomePage 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "本浏览器主页"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "通用主页"
         Height          =   255
         Left            =   5760
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "取消"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5670
         MouseIcon       =   "FormSet.frx":23F3
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   3870
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "保存设置"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4170
         MouseIcon       =   "FormSet.frx":2545
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   3870
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "打开浏览器时是否提醒设为默认浏览器？"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "是"
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "是"
         Height          =   255
         Left            =   4560
         TabIndex        =   9
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "关闭浏览器时是否进行提醒？"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6840
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "主页地址："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.Image ComCancel 
         Height          =   495
         Left            =   5640
         MouseIcon       =   "FormSet.frx":2697
         MousePointer    =   99  'Custom
         Picture         =   "FormSet.frx":27E9
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   1275
      End
      Begin VB.Image ComOK 
         Height          =   495
         Left            =   4080
         MouseIcon       =   "FormSet.frx":6B85
         MousePointer    =   99  'Custom
         Picture         =   "FormSet.frx":6CD7
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   1275
      End
      Begin VB.Image Image3 
         Height          =   4455
         Left            =   -120
         Picture         =   "FormSet.frx":B073
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   7335
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      Picture         =   "FormSet.frx":CB1C
      ScaleHeight     =   4575
      ScaleWidth      =   1530
      TabIndex        =   0
      Top             =   0
      Width           =   1529
      Begin VB.Label ComSkin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "界面修改"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label ComBasic 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "基本设置"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   570
         Left            =   -120
         Picture         =   "FormSet.frx":D1DE
         Stretch         =   -1  'True
         Top             =   75
         Width           =   2010
      End
      Begin VB.Image Image2 
         Height          =   465
         Left            =   0
         Picture         =   "FormSet.frx":D945
         Top             =   720
         Width           =   1995
      End
   End
End
Attribute VB_Name = "FormSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'修改主页配置
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const FO_MOVE = &H1
  Private Const FO_COPY = &H2
  Private Const FO_DELETE = &H3
  Private Const FO_RENAME = &H4
  Private Const FOF_NOCONFIRMATION = &H10
  Private Const FOF_SILENT = &H4
  Private Const FOF_NOERRORUI = &H400
  Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
  Private Type SHFILEOPSTRUCT
                hWnd     As Long
                wFunc     As Long
                pFrom     As String
                pTo     As String
                fFlags     As Integer
                fAnyOperationsAborted     As Long
                hNameMappings     As Long
                lpszProgressTitle     As String           'only     used     if     FOF_SIMPLEPROGRESS
  End Type
    
Private Sub BackDown_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(8).Text = sFile
    End If
If sFile <> "" Then BackDown.Picture = LoadPicture(sFile)
End Sub

Private Sub BackUP_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(7).Text = sFile
    End If
If sFile <> "" Then BackUP.Picture = LoadPicture(sFile)
End Sub

Private Sub ComAdvance_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(2).Text = sFile
    End If
If sFile <> "" Then ComAdvance.Picture = LoadPicture(sFile)
End Sub

Private Sub ComBack_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(1).Text = sFile
    End If
If sFile <> "" Then ComBack.Picture = LoadPicture(sFile)
End Sub

Private Sub ComBasic_Click()
Picture2.Visible = False
Image1.Top = ComBasic.Top - 160
End Sub

Private Sub ComBasic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2.Top <> ComBasic.Top - 100 Then
Image2.Top = ComBasic.Top - 100
End If
End Sub

Private Sub ComCancel_Click()
Unload Me
End Sub

Private Sub ComCell_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(0).Text = sFile
    End If
If sFile <> "" Then ComCell.Picture = LoadPicture(sFile)

End Sub

Private Sub ComDown_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(18).Text = sFile
    End If
If sFile <> "" Then ComDown.Picture = LoadPicture(sFile)
End Sub

Private Sub ComExit_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(13).Text = sFile
    End If
If sFile <> "" Then ComExit.Picture = LoadPicture(sFile)
End Sub

Private Sub ComExitWeb_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(19).Text = sFile
    End If
If sFile <> "" Then ComExitWeb.Picture = LoadPicture(sFile)
End Sub

Private Sub ComFavorite_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(5).Text = sFile
    End If
If sFile <> "" Then ComFavorite.Picture = LoadPicture(sFile)
End Sub

Private Sub ComGO_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(15).Text = sFile
    End If
If sFile <> "" Then ComGO.Picture = LoadPicture(sFile)
End Sub

Private Sub ComHome_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(4).Text = sFile
    End If
If sFile <> "" Then ComHome.Picture = LoadPicture(sFile)
End Sub

Private Sub ComMax_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(12).Text = sFile
    End If
If sFile <> "" Then ComMax.Picture = LoadPicture(sFile)
End Sub

Private Sub ComMin_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(11).Text = sFile
    End If
If sFile <> "" Then ComMin.Picture = LoadPicture(sFile)
End Sub

Private Sub ComOK_Click()
'修改浏览器主页
If TxtHomePage.Text <> "" Then
 Dim b
 Dim i As Integer
 Dim a

  For i = 1 To Len(TxtHomePage.Text)
   a = Asc(Mid(TxtHomePage.Text, i, 1))
 b = b & "," & a * 3
  Next i
Open App.path & "\homepage.dat" For Append As #1
Print #1, b
Close #1
End If
'修改通用主页
If Check2.Value = 1 Then
Dim hKey As Long, s As String, m As Integer
s = TxtHomePage.Text     '默认主页
m = Len(s) + 1
RegCreateKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main\Start", hKey
RegSetValueEx hKey, "Start Page", 0, REG_SZ, ByVal s, m

End If


'修改提示
If OpOpenT.Value = 1 Then
 SetStr(0) = "true"
Else
 SetStr(0) = "false"
End If
If OpExitT.Value = 1 Then
 SetStr(1) = "true"
Else
 SetStr(1) = "false"
End If
Dim SaveLi
i = 0
For i = 0 To UBound(SetStr)
SaveLi = SaveLi & SetStr(i) & ","
Next i
Open App.path & "\set.dat" For Output As #1
Print #1, SaveLi
Close #1
'保存皮肤
On Error Resume Next
If Labelhuanyuan.Visible = True Then
 i = 0
 For i = 0 To 19
  If TextCell(i).Text <> "" Then
   If i = 0 Then
        Kill App.path & "\skin\comcell.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\comcell.gif"
   End If
   If i = 1 Then
        Kill App.path & "\skin\ComBack.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComBack.gif"
   End If
   If i = 2 Then
        Kill App.path & "\skin\ComAdvance.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComAdvance.gif"
   End If
   If i = 3 Then
        Kill App.path & "\skin\ComRefresh.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComRefresh.gif"
   End If
   If i = 4 Then
    Kill App.path & "\skin\ComHome.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComHome.gif"
   End If
   If i = 5 Then
    Kill App.path & "\skin\ComFavorite.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComFavorite.gif"
   End If
   If i = 6 Then
    Kill App.path & "\skin\ComRestore.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComRestore.gif"
   End If
   If i = 7 Then
    Kill App.path & "\skin\BackUP.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\BackUP.gif"
   End If
   If i = 8 Then
    Kill App.path & "\skin\BackDown.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\BackDown.gif"
   End If
   If i = 9 Then
    Kill App.path & "\skin\ComSearch.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComSearch.gif"
   End If
   If i = 10 Then
    Kill App.path & "\skin\ComTranslate.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComTranslate.gif"
   End If
   If i = 11 Then
    Kill App.path & "\skin\ComMin.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComMin.gif"
   End If
   If i = 12 Then
    Kill App.path & "\skin\ComMax.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComMax.gif"
   End If
   If i = 13 Then
    Kill App.path & "\skin\ComExit.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComExit.gif"
   End If
   If i = 14 Then
    Kill App.path & "\skin\P1.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\P1.gif"
   End If
   If i = 15 Then
    Kill App.path & "\skin\ComGO.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComGO.gif"
   End If
   If i = 16 Then
    Kill App.path & "\skin\ComWebList.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComWebList.gif"
   End If
   If i = 17 Then
    Kill App.path & "\skin\ComUP.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComUP.gif"
   End If
   If i = 18 Then
    Kill App.path & "\skin\ComDown.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComDown.gif"
   End If
   If i = 19 Then
    Kill App.path & "\skin\ComExitWeb.gif"
    FileCopy TextCell(i).Text, App.path & "\skin\ComExitWeb.gif"
   End If
   
   
  End If
 Next i
End If
MsgBox ("保存成功")
End Sub

Private Sub ComRefresh_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(3).Text = sFile
    End If
If sFile <> "" Then ComRefresh.Picture = LoadPicture(sFile)
End Sub

Private Sub ComRestore_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(6).Text = sFile
    End If
If sFile <> "" Then ComRestore.Picture = LoadPicture(sFile)
End Sub

Private Sub ComSearch_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(9).Text = sFile
    End If
If sFile <> "" Then ComSearch.Picture = LoadPicture(sFile)
End Sub

Private Sub ComSkin_Click()
Picture2.Visible = True
Image1.Top = ComSkin.Top - 160
End Sub

Private Sub ComSkin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2.Top <> ComSkin.Top - 100 Then
Image2.Top = ComSkin.Top - 100
End If
End Sub

Private Sub ComTranslate_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(10).Text = sFile
    End If
If sFile <> "" Then ComTranslate.Picture = LoadPicture(sFile)
End Sub

Private Sub ComUP_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(17).Text = sFile
    End If
If sFile <> "" Then ComUP.Picture = LoadPicture(sFile)
End Sub

Private Sub ComWebList_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(16).Text = sFile
    End If
If sFile <> "" Then ComWebList.Picture = LoadPicture(sFile)
End Sub

Private Sub Form_Load()
LoadBrowserSet
If SetStr(0) = "true" Then OpOpenT.Value = 1
If SetStr(1) = "true" Then OpExitT.Value = 1

'载入主页
Dim cw As String
Dim T3 As String
Open App.path & "\homepage.dat" For Input As #1
  Do While Not EOF(1)
    Line Input #1, cw$
    T3 = cw
    Loop
Close #1

Dim l
l = Split(T3, ",")
Dim i As Integer
Dim T4 As String
For i = 1 To UBound(l)
T4 = T4 & Chr(l(i) / 3)
Next i
TxtHomePage.Text = T4

LoadComPicture

End Sub



Private Sub Label4_Click()
ComOK_Click
End Sub

Private Sub Label7_Click()
ComCancel_Click
End Sub
Private Sub LoadComPicture()
On Error Resume Next

'Dim rtn As Long
  '  rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
  '  rtn = rtn Or WS_EX_LAYERED
  '  SetWindowLong hwnd, GWL_EXSTYLE, rtn
  '  SetLayeredWindowAttributes hwnd, 0, 245, LWA_ALPHA '窗体透明度
'p2特殊

ComCell.Picture = LoadPicture(App.path & "\skin\comcell.gif")
P1.Picture = LoadPicture(App.path & "\skin\P1.gif")
ComBack.Picture = LoadPicture(App.path & "\skin\comback.gif")
ComAdvance.Picture = LoadPicture(App.path & "\skin\comadvance.gif")
ComRefresh.Picture = LoadPicture(App.path & "\skin\comrefresh.gif")
ComHome.Picture = LoadPicture(App.path & "\skin\comhome.gif")
ComFavorite.Picture = LoadPicture(App.path & "\skin\comfavorite.gif")
ComSearch.Picture = LoadPicture(App.path & "\skin\comsearch.gif")
ComMin.Picture = LoadPicture(App.path & "\skin\commin.gif")
ComMax.Picture = LoadPicture(App.path & "\skin\commax.gif")
ComExit.Picture = LoadPicture(App.path & "\skin\comexit.gif")
ComDown.Picture = LoadPicture(App.path & "\skin\comdown.gif")
ComUP.Picture = LoadPicture(App.path & "\skin\comup.gif")
ComExitWeb.Picture = LoadPicture(App.path & "\skin\comexitweb.gif")
BackDown.Picture = LoadPicture(App.path & "\skin\backdown.gif")
BackUP.Picture = LoadPicture(App.path & "\skin\backup.gif")
ComGO.Picture = LoadPicture(App.path & "\skin\comgo.gif")
ComWebList.Picture = LoadPicture(App.path & "\skin\comweblist.gif")
ComTranslate.Picture = LoadPicture(App.path & "\skin\comtranslate.gif")
ComRestore.Picture = LoadPicture(App.path & "\skin\comrestore.gif")

End Sub

Private Sub Labelhuanyuan_Click()
On Error Resume Next

Kill "\skin\comcell.gif"
Kill "\skin\P1.gif"
Kill "\skin\comback.gif"
Kill "\skin\comadvance.gif"
Kill "\skin\comrefresh.gif"
Kill "\skin\comhome.gif"
Kill "\skin\comfavorite.gif"
Kill "\skin\comsearch.gif"
Kill "\skin\commin.gif"
Kill "\skin\commax.gif"
Kill "\skin\comexit.gif"
Kill "\skin\comdown.gif"
Kill "\skin\comup.gif"
Kill "\skin\comexitweb.gif"
Kill "\skin\backdown.gif"
Kill "\skin\backup.gif"
Kill "\skin\comgo.gif"
Kill "\skin\comweblist.gif"
Kill "\skin\comtranslate.gif"
Kill "\skin\comrestore.gif"


FileCopy App.path & "\skin\recycle\comcell.gif", "\skin\comcell.gif"
FileCopy App.path & "\skin\recycle\P1.gif", "\skin\P1.gif"
FileCopy App.path & "\skin\recycle\comback.gif", "\skin\comback.gif"
FileCopy App.path & "\skin\recycle\comadvance.gif", "\skin\comadvance.gif"
FileCopy App.path & "\skin\recycle\comrefresh.gif", "\skin\comrefresh.gif"
FileCopy App.path & "\skin\recycle\comhome.gif", "\skin\comhome.gif"
FileCopy App.path & "\skin\recycle\comfavorite.gif", "\skin\comfavorite.gif"
FileCopy App.path & "\skin\recycle\comsearch.gif", "\skin\comsearch.gif"
FileCopy App.path & "\skin\recycle\commin.gif", "\skin\commin.gif"
FileCopy App.path & "\skin\recycle\commax.gif", "\skin\commax.gif"
FileCopy App.path & "\skin\recycle\comexit.gif", "\skin\comexit.gif"
FileCopy App.path & "\skin\recycle\comdown.gif", "\skin\comdown.gif"
FileCopy App.path & "\skin\recycle\comup.gif", "\skin\comup.gif"
FileCopy App.path & "\skin\recycle\comexitweb.gif", "\skin\comexitweb.gif"
FileCopy App.path & "\skin\recycle\backdown.gif", "\skin\backdown.gif"
FileCopy App.path & "\skin\recycle\backup.gif", "\skin\backup.gif"
FileCopy App.path & "\skin\recycle\comgo.gif", "\skin\comgo.gif"
FileCopy App.path & "\skin\recycle\comweblist.gif", "\skin\comweblist.gif"
FileCopy App.path & "\skin\recycle\comtranslate.gif", "\skin\comtranslate.gif"
FileCopy App.path & "\skin\recycle\comrestore.gif", "\skin\comrestore.gif"

End Sub

Private Sub P1_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '设置默认要打开文件的扩展名
    file.nMaxFile = 255 '显示文件名的长度
    file.lpstrFileTitle = String$(255, 0) '打开对话框的标题
    file.nMaxFileTitle = 255 '打开对话框的标题的长度
    file.lpstrInitialDir = Environ$("WinDir") '设置盘符
    file.lpstrFilter = "图片文件" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '打开的文件类型
    file.nFilterIndex = 1
    file.lpstrTitle = "打开文件"
    lResult = GetOpenFileName(file) '取得文件名
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        
        TextCell(14).Text = sFile
    End If
If sFile <> "" Then P1.Picture = LoadPicture(sFile)
End Sub

Private Sub TextCell_Change(Index As Integer)
Labelhuanyuan.Visible = True
End Sub
