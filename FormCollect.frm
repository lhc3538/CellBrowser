VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormFavorite 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "收藏夹"
   ClientHeight    =   5715
   ClientLeft      =   8445
   ClientTop       =   3105
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   7435
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FormFavorite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'枚举IE收藏夹里所有网站快捷方式及其网址
'By:NcL_K2
'E-mail:stluo@foxmail.com
'使用及转载请保留上信息,谢谢
'----------------------------------------------------------
'菜单栏>工程>部件 添加microsoft windows common controls 控件
'在窗口上添加一个TreeView控件,名称默认为TreeView1
'菜单栏>工程>引用 添加 Windows Script Host Object Model
'菜单栏>工程>引用 添加 Microsoft Scripting Runtime

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Dim Fso As New FileSystemObject
Dim OnlyKey As Long '节点的Key
Dim Url() As String '保存网址的数组
Dim UrlIndex As Long '节点对应该的网址索引值

Private Sub Form_Load()
Me.Left = 2
Me.Top = Fa.P1.Height + Fa.P2.Height
Me.Height = Fa.Height - Me.Top
TreeView1.LineStyle = 1
Dim FavPath As String '收藏夹路径
Dim GetFavPath As New IWshRuntimeLibrary.WshShell
FavPath = GetFavPath.SpecialFolders("FAVORITES") & "\" '获取收藏夹路径
TreeView1.Nodes.Add , tvwChild, "Root", "IE收藏夹" '添加根节点,Key为"FAVORITES"
GetFiles FavPath, "Root" '枚举根目录下所有Url文件
End Sub

Private Sub GetFolder(ByVal path As String, ByVal key As String) '枚举文件夹
Dim m_Folder As Folder
Dim m_SubFolder As Folder
Set m_Folder = Fso.GetFolder(path)

For Each m_SubFolder In m_Folder.SubFolders
OnlyKey = OnlyKey + 1 '确保当前节点的Key是唯一的
TreeView1.Nodes.Add key, tvwChild, "Key" & OnlyKey, m_SubFolder.Name
GetFiles path & m_SubFolder.Name & "\", "Key" & OnlyKey '枚举当前目录下所有Url文件
Next
End Sub

Private Sub GetFiles(ByVal path As String, ByVal key As String)
Dim m_Folder As Folder
Dim m_File As File
Set m_Folder = Fso.GetFolder(path)
For Each m_File In m_Folder.Files
If Fso.GetExtensionName(path & m_File.Name) = "url" Then '如果是Url文件
TreeView1.Nodes.Add key, tvwChild, "Url" & UrlIndex, m_File.Name '添加到节点
ReDim Preserve Url(UrlIndex) '重新定义数组长度
'获取网址并保存到数组中
Dim Retval As String
Retval = Space(256)
GetPrivateProfileString "InternetShortcut", "URL", "", Retval, 256, path & m_File.Name
Url(UrlIndex) = Retval
UrlIndex = UrlIndex + 1
End If
Next
'当前目录下的Url文件枚举完毕后,还要枚举当然目录下的子文件夹
GetFolder path, key '枚举当前目录下的子文件夹
End Sub


Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Form_Resize()
TreeView1.Height = Me.Height
TreeView1.Width = Me.Width

End Sub



Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Fa.OpenNewPage.Text = Url(CLng(Replace(Node.key, "Url", "")))
End Sub


