VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormFavorite 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "�ղؼ�"
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
         Name            =   "����"
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
'ö��IE�ղؼ���������վ��ݷ�ʽ������ַ
'By:NcL_K2
'E-mail:stluo@foxmail.com
'ʹ�ü�ת���뱣������Ϣ,лл
'----------------------------------------------------------
'�˵���>����>���� ���microsoft windows common controls �ؼ�
'�ڴ��������һ��TreeView�ؼ�,����Ĭ��ΪTreeView1
'�˵���>����>���� ��� Windows Script Host Object Model
'�˵���>����>���� ��� Microsoft Scripting Runtime

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Dim Fso As New FileSystemObject
Dim OnlyKey As Long '�ڵ��Key
Dim Url() As String '������ַ������
Dim UrlIndex As Long '�ڵ��Ӧ�õ���ַ����ֵ

Private Sub Form_Load()
Me.Left = 2
Me.Top = Fa.P1.Height + Fa.P2.Height
Me.Height = Fa.Height - Me.Top
TreeView1.LineStyle = 1
Dim FavPath As String '�ղؼ�·��
Dim GetFavPath As New IWshRuntimeLibrary.WshShell
FavPath = GetFavPath.SpecialFolders("FAVORITES") & "\" '��ȡ�ղؼ�·��
TreeView1.Nodes.Add , tvwChild, "Root", "IE�ղؼ�" '��Ӹ��ڵ�,KeyΪ"FAVORITES"
GetFiles FavPath, "Root" 'ö�ٸ�Ŀ¼������Url�ļ�
End Sub

Private Sub GetFolder(ByVal path As String, ByVal key As String) 'ö���ļ���
Dim m_Folder As Folder
Dim m_SubFolder As Folder
Set m_Folder = Fso.GetFolder(path)

For Each m_SubFolder In m_Folder.SubFolders
OnlyKey = OnlyKey + 1 'ȷ����ǰ�ڵ��Key��Ψһ��
TreeView1.Nodes.Add key, tvwChild, "Key" & OnlyKey, m_SubFolder.Name
GetFiles path & m_SubFolder.Name & "\", "Key" & OnlyKey 'ö�ٵ�ǰĿ¼������Url�ļ�
Next
End Sub

Private Sub GetFiles(ByVal path As String, ByVal key As String)
Dim m_Folder As Folder
Dim m_File As File
Set m_Folder = Fso.GetFolder(path)
For Each m_File In m_Folder.Files
If Fso.GetExtensionName(path & m_File.Name) = "url" Then '�����Url�ļ�
TreeView1.Nodes.Add key, tvwChild, "Url" & UrlIndex, m_File.Name '��ӵ��ڵ�
ReDim Preserve Url(UrlIndex) '���¶������鳤��
'��ȡ��ַ�����浽������
Dim Retval As String
Retval = Space(256)
GetPrivateProfileString "InternetShortcut", "URL", "", Retval, 256, path & m_File.Name
Url(UrlIndex) = Retval
UrlIndex = UrlIndex + 1
End If
Next
'��ǰĿ¼�µ�Url�ļ�ö����Ϻ�,��Ҫö�ٵ�ȻĿ¼�µ����ļ���
GetFolder path, key 'ö�ٵ�ǰĿ¼�µ����ļ���
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


