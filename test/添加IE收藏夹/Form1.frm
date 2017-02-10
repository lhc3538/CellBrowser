VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_PATH As Long = 260
Private Const ERROR_SUCCESS As Long = 0
Private Const S_OK As Long = 0
Private Const S_FALSE As Long = 1
Private Const SHGFP_TYPE_CURRENT As Long = &H0
Private Const SHGFP_TYPE_DEFAULT As Long = &H1
Const CSIDL_FAVORITES As Long = &H6

Private Declare Function DoAddToFavDlg Lib "shdocvw" _
  (ByVal hWnd As Long, _
   ByVal szPath As String, _
   ByVal nSizeOfPath As Long, _
   ByVal szTitle As String, _
   ByVal nSizeOfTitle As Long, _
   ByVal pidl As Long) As Long
   
Private Declare Function DoOrganizeFavDlg Lib "shdocvw" _
  (ByVal hWnd As Long, _
   ByVal lpszRootFolder As String) As Long

Private Declare Function SHGetFolderPath Lib "shfolder" _
   Alias "SHGetFolderPathA" _
  (ByVal hwndOwner As Long, _
   ByVal nFolder As Long, _
   ByVal hToken As Long, _
   ByVal dwReserved As Long, _
   ByVal lpszPath As String) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32" _
  (ByVal hwndOwner As Long, _
   ByVal nFolder As Long, _
   pidl As Long) As Long
   
Private Declare Function WritePrivateProfileString Lib "kernel32" _
   Alias "WritePrivateProfileStringA" _
  (ByVal lpSectionName As String, _
   ByVal lpKeyName As Any, _
   ByVal lpString As Any, _
   ByVal lpFileName As String) As Long
   
Private Declare Sub CoTaskMemFree Lib "ole32" _
   (ByVal pv As Long)

Public Sub ProfileSaveItem(lpSectionName As String, _
                           lpKeyName As String, _
                           lpValue As String, _
                           iniFile As String)

   Call WritePrivateProfileString(lpSectionName, lpKeyName, lpValue, iniFile)

End Sub



Private Function MakeFavouriteEntry(szTitle As String, sURL As String) As String

  '变量定义
   Dim success As Long
   Dim pos As Long
   Dim nSizeOfPath As Long
   Dim nSizeOfTitle As Long
   Dim pidl As Long
   Dim szPath As String
  
  '追加chr$(0)字符
   szTitle = szTitle & Chr$(0)
   nSizeOfTitle = Len(szTitle)
   
  '返回路径的字符串
   szPath = Space$(MAX_PATH) & Chr$(0)
   nSizeOfPath = Len(szPath)
   
  '得到用户“收藏夹”路径的PIDL (pointer to item identifier list)
  '成功后返回值为ERROR_SUCCESS
   If SHGetSpecialFolderLocation(hWnd, _
                                 CSIDL_FAVORITES, _
                                 pidl) = ERROR_SUCCESS Then
        
     '调用“添加到收藏夹”对话框
     'hwnd = 本窗口的句柄
     'szPath =  所选择文件夹的绝对路径，包括文件名和所需的URL
     '                例如，在我的系统里就是C:\Documents and Settings\40Star\Favorites\CSDN.NET--中国最大的开发者网络.url
     'szTitle = 标题
     'pidl    =    PIDL 描述用户的收藏夹的信息
      success = DoAddToFavDlg(hWnd, _
                              szPath, nSizeOfPath, _
                              szTitle, nSizeOfTitle, _
                              pidl)

     '如果路径有效并指定了标题，而且用户选择了“确定”，success 返回 1
      If success = 1 Then
      
        '删除最后的Chr$ (0)
         pos = InStr(szPath, Chr$(0))
         szPath = Left(szPath, pos - 1)
         
         pos = InStr(szTitle, Chr$(0))
         szTitle = Left(szTitle, pos - 1)
      
        '在Text中显示结果
         Text1.Text = szPath
         Text2.Text = szTitle
      
         Call ProfileSaveItem("InternetShortcut", "URL", sURL, szPath)
         
        '返回创建成功的路径
         MakeFavouriteEntry = szPath
      
      End If
      
     '清空PIDL
      Call CoTaskMemFree(pidl)

   End If

End Function



Private Sub Command1_Click()
   Dim szTitle As String
   Dim sURL As String
   Dim sResult As String

  '指定添加到收藏夹后的快捷方式的名称
   szTitle = Text1.Text
   
  '指定添加到收藏夹后的快捷方式的URL
   sURL = Text2.Text
   
  '调用MakeFavouriteEntry函数，打开对话框
   sResult = MakeFavouriteEntry(szTitle, sURL)
End Sub


