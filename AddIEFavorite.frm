VERSION 5.00
Begin VB.Form AddIEFavorite 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
End
Attribute VB_Name = "AddIEFavorite"
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

  '��������
   Dim success As Long
   Dim pos As Long
   Dim nSizeOfPath As Long
   Dim nSizeOfTitle As Long
   Dim pidl As Long
   Dim szPath As String
  
  '׷��chr$(0)�ַ�
   szTitle = szTitle & Chr$(0)
   nSizeOfTitle = Len(szTitle)
   
  '����·�����ַ���
   szPath = Space$(MAX_PATH) & Chr$(0)
   nSizeOfPath = Len(szPath)
   
  '�õ��û����ղؼС�·����PIDL (pointer to item identifier list)
  '�ɹ��󷵻�ֵΪERROR_SUCCESS
   If SHGetSpecialFolderLocation(hWnd, _
                                 CSIDL_FAVORITES, _
                                 pidl) = ERROR_SUCCESS Then
        
     '���á���ӵ��ղؼС��Ի���
     'hwnd = �����ڵľ��
     'szPath =  ��ѡ���ļ��еľ���·���������ļ����������URL
     '                ���磬���ҵ�ϵͳ�����C:\Documents and Settings\40Star\Favorites\CSDN.NET--�й����Ŀ���������.url
     'szTitle = ����
     'pidl    =    PIDL �����û����ղؼе���Ϣ
      success = DoAddToFavDlg(hWnd, _
                              szPath, nSizeOfPath, _
                              szTitle, nSizeOfTitle, _
                              pidl)

     '���·����Ч��ָ���˱��⣬�����û�ѡ���ˡ�ȷ������success ���� 1
      If success = 1 Then
      
        'ɾ������Chr$ (0)
         pos = InStr(szPath, Chr$(0))
         szPath = Left(szPath, pos - 1)
         
         pos = InStr(szTitle, Chr$(0))
         szTitle = Left(szTitle, pos - 1)
      
        '��Text����ʾ���
'         Text1.Text = szPath
'         Text2.Text = szTitle
      
         Call ProfileSaveItem("InternetShortcut", "URL", sURL, szPath)
         
        '���ش����ɹ���·��
         MakeFavouriteEntry = szPath
      
      End If
      
     '���PIDL
      Call CoTaskMemFree(pidl)

   End If

End Function








Private Sub Form_Load()

   Dim szTitle As String
   Dim sURL As String
   Dim sResult As String

  'ָ����ӵ��ղؼк�Ŀ�ݷ�ʽ������
   szTitle = Fa.ActiveForm.WebPage.Document.Title
   
  'ָ����ӵ��ղؼк�Ŀ�ݷ�ʽ��URL
   sURL = Fa.ActiveForm.WebPage.LocationURL
   
  '����MakeFavouriteEntry�������򿪶Ի���
   sResult = MakeFavouriteEntry(szTitle, sURL)
   Unload Me
End Sub
