Attribute VB_Name = "PublicApp"
Option Explicit

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
                             (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const max_path = 260

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'����Ϊɾ������ɫ-------------------------------------------------------------------



Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST& = -1
' �����������б�������λ���κ�������ڵ�ǰ��
Public Const SWP_NOSIZE& = &H1
' ���ִ��ڴ�С
Public Const SWP_NOMOVE& = &H2
' ���ִ���λ��
'����Ϊ����������ǰ��--------------------------------------------------------------
Public FrmWeb(0 To 512) As New WebPage '����������ҳ��
Public PageNum As Integer '����ҳ���ۼ����������ѹرյģ�
Public ActivePage As Integer '��ǰ���ҳ��ID

Public BackRed, BackGreen, BackBlue As Integer

