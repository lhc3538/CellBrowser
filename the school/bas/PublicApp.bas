Attribute VB_Name = "PublicApp"
Option Explicit

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
                             (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const max_path = 260

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'以上为删除背景色-------------------------------------------------------------------



Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Public Const SWP_NOSIZE& = &H1
' 保持窗口大小
Public Const SWP_NOMOVE& = &H2
' 保持窗口位置
'以上为窗体总在最前端--------------------------------------------------------------
Public FrmWeb(0 To 512) As New WebPage '创建、定义页面
Public PageNum As Integer '打开网页的累计数（包括已关闭的）
Public ActivePage As Integer '当前活动网页的ID

Public BackRed, BackGreen, BackBlue As Integer

