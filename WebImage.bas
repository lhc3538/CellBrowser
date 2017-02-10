Attribute VB_Name = "ApiImage"

Option Explicit
Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 _
As String) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOPENFILENAME As OPENFILENAME) As Long

Public LinkPageNum As Integer 'Linkpage>>addpage









 

 
Public Const WM_PRINT As Long = &H317
Public Const PRF_CHECKVISIBLE = &H1
Public Const PRF_NONCLIENT = &H2
Public Const PRF_CLIENT = &H4
Public Const PRF_ERASEBKGND = &H8
Public Const PRF_CHILDREN = &H10
Public Const PRF_OWNED = &H20

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205



