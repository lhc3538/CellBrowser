Attribute VB_Name = "MOUSEWHEEL"
'��ģ����������ֹ�����Ϣ�����Լ��Ĵ���
Public Const PM_REMOVE = &H1

Public Type Msg
 hWnd As Long
 Message As Long
 wParam As Long

 time As Long
End Type

Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Public Declare Function WaitMessage Lib "user32" () As Long
Public bCancel As Boolean
Public Const WM_MOUSEWHEEL = 522

