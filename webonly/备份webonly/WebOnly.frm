VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FormOnly 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7035
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5145
   Icon            =   "WebOnly.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer TimerXY 
      Interval        =   10
      Left            =   3120
      Top             =   3600
   End
   Begin VB.TextBox ComUrl 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer TimerCon 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2040
      Top             =   3600
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   3836
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "FormOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Fnew     As Form

Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
' �����������б�������λ���κ�������ڵ�ǰ��
Private Const SWP_NOSIZE& = &H1
' ���ִ��ڴ�С
Private Const SWP_NOMOVE& = &H2
' ���ִ���λ��
Dim MostTop As Boolean


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1


'��ȡ����ṹ��Ϣ����

Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000
Private Const WS_CAPTION = &HC00000
Private Const WS_SIZEBOX = &H40000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
'Ϊ����ָ��һ����λ�ú�״̬����

Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOREPOSITION = &H200
'�����������Ĵ�С��λ��
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'��ȡwindows�������߶�Ԥ��

Private Declare Function FindWindow Lib "user32 " Alias "FindWindowA " (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32 " (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
                Left   As Long
                Top   As Long
                Right   As Long
                Bottom   As Long
End Type

Dim Change     As Boolean
Public WithEvents M_Dom As MSHTML.HTMLDocument      '����ie���ѡ��
Attribute M_Dom.VB_VarHelpID = -1


Dim mX As Single
Dim mY As Single
Dim DyX, DyY As Single

'����Ϊ�ƶ�Web��������
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'------------------------------------------------------����Ϊ������״̬
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public LMouseState As Long, RMouseState As Long, LastL As Long, LastR As Long
'---------------------------------------------------------------------����Ϊ���λ��
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
x As Long
y As Long
End Type

'���ֶ���
Dim WithEvents Web_V1 As SHDocVwCtl.WebBrowser_V1   '��֤������һ�����ڴ���ض���
Attribute Web_V1.VB_VarHelpID = -1

Private Sub Command2_Click()
ZOOMIE 20

End Sub

Private Sub Command3_Click()
ZOOMIE 100
End Sub

Private Function M_Dom_onselectstart() As Boolean
If Tvalue = True Then
M_Dom_onselectstart = False
Else
M_Dom_onselectstart = True
End If

End Function


Private Sub ZOOMIE(v As Integer)
Dim webdoc As HTMLDocument
    Set webdoc = Web1.Document
    webdoc.parentWindow.execScript "document.body.style.zoom='" & v & "%'"
End Sub



Private Sub AllScreen_Click()
'GetTaskbarHeight
Me.Top = 0
Me.Left = 0
'Me.Height = Screen.Height - GetTaskbarHeight
Me.Width = Screen.Width

   Web1.Width = Me.Width
   Web1.Height = Me.Height

End Sub

Private Sub Command1_Click()
PopupMenu Museum
End Sub

Private Sub Form_Activate()
    Set Web_V1 = Web1.Object
End Sub



Private Sub Form_Load()

AllScreen_Click

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me

End Sub





Private Sub MostTopCom_Click()


If MostTop = False Then
MostTop = True
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' ��������Ϊ�����д���ǰ��
MostTopCom.Caption = "ȡ���ö�"
Else
SetWindowPos Me.hwnd, 1, 0, 0, 0, 0, 3
Me.ZOrder
MostTop = False
MostTopCom.Caption = "�ö�"
Timer1.Enabled = False
End If
End Sub



Private Sub Timer1_Timer()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' ��������Ϊ�����д���ǰ��
End Sub



Private Sub Tuo_Click()
If Tuo.Caption = "��ק���" Then
Tvalue = True
Me.ScaleMode = vbPixels
'Set M_Dom = Web1.Document '��ֹ���ѡ���ı�
Tuo.Caption = "�������"
Else
Tuo.Caption = "��ק���"
Tvalue = False
'Set M_Dom = Web1.Document '�������ѡ���ı�
Me.ScaleMode = 1
Web1.Top = 0
Web1.Left = 0
Web1.Height = Me.Height - 256
Web1.Width = Me.Width - 256
End If

End Sub

Private Sub Web1_DownloadComplete()
Me.Caption = Web1.LocationName
Me.Icon = Web1.DragIcon


Web1.Silent = True
End Sub



Private Sub Web1_DownloadBegin()
 On Error Resume Next
        Web1.Silent = True
End Sub


Private Sub Web_V1_NewWindow(ByVal URL As String, _
    ByVal Flags As Long, ByVal TargetFrameName As String, _
    PostData As Variant, ByVal Headers As String, Processed As Boolean)
    
    Processed = True
    
    Dim FrmWeb As New FormOnly
    FrmWeb.Web1.Navigate (URL)
    FrmWeb.Show
    
End Sub

Private Sub Web1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
     
     
     
     On Error Resume Next
     Dim js As String
    '�ű������ڲ��� ��갴��
     js = "document.body.onmousedown=function()" & vbCrLf & _
       "{location.href='mouse://down|'+window.event.x + '|'+window.event.y;}"
     Web1.Document.parentWindow.execScript js, "javascript"
    '�ű������ڲ��� ����ƶ�
    ' js = "document.body.onmousemove=function()" & vbCrLf & _
    ' "{location.href='mouse://move|'+window.event.x + '|'+window.event.y;}"
    ' Web1.Document.parentWindow.execScript js, "javascript"
    '�ű������ڲ��� ���̧��
     js = "document.body.onmouseup=function()" & vbCrLf & _
      "{location.href='mouse://up|'+window.event.x + '|'+window.event.y;}"
     Web1.Document.parentWindow.execScript js, "javascript"
     
     
End Sub

Private Sub Web1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
   
  Dim Point As POINTAPI
  GetCursorPos Point

    Dim nStr As String
    nStr = URL
    If Left(nStr, 8) <> "mouse://" Then Exit Sub '����ҳ������ת
       
    Cancel = True '����ҳ��Ҫ��ת
    Dim nMouse As String, x As Single, y As Single, S As Single
    nStr = Mid(nStr, 9)
    S = InStr(nStr, "|")
    nMouse = Left(nStr, S - 1): nStr = Mid(nStr, S + 1)
    S = InStr(nStr, "|")
    x = Val(Left(nStr, S - 1)): y = Val(Mid(nStr, S + 1))
   ' Me.Caption = nMouse & "��" & x & " " & y '��ʾ��겶���״̬
    If nMouse = "down" Then
    TimerXY.Enabled = True
    Debug.Print "down"
     mX = Point.x
     mY = Point.y
    End If
    If nMouse = "up" Then
     TimerXY.Enabled = False
    End If
    
'DyX = X
'DyY = Y

If Web1.Top > 0 Then
Web1.Top = 0
Web1.Height = Me.Height
Else
Web1.Height = 0 - Web1.Top + Me.Height - 200
End If

Web1.Width = 1400



    
End Sub
Private Sub TimerXY_Timer()
Dim Point As POINTAPI
GetCursorPos Point

Web1.Top = Web1.Top + Point.y - mY
Web1.Left = Web1.Left + Point.x - mX
'Web1.Document.parentWindow.scrollBy ,
mX = Point.x
mY = Point.y


End Sub
