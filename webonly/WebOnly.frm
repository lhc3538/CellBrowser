VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FormOnly 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5145
   Icon            =   "WebOnly.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   5145
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer TimerXY 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2760
      Top             =   3600
   End
   Begin VB.CommandButton Command1 
      Height          =   105
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "�˵�"
      Top             =   0
      Width           =   105
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
   Begin VB.Menu Museum 
      Caption         =   "�˵�"
      Begin VB.Menu MostTopCom 
         Caption         =   "�ö�"
      End
      Begin VB.Menu AllScreen 
         Caption         =   "ȫ��չʾ"
      End
      Begin VB.Menu Tuo 
         Caption         =   "��ק���"
      End
   End
End
Attribute VB_Name = "FormOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Fnew     As Form

Option Explicit

Dim WithEvents Web_V1 As SHDocVwCtl.WebBrowser_V1  '��֤������һ�����ڴ���ض���
Attribute Web_V1.VB_VarHelpID = -1

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

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


Dim MX As Single
Dim MY As Single
Dim DyX, DyY As Single
Dim Tvalue As Boolean

Private Type POINTAPI
X As Long
Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


'------------------------------------------------------����Ϊ������״̬
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public LMouseState As Long, RMouseState As Long, LastL As Long, LastR As Long
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
If Me.WindowState = 2 Then Me.WindowState = 0

If AllScreen.Caption = "ȫ��չʾ" Then
Me.Top = -450
Me.Left = 0
Me.Height = Screen.Height
Me.Width = Screen.Width
Timer1.Enabled = True
AllScreen.Caption = "ȡ��ȫ��"
Else
Me.Top = 500
Me.Left = 300
Me.Height = 9500
Me.Width = 15600
Timer1.Enabled = False
SetWindowPos Me.hwnd, 1, 0, 0, 0, 0, 3
Me.ZOrder
MostTop = False
MostTopCom.Caption = "�ö�"
AllScreen.Caption = "ȫ��չʾ"
End If


End Sub

Private Sub Command1_Click()
PopupMenu Museum
End Sub

Private Sub Form_Activate()
    Set Web_V1 = Web1.Object
End Sub



Private Sub Form_Load()


Tvalue = False

MostTop = False
Dim URL

Open App.Path & "\OutUrl.dat" For Input As #1
Do While Not EOF(1)
Input #1, URL
Loop
Close #1
Museum.Visible = False
Web1.Navigate URL
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me

End Sub

Private Sub Form_Resize()
If Tvalue = False Then
Web1.Height = Me.Height - 256
Web1.Width = Me.Width - 64
End If

Me.Caption = Web1.LocationName
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
    Dim nMouse As String, X As Single, Y As Single, S As Single
    nStr = Mid(nStr, 9)
    S = InStr(nStr, "|")
    nMouse = Left(nStr, S - 1): nStr = Mid(nStr, S + 1)
    S = InStr(nStr, "|")
    X = Val(Left(nStr, S - 1)): Y = Val(Mid(nStr, S + 1))
   ' Me.Caption = nMouse & "��" & x & " " & y '��ʾ��겶���״̬
    If nMouse = "down" Then
    TimerXY.Enabled = True
    
     MX = Point.X
     MY = Point.Y
    End If
    If nMouse = "up" Then
     TimerXY.Enabled = False
    End If
    
'DyX = X
'DyY = Y
If Tvalue = True Then
If Web1.Top > 0 Then
Web1.Top = 0
Web1.Height = Me.Height
Else
Web1.Height = 0 - Web1.Top + Me.Height - 200
End If

Web1.Width = 1400


End If


    
End Sub
Private Sub TimerXY_Timer()
Dim Point As POINTAPI
GetCursorPos Point

If Tvalue = False Then
Timer1.Enabled = False
Exit Sub
End If

LMouseState = GetAsyncKeyState(vbKeyLButton)
RMouseState = GetAsyncKeyState(vbKeyRButton)
If LMouseState < 0 Or RMouseState < 0 Then
 Web1.Top = Web1.Top + Point.Y - MY
 Web1.Left = Web1.Left + Point.X - MX
'Web1.Document.parentWindow.scrollBy ,
 MX = Point.X
 MY = Point.Y
End If

If LastR < 0 And RMouseState Or LastL < 0 And LMouseState = 0 = 0 Then ' "��� ���Ҽ� �� ̧�� ������һ�Σ�"
Tvalue = False
LastR = RMouseState
LastL = LMouseState
Else
LastR = RMouseState
LastL = LMouseState
End If
End Sub
