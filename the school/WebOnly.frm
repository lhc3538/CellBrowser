VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form WebPage 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5055
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4995
   Icon            =   "WebOnly.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4995
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   -2400
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   4695
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   2955
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   195
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   308
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   4657
      End
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   3015
      LargeChange     =   10
      Left            =   2280
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   -2400
      TabIndex        =   11
      Top             =   4995
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.PictureBox LoadingBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   -1560
      ScaleHeight     =   105
      ScaleWidth      =   3585
      TabIndex        =   8
      Top             =   4080
      Width           =   3615
      Begin VB.PictureBox PMoveBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   735
         TabIndex        =   9
         Top             =   0
         Width           =   735
         Begin VB.Timer TimerLoading 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   120
            Top             =   0
         End
      End
   End
   Begin VB.Timer TimerWebIng 
      Enabled         =   0   'False
      Interval        =   64
      Left            =   1200
      Top             =   1320
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   3240
      ScaleHeight     =   2025
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   720
      Width           =   495
      Begin VB.PictureBox P2 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   0
         ScaleHeight     =   1335
         ScaleWidth      =   495
         TabIndex        =   4
         Top             =   0
         Width           =   495
         Begin VB.Timer TimerP2 
            Enabled         =   0   'False
            Interval        =   64
            Left            =   0
            Top             =   600
         End
      End
   End
   Begin VB.TextBox WebID 
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer TimerXY 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3000
      Top             =   3960
   End
   Begin VB.TextBox ComUrl 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer TimerCon 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2520
      Top             =   3960
   End
   Begin VB.PictureBox WebImBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   4335
      TabIndex        =   5
      Top             =   0
      Width           =   4335
      Begin VB.PictureBox PageEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   10
         Top             =   64
         Width           =   615
         Begin VB.Image Imageend 
            Height          =   375
            Left            =   240
            Picture         =   "WebOnly.frx":324A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.TextBox TWCaption 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   260
         Left            =   120
         TabIndex        =   7
         Text            =   "�ޱ���"
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox TWUrl 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   6
         Text            =   "Http��//"
         Top             =   64
         Width           =   2175
      End
      Begin VB.CommandButton ComGo 
         Caption         =   "Command1"
         Default         =   -1  'True
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   120
         Width           =   255
      End
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   2055
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   2775
      ExtentX         =   4895
      ExtentY         =   3625
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
   Begin VB.CommandButton WebImage 
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "WebPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
' �����������б�������λ���κ�������ڵ�ǰ��
Private Const SWP_NOSIZE& = &H1
' ���ִ��ڴ�С
Private Const SWP_NOMOVE& = &H2
' ���ִ���λ��
Dim MostTop As Boolean


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
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
Private Declare Function GetWindowRect Lib "user32 " (ByVal hWnd As Long, lpRect As RECT) As Long
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

'����Ϊ�ƶ�Web��������
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'------------------------------------------------------����Ϊ������״̬
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public LMouseState As Long, RMouseState As Long, LastL As Long, LastR As Long
'---------------------------------------------------------------------����Ϊ���λ��
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
X As Long
Y As Long
End Type
'------------------------------------------------------����Ϊ��ϵͳ���� ����������
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long

'����
Private mlngLeft As Long
Private mlngTop As Long
Private mlngWidth As Long
Private mlngHeight As Long
'����webbrowser�Ĵ�Сλ��

Dim openID As Integer '��ҳ��ʱ���ض����Ĵ���
Dim WebImageTF As Boolean '�ж��Ƿ��ѽ���


'���ֶ���
Dim WithEvents Web_V1 As SHDocVwCtl.WebBrowser_V1   '��֤������һ�����ڴ���ض���
Attribute Web_V1.VB_VarHelpID = -1
Dim M_y As Integer '������λ��
Dim MoveTime As Integer
Dim CsTop, vTop As Integer
Dim FixX, FixY As Integer '�����ԭʼλ��


Private Sub ComGo_Click()
Web1.navigate TWUrl.Text
End Sub

Private Sub ComUrl_Change()
Web1.navigate ComUrl.Text

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Fcbox.ConBack.Text = "" '�ָ���ť����ɫ
End Sub

Private Sub Imageend_Click()
Unload Me
End Sub

Private Sub Imageend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PageEnd.BackColor = &O0
End Sub



Private Sub Imageend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PageEnd.BackColor = &H404040
End Sub

Private Function M_Dom_onselectstart() As Boolean '?????????
'M_Dom_onselectstart = False
End Function


Private Sub ZOOMIE(v As Integer)
Dim webdoc As HTMLDocument
    Set webdoc = Web1.document
    webdoc.parentWindow.execScript "document.body.style.zoom='" & v & "%'"
End Sub

Private Sub AllScreen_Click()
GetTaskbarHeight
Me.Left = Fcbox.Width - 20
Me.Width = Screen.Width - Me.Left
Me.Top = 0
Me.Height = Screen.Height - GetTaskbarHeight
Me.BackColor = RGB(64, 64, 64)
 WebImBar.Width = Me.Width
  TWCaption.Width = WebImBar.Width / 4
  TWUrl.Left = TWCaption.Left + TWCaption.Width + 256
  TWUrl.Width = WebImBar.Width / 2
  PageEnd.Left = WebImBar.Width - PageEnd.Width - 64
 Web1.Top = WebImBar.Height
 Web1.Left = 0
 LoadingBar.Left = Me.Width / 2 - LoadingBar.Width / 2
 LoadingBar.Top = Me.Height / 2 - LoadingBar.Height
   Web1.Width = Me.Width
 P1.Top = WebImBar.Height
 P1.Height = Me.Height - WebImBar.Height
 P1.Left = Web1.Left + Web1.Width - 260 '������
End Sub

Private Sub Form_Activate()
    Set Web_V1 = Web1.Object 'Web_1��Web1��ϵ
 ActivePage = Val(WebID.Text) '���ݵ�ǰ������ID
MoveTime = 1 '���³���ʱ��Ϊ0
 'ProcessMessages '�������ּ��
FrmList.AllPageback '�б�ѡ��״̬
End Sub



Private Sub Form_Load()
 BackRed = 255
 BackGreen = 255
 BackBlue = 200
AllScreen_Click '���岼��
Me.AutoRedraw = True '������״̬����


'webimage
    VScroll1.Enabled = False
    HScroll1.Enabled = False
   ' WebPage.Width = Screen.Width
    mlngLeft = WebPage.Left
    mlngTop = WebPage.Top
    mlngWidth = WebPage.Width
    mlngHeight = WebPage.Height

WebImageTF = False

'FrmList.PageBack(Val(WebID.Text)).Visible = True '��ʾ��Ӧ���б�
FrmList.TimerCorrect.Enabled = True '�����б�
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim n As Integer
n = Val(WebID.Text)
FrmList.PageBack(n).Visible = False
FrmList.PageBack(n).Cls

'FrmList.TimerCorrect.Enabled = True '�����б�
Unload Me

End Sub


Private Sub P2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
M_y = Y
P2.BackColor = &H0&
TimerP2.Enabled = True '����ҳ�����ƶ�
End Sub


Private Sub P2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Point As POINTAPI
GetCursorPos Point

'���¼�����״̬-----------------------------------------
LMouseState = GetAsyncKeyState(vbKeyLButton)
If LMouseState < 0 Then
'Print "��� ��� ���� ��ס ״̬"
 If P2.Top >= 0 And P2.Top <= P1.Height - P2.Height Then
  P2.Top = (Y - M_y) + P2.Top
  If P2.Top < 0 Then P2.Top = 0
  If P2.Top > P1.Height - P2.Height Then P2.Top = P1.Height - P2.Height
  

 End If

End If
End Sub

Private Sub P2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
P2.BackColor = &H808080
TimerP2.Enabled = False '�ر���ҳ�����ƶ�
End Sub





Private Sub PageEnd_Click()
Unload Me
End Sub

Private Sub PageEnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PageEnd.BackColor = &H0&
End Sub

Private Sub PageEnd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PageEnd.BackColor = &H404040
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub TimerCon_Timer()
Web1.Top = Web1.Top + vTop
If Web1.Top > WebImBar.Height Then  '���Ʊ������ӷ�Χ
  Web1.Top = WebImBar.Height
  TimerCon.Enabled = False
  vTop = 0
  MoveTime = 1
End If
If -Web1.Top > Web1.Height - Me.Height Then
  Web1.Top = -(Web1.Height - Me.Height)
  TimerCon.Enabled = False
  vTop = 0
  MoveTime = 1
End If
If vTop > 0 Then
 vTop = vTop - 8
Else
 vTop = vTop + 8
End If
If Abs(vTop) < 9 Then
 TimerCon.Enabled = False
 vTop = 0
 MoveTime = 1
End If
If Web1.Height = Me.Height Then '��ֹ���
 P1.Visible = False
Else
 P1.Visible = True
 P2.Top = (-Web1.Top + WebImBar.Height) * (P1.Height - P2.Height) / (Web1.Height - Me.Height + WebImBar.Height)  '����������
End If
End Sub


Private Sub TimerLoading_Timer()
PMoveBar.Left = PMoveBar.Left + 32
If PMoveBar.Left > LoadingBar.Width Then
 PMoveBar.Left = -PMoveBar.Width
End If
End Sub

Private Sub TimerP2_Timer()

  Web1.Top = -P2.Top * (Web1.Height - Me.Height + WebImBar.Height) / (P1.Height - P2.Height) + WebImBar.Height '��ҳλ�ø���
End Sub

Private Sub TimerWebIng_Timer()
On Error Resume Next


If Web1.Busy = True Then

 Web1.Left = -Me.Width '������ҳ
 LoadingBar.Visible = True
 TimerLoading.Enabled = True
 TWCaption.Text = "���ڴ򿪣� " & Web1.LocationName
 TWUrl.Text = Web1.LocationURL
Else
 Web1.Left = 0 '��ʾ��ҳ
 LoadingBar.Visible = False
 TimerLoading.Enabled = False
' FopenBox.ShowOpenpic
 TWCaption.Text = Web1.document.Title
 TWUrl.Text = Web1.LocationURL
 TimerWebIng.Enabled = False

 '------���½�ȡ��ҳͼƬ
   DoEvents
   WebImage.SetFocus
   WebImage_Click '�����¼�ֻ��ȡһ����ҳ
End If

End Sub

Private Sub TWUrl_GotFocus()
TWUrl.BackColor = vbWhite
End Sub

Private Sub TWUrl_LostFocus()
TWUrl.BackColor = &HE0E0E0
End Sub



Private Sub Web_V1_BeforeNavigate(ByVal URL As String, ByVal Flags As Long, ByVal TargetFrameName As String, PostData As Variant, ByVal Headers As String, Cancel As Boolean) '��ҳ�ڲ���ת
 openID = 0
 

End Sub

Private Sub Web_V1_DownloadBegin()
Static loadnum As Integer '  ����ֹ�ظ�����
 If loadnum >= 3 Then
  Web1.Left = 0
  Exit Sub
 End If
  loadnum = loadnum + 1
  TimerWebIng.Enabled = True '����ҳ��æ״̬���
End Sub

Private Sub Web_V1_DownloadComplete()
On Error Resume Next
   FrmList.Page(Val(WebID.Text)).Picture = Picture1.image  '����ҳ��ʾ���б���
 If Web1.Height < Me.Height Then

 End If
Me.Caption = Web1.document.Title
FrmList.WebName(Val(WebID.Text)).Caption = Web1.document.Title   '����ҳ������ʾ���б���

End Sub

Private Sub Web1_DownloadComplete()
Me.Caption = Web1.LocationName
End Sub



Private Sub Web1_DownloadBegin()
 On Error Resume Next
        Web1.Silent = True


End Sub


Private Sub Web_V1_NewWindow(ByVal URL As String, _
    ByVal Flags As Long, ByVal TargetFrameName As String, _
    PostData As Variant, ByVal Headers As String, Processed As Boolean)
    
    Processed = True
    

    FrmWeb(PageNum).Web1.navigate (URL)
    FrmWeb(PageNum).Show
    FrmWeb(PageNum).WebID.Text = Str(PageNum)
      FrmList.PageBack(PageNum).Visible = True
    PageNum = PageNum + 1
End Sub

Private Sub Web1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
     
     
     
     On Error Resume Next
     Dim js As String
    '�ű������ڲ��� ��갴��
     js = "document.body.onmousedown=function()" & vbCrLf & _
       "{location.href='mouse://down|'+window.event.x + '|'+window.event.y;}"
     Web1.document.parentWindow.execScript js, "javascript"
    '�ű������ڲ��� ����ƶ�
    ' js = "document.body.onmousemove=function()" & vbCrLf & _
    ' "{location.href='mouse://move|'+window.event.x + '|'+window.event.y;}"
    ' Web1.Document.parentWindow.execScript js, "javascript"
    '�ű������ڲ��� ���̧��
     js = "document.body.onmouseup=function()" & vbCrLf & _
      "{location.href='mouse://up|'+window.event.x + '|'+window.event.y;}"
     Web1.document.parentWindow.execScript js, "javascript"
     
     
End Sub

Private Sub Web1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

 
  Dim Point As POINTAPI
  GetCursorPos Point

    Dim nStr As String
    nStr = URL
    If Left(nStr, 8) <> "mouse://" Then Exit Sub '����ҳ������ת
       
    Cancel = True '����ҳ��Ҫ��ת
    Dim nMouse As String, X As Single, Y As Single, s As Single
    nStr = Mid(nStr, 9)
    s = InStr(nStr, "|")
    nMouse = Left(nStr, s - 1): nStr = Mid(nStr, s + 1)
    s = InStr(nStr, "|")
    X = Val(Left(nStr, s - 1)): Y = Val(Mid(nStr, s + 1))
    ' Me.Caption = nMouse & "��" & x & " " & y '��ʾ��겶���״̬
    If nMouse = "down" Then
     CsTop = Web1.Top
     TimerXY.Enabled = True
     TimerCon.Enabled = False
     'Debug.Print "down"
     MX = Point.X '��̬��¼
     MY = Point.Y
     FixX = Point.X '��̬��¼
     FixY = Point.Y
    End If
    If nMouse = "up" Then
     vTop = (Web1.Top - CsTop) / MoveTime
     TimerCon.Enabled = True
     TimerXY.Enabled = False
     'Debug.Print "up"
     If FixX = Point.X And FixY = Point.Y Then '��ȡ��ʾ�����ذ�ťָ��

      ' If Fcbox.GoLeft = False Then '��������б���ʾ��������
      '   Fcbox.GoLeft = True
      '   Fcbox.Timermove.Enabled = True
      ' End If
       Fcbox.ConBack.Text = "" '�ָ���ť����ɫ

     End If
    End If
    
'DyX = X
'DyY = Y

End Sub
Private Sub TimerXY_Timer()
Dim Point As POINTAPI
GetCursorPos Point

Web1.Top = Web1.Top + (Point.Y - MY) * 15
'Web1.Left = Web1.Left + (Point.X - mX) * 15
'Web1.Document.parentWindow.scrollBy ,

MX = Point.X
MY = Point.Y

'���¼�����״̬-----------------------------------------
LMouseState = GetAsyncKeyState(vbKeyLButton)
If LMouseState < 0 Then
'Print "��� ��� ���� ��ס ״̬"
 MoveTime = MoveTime + 1 '��¼ʱ��
End If
'If LMouseState = 1 Then Print "��� ��� ����һ��"
If LastL < 0 And LMouseState = 0 Then
'Print "��� ��� �� ̧�� ������һ�Σ�"
 vTop = (Web1.Top - CsTop) / MoveTime
 TimerXY.Enabled = False
 'MsgBox (Str(vTop))
 TimerCon.Enabled = True
 LastL = LMouseState
Else
 LastL = LMouseState
End If

RMouseState = GetAsyncKeyState(vbKeyRButton)
'If RMouseState < 0 Then Print "��� �Ҽ� ���� ��ס ״̬"
'If RMouseState = 1 Then Print "��� �Ҽ� ����һ��"
If LastR < 0 And RMouseState = 0 Then
'Print "��� �Ҽ� �� ̧�� ������һ�Σ�"
 TimerXY.Enabled = False
 LastR = RMouseState
Else
 LastR = RMouseState
End If
If Web1.Left + Web1.Width - 256 >= Me.Width - 128 Then
 P1.Left = Me.Width - 128
Else
 P1.Left = Web1.Left + Web1.Width - 256
End If
If Web1.Height = Me.Height Then '��ֹ���
 P1.Visible = False
Else
 P2.Top = (-Web1.Top + WebImBar.Height) * (P1.Height - P2.Height) / (Web1.Height - Me.Height + WebImBar.Height)  '����������
End If
End Sub

'Private Sub ProcessMessages() '�����ֲ�������(���ȶ�)
' Dim Message As Msg

' Do While Not bCancel
' WaitMessage '�ȴ���Ϣ

 'If PeekMessage(Message, Me.hWnd, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then '...when the mousewheel is used...

' If Message.wParam < 0 Then '���Ϲ���
 ' Web1.Top = Web1.Top - 1024
 '  vTop = 0
'   TimerCon.Enabled = True
' Else '���¹���
'  Web1.Top = Web1.Top + 1024
'   vTop = 0
'   TimerCon.Enabled = True
' End If
' End If
  
' DoEvents
' Loop
' End Sub


'��ȡ��ҳͼƬ
Private Sub BrowserFullSize()
    Dim WidthInPixel As Double
    Dim HeightInPixel As Double
    Dim htmDoc As Object

    Set htmDoc = Web1.document

    WidthInPixel = htmDoc.body.scrollWidth + (htmDoc.body.offsetWidth - htmDoc.body.clientWidth)
    HeightInPixel = htmDoc.body.scrollHeight + (htmDoc.body.offsetHeight - htmDoc.body.clientHeight)

    Web1.Left = 0
    Web1.Top = 0
   ' web1.Width = WidthInPixel * Screen.TwipsPerPixelX
   ' web1.Height = HeightInPixel * Screen.TwipsPerPixelY
    Picture1.Width = Web1.Width / 15
    Picture1.Height = Web1.Height / 15
    Set Picture1.Picture = LoadPicture() '����ɵ�ͼ��
    Picture1.AutoRedraw = True
    VScroll1.Max = Picture1.ScaleHeight - Picture2.ScaleHeight
    If VScroll1.Max > 0 Then VScroll1.Enabled = True
    HScroll1.Max = Picture1.ScaleWidth - Picture2.ScaleWidth
    If HScroll1.Max > 0 Then HScroll1.Enabled = True
    Set htmDoc = Nothing
End Sub
Private Sub BrowserNormalSize()
    Web1.Left = mlngLeft
    Web1.Top = mlngTop
    Web1.Width = mlngWidth
    Web1.Height = mlngHeight
End Sub
Private Sub Form_Resize()
Web1.Left = 0
    mlngHeight = Web1.Height
'webimage
    Picture2.Move 0, 0, Me.ScaleWidth - VScroll1.Width, Me.ScaleHeight - HScroll1.Height
    VScroll1.Move Me.ScaleWidth - VScroll1.Width, 0, VScroll1.Width, Me.ScaleHeight - WebImage.Height
    HScroll1.Move 0, Me.ScaleHeight - HScroll1.Height, Me.ScaleWidth - WebImage.Width
    WebImage.Move Me.ScaleWidth - WebImage.Width, Me.ScaleHeight - WebImage.Height
    VScroll1.Max = Picture1.ScaleHeight - Picture2.ScaleHeight
    If VScroll1.Max > 0 Then VScroll1.Enabled = True
    HScroll1.Max = Picture1.ScaleWidth - Picture2.ScaleWidth
    If HScroll1.Max > 0 Then HScroll1.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ShowCaret (HScroll1.hWnd)
    ShowCaret (VScroll1.hWnd)
End Sub

Private Sub WebImage_Click()
   Web1.Height = Me.Height
 
 
On Error Resume Next
Dim lngHwnd As Long
    
    '������ʾ������ҳ
    BrowserFullSize
    
    '���WebBrowser��hwnd
    lngHwnd = FindWindowEx(Me.hWnd, 0, "Shell Embedding", vbNullString)
    lngHwnd = FindWindowEx(lngHwnd, 0, "Shell DocObject View", vbNullString)
    '��ҳ���գ���ʾ��Picture1�У�֮������ͼ�񱣴���ļ�����
    SendMessage lngHwnd, WM_PRINT, Picture1.hdc, PRF_CHILDREN Or PRF_CLIENT Or PRF_ERASEBKGND Or PRF_NONCLIENT Or PRF_OWNED
    DoEvents
   ' Picture1.Refresh
    Picture1.Visible = True
'    mnuSave.Enabled = True
 '   mnuClipboard.Enabled = True
    '......
    
    '�ָ�ԭ���Ĵ�С
    BrowserNormalSize
   ' web1.Visible = firmlistlse
Form_Resize


   Web1.Height = Web1.document.body.scrollHeight * 15 + 256
End Sub

Private Sub vScroll1_GotFocus()
    HideCaret (VScroll1.hWnd)
End Sub

Private Sub vScroll1_LostFocus()
    HideCaret (VScroll1.hWnd)
End Sub

Private Sub HScroll1_GotFocus()
    HideCaret (HScroll1.hWnd)
End Sub

Private Sub HScroll1_LostFocus()
    HideCaret (HScroll1.hWnd)
End Sub

Private Sub HScroll1_Scroll()
Picture1.Left = -HScroll1.Value
 HideCaret (HScroll1.hWnd)
End Sub

Private Sub VScroll1_Scroll()
Picture1.Top = -VScroll1.Value
 HideCaret (VScroll1.hWnd)
End Sub
Private Sub VScroll1_Change()
Picture1.Top = -VScroll1.Value
 HideCaret (VScroll1.hWnd)
End Sub
Private Sub HScroll1_Change()
Picture1.Left = -HScroll1.Value
 HideCaret (HScroll1.hWnd)
End Sub

