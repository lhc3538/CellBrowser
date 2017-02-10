VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "细胞浏览器"
   ClientHeight    =   975
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   2295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "FormDDE"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.PictureBox PicDDE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const COMMANDLINE = "CommandLine="                         '      还是为了省事，定义一个常量
   Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer) 'DDE
   On Error Resume Next
   Static lngCount       As Long

          lngCount = lngCount + 1

    If lngCount = 0 Then
        FrmWeb.WebPage.Navigate CmdStr
    Else
     Fa.OpenNewPage.Text = CmdStr                  '          把信息显示出来
    End If
   Cancel = False
   End Sub
      Private Sub LinkAndSendMessage(ByVal Msg As String) 'DDE
      On Error Resume Next
   Dim t       As Long
   PicDDE.LinkMode = 0                                      '--
   PicDDE.LinkTopic = "CellBrowser|FormDDE"              '      |______连接DDE程序并发送数据/参数
   PicDDE.LinkMode = 2                                      '      |              “|”为管道符，是“退格键”旁边的竖线，
   PicDDE.LinkExecute Command()                             '--                  不是字母或数字！
    
   t = PicDDE.LinkTimeout                  '--
   PicDDE.LinkTimeout = 1                  '      |______终止DDE通道。当然，也可以用别的方法
   PicDDE.LinkMode = 0                        '      |              这里用的是超时强制终止的方法
   PicDDE.LinkTimeout = t                  '--
   End Sub
Private Sub Form_Load()


 If App.PrevInstance Then                  '    程序是否已经运行
    
       Me.LinkTopic = ""                            '    这两行用于清除新运行的程序的DDE服务器属性，
       Me.LinkMode = 0                                '    否则在连接DDE程序时会出乱子的
    
      LinkAndSendMessage "Max"                         '--必须有！
       End
       '      结束新程序的运行
   Else
     Load Fa
   End If

   
End Sub

