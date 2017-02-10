VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkMode        =   1  'Source
   LinkTopic       =   "FormDDE"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicDDE 
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Text            =   "成功"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Opennewpage 
      Height          =   615
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Const COMMANDLINE = "CommandLine="                         '      还是为了省事，定义一个常量
    
   Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
   Static lngCount       As Long
   Dim Info       As String
    
   Info = Opennewpage.Text                  '              保留原有信息
    
   Select Case CmdStr                          '          CmdStr    是DDE程序传送过来的参数
       Case "URLDDE"
           Me.WindowState = 2
           Info = Text2.Text
   End Select
       
   Opennewpage.Text = Info                  '          把信息显示出来
    
   Cancel = False
   End Sub
    
    
   Private Sub LinkAndSendMessage(ByVal Msg As String)
   Dim t       As Long
   PicDDE.LinkMode = 0                                      '--
   PicDDE.LinkTopic = "cellbrowser|FormDDE"              '      |______连接DDE程序并发送数据/参数
   PicDDE.LinkMode = 2                                      '      |              “|”为管道符，是“退格键”旁边的竖线，
   PicDDE.LinkExecute Msg                             '--                  不是字母或数字！
    
   t = PicDDE.LinkTimeout                  '--
   PicDDE.LinkTimeout = 1                  '      |______终止DDE通道。当然，也可以用别的方法
   PicDDE.LinkMode = 0                        '      |              这里用的是超时强制终止的方法
   PicDDE.LinkTimeout = t                  '--
   End Sub
    
    
   Private Sub Form_Load()
   If App.PrevInstance Then                  '    程序是否已经运行
    
       Me.LinkTopic = ""                            '    这两行用于清除新运行的程序的DDE服务器属性，
       Me.LinkMode = 0                                '    否则在连接DDE程序时会出乱子的
    
       LinkAndSendMessage "URLDDE"                         '--
                    '      |-----连接DDE接受程序并传送数据/参数
              '--
    
       If Command <> "" Then                                        '    如果有命令行参数，就传递过去
             LinkAndSendMessage COMMANDLINE + Command
       End If
       End                                                                '      结束新程序的运行
   End If
   End Sub


