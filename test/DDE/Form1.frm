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
   StartUpPosition =   3  '����ȱʡ
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
      Text            =   "�ɹ�"
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
  Const COMMANDLINE = "CommandLine="                         '      ����Ϊ��ʡ�£�����һ������
    
   Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
   Static lngCount       As Long
   Dim Info       As String
    
   Info = Opennewpage.Text                  '              ����ԭ����Ϣ
    
   Select Case CmdStr                          '          CmdStr    ��DDE�����͹����Ĳ���
       Case "URLDDE"
           Me.WindowState = 2
           Info = Text2.Text
   End Select
       
   Opennewpage.Text = Info                  '          ����Ϣ��ʾ����
    
   Cancel = False
   End Sub
    
    
   Private Sub LinkAndSendMessage(ByVal Msg As String)
   Dim t       As Long
   PicDDE.LinkMode = 0                                      '--
   PicDDE.LinkTopic = "cellbrowser|FormDDE"              '      |______����DDE���򲢷�������/����
   PicDDE.LinkMode = 2                                      '      |              ��|��Ϊ�ܵ������ǡ��˸�����Աߵ����ߣ�
   PicDDE.LinkExecute Msg                             '--                  ������ĸ�����֣�
    
   t = PicDDE.LinkTimeout                  '--
   PicDDE.LinkTimeout = 1                  '      |______��ֹDDEͨ������Ȼ��Ҳ�����ñ�ķ���
   PicDDE.LinkMode = 0                        '      |              �����õ��ǳ�ʱǿ����ֹ�ķ���
   PicDDE.LinkTimeout = t                  '--
   End Sub
    
    
   Private Sub Form_Load()
   If App.PrevInstance Then                  '    �����Ƿ��Ѿ�����
    
       Me.LinkTopic = ""                            '    ������������������еĳ����DDE���������ԣ�
       Me.LinkMode = 0                                '    ����������DDE����ʱ������ӵ�
    
       LinkAndSendMessage "URLDDE"                         '--
                    '      |-----����DDE���ܳ��򲢴�������/����
              '--
    
       If Command <> "" Then                                        '    ����������в������ʹ��ݹ�ȥ
             LinkAndSendMessage COMMANDLINE + Command
       End If
       End                                                                '      �����³��������
   End If
   End Sub


