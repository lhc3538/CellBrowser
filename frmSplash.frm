VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "ϸ�������"
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
   StartUpPosition =   2  '��Ļ����
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
Const COMMANDLINE = "CommandLine="                         '      ����Ϊ��ʡ�£�����һ������
   Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer) 'DDE
   On Error Resume Next
   Static lngCount       As Long

          lngCount = lngCount + 1

    If lngCount = 0 Then
        FrmWeb.WebPage.Navigate CmdStr
    Else
     Fa.OpenNewPage.Text = CmdStr                  '          ����Ϣ��ʾ����
    End If
   Cancel = False
   End Sub
      Private Sub LinkAndSendMessage(ByVal Msg As String) 'DDE
      On Error Resume Next
   Dim t       As Long
   PicDDE.LinkMode = 0                                      '--
   PicDDE.LinkTopic = "CellBrowser|FormDDE"              '      |______����DDE���򲢷�������/����
   PicDDE.LinkMode = 2                                      '      |              ��|��Ϊ�ܵ������ǡ��˸�����Աߵ����ߣ�
   PicDDE.LinkExecute Command()                             '--                  ������ĸ�����֣�
    
   t = PicDDE.LinkTimeout                  '--
   PicDDE.LinkTimeout = 1                  '      |______��ֹDDEͨ������Ȼ��Ҳ�����ñ�ķ���
   PicDDE.LinkMode = 0                        '      |              �����õ��ǳ�ʱǿ����ֹ�ķ���
   PicDDE.LinkTimeout = t                  '--
   End Sub
Private Sub Form_Load()


 If App.PrevInstance Then                  '    �����Ƿ��Ѿ�����
    
       Me.LinkTopic = ""                            '    ������������������еĳ����DDE���������ԣ�
       Me.LinkMode = 0                                '    ����������DDE����ʱ������ӵ�
    
      LinkAndSendMessage "Max"                         '--�����У�
       End
       '      �����³��������
   Else
     Load Fa
   End If

   
End Sub

