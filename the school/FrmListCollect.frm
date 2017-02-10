VERSION 5.00
Begin VB.Form FrmListCollect 
   Caption         =   "添加导航条"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   Icon            =   "FrmListCollect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5880
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ComAdd 
      Caption         =   "添加"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "名称："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "地址："
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmListCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComAdd_Click()
Dim s(7) As String
Dim n As Integer
n = 1
Open App.Path & "\collect2.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, s(n)
n = n + 1
Loop
Close #1
For i = 1 To 7
If s(i) = "" Then
s(i) = Text2.Text & "," & Text1.Text
Exit For
End If
Next i
On Error GoTo x1
Kill (App.Path & "\collect2.dat")
x1:
For z = 1 To 7
Open App.Path & "\collect2.dat" For Append As #1
Print #1, s(z)
Close #1
Next z
Msg = MsgBox("添加成功", vbOKOnly, "提示")
Unload Me

End Sub

Private Sub Form_Load()
Text1.Text = FrmWeb(ActivePage).TWUrl.Text
Text2.Text = FrmWeb(ActivePage).TWCaption
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
Me.Height = 2670
Me.Width = 6000
End If
End Sub
