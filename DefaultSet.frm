VERSION 5.00
Begin VB.Form DefaultSet 
   BorderStyle     =   0  'None
   Caption         =   "细胞提示"
   ClientHeight    =   900
   ClientLeft      =   -4005
   ClientTop       =   -180
   ClientWidth     =   7260
   Icon            =   "DefaultSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   480
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   120
      TabIndex        =   2
      Top             =   500
      Width           =   200
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5170
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "不好"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      MouseIcon       =   "DefaultSet.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   555
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "好的"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      MouseIcon       =   "DefaultSet.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   555
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   900
      Left            =   0
      Top             =   0
      Width           =   7260
   End
   Begin VB.Image CancelButton 
      Height          =   375
      Left            =   5760
      MouseIcon       =   "DefaultSet.frx":02B0
      MousePointer    =   99  'Custom
      Picture         =   "DefaultSet.frx":0402
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image OKButton 
      Height          =   375
      Left            =   4320
      MouseIcon       =   "DefaultSet.frx":479E
      MousePointer    =   99  'Custom
      Picture         =   "DefaultSet.frx":48F0
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "我讨厌此提示"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   525
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "您是否喜欢细胞浏览器呢？ 如果你很喜欢用细胞浏览器，那就    将细胞设为为默认吧。"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   1000
      Left            =   -120
      Picture         =   "DefaultSet.frx":8C8C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "DefaultSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
Fa.WindowState = 0
Me.Top = 0
Me.Left = -Me.Width
Timer1.Enabled = True
'Smell.Picture = LoadPicture(App.path & "\image\use\smell1.jpg")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Smell.Picture = LoadPicture(App.path & "\image\use\smell1.jpg")
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
Timer2.Enabled = True
End Sub



Private Sub Label3_Click()
OKButton_Click
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub OKButton_Click()
If Check1.Value = 1 Then
sh = Shell(App.path & "\DefaultBrowser.exe")
End If

If Check2.Value = 1 Then '不提醒
Open App.path & "\set.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, po2$
Loop
Close #1
li = Split(po2, ",")
Dim SaveLi As String
Dim i As Integer
li(0) = "false"
For i = 0 To UBound(li)
SaveLi = SaveLi & li(i) & ","
Next i

Open App.path & "\set.dat" For Output As #1
Print #1, SaveLi
Close #1
End If

Unload Me
End Sub

Private Sub Smell_Click()
OKButton_Click

End Sub

Private Sub Smell_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Smell.Picture = LoadPicture(App.path & "\image\use\smell2.jpg")

End Sub

Private Sub Timer1_Timer()

Me.Left = Me.Left + 256

If Me.Left > 0 Then
Me.Left = 0
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Me.Left = Me.Left - 256
If Me.Left < -Me.Width Then Unload Me
End Sub
