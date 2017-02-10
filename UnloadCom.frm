VERSION 5.00
Begin VB.Form UnloadCom 
   BorderStyle     =   0  'None
   Caption         =   "防止误关"
   ClientHeight    =   1335
   ClientLeft      =   14835
   ClientTop       =   1320
   ClientWidth     =   4260
   Icon            =   "UnloadCom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   120
      TabIndex        =   1
      Top             =   950
      Width           =   200
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "点错了"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      MouseIcon       =   "UnloadCom.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "是的"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      MouseIcon       =   "UnloadCom.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   330
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "我很谨慎，以后不必提醒啦。"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Image CancelButton 
      Height          =   495
      Left            =   2880
      MouseIcon       =   "UnloadCom.frx":02B0
      MousePointer    =   99  'Custom
      Picture         =   "UnloadCom.frx":0402
      Stretch         =   -1  'True
      Top             =   800
      Width           =   1215
   End
   Begin VB.Image OKButton 
      Height          =   615
      Left            =   2880
      MouseIcon       =   "UnloadCom.frx":479E
      MousePointer    =   99  'Custom
      Picture         =   "UnloadCom.frx":48F0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "您确定要关闭细胞浏览器了吗？"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   0
      Picture         =   "UnloadCom.frx":8C8C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4260
   End
End
Attribute VB_Name = "UnloadCom"
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
Me.Left = Fa.Width - Me.Width
End Sub



Private Sub Label3_Click()
OKButton_Click
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub OKButton_Click()


If Check1.Value = 1 Then
'不提醒
Open App.path & "\set.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, po2$
Loop
Close #1
li = Split(po2, ",")
Dim SaveLi As String
Dim i As Integer
li(1) = "false"
For i = 0 To UBound(li)
SaveLi = SaveLi & li(i) & ","
Next i

Open App.path & "\set.dat" For Output As #1
Print #1, SaveLi
Close #1
End If

End
End Sub
