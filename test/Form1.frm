VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   1440
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim f(1 To 5) As New Form2
For i = 1 To 5
f(i).Show
Next i
i = 1
Dim p(1 To 5) As Picture1
For i = 1 To 5
Load p(i)
Me.Controls.Add (p(1))
Next i

End Sub

