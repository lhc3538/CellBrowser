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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents m_MW As CHookMouseWheel
Attribute m_MW.VB_VarHelpID = -1
Private Sub Form_Load()
   Set m_MW = New CHookMouseWheel
   m_MW.hWnd = Me.hWnd
End Sub
Private Sub m_MW_MouseWheel(ByVal hWnd As Long, ByVal Delta As Long, ByVal Shift As Long, ByVal Button As Long, ByVal X As Long, ByVal Y As Long, Cancel As Boolean)

   If TypeOf Screen.ActiveControl Is MSFlexGrid Then sub_MouseWheel Delta, X, Y
End Sub

