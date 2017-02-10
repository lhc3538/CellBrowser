VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FormSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ËÑË÷"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FormSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin SHDocVwCtl.WebBrowser WebPage 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      ExtentX         =   5530
      ExtentY         =   3201
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
      Location        =   "http://www.soso.com/?unc=y400372_2"
   End
End
Attribute VB_Name = "FormSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Dim i As Integer
Dim ofrm As Form
  For Each ofrm In Forms
        If ofrm.Name = "FormSearch" And ofrm.Caption <> Me.Caption Then
         ofrm.Top = i * 450
         ofrm.Height = 450
         i = i + 1
        End If
        
  Next
Me.Top = (SearchNum - 1) * 450
Me.Height = Fa.Height - Fa.P1.Height - Fa.P2.Height - Me.Top
End Sub

Private Sub Form_Load()

Me.Left = 0
Me.Height = Fa.Height - Fa.P1.Height - Fa.P2.Height - Me.Top
Me.Width = Fa.Width * (6 / 19)

Dim ofrm As Form
  For Each ofrm In Forms
          If ofrm.Name = "FrmWeb" Then
                ofrm.Left = Me.Width
                ofrm.Width = Fa.Width - Me.Width
        End If
        
  Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SearchNum = SearchNum - 1
If SearchNum = 0 Then
Dim ofrm As Form
  For Each ofrm In Forms
          If ofrm.Name = "FrmWeb" Then
                ofrm.Left = 0
                ofrm.Width = Fa.Width
        End If
        
  Next
End If
LinkPage.SetFocus
End Sub

Private Sub Form_Resize()
WebPage.Width = Me.Width
WebPage.Height = Me.Height
End Sub



Private Sub WebPage_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
Fa.OpenNewPage.Text = "s," & WebPage.Document.activeelement.href
Me.SetFocus
End Sub

