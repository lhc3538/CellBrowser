VERSION 5.00
Begin VB.Form FrmList 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleMode       =   0  'User
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   52
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   156
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   52
         Left            =   240
         TabIndex        =   157
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   52
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   52
         Left            =   60
         TabIndex        =   158
         Top             =   0
         Width           =   2145
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   52
         Left            =   2280
         MouseIcon       =   "FrmList.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":0152
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   51
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   153
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   51
         Left            =   240
         TabIndex        =   154
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   51
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   51
         Left            =   60
         TabIndex        =   155
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   51
         Left            =   2280
         MouseIcon       =   "FrmList.frx":0A9A
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":0BEC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   50
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   150
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   50
         Left            =   240
         TabIndex        =   151
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   50
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   50
         Left            =   60
         TabIndex        =   152
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   50
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1534
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1686
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   49
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   147
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   49
         Left            =   240
         TabIndex        =   148
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   49
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   49
         Left            =   60
         TabIndex        =   149
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   49
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1FCE
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":2120
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   48
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   144
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   48
         Left            =   240
         TabIndex        =   145
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   48
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   48
         Left            =   60
         TabIndex        =   146
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   48
         Left            =   2280
         MouseIcon       =   "FrmList.frx":2A68
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":2BBA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   47
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   141
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   47
         Left            =   240
         TabIndex        =   142
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   47
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   47
         Left            =   60
         TabIndex        =   143
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   47
         Left            =   2280
         MouseIcon       =   "FrmList.frx":3502
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":3654
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   46
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   138
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   46
         Left            =   240
         TabIndex        =   139
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   46
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   46
         Left            =   60
         TabIndex        =   140
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   46
         Left            =   2280
         MouseIcon       =   "FrmList.frx":3F9C
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":40EE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   45
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   135
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   45
         Left            =   240
         TabIndex        =   136
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   45
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   45
         Left            =   60
         TabIndex        =   137
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   45
         Left            =   2280
         MouseIcon       =   "FrmList.frx":4A36
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":4B88
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   44
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   132
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   44
         Left            =   240
         TabIndex        =   133
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   44
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   44
         Left            =   60
         TabIndex        =   134
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   44
         Left            =   2280
         MouseIcon       =   "FrmList.frx":54D0
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":5622
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   43
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   129
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   43
         Left            =   240
         TabIndex        =   130
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   43
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   43
         Left            =   60
         TabIndex        =   131
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   43
         Left            =   2280
         MouseIcon       =   "FrmList.frx":5F6A
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":60BC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   42
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   126
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   42
         Left            =   240
         TabIndex        =   127
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   42
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   42
         Left            =   60
         TabIndex        =   128
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   42
         Left            =   2280
         MouseIcon       =   "FrmList.frx":6A04
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":6B56
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   41
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   123
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   41
         Left            =   240
         TabIndex        =   124
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   41
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   41
         Left            =   60
         TabIndex        =   125
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   41
         Left            =   2280
         MouseIcon       =   "FrmList.frx":749E
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":75F0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   40
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   120
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   40
         Left            =   240
         TabIndex        =   121
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   40
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   40
         Left            =   60
         TabIndex        =   122
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   40
         Left            =   2280
         MouseIcon       =   "FrmList.frx":7F38
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":808A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   39
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   117
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   39
         Left            =   240
         TabIndex        =   118
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   39
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   39
         Left            =   60
         TabIndex        =   119
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   39
         Left            =   2280
         MouseIcon       =   "FrmList.frx":89D2
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":8B24
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   38
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   114
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   38
         Left            =   240
         TabIndex        =   115
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   38
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   38
         Left            =   60
         TabIndex        =   116
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   38
         Left            =   2280
         MouseIcon       =   "FrmList.frx":946C
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":95BE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   37
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   111
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   37
         Left            =   240
         TabIndex        =   112
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   37
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   37
         Left            =   60
         TabIndex        =   113
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   37
         Left            =   2280
         MouseIcon       =   "FrmList.frx":9F06
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":A058
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   36
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   108
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   36
         Left            =   240
         TabIndex        =   109
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   36
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   36
         Left            =   60
         TabIndex        =   110
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   36
         Left            =   2280
         MouseIcon       =   "FrmList.frx":A9A0
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":AAF2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   35
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   105
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   35
         Left            =   240
         TabIndex        =   106
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   35
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   35
         Left            =   60
         TabIndex        =   107
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   35
         Left            =   2280
         MouseIcon       =   "FrmList.frx":B43A
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":B58C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   34
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   102
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   34
         Left            =   240
         TabIndex        =   103
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   34
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   34
         Left            =   60
         TabIndex        =   104
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   34
         Left            =   2280
         MouseIcon       =   "FrmList.frx":BED4
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":C026
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   33
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   99
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   33
         Left            =   240
         TabIndex        =   100
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   33
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   33
         Left            =   60
         TabIndex        =   101
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   33
         Left            =   2280
         MouseIcon       =   "FrmList.frx":C96E
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":CAC0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   32
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   96
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   32
         Left            =   240
         TabIndex        =   97
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   32
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   32
         Left            =   60
         TabIndex        =   98
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   32
         Left            =   2280
         MouseIcon       =   "FrmList.frx":D408
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":D55A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   31
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   93
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   31
         Left            =   240
         TabIndex        =   94
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   31
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   31
         Left            =   60
         TabIndex        =   95
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   31
         Left            =   2280
         MouseIcon       =   "FrmList.frx":DEA2
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":DFF4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   30
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   90
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   30
         Left            =   240
         TabIndex        =   91
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   30
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   30
         Left            =   60
         TabIndex        =   92
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   30
         Left            =   2280
         MouseIcon       =   "FrmList.frx":E93C
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":EA8E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   29
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   87
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   29
         Left            =   240
         TabIndex        =   88
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   29
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   29
         Left            =   60
         TabIndex        =   89
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   29
         Left            =   2280
         MouseIcon       =   "FrmList.frx":F3D6
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":F528
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   28
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   84
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   28
         Left            =   240
         TabIndex        =   85
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   28
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   28
         Left            =   60
         TabIndex        =   86
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   28
         Left            =   2280
         MouseIcon       =   "FrmList.frx":FE70
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":FFC2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   27
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   81
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   27
         Left            =   240
         TabIndex        =   82
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   27
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   27
         Left            =   60
         TabIndex        =   83
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   27
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1090A
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":10A5C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   26
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   78
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   26
         Left            =   240
         TabIndex        =   79
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   26
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   26
         Left            =   60
         TabIndex        =   80
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   26
         Left            =   2280
         MouseIcon       =   "FrmList.frx":113A4
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":114F6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   25
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   75
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   25
         Left            =   240
         TabIndex        =   76
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   25
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   25
         Left            =   60
         TabIndex        =   77
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   25
         Left            =   2280
         MouseIcon       =   "FrmList.frx":11E3E
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":11F90
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   24
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   72
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   24
         Left            =   240
         TabIndex        =   73
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   24
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   24
         Left            =   60
         TabIndex        =   74
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   24
         Left            =   2280
         MouseIcon       =   "FrmList.frx":128D8
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":12A2A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   23
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   69
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   23
         Left            =   240
         TabIndex        =   70
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   23
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   23
         Left            =   60
         TabIndex        =   71
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   23
         Left            =   2280
         MouseIcon       =   "FrmList.frx":13372
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":134C4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   22
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   66
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   22
         Left            =   240
         TabIndex        =   67
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   22
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   22
         Left            =   60
         TabIndex        =   68
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   22
         Left            =   2280
         MouseIcon       =   "FrmList.frx":13E0C
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":13F5E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   21
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   63
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   21
         Left            =   240
         TabIndex        =   64
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   21
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   21
         Left            =   60
         TabIndex        =   65
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   21
         Left            =   2280
         MouseIcon       =   "FrmList.frx":148A6
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":149F8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   20
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   60
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   20
         Left            =   240
         TabIndex        =   61
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   20
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   20
         Left            =   60
         TabIndex        =   62
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   20
         Left            =   2280
         MouseIcon       =   "FrmList.frx":15340
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":15492
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   19
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   57
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   19
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   19
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   19
         Left            =   60
         TabIndex        =   59
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   19
         Left            =   2280
         MouseIcon       =   "FrmList.frx":15DDA
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":15F2C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   18
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   54
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   18
         Left            =   240
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   18
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   18
         Left            =   60
         TabIndex        =   56
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   18
         Left            =   2280
         MouseIcon       =   "FrmList.frx":16874
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":169C6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   17
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   51
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   17
         Left            =   240
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   17
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   17
         Left            =   60
         TabIndex        =   53
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   17
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1730E
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":17460
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   16
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   48
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   16
         Left            =   240
         TabIndex        =   49
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   16
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   16
         Left            =   60
         TabIndex        =   50
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   16
         Left            =   2280
         MouseIcon       =   "FrmList.frx":17DA8
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":17EFA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   15
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   45
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   15
         Left            =   240
         TabIndex        =   46
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   15
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   15
         Left            =   60
         TabIndex        =   47
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   15
         Left            =   2280
         MouseIcon       =   "FrmList.frx":18842
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":18994
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   14
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   42
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   14
         Left            =   240
         TabIndex        =   43
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   14
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   14
         Left            =   60
         TabIndex        =   44
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   14
         Left            =   2280
         MouseIcon       =   "FrmList.frx":192DC
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1942E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   13
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   39
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   13
         Left            =   240
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   13
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   13
         Left            =   60
         TabIndex        =   41
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   13
         Left            =   2280
         MouseIcon       =   "FrmList.frx":19D76
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":19EC8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   12
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   36
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   12
         Left            =   240
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   12
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   12
         Left            =   60
         TabIndex        =   38
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   12
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1A810
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1A962
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   11
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   33
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   11
         Left            =   240
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   11
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   11
         Left            =   60
         TabIndex        =   35
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   11
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1B2AA
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1B3FC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   10
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   30
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   10
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   10
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   10
         Left            =   60
         TabIndex        =   32
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   10
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1BD44
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1BE96
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   9
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   9
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   9
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   9
         Left            =   60
         TabIndex        =   29
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   9
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1C7DE
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1C930
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   8
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   8
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   8
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   8
         Left            =   60
         TabIndex        =   26
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   8
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1D278
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1D3CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   7
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   7
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   7
         Left            =   60
         TabIndex        =   23
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   7
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1DD12
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1DE64
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   6
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   6
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   6
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   6
         Left            =   60
         TabIndex        =   20
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   6
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1E7AC
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1E8FE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   5
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   5
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         Left            =   60
         TabIndex        =   17
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   5
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1F246
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1F398
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   4
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   4
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   4
         Left            =   2280
         MouseIcon       =   "FrmList.frx":1FCE0
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":1FE32
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   3
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   3
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   60
         TabIndex        =   11
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   3
         Left            =   2280
         MouseIcon       =   "FrmList.frx":2077A
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":208CC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   2
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   2
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   2
         Left            =   2280
         MouseIcon       =   "FrmList.frx":21214
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":21366
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   1
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   1
         Left            =   2280
         MouseIcon       =   "FrmList.frx":21CAE
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":21E00
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PageBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   0
      Left            =   120
      ScaleHeight     =   1935
      ScaleMode       =   0  'User
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox PageUrl 
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image ComEndWeb 
         Height          =   495
         Index           =   0
         Left            =   2280
         MouseIcon       =   "FrmList.frx":22748
         MousePointer    =   99  'Custom
         Picture         =   "FrmList.frx":2289A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
      Begin VB.Label WebName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WebName"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   2265
      End
      Begin VB.Image Page 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2745
      End
   End
   Begin VB.Timer TimerGoRight 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   0
      Top             =   4920
   End
   Begin VB.Timer TimerCorrect 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   5400
   End
End
Attribute VB_Name = "FrmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Êó±êÎ»ÖÃ
Dim MX As Single
Dim MY As Single
Dim DyX, DyY As Single
Dim Tvalue As Boolean
Dim FormMove As Boolean
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
X As Long
Y As Long
End Type
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Í¨ÓÃ±äÁ¿¶¨ÒåÇø:
Dim Xleft As Integer '×î×ó±ßpagebackµÄÎ»ÖÃ
Dim MouseStep As Boolean 'Êó±êÊÇ·ñ°´ÏÂ
Dim Xdistance As Integer 'Êó±ê°´ÏÂÎ»ÖÃµ½xleftµÄ¾àÀë
Dim FirstPageBack As Integer '¾ÀÕýpagebackÎ»ÖÃ£¨µÚÒ»¸ö£©
Dim NearestPageback As Integer '¾ÀÕýpggebackÎ»ÖÃ£¨×î½Ó½ü0£©
Public PublicIndex As Integer 'Í¨ÓÃÐòºÅ

Private Sub ComEndWeb_Click(Index As Integer)
  Unload FrmWeb(Index)
  PageBack(Index + 1).Top = PageBack(Index + 1).Top - PageBack(Index + 1).Height
ListPageBack
End Sub

Private Sub Form_Click()

Dim Lastone As Integer
Dim i As Integer

i = 0
 For i = 0 To 52
  If PageBack(i).Visible = True Then
    Lastone = i
  End If
 Next i
If PageBack(Lastone).Top < 0 Then
  PageBack(Lastone).Top = 0
End If
End Sub

Private Sub Form_Load()
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' ½«´°¿ÚÉèÎªÔÚËùÓÐ´°¿ÚÇ°¶Ë
Me.Top = 0
Me.Left = 0
Me.Height = WebPage.Height
Me.Width = 1

TimerCorrect.Enabled = True 'ÅÅÁÐÁÐ±í
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Fcbox.ConBack.Text = "" '»Ö¸´°´Å¥±³¾°É«
End Sub

Private Sub Page_Click(Index As Integer)
'If Index = 0 Then
'FrmWeb(0).Show
'FrmWeb(0).Web1.navigate ("http://www.1616.net")
'Else
FrmWeb(Index).ZOrder
FrmWeb(Index).SetFocus
'End If
End Sub

Private Sub TimerGoRight_Timer()
If Fcbox.ComFrmList.ToolTipText = "Right" Then
 If Me.Width = 1 Then Me.Width = 60
  Me.Width = Me.Width * 2
  Fcbox.Left = Me.Width
  'FrmWeb(ActivePage).Left = Me.Width + Fcbox.Width - 60
  If Me.Width >= 3000 Then
   Me.Width = 6045
   FrmList.TimerGoRight.Enabled = False
   Fcbox.ComFrmList.ToolTipText = "Left"
  End If
End If

If Fcbox.ComFrmList.ToolTipText = "Left" Then
  Me.Width = Me.Width / 2
  If Me.Width < 400 Then
   Me.Width = 1
   FrmList.TimerGoRight.Enabled = False
   Fcbox.ComFrmList.ToolTipText = "Right"
  End If
  Fcbox.Left = Me.Width
  'FrmWeb(ActivePage).Left = Me.Width + Fcbox.Width - 60
End If
End Sub





Private Sub TimerCorrect_Timer()
PageBack(FirstPageBack).Top = PageBack(FirstPageBack).Top - 560
CorrectPageBack


 If PageBack(NearestPageback).Top < 0 Then
  Dim DistanceNtoF As Integer
  DistanceNtoF = PageBack(NearestPageback).Top - PageBack(FirstPageBack).Top
  PageBack(NearestPageback).Top = 0
  PageBack(FirstPageBack).Top = 0 - DistanceNtoF
  CorrectPageBack
  Xleft = PageBack(FirstPageBack).Top
  TimerCorrect.Enabled = False
 End If
 

End Sub
Private Sub CorrectPageBack() ' ¾ÀÕýÍøÒ³ÁÐ±í
Dim i As Integer
Dim n As Integer
 For i = 0 To 52
  If PageBack(i).Visible = True Then
  
   PageBack(i).Top = PageBack(FirstPageBack).Top + PageBack(0).Height * n
   n = n + 1
  End If
 Next i
 
End Sub

Private Sub Page_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
MouseStep = True
Dim Point As POINTAPI
GetCursorPos Point
Dim i As Integer
 For i = 0 To 52
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
Xdistance = Point.Y * 15 - PageBack(FirstPageBack).Top
End Sub

Private Sub Page_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseStep = True Then
Dim Point As POINTAPI
GetCursorPos Point
Xleft = Point.Y * 15 - Xdistance
ListPageBack
End If
Fcbox.ConBack.Text = "" '»Ö¸´°´Å¥±³¾°É«
End Sub

Private Sub Page_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = False

Dim i As Integer
Dim n As Integer
 For i = 0 To 52
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
 
 For n = 0 To 52
  If PageBack(n).Visible = True And PageBack(n).Top < PageBack(0).Height And PageBack(n).Top >= 0 Then
   NearestPageback = n
   Exit For
  Else
   NearestPageback = FirstPageBack
  End If
 Next n

TimerCorrect.Enabled = True
End Sub






Private Sub PageBack_Click(Index As Integer)
'PageBack(PublicIndex).Picture = LoadPicture(App.Path & "\skin\pageback0.gif")
'PublicIndex = Index
'PageBackClick
End Sub

Private Sub PageBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = True
Dim Point As POINTAPI
GetCursorPos Point
Dim i As Integer
 For i = 0 To 52
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
Xdistance = Point.Y * 15 - PageBack(FirstPageBack).Top
End Sub

Private Sub PageBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseStep = True Then
Dim Point As POINTAPI
GetCursorPos Point
Xleft = Point.Y * 15 - Xdistance
ListPageBack
End If
Fcbox.ConBack.Text = "" '»Ö¸´°´Å¥±³¾°É«
End Sub
Private Sub ListPageBack() 'ÅÅÁÐÍøÒ³ÁÐ±í
On Error Resume Next
Dim Xnum As Integer
Dim i As Integer
Xnum = Xleft

For i = 0 To 52
 If PageBack(i).Visible = True Then
  PageBack(i).Top = Xnum
  Xnum = Xnum + PageBack(i).Height '»áÒç³ö
 End If
Next i
End Sub
Private Sub PageBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = False

Dim i As Integer
Dim n As Integer
 For i = 0 To 52
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
 
 For n = 0 To 5
  If PageBack(n).Visible = True And PageBack(n).Top < PageBack(0).Height And PageBack(n).Top >= 0 Then
   NearestPageback = n
   Exit For
  Else
   NearestPageback = FirstPageBack
  End If
 Next n

TimerCorrect.Enabled = True
End Sub
Public Sub AllPageback() 'È«ÁÐ±í²Ù×÷
Dim i As Integer
For i = 0 To PageBack.UBound
 If i = ActivePage Then
   PageBack(i).BackColor = Fcbox.BackColor
 Else
   If PageBack(i).BackColor <> &H808080 Then
     PageBack(i).BackColor = &H808080
   End If
 End If
Next i
End Sub
