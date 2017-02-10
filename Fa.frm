VERSION 5.00
Begin VB.MDIForm Fa 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "细胞浏览器"
   ClientHeight    =   6630
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   16245
   Icon            =   "Fa.frx":0000
   LinkTopic       =   "Fa"
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.PictureBox P1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   0
      ScaleHeight     =   2760
      ScaleWidth      =   16245
      TabIndex        =   2
      Top             =   0
      Width           =   16245
      Begin VB.ComboBox ComboRestore 
         Height          =   300
         Left            =   7320
         TabIndex        =   209
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox CloseHistory 
         Height          =   300
         Left            =   6600
         TabIndex        =   208
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Timer TimerCorrect 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   8520
         Top             =   1680
      End
      Begin VB.TextBox TextSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   5400
         TabIndex        =   206
         Top             =   200
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   66
         Left            =   2760
         Picture         =   "Fa.frx":000C
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   201
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   66
            Left            =   240
            TabIndex        =   202
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   66
            Left            =   2280
            Picture         =   "Fa.frx":02F5
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   66
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   66
            Left            =   60
            TabIndex        =   203
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   65
         Left            =   2760
         Picture         =   "Fa.frx":0C3D
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   198
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   65
            Left            =   240
            TabIndex        =   199
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   65
            Left            =   2280
            Picture         =   "Fa.frx":0F26
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   65
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   65
            Left            =   60
            TabIndex        =   200
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   64
         Left            =   2760
         Picture         =   "Fa.frx":186E
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   195
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   64
            Left            =   240
            TabIndex        =   196
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   64
            Left            =   2280
            Picture         =   "Fa.frx":1B57
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   64
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   64
            Left            =   60
            TabIndex        =   197
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   63
         Left            =   2760
         Picture         =   "Fa.frx":249F
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   192
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   63
            Left            =   240
            TabIndex        =   193
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   63
            Left            =   2280
            Picture         =   "Fa.frx":2788
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   63
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   63
            Left            =   60
            TabIndex        =   194
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   62
         Left            =   2760
         Picture         =   "Fa.frx":30D0
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   189
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   62
            Left            =   240
            TabIndex        =   190
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   62
            Left            =   2280
            Picture         =   "Fa.frx":33B9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   62
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   62
            Left            =   60
            TabIndex        =   191
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   61
         Left            =   2760
         Picture         =   "Fa.frx":3D01
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   186
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   61
            Left            =   240
            TabIndex        =   187
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   61
            Left            =   2280
            Picture         =   "Fa.frx":3FEA
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   61
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   61
            Left            =   60
            TabIndex        =   188
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   60
         Left            =   2760
         Picture         =   "Fa.frx":4932
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   183
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   60
            Left            =   240
            TabIndex        =   184
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   60
            Left            =   2280
            Picture         =   "Fa.frx":4C1B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   60
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   60
            Left            =   60
            TabIndex        =   185
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   59
         Left            =   2760
         Picture         =   "Fa.frx":5563
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   180
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   59
            Left            =   240
            TabIndex        =   181
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   59
            Left            =   2280
            Picture         =   "Fa.frx":584C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   59
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   59
            Left            =   60
            TabIndex        =   182
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   58
         Left            =   2760
         Picture         =   "Fa.frx":6194
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   177
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   58
            Left            =   240
            TabIndex        =   178
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   58
            Left            =   2280
            Picture         =   "Fa.frx":647D
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   58
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   58
            Left            =   60
            TabIndex        =   179
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   57
         Left            =   2760
         Picture         =   "Fa.frx":6DC5
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   174
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   57
            Left            =   240
            TabIndex        =   175
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   57
            Left            =   2280
            Picture         =   "Fa.frx":70AE
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   57
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   57
            Left            =   60
            TabIndex        =   176
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   56
         Left            =   2760
         Picture         =   "Fa.frx":79F6
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   171
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   56
            Left            =   240
            TabIndex        =   172
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   56
            Left            =   2280
            Picture         =   "Fa.frx":7CDF
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   56
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   56
            Left            =   60
            TabIndex        =   173
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   55
         Left            =   2760
         Picture         =   "Fa.frx":8627
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   168
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   55
            Left            =   240
            TabIndex        =   169
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   55
            Left            =   2280
            Picture         =   "Fa.frx":8910
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   55
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   55
            Left            =   60
            TabIndex        =   170
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   54
         Left            =   2760
         Picture         =   "Fa.frx":9258
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   165
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   54
            Left            =   240
            TabIndex        =   166
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   54
            Left            =   2280
            Picture         =   "Fa.frx":9541
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   54
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   54
            Left            =   60
            TabIndex        =   167
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   53
         Left            =   2760
         Picture         =   "Fa.frx":9E89
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   162
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   53
            Left            =   240
            TabIndex        =   163
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   53
            Left            =   2280
            Picture         =   "Fa.frx":A172
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   53
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   53
            Left            =   60
            TabIndex        =   164
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   52
         Left            =   2760
         Picture         =   "Fa.frx":AABA
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   159
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   52
            Left            =   240
            TabIndex        =   160
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   52
            Left            =   2280
            Picture         =   "Fa.frx":ADA3
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   52
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   52
            Left            =   60
            TabIndex        =   161
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   51
         Left            =   2760
         Picture         =   "Fa.frx":B6EB
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   156
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   51
            Left            =   240
            TabIndex        =   157
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   51
            Left            =   2280
            Picture         =   "Fa.frx":B9D4
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   51
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   51
            Left            =   60
            TabIndex        =   158
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   50
         Left            =   2760
         Picture         =   "Fa.frx":C31C
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   153
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   50
            Left            =   240
            TabIndex        =   154
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   50
            Left            =   2280
            Picture         =   "Fa.frx":C605
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   50
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   50
            Left            =   60
            TabIndex        =   155
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   49
         Left            =   2760
         Picture         =   "Fa.frx":CF4D
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   150
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   49
            Left            =   240
            TabIndex        =   151
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   49
            Left            =   2280
            Picture         =   "Fa.frx":D236
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   49
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   49
            Left            =   60
            TabIndex        =   152
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   48
         Left            =   2760
         Picture         =   "Fa.frx":DB7E
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   147
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   48
            Left            =   240
            TabIndex        =   148
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   48
            Left            =   2280
            Picture         =   "Fa.frx":DE67
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   48
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   48
            Left            =   60
            TabIndex        =   149
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   47
         Left            =   2760
         Picture         =   "Fa.frx":E7AF
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   144
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   47
            Left            =   240
            TabIndex        =   145
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   47
            Left            =   2280
            Picture         =   "Fa.frx":EA98
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   47
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   47
            Left            =   60
            TabIndex        =   146
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   46
         Left            =   2760
         Picture         =   "Fa.frx":F3E0
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   141
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   46
            Left            =   240
            TabIndex        =   142
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   46
            Left            =   2280
            Picture         =   "Fa.frx":F6C9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   46
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   46
            Left            =   60
            TabIndex        =   143
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   45
         Left            =   2760
         Picture         =   "Fa.frx":10011
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   138
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   45
            Left            =   240
            TabIndex        =   139
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   45
            Left            =   2280
            Picture         =   "Fa.frx":102FA
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   45
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   45
            Left            =   60
            TabIndex        =   140
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   44
         Left            =   2760
         Picture         =   "Fa.frx":10C42
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   135
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   44
            Left            =   240
            TabIndex        =   136
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   44
            Left            =   2280
            Picture         =   "Fa.frx":10F2B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   44
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   44
            Left            =   60
            TabIndex        =   137
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   43
         Left            =   2760
         Picture         =   "Fa.frx":11873
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   132
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   43
            Left            =   240
            TabIndex        =   133
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   43
            Left            =   2280
            Picture         =   "Fa.frx":11B5C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   43
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   43
            Left            =   60
            TabIndex        =   134
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   42
         Left            =   2760
         Picture         =   "Fa.frx":124A4
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   129
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   42
            Left            =   240
            TabIndex        =   130
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   42
            Left            =   2280
            Picture         =   "Fa.frx":1278D
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   42
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   42
            Left            =   60
            TabIndex        =   131
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   41
         Left            =   2760
         Picture         =   "Fa.frx":130D5
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   126
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   41
            Left            =   240
            TabIndex        =   127
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   41
            Left            =   2280
            Picture         =   "Fa.frx":133BE
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   41
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   41
            Left            =   60
            TabIndex        =   128
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   40
         Left            =   2760
         Picture         =   "Fa.frx":13D06
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   123
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   40
            Left            =   240
            TabIndex        =   124
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   40
            Left            =   2280
            Picture         =   "Fa.frx":13FEF
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   40
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   40
            Left            =   60
            TabIndex        =   125
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   39
         Left            =   2760
         Picture         =   "Fa.frx":14937
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   120
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   39
            Left            =   240
            TabIndex        =   121
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   39
            Left            =   2280
            Picture         =   "Fa.frx":14C20
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   39
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   39
            Left            =   60
            TabIndex        =   122
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   38
         Left            =   2760
         Picture         =   "Fa.frx":15568
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   117
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   38
            Left            =   240
            TabIndex        =   118
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   38
            Left            =   2280
            Picture         =   "Fa.frx":15851
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   38
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   38
            Left            =   60
            TabIndex        =   119
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   37
         Left            =   2760
         Picture         =   "Fa.frx":16199
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   114
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   37
            Left            =   240
            TabIndex        =   115
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   37
            Left            =   2280
            Picture         =   "Fa.frx":16482
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   37
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   37
            Left            =   60
            TabIndex        =   116
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   36
         Left            =   2760
         Picture         =   "Fa.frx":16DCA
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   111
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   36
            Left            =   240
            TabIndex        =   112
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   36
            Left            =   2280
            Picture         =   "Fa.frx":170B3
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   36
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   36
            Left            =   60
            TabIndex        =   113
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   35
         Left            =   2760
         Picture         =   "Fa.frx":179FB
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   108
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   35
            Left            =   240
            TabIndex        =   109
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   35
            Left            =   2280
            Picture         =   "Fa.frx":17CE4
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   35
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   35
            Left            =   60
            TabIndex        =   110
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   34
         Left            =   2760
         Picture         =   "Fa.frx":1862C
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   105
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   34
            Left            =   240
            TabIndex        =   106
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   34
            Left            =   2280
            Picture         =   "Fa.frx":18915
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   34
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   34
            Left            =   60
            TabIndex        =   107
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   33
         Left            =   2760
         Picture         =   "Fa.frx":1925D
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   102
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   33
            Left            =   240
            TabIndex        =   103
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   33
            Left            =   2280
            Picture         =   "Fa.frx":19546
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   33
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   33
            Left            =   60
            TabIndex        =   104
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   32
         Left            =   2760
         Picture         =   "Fa.frx":19E8E
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   99
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   32
            Left            =   240
            TabIndex        =   100
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   32
            Left            =   2280
            Picture         =   "Fa.frx":1A177
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   32
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   32
            Left            =   60
            TabIndex        =   101
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   31
         Left            =   2760
         Picture         =   "Fa.frx":1AABF
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   96
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   31
            Left            =   240
            TabIndex        =   97
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   31
            Left            =   2280
            Picture         =   "Fa.frx":1ADA8
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   31
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   31
            Left            =   60
            TabIndex        =   98
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   30
         Left            =   2760
         Picture         =   "Fa.frx":1B6F0
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   93
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   30
            Left            =   240
            TabIndex        =   94
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   30
            Left            =   2280
            Picture         =   "Fa.frx":1B9D9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   30
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   30
            Left            =   60
            TabIndex        =   95
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   29
         Left            =   2760
         Picture         =   "Fa.frx":1C321
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   90
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   29
            Left            =   240
            TabIndex        =   91
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   29
            Left            =   2280
            Picture         =   "Fa.frx":1C60A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   29
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   29
            Left            =   60
            TabIndex        =   92
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   28
         Left            =   2760
         Picture         =   "Fa.frx":1CF52
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   87
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   28
            Left            =   240
            TabIndex        =   88
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   28
            Left            =   2280
            Picture         =   "Fa.frx":1D23B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   28
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   28
            Left            =   60
            TabIndex        =   89
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   27
         Left            =   2760
         Picture         =   "Fa.frx":1DB83
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   84
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   27
            Left            =   240
            TabIndex        =   85
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   27
            Left            =   2280
            Picture         =   "Fa.frx":1DE6C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   27
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   27
            Left            =   60
            TabIndex        =   86
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   26
         Left            =   2760
         Picture         =   "Fa.frx":1E7B4
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   81
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   26
            Left            =   240
            TabIndex        =   82
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   26
            Left            =   2280
            Picture         =   "Fa.frx":1EA9D
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   26
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   26
            Left            =   60
            TabIndex        =   83
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   25
         Left            =   2760
         Picture         =   "Fa.frx":1F3E5
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   78
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   25
            Left            =   240
            TabIndex        =   79
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   25
            Left            =   2280
            Picture         =   "Fa.frx":1F6CE
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   25
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   25
            Left            =   60
            TabIndex        =   80
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   24
         Left            =   2760
         Picture         =   "Fa.frx":20016
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   75
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   24
            Left            =   240
            TabIndex        =   76
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   24
            Left            =   2280
            Picture         =   "Fa.frx":202FF
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   24
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   24
            Left            =   60
            TabIndex        =   77
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   23
         Left            =   2760
         Picture         =   "Fa.frx":20C47
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   72
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   23
            Left            =   240
            TabIndex        =   73
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   23
            Left            =   2280
            Picture         =   "Fa.frx":20F30
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   23
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   23
            Left            =   60
            TabIndex        =   74
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   22
         Left            =   2760
         Picture         =   "Fa.frx":21878
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   69
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   22
            Left            =   240
            TabIndex        =   70
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   22
            Left            =   2280
            Picture         =   "Fa.frx":21B61
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   22
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   60
            TabIndex        =   71
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   21
         Left            =   2760
         Picture         =   "Fa.frx":224A9
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   66
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   21
            Left            =   240
            TabIndex        =   67
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   21
            Left            =   2280
            Picture         =   "Fa.frx":22792
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   21
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   21
            Left            =   60
            TabIndex        =   68
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   20
         Left            =   2760
         Picture         =   "Fa.frx":230DA
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   63
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   20
            Left            =   240
            TabIndex        =   64
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   20
            Left            =   2280
            Picture         =   "Fa.frx":233C3
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   20
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   20
            Left            =   60
            TabIndex        =   65
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   19
         Left            =   2760
         Picture         =   "Fa.frx":23D0B
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   60
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   19
            Left            =   240
            TabIndex        =   61
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   19
            Left            =   2280
            Picture         =   "Fa.frx":23FF4
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   19
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   19
            Left            =   60
            TabIndex        =   62
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   18
         Left            =   2760
         Picture         =   "Fa.frx":2493C
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   57
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   18
            Left            =   240
            TabIndex        =   58
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   18
            Left            =   2280
            Picture         =   "Fa.frx":24C25
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   18
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   18
            Left            =   60
            TabIndex        =   59
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   17
         Left            =   2760
         Picture         =   "Fa.frx":2556D
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   54
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   17
            Left            =   240
            TabIndex        =   55
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   17
            Left            =   2280
            Picture         =   "Fa.frx":25856
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   17
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   17
            Left            =   60
            TabIndex        =   56
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   16
         Left            =   2760
         Picture         =   "Fa.frx":2619E
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   51
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   16
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   16
            Left            =   2280
            Picture         =   "Fa.frx":26487
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   16
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   16
            Left            =   60
            TabIndex        =   53
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   15
         Left            =   2760
         Picture         =   "Fa.frx":26DCF
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   48
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   15
            Left            =   240
            TabIndex        =   49
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   15
            Left            =   2280
            Picture         =   "Fa.frx":270B8
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   15
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   15
            Left            =   60
            TabIndex        =   50
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   14
         Left            =   2760
         Picture         =   "Fa.frx":27A00
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   45
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   14
            Left            =   240
            TabIndex        =   46
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   14
            Left            =   2280
            Picture         =   "Fa.frx":27CE9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   14
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   14
            Left            =   60
            TabIndex        =   47
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   13
         Left            =   2760
         Picture         =   "Fa.frx":28631
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   42
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   13
            Left            =   240
            TabIndex        =   43
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   13
            Left            =   2280
            Picture         =   "Fa.frx":2891A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   13
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   13
            Left            =   60
            TabIndex        =   44
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   12
         Left            =   2760
         Picture         =   "Fa.frx":29262
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   39
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   12
            Left            =   240
            TabIndex        =   40
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   12
            Left            =   2280
            Picture         =   "Fa.frx":2954B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   12
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   12
            Left            =   60
            TabIndex        =   41
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   11
         Left            =   2760
         Picture         =   "Fa.frx":29E93
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   36
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   11
            Left            =   240
            TabIndex        =   37
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   11
            Left            =   2280
            Picture         =   "Fa.frx":2A17C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   11
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   60
            TabIndex        =   38
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   10
         Left            =   2760
         Picture         =   "Fa.frx":2AAC4
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   33
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   10
            Left            =   240
            TabIndex        =   34
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   10
            Left            =   2280
            Picture         =   "Fa.frx":2ADAD
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   10
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   60
            TabIndex        =   35
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   9
         Left            =   2760
         Picture         =   "Fa.frx":2B6F5
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   30
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   9
            Left            =   240
            TabIndex        =   31
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   9
            Left            =   2280
            Picture         =   "Fa.frx":2B9DE
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   9
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   60
            TabIndex        =   32
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   8
         Left            =   2760
         Picture         =   "Fa.frx":2C326
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   27
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   8
            Left            =   240
            TabIndex        =   28
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   8
            Left            =   2280
            Picture         =   "Fa.frx":2C60F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   8
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   60
            TabIndex        =   29
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   7
         Left            =   2760
         Picture         =   "Fa.frx":2CF57
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   24
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   7
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   7
            Left            =   2280
            Picture         =   "Fa.frx":2D240
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   7
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   60
            TabIndex        =   26
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   6
         Left            =   2760
         Picture         =   "Fa.frx":2DB88
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   21
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   6
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   6
            Left            =   2280
            Picture         =   "Fa.frx":2DE71
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   6
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   60
            TabIndex        =   23
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   5
         Left            =   2760
         Picture         =   "Fa.frx":2E7B9
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   18
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   5
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   5
            Left            =   2280
            Picture         =   "Fa.frx":2EAA2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   5
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   20
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   4
         Left            =   2760
         Picture         =   "Fa.frx":2F3EA
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   15
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   4
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   4
            Left            =   2280
            Picture         =   "Fa.frx":2F6D3
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   4
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   17
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   3
         Left            =   2760
         Picture         =   "Fa.frx":3001B
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   12
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   3
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   3
            Left            =   2280
            Picture         =   "Fa.frx":30304
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   3
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   14
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   2
         Left            =   2760
         Picture         =   "Fa.frx":30C4C
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   9
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   2
            Left            =   2280
            Picture         =   "Fa.frx":30F35
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   2
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   11
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   1
         Left            =   2760
         Picture         =   "Fa.frx":3187D
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   6
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   1
            Left            =   2280
            Picture         =   "Fa.frx":31B66
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   1
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   8
            Top             =   32
            Width           =   630
         End
      End
      Begin VB.PictureBox PageBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   0
         Left            =   0
         Picture         =   "Fa.frx":324AE
         ScaleHeight     =   1935
         ScaleMode       =   0  'User
         ScaleWidth      =   2745
         TabIndex        =   3
         Top             =   800
         Width           =   2775
         Begin VB.TextBox PageUrl 
            Height          =   270
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ComEndWeb 
            Height          =   495
            Index           =   0
            Left            =   2280
            Picture         =   "Fa.frx":32797
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Label WebName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebName"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   32
            Width           =   630
         End
         Begin VB.Image Page 
            Appearance      =   0  'Flat
            Height          =   1646
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2746
         End
      End
      Begin VB.Image ComTranslate 
         Height          =   615
         Left            =   11160
         Stretch         =   -1  'True
         Top             =   120
         Width           =   615
      End
      Begin VB.Label LabelSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "搜索："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   375
         Left            =   5400
         TabIndex        =   207
         Top             =   200
         Width           =   4215
      End
      Begin VB.Image ComSearch 
         Height          =   615
         Left            =   9720
         Stretch         =   -1  'True
         Top             =   120
         Width           =   615
      End
      Begin VB.Image ComCell 
         Height          =   765
         Left            =   0
         MouseIcon       =   "Fa.frx":330DF
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   0
         Width           =   765
      End
      Begin VB.Image ComBack 
         Height          =   615
         Left            =   960
         Stretch         =   -1  'True
         Top             =   120
         Width           =   615
      End
      Begin VB.Image ComAdvance 
         Height          =   615
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   120
         Width           =   615
      End
      Begin VB.Image ComRefresh 
         Height          =   615
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   120
         Width           =   615
      End
      Begin VB.Image ComHome 
         Height          =   615
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   615
      End
      Begin VB.Image ComFavorite 
         Height          =   615
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   120
         Width           =   615
      End
      Begin VB.Image ComRestore 
         Height          =   615
         Left            =   4560
         Stretch         =   -1  'True
         Top             =   120
         Width           =   615
      End
      Begin VB.Image ComExit 
         Height          =   735
         Left            =   14760
         Stretch         =   -1  'True
         Top             =   0
         Width           =   735
      End
      Begin VB.Image ComMax 
         Height          =   735
         Left            =   14040
         Stretch         =   -1  'True
         Top             =   0
         Width           =   735
      End
      Begin VB.Image ComMin 
         Height          =   735
         Left            =   13320
         Stretch         =   -1  'True
         Top             =   0
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   13560
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   15500
         X2              =   15500
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Image BackUP 
         Height          =   765
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Image BackDown 
         Height          =   765
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin VB.PictureBox P2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   16245
      TabIndex        =   0
      Top             =   2760
      Width           =   16245
      Begin VB.TextBox ActiveUrl 
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   205
         Text            =   "Http:\\"
         Top             =   120
         Width           =   6375
      End
      Begin VB.TextBox OpenNewPage 
         Height          =   375
         Left            =   8400
         TabIndex        =   204
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Timer openListpageback 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   9960
         Top             =   120
      End
      Begin VB.Image ComGO 
         Height          =   615
         Left            =   6600
         MouseIcon       =   "Fa.frx":33231
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image ComWebList 
         Height          =   495
         Left            =   13320
         MouseIcon       =   "Fa.frx":33383
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Image ComExitWeb 
         Height          =   495
         Left            =   14760
         MouseIcon       =   "Fa.frx":334D5
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Image ComUP 
         Height          =   495
         Left            =   14280
         MouseIcon       =   "Fa.frx":33627
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Image ComDown 
         Height          =   495
         Left            =   13800
         MouseIcon       =   "Fa.frx":33779
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   2400
         Y2              =   0
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   15600
         X2              =   15600
         Y1              =   0
         Y2              =   2520
      End
      Begin VB.Label LabelUrl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Http://"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6720
         TabIndex        =   1
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.Menu CellMenu 
      Caption         =   "菜单"
      Begin VB.Menu MenuNewPage 
         Caption         =   "新建空白页"
      End
      Begin VB.Menu MenuSaveWeb 
         Caption         =   "保存网页"
      End
      Begin VB.Menu MenuLine7 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSearchWeb 
         Caption         =   "查找"
      End
      Begin VB.Menu MenuPrintWeb 
         Caption         =   "打印"
      End
      Begin VB.Menu Menuinstrument 
         Caption         =   "工具"
         Begin VB.Menu MenuClear 
            Caption         =   "清理浏览数据"
            Visible         =   0   'False
         End
         Begin VB.Menu MenuShowHtml 
            Caption         =   "查看源代码"
         End
      End
      Begin VB.Menu Menuline5 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAddCollect 
         Caption         =   "添加自定义导航"
      End
      Begin VB.Menu MenuSet 
         Caption         =   "浏览器设置"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "关于"
      End
      Begin VB.Menu MenuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu MenuIECollect 
         Caption         =   "IE收藏夹管理"
      End
      Begin VB.Menu MenuIESte 
         Caption         =   "Internet选项"
      End
      Begin VB.Menu Menuline1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuDefaultBrowser 
         Caption         =   "设为默认浏览器"
      End
      Begin VB.Menu MenuEnd 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu MenuWebList 
      Caption         =   "网页菜单"
      Begin VB.Menu MenuWebOnly 
         Caption         =   "单独展示"
      End
      Begin VB.Menu MenuTimeRefresh 
         Caption         =   "定时刷新"
      End
      Begin VB.Menu Muneline2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuTranslate 
         Caption         =   "翻译此网页"
      End
      Begin VB.Menu Menuline3 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAddFavorite 
         Caption         =   "收藏"
         Begin VB.Menu MenuAddIEFa 
            Caption         =   "添加到IE收藏夹"
         End
         Begin VB.Menu MenuAddMyFa 
            Caption         =   "添加到初始导航"
         End
      End
      Begin VB.Menu Menuline4 
         Caption         =   "-"
      End
      Begin VB.Menu UnloadAllWeb 
         Caption         =   "关闭所有网页"
      End
   End
   Begin VB.Menu MenuRestore 
      Caption         =   "恢复"
      Begin VB.Menu MenuHuiFu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "Fa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
"URLDownloadToFileA" (ByVal pCaller As Long, ByVal _
szURL As String, ByVal szFileName As String, ByVal _
dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5
'--------------------打开记事本+查看网页源代码--------------------------------------------

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Fnew     As Form
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1


'获取窗体结构信息函数

Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000
Private Const WS_CAPTION = &HC00000
Private Const WS_SIZEBOX = &H40000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
'为窗体指定一个新位置和状态函数
Private Declare Function SetWindowPos Lib "user32 " (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOREPOSITION = &H200
'获得整个窗体的大小和位置
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'获取windows任务栏高度预设

Private Declare Function FindWindow Lib "user32 " Alias "FindWindowA " (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32 " (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
                Left   As Long
                Top   As Long
                Right   As Long
                Bottom   As Long
End Type

Dim Change     As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'鼠标位置
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
'通用变量定义区:
Dim Xleft As Integer '最左边pageback的位置
Dim MouseStep As Boolean '鼠标是否按下
Dim Xdistance As Integer '鼠标按下位置到xleft的距离
Dim FirstPageBack As Integer '纠正pageback位置（第一个）
Dim NearestPageback As Integer '纠正pggeback位置（最接近0）
Public PublicIndex As Integer '通用序号


Public Sub PageBackClick() '列表点击
PageBack(PublicIndex).Picture = LoadPicture(App.path & "\skin\pageback1.gif")
Dim ofrm As Form
  For Each ofrm In Forms
        If ofrm.Caption = Trim(Str(PublicIndex)) Then
                ofrm.ZOrder
                ofrm.RefreshWebname.Enabled = True
        End If
        
  Next
End Sub

Private Sub ActiveUrl_Change()
On Error Resume Next
LabelUrl.Caption = ActiveUrl.Text
If ActiveUrl.Width >= Me.Width - (ActiveUrl.Left + ComWebList.Width + ComDown.Width + ComUP.Width + ComExitWeb.Width + ComGO.Width) Then
ActiveUrl.Width = Me.Width - (ActiveUrl.Left + ComWebList.Width + ComDown.Width + ComUP.Width + ComExitWeb.Width + ComGO.Width)
Else
ActiveUrl.Width = LenB(StrConv(ActiveUrl, vbFromUnicode)) * 200
End If
ComGO.Left = ActiveUrl.Left + ActiveUrl.Width
End Sub

Private Sub ActiveUrl_DblClick()
ActiveUrl.SelStart = 0
ActiveUrl.SelLength = Len(ActiveUrl.Text) '全选
End Sub

Private Sub ActiveUrl_GotFocus()
LabelUrl.Visible = False
ComGO.Left = ActiveUrl.Left + ActiveUrl.Width
ComGO.Visible = True
End Sub

Private Sub ActiveUrl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ComGO_Click
End If
End Sub

Private Sub ActiveUrl_LostFocus()
ActiveUrl.Visible = False
LabelUrl.Visible = True
ComGO.Visible = False
End Sub

Private Sub CloseHistory_Change()
On Error Resume Next
'防止重复
Dim i As Integer
For i = 0 To CloseHistory.ListCount
If CloseHistory.List(i) = CloseHistory.Text Then Exit Sub
Next i

If CloseHistory.Text <> "" Then CloseHistory.AddItem (CloseHistory.Text)


For i = 0 To ComboRestore.ListCount - 1
 If i <> 0 Then Load MenuHuiFu(i)
 MenuHuiFu(i).Visible = True
 MenuHuiFu(i).Caption = ComboRestore.List(i)
Next i
End Sub

Private Sub ComAdvance_Click()
On Error Resume Next
ActiveForm.WebPage.GoForward
End Sub

Private Sub ComBack_Click()
On Error Resume Next
ActiveForm.WebPage.GoBack
End Sub

Private Sub ComboRestore_Change()
'防止重复
Dim i As Integer
For i = 0 To ComboRestore.ListCount
If ComboRestore.List(i) = ComboRestore.Text Then Exit Sub
Next i

If ComboRestore.Text <> "" Then ComboRestore.AddItem (ComboRestore.Text)
End Sub

Private Sub ComCell_Click()
 PopupMenu CellMenu
End Sub


Private Sub ComCell_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub

Private Sub ComCell_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Left <> ComCell.Left Then
BackUP.Visible = False
BackUP.Left = ComCell.Left
BackUP.Visible = True
End If

If BackDown.Left <> ComCell.Left Then
BackDown.Visible = False
BackDown.Left = ComCell.Left
BackDown.Visible = True
End If
End Sub

Private Sub ComCell_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub ComFavorite_Click()
FormFavorite.Show
End Sub



Private Sub commin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub

Private Sub commin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Left <> ComMin.Left Then
BackUP.Visible = False
BackUP.Left = ComMin.Left
BackUP.Visible = True
End If

If BackDown.Left <> ComMin.Left Then
BackDown.Visible = False
BackDown.Left = ComMin.Left
BackDown.Visible = True
End If
End Sub

Private Sub commin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub commax_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub

Private Sub commax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Left <> ComMax.Left Then
BackUP.Visible = False
BackUP.Left = ComMax.Left
BackUP.Visible = True
End If

If BackDown.Left <> ComMax.Left Then
BackDown.Visible = False
BackDown.Left = ComMax.Left
BackDown.Visible = True
End If
End Sub

Private Sub commax_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub comexit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub

Private Sub comexit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Left <> ComExit.Left Then
BackUP.Visible = False
BackUP.Left = ComExit.Left
BackUP.Visible = True
End If

If BackDown.Left <> ComExit.Left Then
BackDown.Visible = False
BackDown.Left = ComExit.Left
BackDown.Visible = True
End If
End Sub

Private Sub comexit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub comback_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub
Private Sub comback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Left <> ComBack.Left - 80 Then
BackUP.Visible = False
BackUP.Left = ComBack.Left - 80
BackUP.Visible = True
End If
If BackDown.Left <> ComBack.Left - 80 Then
BackDown.Visible = False
BackDown.Left = ComBack.Left - 80
BackDown.Visible = True
End If
End Sub
Private Sub comback_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub comadvance_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub
Private Sub comadvance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Left <> ComAdvance.Left - 80 Then
BackUP.Visible = False
BackUP.Left = ComAdvance.Left - 80
BackUP.Visible = True
End If
If BackDown.Left <> ComAdvance.Left - 80 Then
BackDown.Visible = False
BackDown.Left = ComAdvance.Left - 80
BackDown.Visible = True
End If
End Sub
Private Sub comadvance_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub comrefresh_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub
Private Sub comrefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Left <> ComRefresh.Left - 80 Then
BackUP.Visible = False
BackUP.Left = ComRefresh.Left - 80
BackUP.Visible = True
End If
If BackDown.Left <> ComRefresh.Left - 80 Then
BackDown.Visible = False
BackDown.Left = ComRefresh.Left - 80
BackDown.Visible = True
End If
End Sub
Private Sub comrefresh_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub comhome_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub
Private Sub comhome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Left <> ComHome.Left - 80 Then
BackUP.Visible = False
BackUP.Left = ComHome.Left - 80
BackUP.Visible = True
End If
If BackDown.Left <> ComHome.Left - 80 Then
BackDown.Visible = False
BackDown.Left = ComHome.Left - 80
BackDown.Visible = True
End If
End Sub
Private Sub comhome_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub comfavorite_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub
Private Sub comfavorite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Left <> ComFavorite.Left - 80 Then
BackUP.Visible = False
BackUP.Left = ComFavorite.Left - 80
BackUP.Visible = True
End If
If BackDown.Left <> ComFavorite.Left - 80 Then
BackDown.Visible = False
BackDown.Left = ComFavorite.Left - 80
BackDown.Visible = True
End If
End Sub
Private Sub comfavorite_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub ComRestore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub

Private Sub ComRestore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Left <> ComRestore.Left - 80 Then
BackUP.Visible = False
BackUP.Left = ComRestore.Left - 80
BackUP.Visible = True
End If
If BackDown.Left <> ComRestore.Left - 80 Then
BackDown.Visible = False
BackDown.Left = ComRestore.Left - 80
BackDown.Visible = True
End If
End Sub

Private Sub ComRestore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub comsearch_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub
Private Sub comsearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Visible = False Then BackUP.Visible = True
If BackDown.Visible = False Then BackDown.Visible = True
If BackUP.Left <> ComSearch.Left - 80 Then
BackUP.Visible = False
BackUP.Left = ComSearch.Left - 80
BackUP.Visible = True
End If
If BackDown.Left <> ComSearch.Left - 80 Then
BackDown.Visible = False
BackDown.Left = ComSearch.Left - 80
BackDown.Visible = True
End If
End Sub
Private Sub comsearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub
Private Sub ComDown_Click()
P1.Height = 2760
LoadBackPicture
Dim ofrm As Form
  For Each ofrm In Forms
        If ofrm.Caption <> "细胞浏览器" Then
                ofrm.Height = Fa.Height - P1.Height - P2.Height
        End If
        
  Next
End Sub

Private Sub ComEndWeb_Click(Index As Integer)
Dim ofrm As Form
  For Each ofrm In Forms
        If ofrm.Caption = Trim(Str(Index)) Then
                Unload ofrm
        End If
        
  Next
ListPageBack
End Sub

Private Sub ComExitWeb_Click()
If ActiveForm.Name = "FrmWeb" Then
Unload ActiveForm
End If
Dim i As Integer
Dim n As Integer
n = 0
For i = 0 To 66
If PageBack(i).Visible = False Then
n = n + 1
End If
If n = 66 Then
LinkPage.Show
End If
Next i
End Sub

Private Sub ComGO_Click()
On Error Resume Next
ActiveForm.WebPage.Navigate ActiveUrl.Text
If Err Then Fa.OpenNewPage.Text = ActiveUrl.Text
End Sub

Private Sub ComHome_Click()
'载入主页
Dim cw As String
Dim T3 As String
Open App.path & "\homepage.dat" For Input As #1
  Do While Not EOF(1)
    Line Input #1, cw$
    T3 = cw
    Loop
Close #1

Dim l
l = Split(T3, ",")
Dim i As Integer
Dim T4 As String
For i = 1 To UBound(l)
T4 = T4 & Chr(l(i) / 3)
Next i
OpenNewPage.Text = T4

End Sub



Private Sub ComMax_Click()
If Me.WindowState = 0 Then
Me.WindowState = 2
Else
Me.WindowState = 0
End If
Dim ofrm As Form
  For Each ofrm In Forms
        If ofrm.Caption <> "细胞浏览器" Then
                ofrm.Height = Fa.Height - P1.Height - P2.Height
        End If
        
  Next
End Sub

Private Sub ComMin_Click()
Me.WindowState = 1
End Sub



Private Sub ComRefresh_Click()
ActiveForm.WebPage.Refresh
End Sub

Private Sub ComSearch_Click()
 Dim frmSearch As FormSearch
 Set frmSearch = New FormSearch

  frmSearch.Show

  frmSearch.Caption = "搜索：" & TextSearch.Text
  frmSearch.WebPage.Navigate "http://www.baidu.com/s?wd=" & TextSearch.Text
SearchNum = SearchNum + 1
End Sub

Private Sub ComTranslate_Click()
Shell App.path & "\translate.exe", vbNormalNoFocus
End Sub

Private Sub ComTranslate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
End Sub

Private Sub ComTranslate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BackUP.Visible = False Then BackUP.Visible = True
If BackDown.Visible = False Then BackDown.Visible = True
If BackUP.Left <> ComTranslate.Left - 80 Then
BackUP.Visible = False
BackUP.Left = ComTranslate.Left - 80
BackUP.Visible = True
End If
If BackDown.Left <> ComTranslate.Left - 80 Then
BackDown.Visible = False
BackDown.Left = ComTranslate.Left - 80
BackDown.Visible = True
End If
End Sub

Private Sub ComTranslate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = True
End Sub

Private Sub ComUP_Click()
P1.Height = 1237
LoadBackPicture
Dim ofrm As Form
  For Each ofrm In Forms
        If ofrm.Caption <> "细胞浏览器" Then
                ofrm.Height = Fa.Height - P1.Height - P2.Height
        End If
        
  Next
End Sub



Private Sub ComWebList_Click()
If ActiveForm.Caption <> "LinkPage" Then
 PopupMenu MenuWebList
End If
End Sub

Private Sub ComRestore_Click()
PopupMenu MenuRestore
End Sub

Private Sub LabelSearch_Click()
TextSearch.Visible = True
TextSearch.SetFocus
End Sub

Private Sub LabelUrl_Click()
ActiveUrl.Visible = True
ActiveUrl.SetFocus
End Sub

Private Sub MDIForm_Load()

Dim lStyle     As Long
Dim MyRect     As RECT
'获取窗体的大小和位置
GetWindowRect Me.hWnd, MyRect
'取得当前窗体信息
lStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
If Change Then
'分别使系统菜单（包括图标）、标题、大小、最大化、最小化显示/有效
'lStyle = lStyle Or WS_SYSMENU
'lStyle = lStyle Or WS_CAPTION
'lStyle = lStyle Or WS_SIZEBOX
'lStyle = lStyle Or WS_MAXIMIZEBOX
'lStyle = lStyle Or WS_MINIMIZEBOX
Else
'分别使系统菜单（包括图标）、标题、大小、最大化、最小化隐藏/无效
'lStyle = lStyle And Not WS_SYSMENU
lStyle = lStyle And Not WS_CAPTION
lStyle = lStyle And Not WS_SIZEBOX
lStyle = lStyle And Not WS_MAXIMIZEBOX
lStyle = lStyle And Not WS_MINIMIZEBOX
End If
'按lStyle的值设置窗体信息
SetWindowLong Me.hWnd, GWL_STYLE, lStyle
'保持窗体的大小与位置不变
SetWindowPos Me.hWnd, 0, MyRect.Left, MyRect.Top, MyRect.Right - MyRect.Left, MyRect.Bottom - MyRect.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
'改变显示/隐藏状态
Change = Not Change
'以上为隐藏Fa的标题栏过程~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CellMenu.Visible = False
MenuWebList.Visible = False
MenuRestore.Visible = False

Xleft = 0
Me.Top = 0
Me.Left = 0
Me.Height = Screen.Height - (MyRect.Bottom - MyRect.Top)
Me.Width = Screen.Width
Dim i As Integer
For i = 0 To 66
'Page(i).Stretch = True
PageBack(i).Visible = False
Next i

If Command() = "" Then
ComHome_Click
Else
OpenNewPage.Text = Command()
End If
ArrayCom '排列
LoadComPicture '载入图片
End Sub
Private Sub LoadComPicture()
On Error Resume Next

'Dim rtn As Long
  '  rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
  '  rtn = rtn Or WS_EX_LAYERED
  '  SetWindowLong hwnd, GWL_EXSTYLE, rtn
  '  SetLayeredWindowAttributes hwnd, 0, 245, LWA_ALPHA '窗体透明度
'p2特殊
Dim TempPic     As StdPicture
Set TempPic = LoadPicture(App.path & "\skin\p1.gif")
P2.PaintPicture TempPic, 0, -P1.Height
LinkPage.Picture = LoadPicture("")
LinkPage.PaintPicture TempPic, 0, -P1.Height - P2.Height
Set TempPic = Nothing   '释放
Me.Picture = LoadPicture(App.path & "\skin\FormBack.gif")
ComCell.Picture = LoadPicture(App.path & "\skin\comcell.gif")
P1.Picture = LoadPicture(App.path & "\skin\P1.gif")
ComBack.Picture = LoadPicture(App.path & "\skin\comback.gif")
ComAdvance.Picture = LoadPicture(App.path & "\skin\comadvance.gif")
ComRefresh.Picture = LoadPicture(App.path & "\skin\comrefresh.gif")
ComHome.Picture = LoadPicture(App.path & "\skin\comhome.gif")
ComFavorite.Picture = LoadPicture(App.path & "\skin\comfavorite.gif")
ComSearch.Picture = LoadPicture(App.path & "\skin\comsearch.gif")
ComMin.Picture = LoadPicture(App.path & "\skin\commin.gif")
ComMax.Picture = LoadPicture(App.path & "\skin\commax.gif")
ComExit.Picture = LoadPicture(App.path & "\skin\comexit.gif")
ComDown.Picture = LoadPicture(App.path & "\skin\comdown.gif")
ComUP.Picture = LoadPicture(App.path & "\skin\comup.gif")
ComExitWeb.Picture = LoadPicture(App.path & "\skin\comexitweb.gif")
BackDown.Picture = LoadPicture(App.path & "\skin\backdown.gif")
BackUP.Picture = LoadPicture(App.path & "\skin\backup.gif")
ComGO.Picture = LoadPicture(App.path & "\skin\comgo.gif")
ComWebList.Picture = LoadPicture(App.path & "\skin\comweblist.gif")
ComTranslate.Picture = LoadPicture(App.path & "\skin\comtranslate.gif")
ComRestore.Picture = LoadPicture(App.path & "\skin\comrestore.gif")

End Sub
Private Sub LoadBackPicture()
'p2特殊
Dim TempPic     As StdPicture
Set TempPic = LoadPicture(App.path & "\skin\p1.gif")
P2.PaintPicture TempPic, 0, -P1.Height
LinkPage.Picture = LoadPicture("")
LinkPage.PaintPicture TempPic, 0, -P1.Height - P2.Height
Set TempPic = Nothing   '释放
End Sub
Private Sub ArrayCom()

ComExit.Left = Me.Width - ComExit.Width
ComMax.Left = Me.Width - ComExit.Width - ComMax.Width
ComMin.Left = Me.Width - ComExit.Width - ComMax.Width - ComMin.Width
ActiveUrl.Left = 128
'ActiveUrl.Width = Me.Width - 5500
LabelUrl.Left = ActiveUrl.Left
LabelUrl.Width = ActiveUrl.Width
ComExitWeb.Left = Me.Width - ComExitWeb.Width
ComUP.Left = ComExitWeb.Left - ComUP.Width
ComDown.Left = ComUP.Left - ComDown.Width
ComWebList.Left = ComDown.Left - ComWebList.Width
'Line1.X2 = Me.Width
Line2.X2 = Me.Width
Line4.x1 = Me.Width - 16
Line4.X2 = Me.Width - 16
Line5.x1 = Me.Width - 16
Line5.X2 = Me.Width - 16
'以上为基本调整及载入过程~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

End Sub

Private Sub LoadNewPage() '载入新窗体
 Static IDPage As Long
 Dim frmPage As FrmWeb
 Set frmPage = New FrmWeb
 frmPage.Caption = IDPage
  frmPage.Show 0
  PageBack(IDPage).Visible = True
   If OpenNewPage.Text = "" Then
    '打开本地导航页
   Else
    If Left(OpenNewPage.Text, 1) = "s" Then
    Dim l() As String
     l = Split(OpenNewPage.Text, ",", 2)
     frmPage.WebPage.Navigate l(1)
     PageUrl(IDPage).Text = l(1)
    Else
     frmPage.WebPage.Navigate OpenNewPage.Text
     PageUrl(IDPage).Text = OpenNewPage.Text
    End If
   End If
 OpenNewPage.Text = ""
 IDPage = IDPage + 1
End Sub
Private Sub ListPageBack() '排列网页列表
On Error Resume Next
Dim Xnum As Integer
Dim i As Integer
Xnum = Xleft

For i = 0 To 66
 If PageBack(i).Visible = True Then
  PageBack(i).Left = Xnum
  Xnum = Xnum + PageBack(i).Width '会溢出
 End If
Next i
End Sub

Private Sub CorrectPageBack() ' 纠正网页列表
Dim i As Integer
Dim n As Integer
 For i = 0 To 66
  If PageBack(i).Visible = True Then
   PageBack(i).Left = PageBack(FirstPageBack).Left + PageBack(0).Width * n
   n = n + 1
  End If
 Next i
 
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmSplash
End Sub

Private Sub MDIForm_Resize()
'ActiveForm.Height = Fa.Height - (Fa.P1.Height + Fa.P2.Height)
'ActiveForm.Width = Fa.Width
End Sub

Private Sub MenuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub MenuAddCollect_Click()
Dim i As Integer
For i = 0 To 8
If LinkPage.WebURL(i).Text = "" Then
FormAddLink.Show
LinkPageNum = i
Exit Sub
End If
Next i
MsgBox ("没有空白导航栏")
End Sub

Private Sub MenuAddIEFa_Click()
AddIEFavorite.Show 0
End Sub

Private Sub MenuAddMyFa_Click()
Dim n As Integer
Dim neirong As String
Dim s(0 To 8) As String
Dim i As Integer
n = 0
Open App.path & "\collect.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, neirong$
s(n) = neirong
n = n + 1
Loop
Close #1
For i = 0 To 8
If s(i) = "" Then
s(i) = WebName(Val(ActiveForm.Caption)).Caption & "," & PageUrl(Val(ActiveForm.Caption)).Text
SavePicture Page(Val(ActiveForm.Caption)).Picture, App.path & "\webimage\" & WebName(Val(ActiveForm.Caption)).Caption & ".gif"
GoTo savethis
End If
Next i

savethis:
i = 0
On Error GoTo x1
Kill (App.path & "\collect.dat")
x1:

Open App.path & "\collect.dat" For Append As #1
For i = 0 To 8
Print #1, s(i)
Next i
Close #1
Unload LinkPage

 MsgBox "保存成功", vbOKOnly, "成功"
End Sub

Private Sub MenuDefaultBrowser_Click()
Shell App.path & "\defaultbrowser.exe"
End Sub

Private Sub MenuEnd_Click()
End
End Sub

Private Sub MenuHuiFu_Click(Index As Integer)
OpenNewPage.Text = CloseHistory.List(Index)
End Sub

Private Sub MenuIECollect_Click()
Shell App.path & "\Favorites.exe", vbNormalNoFocus
End Sub

Private Sub MenuIESte_Click()
Dim ret
 ret = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Sub

Private Sub MenuNewPage_Click()
LinkPage.Show
LinkPage.ZOrder
End Sub




Private Sub MenuPrintWeb_Click()
If ActiveForm.Name = "LinkPage" Then Exit Sub
ActiveForm.WebPage.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub MenuSaveWeb_Click()
If ActiveForm.Name = "LinkPage" Then Exit Sub
ActiveForm.WebPage.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub MenuSearchWeb_Click()
If ActiveForm.Name = "LinkPage" Then Exit Sub
ActiveForm.WebPage.SetFocus
SendKeys "^{f} "
End Sub

Private Sub MenuSet_Click()
FormSet.Show 0
End Sub

Private Sub MenuShowHtml_Click()
If ActiveForm.Name = "LinkPage" Then Exit Sub
URLDownloadToFile 0, ActiveForm.WebPage.LocationURL, App.path & "\ShowHtml.txt", 0, 0
ShellExecute Me.hWnd, "Open", App.path & "\ShowHtml.txt", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub MenuTimeRefresh_Click()

 Dim frmPage As frmBrowser
 Set frmPage = New frmBrowser

  frmPage.Show

End Sub

Private Sub MenuTranslate_Click()
OpenNewPage.Text = "http://translate.googleusercontent.com/translate_c?hl=zh-CN&sl=zh-CN&tl=en&u=" & ActiveForm.WebPage.LocationURL & "&usg=ALkJrhgo8djd0k6-nWl2-0osZzRWAYtrOw"
End Sub

Private Sub MenuWebOnly_Click()
Open App.path & "\OutUrl.dat" For Append As #1
Print #1, ActiveForm.WebPage.LocationURL
Close #1
Shell (App.path & "\WebOnly.exe")
End Sub

Private Sub openListpageback_Timer() 'Frmweb调用时，用来刷新列表位置
ListPageBack
openListpageback.Enabled = False
End Sub

Private Sub OpenNewPage_Change()

If OpenNewPage.Text <> "" Then
 If Me.WindowState = 1 Then Me.WindowState = 0
  LoadNewPage

End If
End Sub







Private Sub P1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackUP.Visible = False
BackDown.Visible = False
End Sub

Private Sub Page_Click(Index As Integer)
PageBack(PublicIndex).Picture = LoadPicture(App.path & "\skin\pageback0.gif")
PublicIndex = Index
PageBackClick
End Sub





Private Sub ComExit_Click()
If SetStr(1) = "true" Then
 UnloadCom.Show 1
Else
 End
End If
End Sub

Private Sub Page_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
MouseStep = True
Dim Point As POINTAPI
GetCursorPos Point
Dim i As Integer
 For i = 0 To 66
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
Xdistance = Point.X * 15 - PageBack(FirstPageBack).Left
End Sub

Private Sub Page_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseStep = True Then
Dim Point As POINTAPI
GetCursorPos Point
Xleft = Point.X * 15 - Xdistance
ListPageBack
End If
End Sub

Private Sub Page_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = False

Dim i As Integer
Dim n As Integer
 For i = 0 To 66
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
 
 For n = 0 To 66
  If PageBack(n).Visible = True And PageBack(n).Left < PageBack(0).Width And PageBack(n).Left >= 0 Then
   NearestPageback = n
   Exit For
  Else
   NearestPageback = FirstPageBack
  End If
 Next n

TimerCorrect.Enabled = True
End Sub






Private Sub PageBack_Click(Index As Integer)
PageBack(PublicIndex).Picture = LoadPicture(App.path & "\skin\pageback0.gif")
PublicIndex = Index
PageBackClick
End Sub

Private Sub PageBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = True
Dim Point As POINTAPI
GetCursorPos Point
Dim i As Integer
 For i = 0 To 66
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
Xdistance = Point.X * 15 - PageBack(FirstPageBack).Left
End Sub

Private Sub PageBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseStep = True Then
Dim Point As POINTAPI
GetCursorPos Point
Xleft = Point.X * 15 - Xdistance
ListPageBack
End If
End Sub

Private Sub PageBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = False

Dim i As Integer
Dim n As Integer
 For i = 0 To 66
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
 
 For n = 0 To 66
  If PageBack(n).Visible = True And PageBack(n).Left < PageBack(0).Width And PageBack(n).Left >= 0 Then
   NearestPageback = n
   Exit For
  Else
   NearestPageback = FirstPageBack
  End If
 Next n

TimerCorrect.Enabled = True
End Sub


Private Sub TextSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ComSearch_Click
End If
End Sub

Private Sub TextSearch_LostFocus()
TextSearch.Visible = False
End Sub



Private Sub TimerCorrect_Timer()
PageBack(FirstPageBack).Left = PageBack(FirstPageBack).Left - 560
CorrectPageBack


 If PageBack(NearestPageback).Left < 0 Then
  Dim DistanceNtoF As Integer
  DistanceNtoF = PageBack(NearestPageback).Left - PageBack(FirstPageBack).Left
  PageBack(NearestPageback).Left = 0
  PageBack(FirstPageBack).Left = 0 - DistanceNtoF
  CorrectPageBack
  Xleft = PageBack(FirstPageBack).Left
  TimerCorrect.Enabled = False
 End If
 
End Sub

Private Sub UnloadAllWeb_Click()

Dim ofrm As Form
  For Each ofrm In Forms
          If ofrm.Name = "FrmWeb" Then
                Unload ofrm
        End If
        
  Next
End Sub

Private Sub WebName_Click(Index As Integer)
PageBack(PublicIndex).Picture = LoadPicture(App.path & "\skin\pageback0.gif")
PublicIndex = Index
PageBackClick

End Sub

Private Sub WebName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = True
Dim Point As POINTAPI
GetCursorPos Point
Dim i As Integer
 For i = 0 To 66
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
Xdistance = Point.X * 15 - PageBack(FirstPageBack).Left
End Sub

Private Sub WebName_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseStep = True Then
Dim Point As POINTAPI
GetCursorPos Point
Xleft = Point.X * 15 - Xdistance
ListPageBack
End If
End Sub

Private Sub WebName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseStep = False

Dim i As Integer
Dim n As Integer
 For i = 0 To 66
  If PageBack(i).Visible = True Then
   FirstPageBack = i
   Exit For
  End If
 Next i
 
 For n = 0 To 66
  If PageBack(n).Visible = True And PageBack(n).Left < PageBack(0).Width And PageBack(n).Left >= 0 Then
   NearestPageback = n
   Exit For
  Else
   NearestPageback = FirstPageBack
  End If
 Next n

TimerCorrect.Enabled = True
End Sub
