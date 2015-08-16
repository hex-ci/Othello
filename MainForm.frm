VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "黑白棋.Net"
   ClientHeight    =   6330
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9480
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "MainForm.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9480
   StartUpPosition =   3  '窗口缺省
   Begin Othello.FlatButton fltbtnExit 
      Height          =   240
      Left            =   9075
      TabIndex        =   9
      Top             =   165
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      MousePointer    =   99
      Style           =   1
      Picture         =   "MainForm.frx":1CFA
      OverPicture     =   "MainForm.frx":1DA0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTip         =   "退出游戏"
      ForeColor       =   0
      AutoSize        =   -1  'True
   End
   Begin Othello.FlatButton fltbtnMin 
      Height          =   240
      Left            =   8835
      TabIndex        =   8
      Top             =   165
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      MousePointer    =   99
      Style           =   1
      Picture         =   "MainForm.frx":202D
      OverPicture     =   "MainForm.frx":2093
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTip         =   "最小化"
      ForeColor       =   0
      AutoSize        =   -1  'True
   End
   Begin VB.Timer tmrOK 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3210
      Top             =   4590
   End
   Begin VB.Timer tmrMusic 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2760
      Top             =   5040
   End
   Begin VB.Timer tmrLightFlash 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3660
      Top             =   4590
   End
   Begin VB.Frame fraMainButton 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   525
      TabIndex        =   24
      Top             =   5580
      Visible         =   0   'False
      Width           =   4545
      Begin Othello.FlatButton fltbtnChat 
         Height          =   345
         Left            =   3735
         TabIndex        =   25
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
         Caption         =   ""
         MousePointer    =   99
         Style           =   1
         Picture         =   "MainForm.frx":2244
         DownPicture     =   "MainForm.frx":2847
         OverPicture     =   "MainForm.frx":2E4C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTip         =   "显示/隐藏聊天窗口"
         ForeColor       =   0
         AutoSize        =   -1  'True
      End
      Begin Othello.FlatButton fltbtnTable 
         Height          =   345
         Left            =   2580
         TabIndex        =   26
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
         Caption         =   ""
         MousePointer    =   99
         Style           =   1
         Picture         =   "MainForm.frx":3400
         DownPicture     =   "MainForm.frx":3A16
         OverPicture     =   "MainForm.frx":4037
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTip         =   "显示/隐藏棋局列表窗口"
         ForeColor       =   0
         AutoSize        =   -1  'True
      End
      Begin Othello.FlatButton fltbtnOnline 
         Height          =   345
         Left            =   1440
         TabIndex        =   27
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
         Caption         =   ""
         MousePointer    =   99
         Style           =   1
         Picture         =   "MainForm.frx":45F8
         DownPicture     =   "MainForm.frx":4BCE
         OverPicture     =   "MainForm.frx":51AB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTip         =   "显示/隐藏在线用户窗口"
         ForeColor       =   0
         AutoSize        =   -1  'True
      End
      Begin Othello.FlatButton fltbtnMenu 
         Height          =   345
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         Caption         =   ""
         MousePointer    =   99
         Style           =   1
         Picture         =   "MainForm.frx":5597
         DownPicture     =   "MainForm.frx":5C0A
         OverPicture     =   "MainForm.frx":627D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTip         =   "游戏主菜单"
         ForeColor       =   0
         AutoSize        =   -1  'True
      End
   End
   Begin VB.Timer tmrFaceFlash 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4110
      Top             =   4590
   End
   Begin VB.Frame fraStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   1635
      Left            =   5715
      TabIndex        =   10
      Top             =   4275
      Visible         =   0   'False
      Width           =   3285
      Begin VB.Label lblTime1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   2355
         TabIndex        =   33
         Top             =   840
         Width           =   75
      End
      Begin VB.Image imgLight 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   0
         Top             =   1245
         Width           =   300
      End
      Begin VB.Image imgLight 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   0
         Top             =   600
         Width           =   300
      End
      Begin VB.Image imgLight 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   0
         Top             =   300
         Width           =   300
      End
      Begin VB.Image imgLight 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   300
      End
      Begin VB.Label lblTips 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请登陆或立即注册。"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   945
         TabIndex        =   20
         Top             =   1305
         Width           =   1620
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--:--"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   945
         TabIndex        =   19
         Top             =   990
         Width           =   450
      End
      Begin VB.Label lblTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "无棋局"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   945
         TabIndex        =   18
         Top             =   675
         Width           =   540
      End
      Begin VB.Label lblConnect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "无连接"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   945
         TabIndex        =   17
         Top             =   375
         Width           =   540
      End
      Begin VB.Label lblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未登录"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   945
         TabIndex        =   16
         Top             =   75
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "登录:"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   405
         TabIndex        =   15
         Top             =   75
         Width           =   450
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "棋局:"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   405
         TabIndex        =   14
         Top             =   675
         Width           =   450
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "连接:"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   405
         TabIndex        =   13
         Top             =   375
         Width           =   450
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "提示:"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   405
         TabIndex        =   12
         Top             =   1305
         Width           =   450
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "时间:"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   405
         TabIndex        =   11
         Top             =   990
         Width           =   450
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   2130
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":66FD
            Key             =   ""
            Object.Tag             =   "注册向导(&R)..."
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":67F9
            Key             =   ""
            Object.Tag             =   "登录(&L)..."
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6891
            Key             =   ""
            Object.Tag             =   "注销(&T)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6935
            Key             =   ""
            Object.Tag             =   "马上玩游戏(&P)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":69CD
            Key             =   ""
            Object.Tag             =   "公共聊天区(&U)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6A69
            Key             =   ""
            Object.Tag             =   "棋局(&B)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6BF1
            Key             =   ""
            Object.Tag             =   "查看资料(&V)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6C7D
            Key             =   ""
            Object.Tag             =   "修改资料(&I)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6D01
            Key             =   ""
            Object.Tag             =   "音乐(&M)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6D9D
            Key             =   ""
            Object.Tag             =   "播放(&P)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6DF5
            Key             =   ""
            Object.Tag             =   "停止(&S)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6E4D
            Key             =   ""
            Object.Tag             =   "暂停(&A)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6EA9
            Key             =   ""
            Object.Tag             =   "上一首(&E)"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6F09
            Key             =   ""
            Object.Tag             =   "下一首(&L)"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6F65
            Key             =   ""
            Object.Tag             =   "设置(&S)..."
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":7029
            Key             =   ""
            Object.Tag             =   "帮助(&H)..."
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":70ED
            Key             =   ""
            Object.Tag             =   "反馈(&F)"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":716D
            Key             =   ""
            Object.Tag             =   "改进意见(&S)..."
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":71F1
            Key             =   ""
            Object.Tag             =   "&Bug 报告..."
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":7275
            Key             =   ""
            Object.Tag             =   "技术支持(&U)..."
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":72F9
            Key             =   ""
            Object.Tag             =   "幸福家园论坛(&G)"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":7401
            Key             =   ""
            Object.Tag             =   "关于黑白棋.Net(&A)..."
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":7489
            Key             =   ""
            Object.Tag             =   "退出(&X)"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00E0E0E0&
      Height          =   1590
      Left            =   5580
      ScaleHeight     =   1590
      ScaleWidth      =   3285
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   3285
   End
   Begin MSWinsockLib.Winsock tcpGame 
      Left            =   4560
      Top             =   4590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpListen 
      Left            =   5010
      Top             =   4590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet GameInet 
      Left            =   915
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   2760
      Top             =   4590
   End
   Begin MSComctlLib.ImageList ilsFace 
      Left            =   1530
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   8421376
      _Version        =   393216
   End
   Begin VB.PictureBox MainPicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H00E0E0E0&
      Height          =   2325
      Left            =   255
      ScaleHeight     =   2325
      ScaleWidth      =   2670
      TabIndex        =   0
      Top             =   255
      Width           =   2670
   End
   Begin VB.Frame fraPlayer 
      BackColor       =   &H00000000&
      Caption         =   "黑方"
      ForeColor       =   &H00E0E0E0&
      Height          =   945
      Index           =   0
      Left            =   5610
      TabIndex        =   2
      Top             =   1890
      Visible         =   0   'False
      Width           =   3345
      Begin VB.PictureBox picPlayerFace 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   0
         Left            =   135
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   34
         Top             =   300
         Width           =   480
      End
      Begin Othello.FlatButton fltbtnStart 
         Height          =   300
         Index           =   0
         Left            =   2655
         TabIndex        =   31
         Top             =   195
         Visible         =   0   'False
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Caption         =   "开始"
         MousePointer    =   99
         Style           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HotColor        =   0
         EnableHot       =   -1  'True
         OverBorderColor =   8779006
         DownBorderColor =   13499390
         BorderColor     =   16777215
         OverBackColor   =   13499390
         DownBackColor   =   8779006
         BackColor       =   14737632
         ForeColor       =   0
      End
      Begin Othello.FlatButton fltbtnSitDown 
         Height          =   300
         Index           =   0
         Left            =   690
         TabIndex        =   29
         Top             =   195
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   529
         Caption         =   "坐在黑方(&B)"
         MousePointer    =   99
         Style           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HotColor        =   0
         EnableHot       =   -1  'True
         OverBorderColor =   13499390
         DownBorderColor =   13499390
         BorderColor     =   16777215
         OverBackColor   =   13499390
         DownBackColor   =   8779006
         BackColor       =   14737632
         ForeColor       =   0
      End
      Begin Othello.FlatButton fltbtnReadyStart 
         Height          =   300
         Index           =   0
         Left            =   690
         TabIndex        =   36
         Top             =   195
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   529
         Caption         =   "我准备好了！"
         MousePointer    =   99
         Style           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableHot       =   -1  'True
         ForeColor       =   0
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   735
         TabIndex        =   3
         Top             =   210
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblPlayerTips 
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   735
         TabIndex        =   22
         Top             =   615
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label lblChessNum 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   2655
         TabIndex        =   4
         Top             =   210
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Shape shpChessNum 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   300
         Index           =   0
         Left            =   2655
         Shape           =   2  'Oval
         Top             =   195
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Shape shpPlayerTips 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   285
         Index           =   0
         Left            =   690
         Top             =   570
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Shape shpPlayerName 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   300
         Index           =   0
         Left            =   690
         Top             =   195
         Visible         =   0   'False
         Width           =   1920
      End
   End
   Begin VB.Frame fraPlayer 
      BackColor       =   &H00000000&
      Caption         =   "白方"
      ForeColor       =   &H00E0E0E0&
      Height          =   945
      Index           =   1
      Left            =   5610
      TabIndex        =   5
      Top             =   3030
      Visible         =   0   'False
      Width           =   3345
      Begin VB.PictureBox picPlayerFace 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   1
         Left            =   135
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   35
         Top             =   300
         Width           =   480
      End
      Begin Othello.FlatButton fltbtnSitDown 
         Height          =   300
         Index           =   1
         Left            =   690
         TabIndex        =   30
         Top             =   195
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   529
         Caption         =   "坐在白方(&W)"
         MousePointer    =   99
         Style           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HotColor        =   0
         EnableHot       =   -1  'True
         OverBorderColor =   13499390
         DownBorderColor =   13499390
         BorderColor     =   16777215
         OverBackColor   =   13499390
         DownBackColor   =   8779006
         BackColor       =   14737632
         ForeColor       =   0
      End
      Begin Othello.FlatButton fltbtnStart 
         Height          =   300
         Index           =   1
         Left            =   2655
         TabIndex        =   32
         Top             =   195
         Visible         =   0   'False
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Caption         =   "开始"
         MousePointer    =   99
         Style           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HotColor        =   0
         EnableHot       =   -1  'True
         OverBorderColor =   8779006
         DownBorderColor =   13499390
         BorderColor     =   16777215
         OverBackColor   =   13499390
         DownBackColor   =   8779006
         BackColor       =   14737632
         ForeColor       =   0
      End
      Begin Othello.FlatButton fltbtnReadyStart 
         Height          =   300
         Index           =   1
         Left            =   690
         TabIndex        =   37
         Top             =   195
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   529
         Caption         =   "我准备好了！"
         MousePointer    =   99
         Style           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableHot       =   -1  'True
         ForeColor       =   0
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   735
         TabIndex        =   6
         Top             =   225
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblPlayerTips 
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   735
         TabIndex        =   23
         Top             =   615
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label lblChessNum 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   2655
         TabIndex        =   7
         Top             =   210
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Shape shpChessNum 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   2655
         Shape           =   2  'Oval
         Top             =   195
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Shape shpPlayerTips 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   285
         Index           =   1
         Left            =   690
         Top             =   570
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Shape shpPlayerName 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   690
         Top             =   195
         Visible         =   0   'False
         Width           =   1920
      End
   End
   Begin VB.Label lblFooter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Happy Family"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   8070
      TabIndex        =   21
      Top             =   5955
      Width           =   1260
   End
   Begin VB.Image imgBottomBorder 
      Enabled         =   0   'False
      Height          =   150
      Left            =   0
      Picture         =   "MainForm.frx":7525
      Top             =   6225
      Width           =   9525
   End
   Begin VB.Image imgRightBorder 
      Enabled         =   0   'False
      Height          =   6375
      Left            =   9375
      Picture         =   "MainForm.frx":778E
      Top             =   0
      Width           =   150
   End
   Begin VB.Image imgTopBorder 
      Enabled         =   0   'False
      Height          =   150
      Left            =   0
      Picture         =   "MainForm.frx":7A33
      Top             =   0
      Width           =   9525
   End
   Begin VB.Image imgLeftBorder 
      Enabled         =   0   'False
      Height          =   6375
      Left            =   0
      Picture         =   "MainForm.frx":7CCF
      Top             =   0
      Width           =   150
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu S0 
         Caption         =   "!"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "注册向导(&R)..."
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "登录(&L)..."
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "注销(&T)"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayGame 
         Caption         =   "马上玩游戏(&P)"
      End
      Begin VB.Menu mnuPublicChat 
         Caption         =   "公共聊天区(&U)"
      End
      Begin VB.Menu S9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTable 
         Caption         =   "棋局(&B)"
         Begin VB.Menu mnuTableCreate 
            Caption         =   "创建棋局(&C)..."
         End
         Begin VB.Menu mnuTableExit 
            Caption         =   "退出棋局(&E)"
         End
         Begin VB.Menu mnuTableCancel 
            Caption         =   "取消本局(&N)"
         End
         Begin VB.Menu mnuTableDraw 
            Caption         =   "请求和棋(&D)..."
         End
         Begin VB.Menu mnuTableLose 
            Caption         =   "认输(&L)"
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "查看资料(&V)"
         Begin VB.Menu mnuViewMyInfo 
            Caption         =   "查看个人资料(&I)..."
         End
         Begin VB.Menu mnuViewInfo 
            Caption         =   "查看对手资料(&V)..."
         End
         Begin VB.Menu mnuViewTableInfo 
            Caption         =   "查看棋局信息(&W)..."
         End
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "修改资料(&I)"
         Begin VB.Menu mnuEditInfo 
            Caption         =   "修改个人资料(&T)..."
         End
         Begin VB.Menu mnuEditTableInfo 
            Caption         =   "修改棋局信息(&B)..."
         End
      End
      Begin VB.Menu S4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMusic 
         Caption         =   "音乐(&M)"
         Begin VB.Menu mnuMusicPlay 
            Caption         =   "播放(&P)"
         End
         Begin VB.Menu mnuMusicStop 
            Caption         =   "停止(&S)"
         End
         Begin VB.Menu mnuMusicPause 
            Caption         =   "暂停(&A)"
         End
         Begin VB.Menu mnuMusicPrevious 
            Caption         =   "上一首(&E)"
         End
         Begin VB.Menu mnuMusicLast 
            Caption         =   "下一首(&L)"
         End
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "设置(&S)..."
      End
      Begin VB.Menu S10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "帮助(&H)..."
      End
      Begin VB.Menu mnuFeedback 
         Caption         =   "反馈(&F)"
         Begin VB.Menu mnuSuggestion 
            Caption         =   "改进意见(&S)..."
         End
         Begin VB.Menu mnuBugReport 
            Caption         =   "&Bug 报告..."
         End
         Begin VB.Menu mnuSupport 
            Caption         =   "技术支持(&U)..."
         End
      End
      Begin VB.Menu mnuHomePage 
         Caption         =   "幸福家园论坛(&G)"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于黑白棋.Net(&A)..."
      End
      Begin VB.Menu S6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "隐藏(Ctrl+`)"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 主窗体

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const RSP_SIMPLE_SERVICE = 1
Private Const RSP_UNREGISTER_SERVICE = 0
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
'Private Const GW_OWNER = 4

Private WithEvents Table As CTable      ' 棋局对象
Attribute Table.VB_VarHelpID = -1

Dim MouseX As Single
Dim MouseY As Single
Dim FormMoving As Boolean

Dim Starting As Boolean

Dim mblnServerOK As Boolean

Dim mblnSitDown(PLY_NUMBER) As Boolean
Dim mblnStart(PLY_NUMBER) As Boolean
Dim mblnReadyStart(PLY_NUMBER) As Boolean

Dim mlngPlayer(PLY_NUMBER) As Long

Dim mblnBusy As Boolean

' 重试步骤
Dim mlngTryJoin As Long
Dim mstrTryLANIP As String
Dim mstrTryIP As String
Dim mlngTryPort As Long

' 在局域网中创建棋局的提示
Dim mblnLANMessage As Boolean

' 游戏焦点（也就是当前是否轮到你下）
Dim mblnGameFocus As Boolean

' 秒增量器，从1到15则秒减1
Dim mlngSecond As Long
' 游戏总时间
Dim mlngGameSecond As Long

' 电脑秒增量器，从1到15则秒减1（用于单机模式）
Dim mlngComputerSecond As Long
' 电脑游戏总时间（用于单机模式）
Dim mlngComputerGameSecond As Long

' 灯位置
Dim mlngLightWhere As Long

' 确定是否可以切断连接
Dim mblnAgreeDisconnect As Boolean

' 存储远程地址与端口
Dim mstrRemoteHostIP As String
Dim mlngRemotePort As Long
Dim mlngSocks5Status As Long

Dim mblnSocks5Connected As Boolean

' 是否在隐藏状态，用于 BossKey 功能。
Dim mblnHide As Boolean

' 游戏步数（用于单机模式）
Dim mlngStep As Long


Private Sub Form_Activate()
    On Error Resume Next

    If Starting Then
        Me.Enabled = False
        Call DisplayPic
        ' 判断是否是单机模式
        If gblnOfflineMode Then
            ' 单机模式
            lblTips.Caption = LoadString(259) '"单机模式，请创建棋局！"
        Else
            ' 网络模式需要登陆
            Call frmLogin.ShowEx
        End If
        Me.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim Temp As String

    On Error Resume Next

    Starting = True
    Me.WindowState = vbNormal

    For i = 1 To glngSave_FaceNumber
        Temp = gstrAppPath & "\Images\Image" & CStr(i) & ".gif"
        If FileExisted(Temp) Then
            Call ilsFace.ListImages.Add(i, , LoadPicture(Temp))
        Else
            Call ilsFace.ListImages.Add(i, , LoadResPicture("Face", vbResBitmap))
        End If
    Next i
    ilsFace.Tag = CStr(glngSave_FaceNumber)

    ' 控件初始化开始

    If Not OTHELLO_DEBUG Then
        ' 菜单初始化，注意退出时清除。
        Call mCoolMenu.Install(Me.hWnd, ilsMenu, frmResource.imgMenuSide.Picture)
        Call mCoolMenu.SelectColor(Me.hWnd, CLR_SELECT_MENU)
        ''Call mCoolMenu.ComplexChecks(Me.hWnd, False)
    End If

    ' 设置热键
    Call RegisterHotKey(Me.hWnd, HOTKEY_ID, MOD_CONTROL, KEY_HOTKEY)

    MainPicture.Width = 5100
    MainPicture.Height = 5100

    Set fltbtnMenu.MouseIcon = HandCursor
    Set fltbtnOnline.MouseIcon = HandCursor
    Set fltbtnTable.MouseIcon = HandCursor
    Set fltbtnChat.MouseIcon = HandCursor
    Set fltbtnMin.MouseIcon = HandCursor
    Set fltbtnExit.MouseIcon = HandCursor
    Set fltbtnSitDown(0).MouseIcon = HandCursor
    Set fltbtnSitDown(1).MouseIcon = HandCursor
    Set fltbtnStart(0).MouseIcon = HandCursor
    Set fltbtnStart(1).MouseIcon = HandCursor

    For i = 0 To 2
        Call SetLight(i, lsLightRed)
    Next i

    Call tcpListen.Close
    Call tcpGame.Close

    ' 控件初始化结束

    Set gBackgroundMusic = New Mmedia
    Set gSoundEffects = New Mmedia

    Call CreateRectForm(Me.hWnd, GetPixelX(Width), GetPixelY(Height))

    If gwifSave_MainWindow.Center Then
        Call Me.Move((Screen.Width - Width) \ 2, (Screen.Height * 0.9 - Height) \ 2)
    Else
        Call Me.Move(gwifSave_MainWindow.Left, gwifSave_MainWindow.Top)
    End If

    Call Sleep(1000)
    Call Unload(frmSplash)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next

    If UnloadMode = vbFormCode Or UnloadMode = vbFormControlMenu Then
        If gblnGameStart And (Not gblnOfflineMode) Then
            Cancel = True
            Call frmRequest.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name, gYourUserInfo.Face, CMD_RequestExitGame, , True)
            Exit Sub
        Else
            Me.WindowState = vbNormal
            If MessageBox(Me.hWnd, LoadString(129), vbQuestion Or vbYesNo) <> vbYes Then
                Cancel = True
                Exit Sub
            End If
            Call ExitTable(True)
            Call Disconnect
            Call Logout(True)
        End If
    End If

    'If Not (Table Is Nothing) Then
    '    Set Table = Nothing
    'End If
    'Call mCoolMenu.Uninstall(Me.hWnd)
    'Call UnregisterHotKey(Me.hWnd, HOTKEY_ID)
    'Call Quit
    'Debug.Print "1",
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    If Not (Table Is Nothing) Then
        Set Table = Nothing
    End If

    If Not OTHELLO_DEBUG Then
        Call mCoolMenu.Uninstall(Me.hWnd)
    End If
    
    Call UnregisterHotKey(Me.hWnd, HOTKEY_ID)
    Call DeleteRgn
    Call Quit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    If (Shift And vbCtrlMask) = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyL
                Call mnuLogin_Click
            Case vbKeyT
                Call mnuLogout_Click
            Case vbKeyE
                Call mnuTableCreate_Click
            Case vbKeyP
                Call mnuPlayGame_Click
            Case vbKeyG
                Call mnuHomePage_Click
            Case vbKeyS
                Call mnuOption_Click
            Case vbKeyX
                Call mnuExit_Click
        End Select
        KeyCode = 0
        Shift = 0
    ElseIf KeyCode = 93 Or KeyCode = vbKeyEscape Then
        Call Me.PopupMenu(mnuPop, , 0, 0)
        KeyCode = 0
    End If
End Sub

Private Sub fltbtnSitDown_Click(Index As Integer, Button As Integer)
    Dim intPartner As Integer
    Dim Face As Long

    On Error Resume Next

    ' 隐藏“坐下”按钮。
    fltbtnSitDown(0).Visible = False
    fltbtnSitDown(1).Visible = False
    ' 设置你选择了哪一方，并且标记你已经坐下。
    mlngPlayer(PLY_ME) = Index + 1
    mblnSitDown(Index) = True
    shpPlayerName(Index).Visible = True

    intPartner = ToPartner(Index, 1)

    If gblnOfflineMode Then
        ' 人脑设置
        mblnStart(GetPlayerIndex(PLY_ME)) = True

        shpChessNum(Index).Visible = True
        Call DisplayFace(Index, glngSave_OfflineFace)
        Call SetLabel(lblChessNum(Index), "00", "", True)
        Call SetLabel(lblPlayerName(Index), "人脑", "", True)
        lblPlayerTips(Index).Caption = ""

        ' 电脑设置
        mlngPlayer(PLY_YOU) = intPartner + 1
        mblnSitDown(intPartner) = True
        mblnStart(GetPlayerIndex(PLY_YOU)) = True
        ' 随机电脑头像，但不和人脑头像相同
        Do
            Face = Int((Val(MainForm.ilsFace.Tag)) * Rnd + 1)
        Loop While Face = glngSave_OfflineFace
        shpPlayerName(intPartner).Visible = True
        shpChessNum(intPartner).Visible = True
        Call DisplayFace(intPartner, Face)
        Call SetLabel(lblChessNum(intPartner), "00", "", True)
        Call SetLabel(lblPlayerName(intPartner), "电脑", "", True)
        lblPlayerTips(intPartner).Caption = ""

        ' 开始单机模式棋局
        Call StartOfflineGame
    Else
    
        If gblnConnect Then
            ' 如果程序已连接，则发送给对方你的选择。
            Call SendCommand(CMD_SitDown & CStr(mlngPlayer(PLY_ME)) & "|" & gMyUserInfo.UserName & "|" & CStr(gMyUserInfo.Face) & "|" & gMyUserInfo.Name)
        End If
    
        ' 单击“坐下”按钮以后进行的操作。
        ' 设置你的名字和头像。
        Call SetLabel(lblPlayerName(Index), GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name), "", True)
    
        Call DisplayFace(Index, gMyUserInfo.Face)

        If CheckArray(mblnSitDown()) Then
            ' 如果双方都已做好，则显示你的“开始”按钮，
            ' 并显示对方的“开始”提示。
            fltbtnStart(GetPlayerIndex(PLY_ME)).Visible = True
            Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_ME)), LoadString(115), "", True)
            Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_YOU)), _
                          LoadString(118), LoadString(119), _
                          True)
        Else
            ' 否则只显示对方的“等待”提示。
            'Call AlignControl(lblPlayerTips(intPartner), lblPlayerName(intPartner))
            lblPlayerTips(intPartner).Visible = False
            shpPlayerTips(intPartner).Visible = False
            If gblnConnect Then
                ' 如果已连接，则显示对方的“等待坐下”提示。
                Call SetLabel(lblPlayerTips(Index), _
                              LoadString(120), LoadString(121), _
                              True)
            Else
                ' 无连接，则显示对方的“等待加入”提示。
                Call SetLabel(lblPlayerTips(Index), LoadString(122), LoadString(123), True)
            End If
        End If
    End If
End Sub

Private Sub fltbtnStart_Click(Index As Integer, Button As Integer)
    On Error Resume Next

    mblnStart(GetPlayerIndex(PLY_ME)) = True
    Call SendCommand(CMD_GameStart)
    fltbtnStart(Index).Visible = False
    shpChessNum(Index).Visible = True
    Call SetLabel(lblChessNum(Index), "00", "", True)

    lblPlayerTips(Index).Caption = ""
    If CheckArray(mblnStart()) Then
        ' 游戏真正开始。
        If Not GameStart() Then
            Call SendCommand(CMD_AgainStart)
            If mblnStart(GetPlayerIndex(PLY_ME)) Then
                Call AgainStart
            End If
            Call MessageBox(Me.hWnd, LoadString(130), vbExclamation, LoadString(177))
            ' 注意：这里可能需要复位游戏和发送一些消息给对方
        End If
    End If
End Sub

Private Sub fltbtnReadyStart_Click(Index As Integer, Button As Integer)
    On Error Resume Next

    mblnReadyStart(GetPlayerIndex(PLY_ME)) = True
    Call SendCommand(CMD_GameReadyStart)
    fltbtnReadyStart(Index).Visible = False
    lblPlayerTips(Index).Caption = LoadString(188)

    If CheckArray(mblnReadyStart()) Then
        Call AgainStart
    End If
End Sub

Private Sub fraPlayer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, fraPlayer(Index).Left + X, fraPlayer(Index).Top + Y)
End Sub

Private Sub fraPlayer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseUp(Button, Shift, fraPlayer(Index).Left + X, fraPlayer(Index).Top + Y)
End Sub

Private Sub lblFooter_DblClick()
    Call mnuAbout_Click
End Sub

Private Sub lblPlayerName_DblClick(Index As Integer)
    On Error Resume Next

    If Not gblnLogin Then Exit Sub

    If Index = GetPlayerIndex(PLY_ME) And mblnSitDown(GetPlayerIndex(PLY_ME)) Then
        Call frmUserInfo.ShowEx(gMyUserInfo.UserName, gMyUserInfo.Name)
    ElseIf Index = GetPlayerIndex(PLY_YOU) And mblnSitDown(GetPlayerIndex(PLY_YOU)) Then
        Call frmUserInfo.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name)
    End If
End Sub

Private Sub mnuEditTableInfo_Click()
    If (Not gblnLogin) Or (Not gblnCreator) Or gblnGameStart Then Exit Sub

    Call frmTableInfo.ShowEx(gMainTableInfo.TableName, True)
End Sub

Private Sub mnuTableCancel_Click()
    If gblnGameStart Then
        If gblnOfflineMode Then
            Call ExitTable(False)
            Call Disconnect
        Else
            Call frmRequest.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name, gYourUserInfo.Face, CMD_RequestCancelGame, , True)
        End If
    End If
End Sub

Private Sub mnuHelp_Click()
    On Error Resume Next

    If FileExisted(gstrAppPath & "\Othello.chm") Then
        Call ShellExecute(Me.hWnd, "open", gstrAppPath & "\Othello.chm", 0, 0, 1)
    Else
        mnuHelp.Enabled = False
    End If
End Sub

Private Sub mnuHide_Click()
    Call BossKey
End Sub

Private Sub mnuTableLose_Click()
    Dim mbrMsgBox As VbMsgBoxResult

    If gblnGameStart Then
        mbrMsgBox = MessageBox(Me.hWnd, LoadString(131), vbQuestion Or vbYesNo, LoadString(178))
        If mbrMsgBox <> vbYes Then Exit Sub
        ' 如果游戏还没有结束
        If gblnGameStart Then
            If gblnOfflineMode Then
                Call OfflineGameOver(True)
            Else
                Call SendCommand(CMD_GameOver & GAME_WIN, True)
                Call GameOver(GAME_LOSE)
            End If
        End If
    End If
End Sub

Private Sub mnuMusicLast_Click()
    Dim LastStatus As String

    On Error Resume Next

    If glngSave_PlayListPosition >= glngSave_PlayListNumber Then _
        glngSave_PlayListPosition = 0

    glngSave_PlayListPosition = glngSave_PlayListPosition + 1
    LastStatus = gBackgroundMusic.Status
    If LastStatus <> "stopped" Then
        Call gBackgroundMusic.mmStop
        Call gBackgroundMusic.mmClose
    End If
    If LastStatus = "playing" Or LastStatus = "paused" Then
        Call mnuMusicPlay_Click
    End If
End Sub

Private Sub mnuMusicPause_Click()
    On Error Resume Next

    Call gBackgroundMusic.mmPause
End Sub

Private Sub mnuMusicPlay_Click()
    On Error Resume Next

    If glngSave_PlayListPosition > glngSave_PlayListNumber Then _
        glngSave_PlayListPosition = 1

    If gstrSave_PlayListName(glngSave_PlayListPosition) = "" Then Exit Sub

    If gBackgroundMusic.Status <> "paused" Then
        Call gBackgroundMusic.mmOpen(gstrSave_PlayListName(glngSave_PlayListPosition))
    End If

    Call gSoundEffects.mmStop
    Call gBackgroundMusic.mmPlay
    tmrMusic.Enabled = True
End Sub

Private Sub mnuMusicPrevious_Click()
    Dim LastStatus As String

    On Error Resume Next

    If glngSave_PlayListPosition < 2 Then _
        glngSave_PlayListPosition = glngSave_PlayListNumber + 1

    glngSave_PlayListPosition = glngSave_PlayListPosition - 1
    LastStatus = gBackgroundMusic.Status
    If LastStatus <> "stopped" Then
        Call gBackgroundMusic.mmStop
        Call gBackgroundMusic.mmClose
    End If
    If LastStatus = "playing" Or LastStatus = "paused" Then
        Call mnuMusicPlay_Click
    End If
End Sub

Private Sub mnuMusicStop_Click()
    On Error Resume Next

    tmrMusic.Enabled = False
    Call gSoundEffects.mmStop
    Call gBackgroundMusic.mmStop
End Sub

Private Sub mnuPlayGame_Click()
    ' 如果已加入，则返回。
    If (Not gblnLogin) Or gblnConnect Or gblnCreator Then Exit Sub

    ' 自动加入棋局。
    If Not frmTable.AutoJoin() Then
        Call MessageBox(Me.hWnd, LoadString(132), vbExclamation, LoadString(181))
    End If
End Sub

Private Sub mnuPublicChat_Click()
    Call frmPublicChat.ShowEx
End Sub

Private Sub mnuTableDraw_Click()
    If gblnGameStart And Not gblnOfflineMode Then
        Call frmRequest.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name, gYourUserInfo.Face, CMD_RequestDrawGame, , True)
    End If
End Sub

Private Sub mnuBugReport_Click()
    On Error Resume Next

    Call ShellExecute(Me.hWnd, "open", "mailto:" & HAPPY_FAMILY_BBS_MAIL & "?subject=《黑白棋.Net》Bug 报告 -- 版本: " & App.Major & "." & App.Minor & "." & App.Revision, 0, 0, 1)
End Sub

Private Sub mnuSuggestion_Click()
    On Error Resume Next

    Call ShellExecute(Me.hWnd, "open", "mailto:" & HAPPY_FAMILY_BBS_MAIL & "?subject=《黑白棋.Net》改进意见 -- 版本: " & App.Major & "." & App.Minor & "." & App.Revision, 0, 0, 1)
End Sub

Private Sub mnuSupport_Click()
    On Error Resume Next

    Call ShellExecute(Me.hWnd, "open", "mailto:" & HAPPY_FAMILY_BBS_MAIL & "?subject=《黑白棋.Net》技术支持 -- 版本: " & App.Major & "." & App.Minor & "." & App.Revision, 0, 0, 1)
End Sub

Private Sub mnuTableExit_Click()
    If ((Not gblnLogin) Or (Not gblnCreator And Not gblnConnect)) And (Not gblnOfflineMode Or Not gblnCreator) Then Exit Sub

    If gblnGameStart And Not gblnOfflineMode Then
        Call frmRequest.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name, gYourUserInfo.Face, CMD_RequestExitTable, , True)
    Else
        Dim mbrMsgBox As VbMsgBoxResult
        mbrMsgBox = MessageBox(Me.hWnd, LoadString(133), vbQuestion Or vbYesNo, LoadString(178))
        If mbrMsgBox <> vbYes Then Exit Sub
        Call ExitTable(False)
        Call Disconnect
    End If
End Sub

Private Sub mnuViewMyInfo_Click()
    If Not gblnLogin Then Exit Sub

    Call frmUserInfo.ShowEx(gMyUserInfo.UserName, gMyUserInfo.Name)
End Sub

Private Sub mnuViewTableInfo_Click()
    If (Not gblnLogin) Or (Not gblnCreator And Not gblnConnect) Then Exit Sub

    Call frmTableInfo.ShowEx(gMainTableInfo.TableName, False)
End Sub

Private Sub tmrFaceFlash_Timer()
    On Error Resume Next

    If mblnGameFocus Then
        picPlayerFace(GetPlayerIndex(PLY_ME)).Visible = Not picPlayerFace(GetPlayerIndex(PLY_ME)).Visible
    Else
        picPlayerFace(GetPlayerIndex(PLY_YOU)).Visible = Not picPlayerFace(GetPlayerIndex(PLY_YOU)).Visible
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If (Button And vbLeftButton) = vbLeftButton Then
        If gblnMenuDisplay Then
            Call ReleaseCapture
            Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        Else
            Call SetCapture(Me.hWnd)
            MouseX = X: MouseY = Y
            FormMoving = True
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If (Button And vbLeftButton) = vbLeftButton And FormMoving And GetCapture() = Me.hWnd And Not gblnMenuDisplay Then
        Call Me.Move(X - MouseX + Me.Left, Y - MouseY + Me.Top)
        Call frmOnline.FormMove(False)
        Call frmTable.FormMove(False)
        Call frmChat.FormMove(False)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If (Button And vbLeftButton) = vbLeftButton And FormMoving Then
        Call ReleaseCapture
        FormMoving = False
    End If
    If (Button And vbRightButton) = vbRightButton Then
        gblnMenuDisplay = True
        Call Me.PopupMenu(mnuPop)
        gblnMenuDisplay = False
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        Call frmChat.FormNormal
        Call frmTable.FormNormal
        Call frmOnline.FormNormal
        Call frmEditInfo.FormNormal
        Call frmTableInfo.FormNormal
        Call frmUserInfo.FormNormal
        Call frmPublicChat.FormNormal
        Call frmRequest.FormNormal
        Call frmProgress.FormNormal
    ElseIf Me.WindowState = vbMinimized Then
        Call frmOnline.FormMinimize
        Call frmTable.FormMinimize
        Call frmChat.FormMinimize
        Call frmEditInfo.FormMinimize
        Call frmTableInfo.FormMinimize
        Call frmUserInfo.FormMinimize
        Call frmPublicChat.FormMinimize
        Call frmRequest.FormMinimize
        Call frmProgress.FormMinimize
    End If
End Sub

Private Sub fraMainButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, fraMainButton.Left + X, fraMainButton.Top + Y)
End Sub

Private Sub fraMainButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseUp(Button, Shift, fraMainButton.Left + X, fraMainButton.Top + Y)
End Sub

Private Sub fraStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, fraStatus.Left + X, fraStatus.Top + Y)
End Sub

Private Sub fraStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseUp(Button, Shift, fraStatus.Left + X, fraStatus.Top + Y)
End Sub

Private Sub GameInet_StateChanged(ByVal State As Integer)
    Select Case State
        Case icResponseReceived
            mblnServerOK = True
        Case icError  '11
            mblnServerOK = False
    End Select
End Sub

Private Sub tmrGame_Timer()
    On Error Resume Next

    ' 游戏计时
    mlngSecond = mlngSecond + 1
    If mlngSecond > PER_SECOND Then
        mlngSecond = 0
        mlngGameSecond = mlngGameSecond - 1
        lblTime.Caption = GetTime(mlngGameSecond)
        lblTime1.Caption = lblTime.Caption
        If mlngGameSecond < 1 Then
            tmrGame.Enabled = False
            Call SendCommand(CMD_TimeOver)
            Call TimeOver(GAME_LOSE)
        End If
    End If
End Sub

Private Sub MainPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Col As Integer
    Dim Row As Integer
    Dim XX As Single
    Dim YY As Single

    On Error Resume Next

    If gblnGameStart And mblnGameFocus And Button = vbLeftButton Then
        Col = Table.GetCol(X)
        Row = Table.GetRow(Y)
        If Col < 0 Or Col > 7 Or Row < 0 Or Row > 7 Then Exit Sub
        If Table.IsDown(Col, Row, mlngPlayer(PLY_ME)) Then
            XX = Table.GetX(Col)
            YY = Table.GetY(Row)
            Call MainPicture.PaintPicture(SelectDown, XX, YY)
        End If
    End If
    If Not gblnGameStart Then
        Call Form_MouseDown(Button, Shift, MainPicture.Left + X, MainPicture.Top + Y)
    End If
End Sub

Private Sub MainPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Col As Integer
    Dim Row As Integer
    Dim Num As Integer

    On Error Resume Next

    If gblnGameStart And mblnGameFocus And Button = vbLeftButton Then
        Col = Table.GetCol(X)
        Row = Table.GetRow(Y)
        If Col < 0 Or Col > 7 Or Row < 0 Or Row > 7 Then Exit Sub
        If Table.IsDown(Col, Row, mlngPlayer(PLY_ME)) Then
            If Not gblnOfflineMode Then
                ' 发送落子数据
                Call SendCommand(CMD_DownChessMan & Chr(SetDownManCommand(mlngPlayer(PLY_ME), Col, Row)), True)
            End If

            mblnGameFocus = False
            ' 播放音效
            Call PlaySoundEffects(SOUND_DOWN_MAN, gstrSave_SoundValue(SOUND_DOWN_MAN))

            LastDown.Col = Col: LastDown.Row = Row: LastMan = mlngPlayer(PLY_ME)
            Call SetMousePointer(psPointer)
            Call Table.DownMan(Col, Row, mlngPlayer(PLY_ME))
            tmrFaceFlash.Enabled = False
            picPlayerFace(0).Visible = True
            picPlayerFace(1).Visible = True

            ' 对方接收成功才提示对方落子

            ' 重画棋盘
            Call DrawTable
            ' 停止记时
            If gMainTableInfo.Timer > 0 Then tmrGame.Enabled = False

            If gblnOfflineMode Then
                mlngStep = mlngStep + 1

                ' 这里需要判断电脑是否可落子
                If Not Table.CanDown(mlngPlayer(PLY_YOU)) Then
                    If Not Table.CanDown(mlngPlayer(PLY_ME)) Then
                        Call OfflineGameOver(False)
                    Else
                        mblnGameFocus = True
                        Call DrawTable
                    End If
                Else
                    ' 电脑开始计算，计算完成将发送 WM_THINKEND 消息。
                    Call DisplayCurrentPlayer(mlngPlayer(PLY_YOU))
                    If Not ComputerThink() Then
                        mblnGameFocus = False
                        gblnGameStart = False
                    End If
                End If
            End If
        Else
            ' 播放音效
            Call PlaySoundEffects(SOUND_DOWN_ERROR, gstrSave_SoundValue(SOUND_DOWN_ERROR))
        End If
    End If
End Sub

Private Sub MainPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Col As Integer
    Dim Row As Integer
    Dim NoHave As Boolean

    On Error Resume Next

    If gblnGameStart And mblnGameFocus Then
        NoHave = True
        Col = Table.GetCol(X)
        Row = Table.GetRow(Y)

        If Col >= 0 And Col <= 7 And Row >= 0 And Row <= 7 Then
            If Current.Col <> Col Or Current.Row <> Row Then
                If Current.Col >= 0 And Current.Row >= 0 Then
                    If Table.GetMan(Current.Col, Current.Row) = T_NONE Then
                        Call LostTableFocus(Current.Col, Current.Row)
                        Call ReleaseCapture
                    End If
                End If
                If Table.GetMan(Col, Row) = T_NONE And Table.IsDown(Col, Row, mlngPlayer(PLY_ME)) Then
                    Call SetCapture(MainPicture.hWnd)
                    Call SetTableFocus(Col, Row, Button)
                End If
                Current.Col = Col: Current.Row = Row
            End If
            NoHave = False
        End If
        
        If NoHave And Current.Col >= 0 And Current.Col <= 7 And Current.Row >= 0 And Current.Row <= 7 Then
            If Table.GetMan(Current.Col, Current.Row) = T_NONE Then
                Call LostTableFocus(Current.Col, Current.Row)
                Call ReleaseCapture
            End If
            Current.Col = -1: Current.Row = -1
        End If
    End If
End Sub

Private Sub mnuEditInfo_Click()
    If Not gblnLogin Then Exit Sub

    Call frmEditInfo.ShowEx
End Sub

Private Sub mnuHomePage_Click()
    On Error Resume Next

    Call ShellExecute(Me.hWnd, "open", HAPPY_FAMILY_BBS, 0, 0, 1)
End Sub

Private Sub mnuTableCreate_Click()
    If (Not gblnOfflineMode Or gblnGameStart) And ((Not gblnLogin) Or gblnCreator Or gblnConnect) Then Exit Sub

    If gblnOfflineMode Then
        gMainTableInfo.Timer = 0
        Call CreateOfflineTable
    Else
        Call frmCreateTable.Show(vbModal)
    End If
End Sub

Private Sub mnuViewInfo_Click()
    If (Not gblnLogin) Or (Not gblnConnect) Then Exit Sub

    Call frmUserInfo.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name)
End Sub

Private Sub fltbtnChat_ValueChange(Value As Boolean)
    Value = Not Value
    If Value Then
        Call frmChat.ShowEx(Me)
    Else
        Call frmChat.HideEx
    End If
End Sub

Private Sub fltbtnExit_Click(Button As Integer)
    Call Unload(Me)
End Sub

Private Sub fltbtnMenu_Click(Button As Integer)
    On Error GoTo ErrHandler

    Call Me.PopupMenu(mnuPop)

    Exit Sub

ErrHandler:
    Call mCoolMenu.Uninstall(Me.hWnd)
    Call Me.PopupMenu(mnuPop)
End Sub

Private Sub fltbtnMin_Click(Button As Integer)
    Me.WindowState = vbMinimized
End Sub

Private Sub fltbtnTable_ValueChange(Value As Boolean)
    Value = Not Value
    If Value Then
        Call frmTable.ShowEx(Me)
    Else
        Call frmTable.HideEx
    End If
End Sub

Private Sub tmrLightFlash_Timer()
    On Error Resume Next

    imgLight(mlngLightWhere).Visible = Not imgLight(mlngLightWhere).Visible
End Sub

Private Sub tmrMusic_Timer()
    On Error Resume Next

    If gBackgroundMusic.Status = "stopped" Or gBackgroundMusic.Status = "" Then
        glngSave_PlayListPosition = glngSave_PlayListPosition + 1
        Call mnuMusicPlay_Click
    End If
End Sub

Private Sub tmrOK_Timer()
    ' 如果准备重新发送的命令为空，则直接返回
    If tmrOK.Tag = "" Then Exit Sub
    Call SendCommand(tmrOK.Tag)
End Sub

Private Sub fltbtnOnline_ValueChange(Value As Boolean)
    Value = Not Value
    If Value Then
        Call frmOnline.ShowEx(Me)
    Else
        frmOnline.HideEx
    End If
End Sub

' 对方关闭连接
Private Sub tcpGame_Close()
    On Error Resume Next

    If mblnAgreeDisconnect And gblnGameStart Then
        ' 正常离开
        Call CancelGame
    ElseIf gblnGameStart Then
        ' 对方断线
        Call CancelGame
        Call MessageBox(Me.hWnd, LoadString(134), vbInformation, LoadString(180))
    End If

    If gblnCreator Then
        Dim mbrMsgBox As VbMsgBoxResult
        Dim TableInfo As tagTableInfo

        TableInfo = gMainTableInfo
        If gblnConnect Then
            Call ExitTable(False)
            Call Disconnect
            mbrMsgBox = MessageBox(Me.hWnd, LoadString(135), vbQuestion Or vbYesNo, LoadString(178))
            If mbrMsgBox = vbYes Then
                gMainTableInfo = TableInfo
                Call ReadyCreateTable
            End If
        Else
            Call tcpGame.Close
            Call StartListen(glngSave_GamePort)
        End If
    Else
        Call ExitTable(False)
        If gblnConnect Then
            Call Disconnect
            Call MessageBox(Me.hWnd, LoadString(136), vbInformation, LoadString(180))
        Else
            Call Disconnect
        End If
    End If
End Sub

' 客户程序连接完毕，在客户端产生此事件。
Private Sub tcpGame_Connect()
    On Error Resume Next

    If gblnSave_UseProxy And (Not mblnSocks5Connected) Then
        Dim Buffers(3) As Byte

        Buffers(0) = 5  ' 代理版本
        Buffers(1) = 2  ' 方法数量
        Buffers(2) = 0  ' 方法: 无用户检验
        Buffers(3) = 2  ' 方法: 用户检验

        mlngSocks5Status = 2
        Call tcpGame.SendData(Buffers())
    Else
        ' 发送加入请求
        If gMainTableInfo.TableType = TABLE_LIMIT Then
            lblTips.Caption = LoadString(189)
        End If
        Call SendCommand(CMD_Request & CMD_RequestJoin & gMainTableInfo.TableName & "|" & gMyUserInfo.UserName & "|" & gMyUserInfo.Name)
    End If
End Sub

Private Sub tcpGame_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String

    On Error Resume Next

    If bytesTotal < 1 Then Exit Sub

    Call tcpGame.GetData(strData, vbString)

    If gblnSave_UseProxy And (Not mblnSocks5Connected) And (Not gblnCreator) Then
        Select Case mlngSocks5Status
            Case 2
                ' 需要验证，则发送验证信息
                If GetData(strData, 2) = 2 Then
                    Dim Anthreq As String

                    Anthreq = Anthreq & Chr(5)
                    Anthreq = Anthreq & Chr(Len(gstrSave_Socks5Username))
                    Anthreq = Anthreq & gstrSave_Socks5Username
                    Anthreq = Anthreq & Chr(Len(Decipher(gstrLocalPassword, gstrSave_Socks5Password)))
                    Anthreq = Anthreq & Decipher(gstrLocalPassword, gstrSave_Socks5Password)

                    mlngSocks5Status = 3
                    Call tcpGame.SendData(Anthreq)
                    Exit Sub
                End If

                mlngSocks5Status = 4
                Call SendSocks5Request(mstrRemoteHostIP, mlngRemotePort)

            Case 3
                If GetData(strData, 2) <> 0 Then
                    Call tcpGame.Close
                    lblTips.Caption = LoadString(207)
                    Call MessageBox(Me.hWnd, LoadString(164), vbExclamation, LoadString(186))
                    Exit Sub
                End If

                mlngSocks5Status = 4
                Call SendSocks5Request(mstrRemoteHostIP, mlngRemotePort)

            Case 4
                If GetData(strData, 2) <> 0 Then
                    Call tcpGame.Close
                    If mlngTryJoin = 1 Then
                        Call ReadyTryJoin(mstrTryLANIP, mstrTryIP, mlngTryPort)
                        Exit Sub
                    End If
                    lblTips.Caption = LoadString(207)
                    Call MessageBox(Me.hWnd, LoadString(165), vbExclamation, LoadString(186))
                    Exit Sub
                Else
                    ' 代理连接成功！！！可以正常发送数据了！
                    mblnSocks5Connected = True
                    Call tcpGame_Connect
                End If
        End Select
    Else
        Call Command(Left(strData, 1), Mid(strData, 2))
    End If
End Sub

Private Sub tcpGame_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next

    If Number <> sckSuccess Then
        If tcpGame.State <> sckClosed Then tcpGame.Close
        If Not gblnGameStart Then
            If mlngTryJoin = 1 Then
                Call ReadyTryJoin(mstrTryLANIP, mstrTryIP, mlngTryPort)
                Exit Sub
            End If
            lblTips.Caption = LoadString(207)
            Call MessageBox(Me.hWnd, LoadString(114), vbExclamation, LoadString(181))
        End If
        Call tcpGame_Close
    Else
        mlngTryJoin = 0
    End If
End Sub

' 服务器端接收连接，并建立与客户端的连接。
Private Sub tcpListen_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next

    If tcpListen.State <> sckClosed Then tcpListen.Close
    tcpListen.LocalPort = 0
    mlngTryJoin = 0
    Call tcpGame.Accept(requestID)     ' 建立连接
End Sub

Private Sub mnuLogout_Click()
    If Not gblnLogin Then Exit Sub

    If gblnGameStart Then
        Call frmRequest.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name, gYourUserInfo.Face, CMD_RequestLogout, , True)
    Else
        Dim mbrMsgBox As VbMsgBoxResult

        mbrMsgBox = MessageBox(Me.hWnd, LoadString(137), vbYesNo Or vbQuestion, LoadString(178))
        If mbrMsgBox <> vbYes Then Exit Sub

        Call ExitTable(True)
        Call Disconnect

        If Logout(False) Then
            gblnLogin = False
            Call LoginSet
        Else
            Call MessageBox(Me.hWnd, LoadString(138), vbCritical, LoadString(181))
        End If
    End If
End Sub

Private Function StartListen(ByVal Port As Long) As Boolean
    On Error GoTo ErrorHandler

    If tcpListen.State <> sckClosed Then tcpListen.Close
    tcpListen.LocalPort = Port
    Call tcpListen.Listen
    StartListen = True

    Exit Function

ErrorHandler:
    Call MessageBox(Me.hWnd, LoadString(139), vbCritical, LoadString(181))
    Call tcpListen.Close
    StartListen = False
End Function

Private Sub mnuExit_Click()
    Call Unload(Me)
End Sub

Private Sub mnuLogin_Click()
    If gblnLogin Or gblnOfflineMode Then Exit Sub

    Call frmLogin.Show(vbModal)
End Sub

Private Sub mnuRegister_Click()
    Call frmRegister.Show(vbModal)
End Sub

Private Sub mnuAbout_Click()
    Call frmAbout.ShowEx
End Sub

Private Sub mnuOption_Click()
    glngSave_OptionPage = frmOption.ShowEx(glngSave_OptionPage)
End Sub

' 游戏真正开始子程序
Private Function GameStart() As Boolean
    Dim strUrl As String
    Dim strStatus As String

    On Error Resume Next

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_GAME_START & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If (Not ServerCommand(GameInet, mblnServerOK, strUrl, strStatus, , True)) And strStatus <> STATUS_OK Then
        GameStart = False
        Exit Function
    End If

    gblnGameStart = True
    mblnAgreeDisconnect = False

    If Not (Table Is Nothing) Then
        Set Table = Nothing
    End If

    ' 播放音效
    Call PlaySoundEffects(SOUND_GAME_START, gstrSave_SoundValue(SOUND_GAME_START))

    ' 创建新的棋局对象
    Set Table = New CTable

    mlngGameSecond = gMainTableInfo.Timer * 60
    mlngSecond = 0

    If gMainTableInfo.Timer > 0 Then
        lblTime.Caption = GetTime(mlngGameSecond)
        lblTime1.Caption = lblTime.Caption
    End If

    If mlngPlayer(PLY_ME) = T_BLACK Then
        mblnGameFocus = True
        Call SetMousePointer(psDefault)
        Call DisplayCurrentPlayer(T_BLACK)
        ' 开始记时
        If gMainTableInfo.Timer > 0 Then tmrGame.Enabled = True
    Else
        mblnGameFocus = False
        Call SetMousePointer(psPointer)
        Call DisplayCurrentPlayer(T_BLACK)
    End If
    ' 开始闪烁当前游戏者的头像
    tmrFaceFlash.Enabled = True

    ' 判断自己是否可以落子
    If Not Table.CanDown(mlngPlayer(PLY_ME)) Then
        Call SendCommand(CMD_NoDown)
        mblnGameFocus = False
        Call SetMousePointer(psPointer)
        Call DisplayCurrentPlayer(mlngPlayer(PLY_YOU))
        tmrGame.Enabled = False
    End If

    ' 画棋盘
    Call DrawTable

    GameStart = True
End Function

Private Sub DisplayPic()
    Dim Block As Single
    Dim OneBlock As Long
    Dim i As Single
    Dim j As Single

    On Error Resume Next

    DoEvents

    MainPicture.AutoRedraw = False
    Call Sleep(100)
    Block = GetTwipX(10)
    OneBlock = MainPicture.ScaleWidth \ Block
    For i = GetTwipX(1) To GetTwipX(OneBlock) Step GetTwipX(1)
        For j = 0 To Block - GetTwipX(1) Step GetTwipX(1)
            Call MainPicture.PaintPicture(ChessBoard, j * OneBlock, 0, i, MainPicture.ScaleHeight, j * OneBlock, 0, i, MainPicture.ScaleHeight)
        Next j
        Call Sleep(8)
    Next i
    MainPicture.AutoRedraw = True
    Set MainPicture.Picture = ChessBoard

    picTitle.AutoRedraw = False
    For i = -picTitle.ScaleWidth + GetTwipX(100) To 0 Step GetTwipX(2)
        Call picTitle.PaintPicture(GameTitle, 0, i)
        Call Sleep(5)
    Next i
    picTitle.AutoRedraw = True
    Set picTitle.Picture = GameTitle

    fltbtnExit.Visible = True
    fltbtnMin.Visible = True
    fraMainButton.Visible = True
    fraPlayer(0).Visible = True
    fraPlayer(1).Visible = True
    fraStatus.Visible = True

    ' 设置在线用户窗口
    frmOnline.DockStyle = gwifSave_OnlineWindow.DockStyle
    If gwifSave_OnlineWindow.Show Then
        fltbtnOnline.Value = True
        Call frmOnline.ShowEx(Me)
    End If

    ' 设置棋局列表窗口
    frmTable.DockStyle = gwifSave_TableWindow.DockStyle
    If gwifSave_TableWindow.Show Then
        fltbtnTable.Value = True
        Call frmTable.ShowEx(Me)
    End If

    ' 设置聊天窗口
    frmChat.DockStyle = gwifSave_ChatWindow.DockStyle
    If gwifSave_ChatWindow.Show Then
        fltbtnChat.Value = True
        Call frmChat.ShowEx(Me)
    End If

    ' 设置公共聊天区
    If gwifSave_PublicChatWindow.Show Then
        Call frmPublicChat.ShowEx
    End If

    DoEvents
    Starting = False

    Me.WindowState = vbNormal
    Screen.MousePointer = vbDefault
End Sub

Private Sub SetTableFocus(ByVal Col As Integer, ByVal Row As Integer, ByVal Style As Integer)
    Dim X As Single
    Dim Y As Single

    On Error Resume Next

    If mlngPlayer(PLY_ME) = T_BLACK Then
        Call SetMousePointer(psBlack)
    Else
        Call SetMousePointer(psWhite)
    End If
    X = Table.GetX(Col)
    Y = Table.GetY(Row)
    If Style = 1 Then
        Call MainPicture.PaintPicture(SelectDown, X, Y)
    Else
        Call MainPicture.PaintPicture(SelectIcon, X, Y)
    End If
End Sub

Private Sub LostTableFocus(ByVal Col As Integer, ByVal Row As Integer)
    Dim X As Single
    Dim Y As Single

    On Error Resume Next

    Call SetMousePointer(psDefault)
    X = Table.GetX(Col)
    Y = Table.GetY(Row)
    If Table.GetMan(Col, Row) = T_NONE And gblnSave_DownTip And Table.IsDown(Col, Row, mlngPlayer(PLY_ME)) Then
        Call MainPicture.PaintPicture(TipsBitmap, Table.GetX(Col), Table.GetY(Row))
    Else
        MainPicture.Line (X, Y)-(X + Table.PanelWidth - 32, Y + Table.PanelHeight - 20), CLR_TABLE_NORMAL, BF
    End If
End Sub

Public Sub DrawTable()
    Dim i As Long
    Dim j As Long

    On Error Resume Next

    Call MainPicture.PaintPicture(ChessBoard, TablePos.X, TablePos.Y)
    If gblnGameStart Then
        For i = 0 To 7
            For j = 0 To 7
                If gblnSave_DownTip And mblnGameFocus And Table.IsDown(j, i, mlngPlayer(PLY_ME)) Then
                    Call MainPicture.PaintPicture(TipsBitmap, Table.GetX(j), Table.GetY(i))
                ElseIf j = LastDown.Col And i = LastDown.Row Then
                    Call DrawChessMan(j, i, Table.GetMan(j, i), STY_SELECT_MAN)
                Else
                    Call DrawChessMan(j, i, Table.GetMan(j, i), STY_NORMAL_MAN)
                End If
            Next j
        Next i
        If mlngPlayer(PLY_ME) > 0 Then
            lblChessNum(GetPlayerIndex(PLY_ME)).Caption = Format(Table.GetTotal(mlngPlayer(PLY_ME)), "0#")
        End If
        If mlngPlayer(PLY_YOU) > 0 Then
            lblChessNum(GetPlayerIndex(PLY_YOU)).Caption = Format(Table.GetTotal(mlngPlayer(PLY_YOU)), "0#")
        End If
    End If
End Sub

Private Sub DrawChessMan(ByVal Col As Long, ByVal Row As Long, Man As Byte, Optional ByVal Style As Integer = STY_NORMAL_MAN)
    Dim X As Single
    Dim Y As Single

    On Error Resume Next

    X = Table.GetX(Col)
    Y = Table.GetY(Row)
    If Style = STY_NORMAL_MAN Then
        Select Case Man
            Case T_BLACK
                Call MainPicture.PaintPicture(BlackMan, Table.GetX(Col), Table.GetY(Row))
            Case T_WHITE
                Call MainPicture.PaintPicture(WhiteMan, Table.GetX(Col), Table.GetY(Row))
            Case Else
                MainPicture.Line (X, Y)-(X + Table.PanelWidth - 32, Y + Table.PanelHeight - 20), CLR_TABLE_NORMAL, BF
        End Select
    Else
        Select Case Man
            Case T_BLACK
                Call MainPicture.PaintPicture(SelBlackMan, Table.GetX(Col), Table.GetY(Row))
            Case T_WHITE
                Call MainPicture.PaintPicture(SelWhiteMan, Table.GetX(Col), Table.GetY(Row))
        End Select
    End If
End Sub

Public Sub LoginSet()
    On Error Resume Next

    If gblnLogin Then
        ' 已经登陆
        Call frmOnline.ReloadOnline
        Call frmTable.ReloadTable(True)
        Call SetLight(0, lsLightGreen)
        Call frmChat.EnableChat(GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name), False)
        Call frmPublicChat.EnableChat(GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name))
        lblLogin.Caption = GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name)
        lblTips.Caption = LoadString(190)
    Else
        ' 还没有登陆
        Call frmOnline.lvwOnline.ListItems.Clear
        Call frmTable.lvwTable.ListItems.Clear
        Call SetLight(0, lsLightRed)
        Call SetLight(1, lsLightRed)
        Call SetLight(2, lsLightRed)
        Call frmChat.DisableChat(True)
        Call frmPublicChat.DisableChat
        Call frmUserInfo.Hide
        Call frmEditInfo.Hide
        Call frmTableInfo.Hide
        lblLogin.Caption = LoadString(191)
        lblTips.Caption = LoadString(192)
        gMyUserInfo.UserName = ""
        gYourUserInfo.UserName = ""
    End If
End Sub

Private Sub DisplayCurrentPlayer(ByVal Player As Byte)
    ' 再次开始闪烁头像
    On Error Resume Next

    picPlayerFace(PLY_ME).Visible = True
    picPlayerFace(PLY_YOU).Visible = True
    tmrFaceFlash.Enabled = True

    If Player = mlngPlayer(PLY_ME) Then
        Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_ME)), LoadString(124), "", True)
        Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_YOU)), "", "", True)
    Else
        Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_YOU)), LoadString(125), "", True)
        Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_ME)), "", "", True)
    End If
End Sub

' 准备创建棋局，用于判断创建是否成功。
Public Sub ReadyCreateTable()
    On Error Resume Next

    If StartListen(glngSave_GamePort) Then
        If CreateTable() Then
            Call frmTable.ReloadTable(True)
            gblnCreator = True
            frmTable.fltbtnJoin.Enabled = False
            shpPlayerTips(0).Visible = True
            shpPlayerTips(1).Visible = True
            Call SetLabel(lblPlayerTips(0), LoadString(126), "", True)
            Call SetLabel(lblPlayerTips(1), LoadString(127), "", True)
            fltbtnSitDown(0).Visible = True
            fltbtnSitDown(1).Visible = True
            lblTable.Caption = gMainTableInfo.TableName
            Call SetLight(2, lsLightGreen)
            lblTips.Caption = LoadString(193)
            ' 显示时间
            If gMainTableInfo.Timer > 0 Then
                lblTime.Caption = GetTime(gMainTableInfo.Timer * 60)
            End If
        Else
            Call tcpListen.Close
        End If
    End If
End Sub

Private Function CreateTable() As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    On Error Resume Next

    If gMainTableInfo.TableName = "" Then
        CreateTable = False
        Exit Function
    End If

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_TABLE_CREATE & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&creator=" & ToUrlString(gMainTableInfo.Creator) & _
                                  "&nickname=" & ToUrlString(gMyUserInfo.Name) & _
                                  "&name=" & ToUrlString(gMainTableInfo.TableName) & _
                                  "&type=" & CStr(gMainTableInfo.TableType) & _
                                  "&timer=" & CStr(gMainTableInfo.Timer) & _
                                  "&level=" & IIf(gMainTableInfo.UpLevel, "1", "0") & _
                                  "&lanip=" & gstrLocalIP & _
                                  "&port=" & CStr(glngSave_GamePort) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If ServerCommand(GameInet, mblnServerOK, strUrl, strStatus, strData) Then
        Select Case strStatus
            Case STATUS_OK
                Call LoadTableInfo(gMainTableInfo, strData)
                If gMainTableInfo.LANIP <> "" And Not mblnLANMessage Then
                    mblnLANMessage = True
                    Call MessageBox(Me.hWnd, LoadString(140), vbInformation, LoadString(180))
                End If
                CreateTable = True
                Exit Function
            Case STATUS_ERROR
                Call MessageBox(Me.hWnd, LoadString(141) & strData, vbExclamation, LoadString(181))
            Case Else
                Call MessageBox(Me.hWnd, LoadString(142), vbExclamation, LoadString(181))
        End Select
    End If
    CreateTable = False
End Function

Public Sub ReadyTryJoin(ByVal LANIP As String, ByVal ip As String, ByVal Port As Long)
    If mlngTryJoin = 0 Then
        mlngTryJoin = 1
        mstrTryLANIP = LANIP
        mstrTryIP = ip
        mlngTryPort = Port
        Call ReadyJoinTable(mstrTryLANIP, mlngTryPort)
    ElseIf mlngTryJoin = 1 Then
        mlngTryJoin = 0
        Call ReadyJoinTable(mstrTryIP, mlngTryPort)
    End If
End Sub

' 准备加入棋局公共函数，此时并没有真正加入。
Public Sub ReadyJoinTable(ByVal RemoteHostIP As String, ByVal RemotePort As Long)
    On Error GoTo ErrorHandler

    If tcpGame.State <> sckClosed Then tcpGame.Close

    ' 连接到服务器端。
    If gblnSave_UseProxy Then
        mstrRemoteHostIP = RemoteHostIP
        mlngRemotePort = RemotePort
        mblnSocks5Connected = False
        mlngSocks5Status = 1
        Call tcpGame.Connect(gstrSave_Socks5ProxyIP, glngSave_Socks5ProxyPort)
    Else
        Call tcpGame.Connect(RemoteHostIP, RemotePort)
    End If

    lblTips.Caption = LoadString(194)

    Exit Sub

ErrorHandler:
    Call tcpGame.Close
    If mlngTryJoin = 1 Then
        Call ReadyTryJoin(mstrTryLANIP, mstrTryIP, mlngTryPort)
    Else
        Call MessageBox(Me.hWnd, LoadString(114), vbExclamation, LoadString(181))
    End If
End Sub

Public Function JoinTable(ByVal TableName As String) As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    On Error Resume Next

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_TABLE_JOIN & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&name=" & ToUrlString(TableName) & _
                                  "&visitor=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&nickname=" & ToUrlString(gMyUserInfo.Name) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If ServerCommand(GameInet, mblnServerOK, strUrl, strStatus, strData) Then
        Select Case strStatus
            Case STATUS_OK
                Call LoadTableInfo(gMainTableInfo, strData)
                JoinTable = True
                Exit Function
            Case STATUS_ERROR
                Call MessageBox(Me.hWnd, LoadString(143) & strData, vbExclamation, LoadString(181))
            Case Else
                Call MessageBox(Me.hWnd, LoadString(144), vbCritical, LoadString(181))
        End Select
    End If
    JoinTable = False
End Function

Public Sub LostFocus()
    If Not Starting Then
        Set picTitle.Picture = NoFocusTitle
    End If
End Sub

Public Sub GetFocus()
    If Not Starting Then
        Set picTitle.Picture = GameTitle
    End If
End Sub

Private Sub Table_DownChessMan(Col As Integer, Row As Integer, Man As Byte)
    'Call Sleep(100)
    'Call DrawChessMan(Col, Row, Man)
    'Call MainPicture.Refresh
End Sub

Public Function GetFace(ByVal FaceNumber As Long) As StdPicture
    On Error Resume Next

    If FaceNumber < 1 Then Set GetFace = Nothing
    Set GetFace = ilsFace.ListImages(FaceNumber).Picture
End Function

' 处理对方加入
Private Sub VisitorJoin()
    On Error Resume Next

    mblnSitDown(GetPlayerIndex(PLY_YOU)) = True
    fltbtnSitDown(GetPlayerIndex(PLY_YOU)).Visible = False

    shpPlayerTips(GetPlayerIndex(PLY_YOU)).Visible = True
    lblPlayerTips(GetPlayerIndex(PLY_YOU)).Caption = ""
    shpPlayerName(GetPlayerIndex(PLY_YOU)).Visible = True

    Call SetLabel(lblPlayerName(GetPlayerIndex(PLY_YOU)), GetDisplayName(gYourUserInfo.UserName, gYourUserInfo.Name), "", True)

    Call DisplayFace(GetPlayerIndex(PLY_YOU), gYourUserInfo.Face)

    If CheckArray(mblnSitDown()) Then
        fltbtnStart(GetPlayerIndex(PLY_ME)).Visible = True
        Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_ME)), LoadString(115), "", True)
        shpPlayerTips(GetPlayerIndex(PLY_YOU)).Visible = True
        Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_YOU)), _
                      LoadString(118), LoadString(119), _
                      True)
    End If
End Sub

Public Sub SendTalk(ByVal Talk As String)
    Call SendCommand(CMD_Talk & Talk)
End Sub

Private Sub ConnectStatus(ByVal Status As Boolean)
    gblnConnect = Status

    If gblnConnect Then
        Call frmChat.EnableChat(GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name), True)
    Else
        Call frmChat.DisableChat(False)
    End If
End Sub

' 组合落子命令
Private Function SetDownManCommand(ByVal Man As Byte, ByVal Col As Integer, ByVal Row As Integer) As Long
    Dim Ret As Long

    If Man < 1 Or Man > 2 Then Exit Function

    ' 置位 Ret 为 8000H
    ' 原因:必须置第 16 位为 1，才能正常传输。
    Ret = &H8000
    ' 棋子左移 8 位
    Ret = Ret + Man * 256
    ' 填补第 8 位为 1
    ' 原因:必须置第 8 位为 1，才能正常传输。
    Ret = Ret + 128
    ' 列位置左移 3 位
    Ret = Ret + (Col * 8)
    ' 行位置
    Ret = Ret + Row

    SetDownManCommand = Ret
End Function

' 拆分落子命令
' 使用: Command - 原始命令（数值，准备拆分的命令）
'       Col     - 返回的列位置（拆分出的列位置）
'       Row     - 返回的行位置（拆分出的行位置）
' 返回值: 棋子（就是下的哪个棋）
Private Function GetDownManCommand(ByVal Command As Integer, ByRef Col As Integer, ByRef Row As Integer) As Byte
    ' 和二进制数 0000000000111000 进行与操作（屏蔽其它位），
    ' 并右移 3 位
    Col = (Command And 56) \ 8

    ' 和二进制数 0000000000000111 进行与操作（屏蔽其它位），
    Row = (Command And 7)

    ' 和二进制数 0000001100000000 进行与操作（屏蔽其它位），
    ' 并右移 8 位
    GetDownManCommand = (Command And 768) \ 256
End Function

Private Function VisitorDownMan(ByVal Col As Integer, ByVal Row As Integer, ByVal Man As Byte) As Boolean
    Dim i As Long

    On Error Resume Next

    If Not Table.IsDown(Col, Row, Man) Then
        VisitorDownMan = False
        Exit Function
    End If

    ' 播放音效
    Call PlaySoundEffects(SOUND_DOWN_MAN, gstrSave_SoundValue(SOUND_DOWN_MAN))

    LastDown.Col = Col: LastDown.Row = Row: LastMan = Man
    Call SetMousePointer(psPointer)
    For i = 0 To 1
        Call DrawChessMan(Col, Row, Man, 0)
        Call MainPicture.Refresh
        Call Sleep(100)
        Call DrawChessMan(Col, Row, T_NONE, 0)
        Call MainPicture.Refresh
        Call Sleep(100)
    Next i
    Call Table.DownMan(Col, Row, Man)

    ' 判断你是否可以落子
    If Not Table.CanDown(mlngPlayer(PLY_ME)) Then
        Call SendCommand(CMD_NoDown, True)
        Call DrawTable
        ' 播放音效
        Call PlaySoundEffects(SOUND_NOT_DOWN, gstrSave_SoundValue(SOUND_NOT_DOWN))
        VisitorDownMan = True
        Exit Function
    End If

    mblnGameFocus = True
    Call SetMousePointer(psDefault)
    ' 提示你落子
    Call DisplayCurrentPlayer(mlngPlayer(PLY_ME))
    ' 开始记时
    If gMainTableInfo.Timer > 0 Then tmrGame.Enabled = True
    ' 重画棋盘
    Call DrawTable

    VisitorDownMan = True
End Function

Private Function CheckArray(blnArray() As Boolean) As Boolean
    Dim i As Variant

    For Each i In blnArray
        If i = False Then
            CheckArray = False
            Exit Function
        End If
    Next i

    CheckArray = True
End Function

Public Sub SendCommand(ByVal Command As String, Optional ByVal WaitOK As Boolean = False)
    On Error Resume Next

    If tcpGame.State = sckConnected Then
        Call tcpGame.SendData(Command)
        If WaitOK Then
            tmrOK.Tag = Command
            tmrOK.Enabled = True
        End If
    End If
End Sub

Private Function Command(ByVal strCommand As String, ByVal strData As String) As Boolean
    On Error Resume Next

    Select Case strCommand

        Case CMD_OK
            Call ReceiveOK(strData)

        Case CMD_Connected  ' 对方允许加入，在服务器端接收此命令。
            Call ConnectStatus(True)
            Call frmTable.ReloadTable(True)
            Call frmOnline.ReloadOnline
            ' 播放音效
            Call PlaySoundEffects(SOUND_JOIN_TABLE, gstrSave_SoundValue(SOUND_JOIN_TABLE))
            gYourUserInfo.UserName = GetRecord(strData, 1)
            gYourUserInfo.Name = GetRecord(strData, 2)
            lblConnect.Caption = GetDisplayName(gYourUserInfo.UserName, gYourUserInfo.Name)
            lblTips.Caption = LoadString(195)
            Call SetLight(1, lsLightGreen)
            If mblnSitDown(GetPlayerIndex(PLY_ME)) Then
                Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_ME)), _
                              LoadString(120), LoadString(121), _
                              True)
                ' 如果已坐下，则发送此命令给客户端。
                ' 参数: 我的棋子、用户名、头像、昵称。
                Call SendCommand(CMD_SitDown & CStr(mlngPlayer(PLY_ME)) & "|" & gMyUserInfo.UserName & "|" & CStr(gMyUserInfo.Face) & "|" & gMyUserInfo.Name)
            Else
                ' 如果没有人坐下，发送此命令给客户端。
                ' 参数: 我的用户名、昵称
                Call SendCommand(CMD_NoneSitDown & gMyUserInfo.UserName & "|" & gMyUserInfo.Name)
            End If

        Case CMD_GameStart
            Call FlashWindow(Me.hWnd)
            mblnStart(GetPlayerIndex(PLY_YOU)) = True
            lblPlayerTips(GetPlayerIndex(PLY_YOU)).Caption = ""
            shpChessNum(GetPlayerIndex(PLY_YOU)).Visible = True
            Call SetLabel(lblChessNum(GetPlayerIndex(PLY_YOU)), "00", "", True)
            If CheckArray(mblnStart()) Then
                ' 游戏真正开始。
                If Not GameStart() Then
                    Call SendCommand(CMD_AgainStart)
                    If mblnStart(GetPlayerIndex(PLY_ME)) Then
                        Call AgainStart
                    End If
                    Call MessageBox(Me.hWnd, LoadString(130), vbExclamation, LoadString(181))
                End If
            End If

        Case CMD_AgainStart
            If gblnGameStart Then
                Call AgainStart
                Call MessageBox(Me.hWnd, LoadString(145), vbExclamation, LoadString(181))
            End If

        Case CMD_GameReadyStart
            Call FlashWindow(Me.hWnd)
            mblnReadyStart(GetPlayerIndex(PLY_YOU)) = True
            lblPlayerTips(GetPlayerIndex(PLY_YOU)).Caption = LoadString(196)

            If CheckArray(mblnReadyStart()) Then
                Call AgainStart
            End If

        Case CMD_SitDown    ' 客人加入(坐下)，双方互相传送。
            Call FlashWindow(Me.hWnd)
            mlngPlayer(PLY_YOU) = CLng(GetRecord(strData, 1))
            gYourUserInfo.UserName = GetRecord(strData, 2)
            gYourUserInfo.Face = CLng(GetRecord(strData, 3))
            gYourUserInfo.Name = GetRecord(strData, 4)
            If Not mblnSitDown(ToPartner(mlngPlayer(PLY_YOU)) - 1) Then
                shpPlayerTips(ToPartner(mlngPlayer(PLY_YOU)) - 1).Visible = True
                Call SetLabel(lblPlayerTips(ToPartner(mlngPlayer(PLY_YOU)) - 1), LoadString(128), "", True)
                fltbtnSitDown(ToPartner(mlngPlayer(PLY_YOU)) - 1).Visible = True
            End If
            Call VisitorJoin

        Case CMD_NoneSitDown
            Call FlashWindow(Me.hWnd)
            gYourUserInfo.UserName = GetRecord(strData, 1)
            gYourUserInfo.Name = GetRecord(strData, 2)
            shpPlayerTips(0).Visible = True
            shpPlayerTips(1).Visible = True
            Call SetLabel(lblPlayerTips(0), LoadString(126), "", True)
            Call SetLabel(lblPlayerTips(1), LoadString(127), "", True)
            fltbtnSitDown(0).Visible = True
            fltbtnSitDown(1).Visible = True

        Case CMD_DownChessMan
            Dim Col As Integer
            Dim Row As Integer
            Dim Man As Byte

            ' 收到则发送 OK 消息，并附带本次命令及参数。
            Call FlashWindow(Me.hWnd)
            Call SendCommand(CMD_OK & strCommand & strData)
            Man = GetDownManCommand(Asc(strData), Col, Row)
            Call VisitorDownMan(Col, Row, Man)

        Case CMD_NoDown
            Call SendCommand(CMD_OK & strCommand)
            If Not Table.CanDown(mlngPlayer(PLY_ME)) Then
                ' 如果双方都无路可走
                Call SendCommand(CMD_GameOver, True)
                Call GameOver
            Else
                ' 对方无路可走时
                mblnGameFocus = True
                Call SetMousePointer(psDefault)
                ' 提示你落子
                Call DisplayCurrentPlayer(mlngPlayer(PLY_ME))
                ' 开始记时
                If gMainTableInfo.Timer > 0 Then tmrGame.Enabled = True
                ' 重画棋盘
                Call DrawTable
                ' 播放音效
                Call PlaySoundEffects(SOUND_NOT_DOWN, gstrSave_SoundValue(SOUND_NOT_DOWN))
            End If

        Case CMD_GameOver
            Call SendCommand(CMD_OK & strCommand)
            Call GameOver(strData)

        Case CMD_TimeOver
            Call TimeOver(GAME_WIN)

        Case CMD_Talk
            Call FlashWindow(Me.hWnd)
            If Not fltbtnChat.Value Then
                fltbtnChat.Value = True
                Call CloseModal
                Call frmChat.ShowEx(Me)
            End If
            Call frmChat.Chat(GetDisplayName(gYourUserInfo.UserName, gYourUserInfo.Name), strData)
            ' 播放音效
            Call PlaySoundEffects(SOUND_CHAT, gstrSave_SoundValue(SOUND_CHAT))

        Case CMD_InfoChanged
            Call FlashWindow(Me.hWnd)
            gYourUserInfo.Face = Val(GetRecord(strData, 1))
            gYourUserInfo.Name = GetRecord(strData, 2)
            lblConnect.Caption = GetDisplayName(gYourUserInfo.UserName, gYourUserInfo.Name)
            Call DisplayFace(GetPlayerIndex(PLY_YOU), gYourUserInfo.Face)
            Call SetLabel(lblPlayerName(GetPlayerIndex(PLY_YOU)), GetDisplayName(gYourUserInfo.UserName, gYourUserInfo.Name), "", True)
            ' 注意：这里要显示一些提示！

        Case CMD_TableChanged
            Call FlashWindow(Me.hWnd)
            gMainTableInfo.TableType = CLng(GetRecord(strData, 1))
            gMainTableInfo.Timer = CLng(GetRecord(strData, 2))
            gMainTableInfo.UpLevel = CBool(GetRecord(strData, 3))
            If gMainTableInfo.Timer < 1 Then
                lblTime.Caption = "--:--"
            Else
                lblTime.Caption = GetTime(gMainTableInfo.Timer * 60)
            End If
            Call frmTable.ReloadTable(True)
            ' 注意：这里要显示一些提示！

        ' 收到请求命令
        Case CMD_Request
            Call FlashWindow(Me.hWnd)
            Call ReceiveRequest(strData)

        Case CMD_ReRequest
            Call FlashWindow(Me.hWnd)
            Call ReceiveReRequest(strData)

    End Select
End Function

Private Function Logout(ByVal Quiet As Boolean) As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    On Error Resume Next

    If Not gblnLogin Then
        Logout = False
        Exit Function
    End If

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_LOGOUT & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If Quiet Then
        Logout = ServerExecute(GameInet, strUrl)
    Else
        ' 播放音效
        Call PlaySoundEffects(SOUND_LOGOUT, gstrSave_SoundValue(SOUND_LOGOUT))
        If ServerCommand(GameInet, mblnServerOK, strUrl, strStatus, strData, Quiet, True) Then
            If strStatus <> STATUS_OK Then
                Logout = False
            Else
                Logout = True
            End If
        Else
            Logout = False
        End If
    End If
End Function

Private Sub StartFlashLight(ByVal Where As Long, ByVal Light As eLightStyle)
    On Error Resume Next

    mlngLightWhere = Where
    Call SetLight(Where, Light)
    tmrLightFlash.Enabled = True
End Sub

Private Sub StopFlashLight()
    On Error Resume Next

    tmrLightFlash.Enabled = False
    imgLight(mlngLightWhere).Visible = True
End Sub

Private Sub Disconnect()
    On Error Resume Next

    If tcpListen.State <> sckClosed Then Call tcpListen.Close
    If tcpGame.State <> sckClosed Then Call tcpGame.Close

    gblnGameStart = False
    gblnConnect = False
    gblnCreator = False
    lblConnect.Caption = LoadString(197)
    mlngTryJoin = 0
    Call ConnectStatus(False)
    Call SetLight(1, lsLightRed)
    gYourUserInfo.UserName = ""
End Sub

Private Sub SetLight(ByVal Where As Long, ByVal Light As eLightStyle)
    On Error Resume Next

    Select Case Light
        Case lsLightGreen
            Set imgLight(Where).Picture = objLightOn
        Case lsLightRed
            Set imgLight(Where).Picture = objLightOff
        Case lsLightYellow
            Set imgLight(Where).Picture = objLightYellow
    End Select
    Call imgLight(Where).Refresh
    imgLight(Where).Visible = True
End Sub

Public Sub AgreeRequest(ByVal strCommand As String)
    On Error Resume Next

    gblnGameStart = False
    Select Case strCommand
        Case CMD_RequestExitGame
            mblnAgreeDisconnect = True
            Call CancelGame

        Case CMD_RequestExitTable
            mblnAgreeDisconnect = True
            Call CancelGame

        Case CMD_RequestCancelGame
            Call CancelGame

            Call GameFinish
            If Not (Table Is Nothing) Then
                Call Table.Clear
                Call DrawTable
                Set Table = Nothing
            End If
            lblPlayerTips(GetPlayerIndex(PLY_ME)).Caption = LoadString(198)
            If Not mblnReadyStart(GetPlayerIndex(PLY_YOU)) Then
                lblPlayerTips(GetPlayerIndex(PLY_YOU)).Caption = LoadString(199)
            End If

            fltbtnReadyStart(GetPlayerIndex(PLY_ME)).Visible = True

        Case CMD_RequestLogout
            mblnAgreeDisconnect = True
            Call CancelGame

        Case CMD_RequestDrawGame
            Call GameOver(GAME_DRAW)

    End Select
End Sub

Public Sub DisagreeRequest(ByVal strCommand As String)
    Select Case strCommand
        Case CMD_RequestExitGame
            mblnAgreeDisconnect = False

        Case CMD_RequestExitTable
            mblnAgreeDisconnect = False

        Case CMD_RequestCancelGame
            mblnAgreeDisconnect = False

        Case CMD_RequestLogout
            mblnAgreeDisconnect = False
    End Select
End Sub

Private Sub ExitTable(ByVal Quiet As Boolean)
    Dim i As Long
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    On Error Resume Next

    If (Not gblnConnect) And (Not gblnCreator) Then
        Exit Sub
    End If

    Call frmTableInfo.Hide
    tmrGame.Enabled = False

    If Not gblnOfflineMode Then
        strUrl = gstrSave_ServerUrl & SERVER_ACTION_TABLE_EXIT & _
                                      "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                      "&password=" & MD5(gMyUserInfo.Password) & _
                                      "&name=" & ToUrlString(gMainTableInfo.TableName) & _
                                      "&visitor=" & ToUrlString(gYourUserInfo.UserName) & _
                                      "&" & MakeServerPassword() & _
                                      "&" & MakeVersion()
    
        Call ServerCommand(GameInet, mblnServerOK, strUrl, strStatus, strData, Quiet, True)
    
        If gblnCreator Then
            strUrl = gstrSave_ServerUrl & SERVER_ACTION_TABLE_REMOVE & _
                                          "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                          "&password=" & MD5(gMyUserInfo.Password) & _
                                          "&name=" & ToUrlString(gMainTableInfo.TableName) & _
                                          "&" & MakeServerPassword() & _
                                          "&" & MakeVersion()
    
            Call ServerCommand(GameInet, mblnServerOK, strUrl, strStatus, strData, Quiet, True)
        End If
    
        Call frmTable.ReloadTable(True)
    End If

    ' 播放音效
    Call PlaySoundEffects(SOUND_EXIT_TABLE, gstrSave_SoundValue(SOUND_EXIT_TABLE))

    ' 停止闪烁头像
    tmrFaceFlash.Enabled = False

    For i = 0 To 1
        Call SetLabel(lblPlayerName(i), "", "", False)
        Call SetLabel(lblChessNum(i), "", "", False)
        Call SetLabel(lblPlayerTips(i), "", "", False)

        ' 隐藏头像
        picPlayerFace(i).Visible = False
        ' 隐藏名称等的背景
        shpPlayerName(i).Visible = False
        shpChessNum(i).Visible = False
        shpPlayerTips(i).Visible = False
        ' 隐藏坐下与开始按钮
        fltbtnSitDown(i).Visible = False
        fltbtnStart(i).Visible = False
        fltbtnReadyStart(i).Visible = False

        mblnStart(i) = False
        mblnReadyStart(i) = False
        mblnSitDown(i) = False
        mlngPlayer(i) = T_NONE
    Next i

    ' 复位棋盘光标
    Call SetMousePointer(psPointer)

    ' 关灯
    Call SetLight(2, lsLightRed)
    lblTable.Caption = LoadString(200)

    lblTime.Caption = "--:--"
    lblTime1.Caption = ""

    If gblnOfflineMode Then
        lblTips.Caption = LoadString(259)
    End If

    gMainTableInfo.TableName = ""

    If Not (Table Is Nothing) Then
        Call Table.Clear
        Call DrawTable
        Set Table = Nothing
    End If
End Sub

Private Function CancelGame() As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    'If Not gblnGameStart Then Exit Function

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_GAME_CANCEL & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If ServerCommand(GameInet, mblnServerOK, strUrl, strStatus, strData, True) Then
        If strStatus <> STATUS_OK Then
            CancelGame = False
        Else
            CancelGame = True
        End If
    Else
        CancelGame = False
    End If
End Function

Private Sub AgainStart()
    ' 关闭游戏记时
    On Error Resume Next

    tmrGame.Enabled = False

    lblTime.Caption = "--:--"
    lblTime1.Caption = ""

    gblnGameStart = False

    mblnStart(GetPlayerIndex(PLY_ME)) = False

    fltbtnStart(GetPlayerIndex(PLY_ME)).Visible = True
    fltbtnReadyStart(GetPlayerIndex(PLY_ME)).Visible = False
    shpChessNum(GetPlayerIndex(PLY_ME)).Visible = False
    Call SetLabel(lblChessNum(GetPlayerIndex(PLY_ME)), "", "", False)

    Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_ME)), LoadString(115), "", True)

    mblnStart(GetPlayerIndex(PLY_YOU)) = False

    fltbtnStart(GetPlayerIndex(PLY_YOU)).Visible = False
    fltbtnReadyStart(GetPlayerIndex(PLY_YOU)).Visible = False
    shpChessNum(GetPlayerIndex(PLY_YOU)).Visible = False
    Call SetLabel(lblChessNum(GetPlayerIndex(PLY_YOU)), "", "", False)

    Call SetLabel(lblPlayerTips(GetPlayerIndex(PLY_YOU)), LoadString(118), LoadString(119), True)

    If Not (Table Is Nothing) Then
        Call Table.Clear
        Call DrawTable
        Set Table = Nothing
    End If
End Sub

Private Function GameOver(Optional ByVal SetGameState As String = "") As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    Dim GameState As Long
    Dim strMessage As String
    Dim strMessage1 As String

    Dim BlackChessNumber As Long
    Dim WhiteChessNumber As Long

    On Error Resume Next

    If SetGameState = "" Then
        GameState = Table.Umpire(mlngPlayer(PLY_ME))
        BlackChessNumber = Table.GetTotal(T_BLACK)
        WhiteChessNumber = Table.GetTotal(T_WHITE)
    Else
        GameState = SetGameState
    End If

    Call GameFinish

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_GAME_OVER & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&partner=" & ToUrlString(gYourUserInfo.UserName) & _
                                  "&tablename=" & ToUrlString(gMainTableInfo.TableName) & _
                                  "&state=" & CStr(GameState) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If ServerCommand(GameInet, mblnServerOK, strUrl, strStatus, strData, False) Then
        If strStatus <> STATUS_OK Then
            Call MessageBox(Me.hWnd, LoadString(260), vbExclamation, LoadString(177))
            GameOver = False
        Else
            ' 刷新在线及棋局信息
            Call frmOnline.ReloadOnline
            Call frmTable.ReloadTable(True)

            If mlngPlayer(PLY_ME) = T_BLACK Then
                strMessage = GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name) & LoadString(202) & CStr(BlackChessNumber) & vbCr _
                             & GetDisplayName(gYourUserInfo.UserName, gYourUserInfo.Name) & LoadString(203) & CStr(WhiteChessNumber) & vbCr
            Else
                strMessage = GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name) & LoadString(203) & CStr(WhiteChessNumber) & vbCr _
                             & GetDisplayName(gYourUserInfo.UserName, gYourUserInfo.Name) & LoadString(202) & CStr(BlackChessNumber) & vbCr
            End If
            strMessage1 = LoadString(201) & GetRecord(strData, 1)

            If GameState = GAME_WIN Then
                ' 播放音效
                Call PlaySoundEffects(SOUND_GAME_WIN, gstrSave_SoundValue(SOUND_GAME_WIN))
                If SetGameState = "" Then
                    Call MessageBox(Me.hWnd, LoadString(146) & strMessage & vbCr & LoadString(147) & GetRecord(strData, 2) & strMessage1, vbInformation, LoadString(179))
                Else
                    Call MessageBox(Me.hWnd, LoadString(146) & LoadString(148) & LoadString(147) & GetRecord(strData, 2) & strMessage1, vbInformation, LoadString(179))
                End If
            ElseIf GameState = GAME_LOSE Then
                ' 播放音效
                Call PlaySoundEffects(SOUND_GAME_LOSE, gstrSave_SoundValue(SOUND_GAME_LOSE))
                If SetGameState = "" Then
                    Call MessageBox(Me.hWnd, LoadString(149) & strMessage & vbCr & LoadString(150) & GetRecord(strData, 2) & strMessage1, vbInformation, LoadString(179))
                Else
                    Call MessageBox(Me.hWnd, LoadString(149) & LoadString(151) & LoadString(150) & GetRecord(strData, 2) & strMessage1, vbInformation, LoadString(179))
                End If
            Else
                ' 播放音效
                Call PlaySoundEffects(SOUND_GAME_DRAW, gstrSave_SoundValue(SOUND_GAME_DRAW))
                If SetGameState = "" Then
                    Call MessageBox(Me.hWnd, LoadString(152) & strMessage & vbCr & strMessage1, vbInformation, LoadString(179))
                Else
                    Call MessageBox(Me.hWnd, LoadString(152) & LoadString(153) & strMessage1, vbInformation, LoadString(179))
                End If
            End If
            GameOver = True
        End If
    Else
        Call MessageBox(Me.hWnd, LoadString(260), vbExclamation, LoadString(177))
        GameOver = False
    End If

    If Not gblnConnect Then
        GameOver = True
        Exit Function
    End If

    lblPlayerTips(GetPlayerIndex(PLY_ME)).Caption = LoadString(198)
    If Not mblnReadyStart(GetPlayerIndex(PLY_YOU)) Then
        lblPlayerTips(GetPlayerIndex(PLY_YOU)).Caption = LoadString(199)
    End If

    fltbtnReadyStart(GetPlayerIndex(PLY_ME)).Visible = True
End Function

Private Function TimeOver(ByVal GameState As Byte) As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    Dim strMessage As String

    On Error Resume Next

    Call GameFinish

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_GAME_OVER & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&partner=" & ToUrlString(gYourUserInfo.UserName) & _
                                  "&tablename=" & ToUrlString(gMainTableInfo.TableName) & _
                                  "&state=" & CStr(GameState) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If ServerCommand(GameInet, mblnServerOK, strUrl, strStatus, strData, False) Then
        If strStatus <> STATUS_OK Then
            Call MessageBox(Me.hWnd, LoadString(260), vbExclamation, LoadString(177))
            TimeOver = False
        Else
            ' 刷新在线及棋局信息
            Call frmOnline.ReloadOnline
            Call frmTable.ReloadTable(True)

            strMessage = LoadString(201) & GetRecord(strData, 1)

            If GameState = GAME_WIN Then
                ' 播放音效
                Call PlaySoundEffects(SOUND_GAME_WIN, gstrSave_SoundValue(SOUND_GAME_WIN))
                Call MessageBox(Me.hWnd, LoadString(146) & LoadString(154) & LoadString(147) & GetRecord(strData, 2) & strMessage, vbInformation, LoadString(179))
            Else
                ' 播放音效
                Call PlaySoundEffects(SOUND_GAME_LOSE, gstrSave_SoundValue(SOUND_GAME_LOSE))
                Call MessageBox(Me.hWnd, LoadString(149) & LoadString(155) & LoadString(150) & GetRecord(strData, 2) & strMessage, vbInformation, LoadString(179))
            End If
            TimeOver = True
        End If
    Else
        Call MessageBox(Me.hWnd, LoadString(260), vbExclamation, LoadString(177))
        TimeOver = False
    End If

    If Not gblnConnect Then
        TimeOver = True
        Exit Function
    End If

    lblPlayerTips(GetPlayerIndex(PLY_ME)).Caption = LoadString(198)
    If Not mblnReadyStart(GetPlayerIndex(PLY_YOU)) Then
        lblPlayerTips(GetPlayerIndex(PLY_YOU)).Caption = LoadString(199)
    End If

    fltbtnReadyStart(GetPlayerIndex(PLY_ME)).Visible = True
End Function

Private Sub DisplayFace(ByVal Index As Long, ByVal Number As Long)
    On Error Resume Next

    Call picPlayerFace(Index).Cls
    If Number > Val(ilsFace.Tag) Or Number < 1 Then
        Set picPlayerFace(Index).Picture = frmResource.imgResDefaultFace.Picture
    Else
        Call ilsFace.ListImages(Number).Draw(picPlayerFace(Index).hDC, 0, 0, imlTransparent)
    End If
    Call picPlayerFace(Index).Refresh
    picPlayerFace(Index).Visible = True
End Sub

Public Sub RefreshMyFace(ByVal Number As Long)
    On Error Resume Next

    If mblnSitDown(GetPlayerIndex(PLY_ME)) Then
        Call DisplayFace(GetPlayerIndex(PLY_ME), Number)
    End If
End Sub

Public Sub RefreshMyName()
    On Error Resume Next

    If mblnSitDown(GetPlayerIndex(PLY_ME)) Then
        Call SetLabel(lblPlayerName(GetPlayerIndex(PLY_ME)), GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name), "", True)
    End If
    If gblnLogin Then
        lblLogin.Caption = GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name)
    End If
End Sub

Private Sub GameFinish()
    On Error Resume Next

    tmrGame.Enabled = False
    tmrFaceFlash.Enabled = False
    mblnGameFocus = False
    Call SetMousePointer(psPointer)
    picPlayerFace(0).Visible = True
    picPlayerFace(1).Visible = True
    lblPlayerTips(0).Caption = ""
    lblPlayerTips(1).Caption = ""
    lblChessNum(0).Caption = ""
    lblChessNum(1).Caption = ""
    shpChessNum(0).Visible = False
    shpChessNum(1).Visible = False

    gblnGameStart = False

    mblnReadyStart(PLY_ME) = False
    mblnReadyStart(PLY_YOU) = False
End Sub

Private Function GetData(ByVal strData As String, ByVal Index As Integer) As Byte
    GetData = AscW(Mid(strData, Index, 1))
End Function

Private Sub SendSocks5Request(ByVal ip As String, ByVal Port As Long)
    Dim Buffers(9) As Byte
    Dim p(4) As Byte
    Dim i As Long

    On Error Resume Next

    For i = 1 To 4
        p(i) = Int(GetInfo(ip, i, "."))
    Next i

    Buffers(0) = 5
    Buffers(1) = 1
    Buffers(2) = 0
    Buffers(3) = 1
    Buffers(4) = p(1)
    Buffers(5) = p(2)
    Buffers(6) = p(3)
    Buffers(7) = p(4)
    Buffers(8) = (&HFF00 And Port) \ &H100
    Buffers(9) = &HFF And Port

    Call tcpGame.SendData(Buffers())
End Sub

Private Sub ReceiveOK(ByVal strData As String)
    Dim strCommand As String

    On Error Resume Next

    tmrOK.Enabled = False
    tmrOK.Tag = ""

    strCommand = Left(strData, 1)

    Select Case strCommand
        Case CMD_DownChessMan
            Call DisplayCurrentPlayer(mlngPlayer(PLY_YOU))
    End Select
End Sub

Private Sub ReceiveRequest(ByVal Data As String)
    Dim strCommand As String
    Dim strData As String
    Dim CommandHead As String

    On Error Resume Next

    strCommand = Left(Data, 1)
    If Len(Data) > 1 Then strData = Right(Data, Len(Data) - 1)

    CommandHead = CMD_ReRequest & strCommand

    Select Case strCommand
        ' 服务器端(创建棋局的一方)接收此命令
        Case CMD_RequestJoin
            ' 根据棋局类型判断是否同意加入
            If Not mblnBusy Then
                mblnBusy = True
                If gMainTableInfo.TableType = TABLE_LIMIT Then
                    Dim mbrMsgBox As VbMsgBoxResult
                    mbrMsgBox = MessageBox(Me.hWnd, GetDisplayName(GetRecord(strData, 2), GetRecord(strData, 3)) & " 请求加入棋局！" & vbCr & vbCr & "您同意加入吗？", vbQuestion Or vbYesNo, LoadString(182))
                    If mbrMsgBox = vbYes Then
                        Call SendCommand(CommandHead & CMD_Agree & gMainTableInfo.TableName & "|" & gMyUserInfo.UserName & "|" & gMyUserInfo.Name)
                    Else
                        Call SendCommand(CommandHead & CMD_Disagree & LoadString(204))
                    End If
                Else
                    If GetRecord(strData, 1) = gMainTableInfo.TableName Then
                        Call SendCommand(CommandHead & CMD_Agree & gMainTableInfo.TableName & "|" & gMyUserInfo.UserName & "|" & gMyUserInfo.Name)
                    Else
                        Call SendCommand(CommandHead & CMD_Disagree & LoadString(205))
                    End If
                End If
                mblnBusy = False
            End If

        ' 对方接收此命令
        Case CMD_RequestExitGame
            Call CloseModal
            Call frmRequest.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name, gYourUserInfo.Face, CMD_RequestExitGame, strData, False)

        ' -------------------------
        Case CMD_RequestExitTable
            Call CloseModal
            Call frmRequest.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name, gYourUserInfo.Face, CMD_RequestExitTable, strData, False)

        Case CMD_RequestCancelGame
            Call CloseModal
            Call frmRequest.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name, gYourUserInfo.Face, CMD_RequestCancelGame, strData, False)

        ' 对方接收此命令
        Case CMD_RequestLogout
            Call CloseModal
            Call frmRequest.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name, gYourUserInfo.Face, CMD_RequestLogout, strData, False)

        Case CMD_RequestDrawGame
            Call CloseModal
            Call frmRequest.ShowEx(gYourUserInfo.UserName, gYourUserInfo.Name, gYourUserInfo.Face, CMD_RequestDrawGame, strData, False)

    End Select
End Sub

Private Sub ReceiveReRequest(ByVal Data As String)
    Dim strCommand As String
    Dim strData As String
    Dim IsAgree As Boolean

    On Error Resume Next

    strCommand = Left(Data, 1)

    If Mid(Data, 2, 1) = CMD_Agree Then
        IsAgree = True
    Else
        IsAgree = False
    End If

    If Len(Data) > 2 Then strData = Right(Data, Len(Data) - 2)

    Select Case strCommand
        Case CMD_RequestJoin

            ' 客户端接收此命令
            If IsAgree Then
                If JoinTable(gMainTableInfo.TableName) Then
                    Call ConnectStatus(True)
                    Call frmTable.ReloadTable(True)
                    Call frmOnline.ReloadOnline
                    gYourUserInfo.UserName = GetRecord(strData, 2)
                    gYourUserInfo.Name = GetRecord(strData, 3)
                    lblConnect.Caption = GetDisplayName(gYourUserInfo.UserName, gYourUserInfo.Name)
                    lblTable.Caption = gMainTableInfo.TableName
                    lblTips.Caption = LoadString(206)
                    Call SetLight(1, lsLightGreen)
                    Call SetLight(2, lsLightGreen)

                    ' 播放音效
                    Call PlaySoundEffects(SOUND_JOIN_TABLE, gstrSave_SoundValue(SOUND_JOIN_TABLE))

                    ' 显示时间
                    If gMainTableInfo.Timer > 0 Then
                        lblTime.Caption = GetTime(gMainTableInfo.Timer * 60)
                    End If
                    Call SendCommand(CMD_Connected & gMyUserInfo.UserName & "|" & gMyUserInfo.Name)
                Else
                    Call Disconnect
                End If
            Else
                lblTips.Caption = LoadString(207)
                Call MessageBox(Me.hWnd, LoadString(156) & strData, vbInformation, LoadString(183))
                Call Disconnect
            End If

        ' 对方接收此命令
        Case CMD_RequestExitGame
            If IsAgree Then
                mblnAgreeDisconnect = True
                Call CancelGame
                Call ExitTable(True)
                Call Disconnect
                Call MessageBox(Me.hWnd, LoadString(157), vbInformation, LoadString(184))
                Call Unload(Me)
            Else
                mblnAgreeDisconnect = False
                Call MessageBox(Me.hWnd, LoadString(158), vbInformation, LoadString(183))
            End If
        ' -------------------------
        Case CMD_RequestExitTable
            If IsAgree Then
                mblnAgreeDisconnect = True
                Call CancelGame
                Call ExitTable(True)
                Call Disconnect
                Call MessageBox(Me.hWnd, LoadString(157), vbInformation, LoadString(184))
            Else
                mblnAgreeDisconnect = False
                Call MessageBox(Me.hWnd, LoadString(158), vbInformation, LoadString(183))
            End If
        ' -------------------------
        Case CMD_RequestCancelGame
            If IsAgree Then
                Call CancelGame
                Call GameFinish
                If Not (Table Is Nothing) Then
                    Call Table.Clear
                    Call DrawTable
                    Set Table = Nothing
                End If
                Call MessageBox(Me.hWnd, LoadString(159), vbInformation, LoadString(184))
                If Not gblnConnect Then Exit Sub
                lblPlayerTips(GetPlayerIndex(PLY_ME)).Caption = LoadString(198)
                If Not mblnReadyStart(GetPlayerIndex(PLY_YOU)) Then
                    lblPlayerTips(GetPlayerIndex(PLY_YOU)).Caption = LoadString(199)
                End If
                fltbtnReadyStart(GetPlayerIndex(PLY_ME)).Visible = True
            Else
                Call MessageBox(Me.hWnd, LoadString(158), vbInformation, LoadString(183))
            End If

        ' 对方接收此命令
        Case CMD_RequestLogout
            If IsAgree Then
                mblnAgreeDisconnect = True
                Call CancelGame
                Call ExitTable(True)
                Call Disconnect
                Call MessageBox(Me.hWnd, LoadString(157), vbInformation, LoadString(184))
                Call mnuLogout_Click
            Else
                mblnAgreeDisconnect = False
                Call MessageBox(Me.hWnd, LoadString(158), vbInformation, LoadString(183))
            End If

        Case CMD_RequestDrawGame
            If IsAgree Then
                mblnAgreeDisconnect = True
                Call GameOver(GAME_DRAW)
            Else
                mblnAgreeDisconnect = False
                Call MessageBox(Me.hWnd, LoadString(158), vbInformation, LoadString(183))
            End If

    End Select
End Sub

Private Function GetPlayerIndex(ByVal Player As Integer) As Integer
    On Error Resume Next

    If mlngPlayer(Player) < 1 Then
        GetPlayerIndex = 0
    Else
        GetPlayerIndex = mlngPlayer(Player) - 1
    End If
End Function

Private Function SetMousePointer(ByVal Pointer As ePointerStyle)
    On Error Resume Next

    Select Case Pointer
        Case psDefault:
            Set MainPicture.MouseIcon = DefaultCursor
            MainPicture.MousePointer = vbCustom
        Case psBlack:
            Set MainPicture.MouseIcon = BlackCursor
            MainPicture.MousePointer = vbCustom
        Case psWhite:
            Set MainPicture.MouseIcon = WhiteCursor
            MainPicture.MousePointer = vbCustom
        Case psHourglass:
            MainPicture.MousePointer = vbHourglass
        Case Else:
            MainPicture.MousePointer = vbDefault
    End Select
End Function

Public Sub CreateOfflineTable()
    gblnCreator = True
    ' 清除棋局
    If Not (Table Is Nothing) Then
        Set Table = Nothing
    End If
    ' 一些初始化工作
    Call picPlayerFace(0).Cls
    Call picPlayerFace(1).Cls
    shpPlayerTips(0).Visible = True
    shpPlayerTips(1).Visible = True
    Call SetLabel(lblPlayerTips(0), LoadString(126), "", True)
    Call SetLabel(lblPlayerTips(1), LoadString(127), "", True)
    fltbtnSitDown(0).Visible = True
    fltbtnSitDown(1).Visible = True
    Call SetLight(2, lsLightGreen)
    lblTable.Caption = "单机模式棋局"
    lblTips.Caption = LoadString(193)
    ' 显示时间
    lblTime1.Caption = ""
    If gMainTableInfo.Timer > 0 Then
        lblTime.Caption = GetTime(gMainTableInfo.Timer * 60)
    End If
    Call DrawTable
End Sub

Private Sub StartOfflineGame()
    On Error Resume Next

    gblnGameStart = True

    If Not (Table Is Nothing) Then
        Set Table = Nothing
    End If

    ' 播放音效
    Call PlaySoundEffects(SOUND_GAME_START, gstrSave_SoundValue(SOUND_GAME_START))

    ' 创建新的棋局对象
    Set Table = New CTable

    mlngGameSecond = gMainTableInfo.Timer * 60
    mlngSecond = 0
    mlngComputerGameSecond = gMainTableInfo.Timer * 60
    mlngComputerSecond = 0

    mlngStep = 1

    If gMainTableInfo.Timer > 0 Then
        'lblTime.Caption = GetTime(mlngGameSecond)
        'lblTime1.Caption = lblTime.Caption
    End If

    If mlngPlayer(PLY_ME) = T_BLACK Then
        mblnGameFocus = True
        Call SetMousePointer(psDefault)
        Call DisplayCurrentPlayer(T_BLACK)
        ' 开始记时
        If gMainTableInfo.Timer > 0 Then tmrGame.Enabled = True
    Else
        mblnGameFocus = False
        Call SetMousePointer(psPointer)
        Call DisplayCurrentPlayer(T_WHITE)
        Call DrawTable
        If Not ComputerThink() Then
            mblnGameFocus = False
            gblnGameStart = False
        End If
    End If
    ' 开始闪烁当前游戏者的头像
    'tmrFaceFlash.Enabled = True

    ' 判断自己是否可以落子
    If Not Table.CanDown(mlngPlayer(PLY_ME)) Then
        mblnGameFocus = False
        Call SetMousePointer(psPointer)
        Call DisplayCurrentPlayer(mlngPlayer(PLY_YOU))
        tmrGame.Enabled = False
    End If

    ' 画棋盘
    Call DrawTable

    'GameStart = True
End Sub

Private Function OfflineGameOver(ByVal Lose As Boolean) As Boolean
    Dim GameState As Long
    Dim strMessage As String
    'Dim strMessage1 As String

    Dim BlackChessNumber As Long
    Dim WhiteChessNumber As Long

    On Error Resume Next

    If Lose Then
        GameState = GAME_LOSE
    Else
        GameState = Table.Umpire(mlngPlayer(PLY_ME))
    End If
    BlackChessNumber = Table.GetTotal(T_BLACK)
    WhiteChessNumber = Table.GetTotal(T_WHITE)

    Call GameFinish

    If mlngPlayer(PLY_ME) = T_BLACK Then
        strMessage = "人脑" & LoadString(202) & CStr(BlackChessNumber) & vbCr _
                     & "电脑" & LoadString(203) & CStr(WhiteChessNumber) & vbCr
    Else
        strMessage = "人脑" & LoadString(203) & CStr(WhiteChessNumber) & vbCr _
                     & "电脑" & LoadString(202) & CStr(BlackChessNumber) & vbCr
    End If
    'strMessage1 = LoadString(201) & GetRecord(strData, 1)

    If GameState = GAME_WIN Then
        ' 播放音效
        Call PlaySoundEffects(SOUND_GAME_WIN, gstrSave_SoundValue(SOUND_GAME_WIN))
        Call MessageBox(Me.hWnd, LoadString(146) & strMessage, vbInformation, LoadString(179))
    ElseIf GameState = GAME_LOSE Then
        ' 播放音效
        Call PlaySoundEffects(SOUND_GAME_LOSE, gstrSave_SoundValue(SOUND_GAME_LOSE))
        Call MessageBox(Me.hWnd, LoadString(149) & strMessage, vbInformation, LoadString(179))
    Else
        ' 播放音效
        Call PlaySoundEffects(SOUND_GAME_DRAW, gstrSave_SoundValue(SOUND_GAME_DRAW))
        Call MessageBox(Me.hWnd, LoadString(152) & strMessage, vbInformation, LoadString(179))
    End If
    OfflineGameOver = True
End Function

Public Sub ComputerDownMan(ByVal Col As Long, ByVal Row As Long)
    Dim i As Long
    Dim Man As Byte

    On Error Resume Next

    Col = Col - 1: Row = Row - 1
    Man = CByte(mlngPlayer(PLY_YOU))
    If (Not gblnGameStart) Or (Not Table.IsDown(Col, Row, Man)) Then
        Exit Sub
    End If

    ' 播放音效
    Call PlaySoundEffects(SOUND_DOWN_MAN, gstrSave_SoundValue(SOUND_DOWN_MAN))

    LastDown.Col = Col: LastDown.Row = Row: LastMan = Man
    For i = 0 To 1
        Call DrawChessMan(Col, Row, Man, 0)
        Call MainPicture.Refresh
        Call Sleep(100)
        Call DrawChessMan(Col, Row, T_NONE, 0)
        Call MainPicture.Refresh
        Call Sleep(100)
    Next i
    Call Table.DownMan(Col, Row, Man)
    Call DrawTable

    mlngStep = mlngStep + 1

    Call SetMousePointer(psPointer)

    ' 判断你是否可以落子
    If Not Table.CanDown(mlngPlayer(PLY_ME)) Then
        If Not Table.CanDown(mlngPlayer(PLY_YOU)) Then
            Call OfflineGameOver(False)
        Else
            If Not ComputerThink() Then
                mblnGameFocus = False
                gblnGameStart = False
            End If
            ' 播放音效
            Call PlaySoundEffects(SOUND_NOT_DOWN, gstrSave_SoundValue(SOUND_NOT_DOWN))
        End If
        Exit Sub
    End If

    mblnGameFocus = True
    Call SetMousePointer(psDefault)
    ' 提示你落子
    Call DisplayCurrentPlayer(mlngPlayer(PLY_ME))
    ' 开始记时
    If gMainTableInfo.Timer > 0 Then tmrGame.Enabled = True
    ' 重画棋盘
    Call DrawTable
End Sub

Private Function ComputerThink() As Boolean
    Dim i As Long
    Dim j As Long
    Dim Board(7, 7) As Long

    On Error GoTo ErrHandler

    ' 电脑开始计算，计算完成将发送 WM_THINKEND 消息。
    Call SetMousePointer(psHourglass)
    For i = 0 To 7
        For j = 0 To 7
            Board(i, j) = Table.GetMan(i, j)
        Next j
    Next i
    
    Call Think(Me.hWnd, Board(0, 0), glngSave_Level, mlngPlayer(PLY_YOU), mlngStep)
    ComputerThink = True

    Exit Function

ErrHandler:
    Call SetMousePointer(psPointer)
    Call MessageBox(Me.hWnd, "未找到黑白棋.Net动态链接库(Othello.DLL)文件！请立即退出，并重新安装游戏！", vbCritical, "严重")
    ComputerThink = False
End Function

Public Sub BossKey(Optional ByVal ForceShow As Boolean)
    On Error Resume Next

    If mblnHide Or ForceShow Then
        mblnHide = False
        App.TaskVisible = True
        Call RegisterServiceProcess(GetCurrentProcessId(), RSP_UNREGISTER_SERVICE)
        Call ShowWindow(Me.hWnd, SW_SHOW)
        Me.WindowState = vbNormal
        'Me.Visible = True
        Call SetForegroundWindow(Me.hWnd)
        'Call SetActiveWindow(Me.hwnd)
    Else
        mblnHide = True
        Call fltbtnMin_Click(vbLeftButton)
        ' 隐藏任务
        Call ShowWindow(Me.hWnd, SW_HIDE)
        App.TaskVisible = False
        Call RegisterServiceProcess(GetCurrentProcessId(), RSP_SIMPLE_SERVICE)
        Me.Visible = False
    End If
End Sub

