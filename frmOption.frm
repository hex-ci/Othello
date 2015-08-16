VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmOption 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   4980
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7845
   FillStyle       =   0  'Solid
   Icon            =   "frmOption.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraOption 
      Caption         =   "声音设置"
      Height          =   3135
      Index           =   3
      Left            =   -1665
      TabIndex        =   47
      Top             =   4860
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Timer tmrSound 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4320
         Top             =   135
      End
      Begin VB.CommandButton cmdPlay 
         Height          =   300
         Left            =   4275
         Picture         =   "frmOption.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2745
         Width           =   375
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   300
         Left            =   3825
         Picture         =   "frmOption.frx":0E96
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2745
         Width           =   375
      End
      Begin VB.ComboBox cboSound 
         Height          =   300
         Left            =   930
         TabIndex        =   30
         Top             =   2745
         Width           =   2790
      End
      Begin MSComctlLib.ListView lvwSound 
         Height          =   2205
         Left            =   120
         TabIndex        =   28
         Top             =   465
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   3889
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "File"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList ilsSound 
         Left            =   3450
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOption.frx":0F11
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "文件(&F):"
         Height          =   180
         Left            =   165
         TabIndex        =   29
         Top             =   2790
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "事件(&E):"
         Height          =   180
         Left            =   165
         TabIndex        =   27
         Top             =   255
         Width           =   720
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "音乐设置"
      Height          =   3195
      Index           =   2
      Left            =   3285
      TabIndex        =   46
      Top             =   4890
      Visible         =   0   'False
      Width           =   4860
      Begin MSComctlLib.ListView lvwPlayList 
         Height          =   2475
         Left            =   150
         TabIndex        =   34
         Top             =   570
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   4366
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   1589
         EndProperty
      End
      Begin Othello.FlatButton fltbtnDown 
         Height          =   345
         Left            =   3975
         TabIndex        =   39
         Top             =   2670
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   609
         Caption         =   "向下"
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
         ToolTip         =   "向下移动(Ctrl+↓)"
         ForeColor       =   0
      End
      Begin Othello.FlatButton fltbtnUp 
         Height          =   345
         Left            =   3975
         TabIndex        =   38
         Top             =   2235
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   609
         Caption         =   "向上"
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
         ToolTip         =   "向上移动(Ctrl+↑)"
         ForeColor       =   0
      End
      Begin Othello.FlatButton fltbtnRemove 
         Height          =   345
         Left            =   3975
         TabIndex        =   37
         Top             =   1455
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   609
         Caption         =   "删除"
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
         ToolTip         =   "删除项目(Delete)"
         ForeColor       =   0
      End
      Begin Othello.FlatButton fltbtnAdd 
         Height          =   345
         Left            =   3975
         TabIndex        =   35
         Top             =   585
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   609
         Caption         =   "添加"
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
         ToolTip         =   "添加歌曲(Ctrl+A)"
         ForeColor       =   0
      End
      Begin Othello.FlatButton fltbtnEdit 
         Height          =   345
         Left            =   3975
         TabIndex        =   36
         Top             =   1020
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   609
         Caption         =   "编辑"
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
         ToolTip         =   "编辑播放列表项目"
         ForeColor       =   0
      End
      Begin VB.Label lblPlayListNumber 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   180
         Left            =   3210
         TabIndex        =   49
         Top             =   300
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "歌曲:"
         Height          =   180
         Left            =   2700
         TabIndex        =   48
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "播放列表(&P):"
         Height          =   180
         Left            =   180
         TabIndex        =   33
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "网络设置"
      Height          =   3105
      Index           =   1
      Left            =   7695
      TabIndex        =   45
      Top             =   1110
      Visible         =   0   'False
      Width           =   4785
      Begin MSWinsockLib.Winsock tcpSocks5 
         Left            =   2745
         Top             =   765
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtHttpIP 
         Height          =   285
         Left            =   180
         TabIndex        =   17
         Top             =   1440
         Width           =   3180
      End
      Begin VB.TextBox txtHttpPort 
         Height          =   285
         Left            =   3420
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
      Begin Othello.FlatButton fltbtnTest 
         Height          =   295
         Left            =   3420
         TabIndex        =   26
         Top             =   2610
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "测试"
         MousePointer    =   99
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
      Begin VB.TextBox txtSocks5Password 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1935
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   2610
         Width           =   1410
      End
      Begin VB.TextBox txtSocks5Username 
         Height          =   285
         Left            =   180
         TabIndex        =   24
         Top             =   2610
         Width           =   1695
      End
      Begin VB.TextBox txtSocks5Port 
         Height          =   285
         Left            =   3420
         TabIndex        =   23
         Top             =   2025
         Width           =   1215
      End
      Begin VB.TextBox txtSocks5IP 
         Height          =   285
         Left            =   180
         TabIndex        =   21
         Top             =   2025
         Width           =   3180
      End
      Begin VB.CheckBox chkProxy 
         Caption         =   "使用代理服务器(&X)"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   855
         Width           =   1890
      End
      Begin VB.TextBox txtServerURL 
         Height          =   285
         Left            =   1485
         TabIndex        =   14
         Top             =   360
         Width           =   3165
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&HTTP 代理地址:"
         Height          =   180
         Left            =   225
         TabIndex        =   16
         Top             =   1215
         Width           =   1260
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "端口号(&P):"
         Height          =   180
         Left            =   3465
         TabIndex        =   18
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "校验用户密码:"
         Height          =   180
         Left            =   1980
         TabIndex        =   51
         Top             =   2385
         Width           =   1170
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "校验用户名:"
         Height          =   180
         Left            =   225
         TabIndex        =   50
         Top             =   2385
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "端口号(&T):"
         Height          =   180
         Left            =   3465
         TabIndex        =   22
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Socks&5 代理地址:"
         Height          =   180
         Left            =   225
         TabIndex        =   20
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "服务器地址(&S):"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   420
         Width           =   1275
      End
   End
   Begin Othello.FlatButton fltbtnApply 
      Height          =   360
      Left            =   6210
      TabIndex        =   42
      Top             =   4455
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      Caption         =   "应用(&A)"
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
   Begin Othello.FlatButton fltbtnCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   4875
      TabIndex        =   41
      Top             =   4455
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      Caption         =   "取消(&C)"
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
   Begin Othello.FlatButton fltbtnOK 
      Default         =   -1  'True
      Height          =   360
      Left            =   3540
      TabIndex        =   40
      Top             =   4455
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      Caption         =   "确定(&O)"
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
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2535
      ScaleHeight     =   615
      ScaleWidth      =   5025
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   150
      Width           =   5025
   End
   Begin VB.Frame fraOption 
      Caption         =   "常规设置"
      Height          =   3105
      Index           =   0
      Left            =   2655
      TabIndex        =   43
      Top             =   1080
      Visible         =   0   'False
      Width           =   4710
      Begin MSComctlLib.ImageCombo imgcboFace 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1710
         TabIndex        =   12
         Top             =   2385
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin VB.ComboBox cboLevel 
         Height          =   300
         ItemData        =   "frmOption.frx":1075
         Left            =   1710
         List            =   "frmOption.frx":1094
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1995
         Width           =   1650
      End
      Begin VB.CheckBox chkOfflineMode 
         Caption         =   "单机游戏模式(需要重新启动游戏)(&S)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   1680
         Width           =   3585
      End
      Begin VB.CheckBox chkUserReload 
         Caption         =   "自动刷新用户列表(&U)"
         Height          =   195
         Left            =   270
         TabIndex        =   2
         Top             =   787
         Width           =   2010
      End
      Begin VB.TextBox txtTableTime 
         Height          =   285
         Left            =   3660
         TabIndex        =   7
         Top             =   1125
         Width           =   570
      End
      Begin VB.CheckBox chkTableReload 
         Caption         =   "自动刷新棋局列表(&B)"
         Height          =   195
         Left            =   270
         TabIndex        =   5
         Top             =   1155
         Width           =   2010
      End
      Begin VB.TextBox txtUserTime 
         Height          =   285
         Left            =   3660
         TabIndex        =   4
         Top             =   750
         Width           =   570
      End
      Begin VB.CheckBox chkDownTip 
         Caption         =   "显示落子提示(&D)"
         Height          =   195
         Left            =   270
         TabIndex        =   1
         Top             =   420
         Width           =   1650
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "我的头像(&F):"
         Height          =   180
         Left            =   570
         TabIndex        =   11
         Top             =   2550
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "难度级别(&L):"
         Height          =   180
         Left            =   570
         TabIndex        =   9
         Top             =   2070
         Width           =   1080
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "秒"
         Height          =   180
         Left            =   4320
         TabIndex        =   53
         Top             =   1185
         Width           =   180
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "刷新间隔(&M):"
         Height          =   180
         Left            =   2535
         TabIndex        =   6
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "秒"
         Height          =   180
         Left            =   4320
         TabIndex        =   52
         Top             =   810
         Width           =   180
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "刷新间隔(&T):"
         Height          =   180
         Left            =   2535
         TabIndex        =   3
         Top             =   795
         Width           =   1080
      End
   End
   Begin MSComctlLib.TreeView treOption 
      Height          =   4215
      Left            =   165
      TabIndex        =   0
      Top             =   135
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   7435
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   618
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
   End
   Begin VB.Image imgFooter 
      Height          =   285
      Left            =   315
      Top             =   4530
      Width           =   1650
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngPlayListNumber As Long
Dim mlngOptionPage As Long
Dim SoundEffects As Mmedia
Dim Socks5Status As Long

Public Function ShowEx(ByVal Page As Long) As Long
    Dim Temp As Node

    On Error Resume Next

    If Page < 1 Or Page > treOption.Nodes.Count Then
        Page = 1
    End If
    Set Temp = treOption.Nodes.Item(Page)
    Temp.Selected = True
    Call treOption_NodeClick(Temp)

    Call Me.Show(vbModal, MainForm)

    ShowEx = mlngOptionPage
End Function

Private Sub Form_Load()
    Dim i As Long
    Dim j As Long
    Dim Temp As Node
    Dim itmX As ListItem
    Dim TempSound As String
    Dim Sound(MAX_SOUND) As String

    On Error GoTo ErrorHandler

    Set imgFooter.Picture = frmResource.imgResFooter.Picture

    Set Temp = treOption.Nodes.Add(, tvwFirst, "R1", LoadString(236))
    Temp.Tag = LoadString(237)
    Set Temp = treOption.Nodes.Add("R1", tvwChild, "Set1", LoadString(238))
    Temp.Tag = LoadString(239)
    Set Temp = treOption.Nodes.Add("R1", tvwChild, "Set2", LoadString(240))
    Temp.Tag = LoadString(241)
    Set Temp = treOption.Nodes.Add("R1", tvwChild, "Set3", LoadString(242))
    Temp.Tag = LoadString(243)
    Temp.EnsureVisible

    ' 初始化声音对象
    Set SoundEffects = New Mmedia

    ' 装载设置信息
    chkOfflineMode.Value = Abs(gblnSave_OfflineMode)
    cboLevel.ListIndex = glngSave_Level - 1
    chkDownTip.Value = Abs(gblnSave_DownTip)
    chkUserReload.Value = Abs(gblnSave_OnlineAutoReload)
    chkTableReload.Value = Abs(gblnSave_TableAutoReload)
    txtUserTime.Text = CStr(glngSave_OnlineAutoReloadTime)
    txtTableTime.Text = CStr(glngSave_TableAutoReloadTime)

    ' 初始化单机模式头像下拉框
    Set imgcboFace.ImageList = MainForm.ilsFace
    For i = 1 To Val(MainForm.ilsFace.Tag)
        imgcboFace.ComboItems.Add.Image = i
    Next i
    imgcboFace.ComboItems.Item(glngSave_OfflineFace).Selected = True

    Set lvwSound.SmallIcons = ilsSound

    ' 装载网络设置
    If gstrSave_ServerUrl = "" Then
        txtServerURL.Text = DEFAULT_gServerUrl
    Else
        txtServerURL.Text = ParseURL(gstrSave_ServerUrl, True)
    End If
    chkProxy.Value = Abs(gblnSave_UseProxy)
    txtHttpIP.Text = gstrSave_HttpProxyIP
    txtHttpPort.Text = CStr(glngSave_HttpProxyPort)
    txtSocks5IP.Text = gstrSave_Socks5ProxyIP
    txtSocks5Port.Text = CStr(glngSave_Socks5ProxyPort)
    txtSocks5Username.Text = gstrSave_Socks5Username
    txtSocks5Password.Text = Decipher(gstrLocalPassword, gstrSave_Socks5Password)
    Call chkProxy_Click

    ' 装载播放列表
    lvwPlayList.ColumnHeaders(1).Width = lvwPlayList.Width - 310
    mlngPlayListNumber = glngSave_PlayListNumber
    For i = 1 To mlngPlayListNumber
        Set itmX = lvwPlayList.ListItems.Add(, , ExtractName(gstrSave_PlayListName(i)))
        itmX.Tag = gstrSave_PlayListName(i)
    Next i
    If glngSave_PlayListPosition = 0 And glngSave_PlayListNumber > 0 Then
        glngSave_PlayListPosition = 1
    End If
    If glngSave_PlayListPosition <= lvwPlayList.ListItems.Count And glngSave_PlayListPosition > 0 And mlngPlayListNumber > 0 Then
        lvwPlayList.ListItems(glngSave_PlayListPosition).ForeColor = vbBlue
        lvwPlayList.ListItems(glngSave_PlayListPosition).Selected = True
        Call lvwPlayList.ListItems(glngSave_PlayListPosition).EnsureVisible
    End If
    lblPlayListNumber.Caption = CStr(mlngPlayListNumber)

    ' 装载声音列表
    For i = 1 To MAX_SOUND
        Set itmX = lvwSound.ListItems.Add(, , gstrSave_SoundName(i))
        If gstrSave_SoundValue(i) = DEFAULT_SOUND Then
            itmX.Tag = DEFAULT_SOUND
            itmX.SubItems(1) = LoadString(244)
            itmX.SmallIcon = 1
        ElseIf gstrSave_SoundValue(i) = "" Then
            itmX.Tag = ""
            itmX.SubItems(1) = ""
        Else
            itmX.Tag = gstrSave_SoundValue(i)
            itmX.SubItems(1) = ExtractName(gstrSave_SoundValue(i))
            itmX.SmallIcon = 1
        End If
    Next i
    Call cboSound.AddItem(LoadString(244))
    Call cboSound.AddItem(LoadString(245))
    For i = 1 To MAX_SOUND
        TempSound = gstrSave_SoundValue(i)
        If TempSound <> DEFAULT_SOUND And TempSound <> "" Then
            Sound(i) = TempSound
        End If
    Next i
    ' 处理声音列表
    For i = 1 To MAX_SOUND - 1
        For j = i + 1 To MAX_SOUND
            If Sound(i) = Sound(j) Then Sound(j) = ""
        Next j
    Next i
    For i = 1 To MAX_SOUND - 1
        For j = i + 1 To MAX_SOUND
            If Sound(i) = "" And Sound(j) <> "" Then
                Call Swap(Sound(i), Sound(j))
                Exit For
            End If
        Next j
    Next i
    For i = 1 To MAX_SOUND
        If Sound(i) = "" Then Exit For
        Call cboSound.AddItem(Sound(i))
    Next i
    Call lvwSound_ItemClick(lvwSound.ListItems(1))
    Call ColumnSize(lvwSound, 1, 38)
    Call ColumnSize(lvwSound, 2, 55)

    ' 更新控件
    Call chkOfflineMode_Click
    Call chkUserReload_Click
    Call chkTableReload_Click
    Call chkProxy_Click

    Exit Sub

ErrorHandler:
    Call MessageBox(Me.hWnd, Err.Description, vbCritical, LoadString(181))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next

    Call tcpSocks5.Close
    If Not (treOption.SelectedItem Is Nothing) Then
        mlngOptionPage = treOption.SelectedItem.Index
    Else
        mlngOptionPage = 1
    End If

    Call SoundEffects.mmClose
    Set SoundEffects = Nothing
End Sub

' 使所有更改生效！
Private Sub fltbtnApply_Click(Button As Integer)
    Dim i As Long

    On Error Resume Next

    ' 保存常规设置
    gblnSave_DownTip = CBool(chkDownTip.Value)
    gblnSave_OfflineMode = CBool(chkOfflineMode.Value)
    glngSave_Level = cboLevel.ListIndex + 1
    glngSave_OfflineFace = imgcboFace.SelectedItem.Index
    gblnSave_OnlineAutoReload = CBool(chkUserReload.Value)
    gblnSave_TableAutoReload = CBool(chkTableReload.Value)
    glngSave_OnlineAutoReloadTime = Int(txtUserTime.Text)
    glngSave_TableAutoReloadTime = Int(txtTableTime.Text)

    ' 保存网络设置
    If txtServerURL.Text = "" Then
        gstrSave_ServerUrl = ParseURL(DEFAULT_gServerUrl, False)
    Else
        gstrSave_ServerUrl = ParseURL(txtServerURL.Text, False)
    End If
    gblnSave_UseProxy = CBool(chkProxy.Value)
    gstrSave_HttpProxyIP = txtHttpIP.Text
    glngSave_HttpProxyPort = Int(txtHttpPort.Text)
    gstrSave_Socks5ProxyIP = txtSocks5IP.Text
    glngSave_Socks5ProxyPort = Int(txtSocks5Port.Text)
    gstrSave_Socks5Username = txtSocks5Username.Text
    gstrSave_Socks5Password = Cipher(gstrLocalPassword, txtSocks5Password.Text)

    ' 保存播放列表
    glngSave_PlayListNumber = lvwPlayList.ListItems.Count
    For i = 1 To glngSave_PlayListNumber
        gstrSave_PlayListName(i) = lvwPlayList.ListItems(i).Tag
    Next i
    If lvwPlayList.SelectedItem Is Nothing Then
        glngSave_PlayListPosition = 1
    Else
        glngSave_PlayListPosition = lvwPlayList.SelectedItem.Index
    End If

    ' 保存声音列表
    For i = 1 To MAX_SOUND
        gstrSave_SoundValue(i) = lvwSound.ListItems.Item(i).Tag
    Next i

    frmOnline.tmrReload.Enabled = gblnSave_OnlineAutoReload
    frmTable.tmrReload.Enabled = gblnSave_TableAutoReload
    Call MainForm.DrawTable
End Sub

Private Sub cboSound_Click()
    Call SetSoundList(cboSound.Text)
End Sub

Private Sub cboSound_LostFocus()
    Call SetSoundList(cboSound.Text)
End Sub

Private Sub chkOfflineMode_Click()
    If (chkOfflineMode.Value = vbChecked) And (Not gblnGameStart) Then
        cboLevel.Enabled = True
        imgcboFace.Enabled = True
    Else
        cboLevel.Enabled = False
        imgcboFace.Enabled = False
    End If
End Sub

Private Sub chkProxy_Click()
    If chkProxy.Value = vbChecked Then
        Call SetEnabled(txtHttpIP, True)
        Call SetEnabled(txtHttpPort, True)
        Call SetEnabled(txtSocks5IP, True)
        Call SetEnabled(txtSocks5Port, True)
        Call SetEnabled(txtSocks5Username, True)
        Call SetEnabled(txtSocks5Password, True)
        fltbtnTest.Enabled = True
    Else
        Call SetEnabled(txtHttpIP, False)
        Call SetEnabled(txtHttpPort, False)
        Call SetEnabled(txtSocks5IP, False)
        Call SetEnabled(txtSocks5Port, False)
        Call SetEnabled(txtSocks5Username, False)
        Call SetEnabled(txtSocks5Password, False)
        fltbtnTest.Enabled = False
    End If
End Sub

Private Sub chkTableReload_Click()
    If chkTableReload.Value = vbChecked Then
        Call SetEnabled(txtTableTime, True)
    Else
        Call SetEnabled(txtTableTime, False)
    End If
End Sub

Private Sub chkUserReload_Click()
    If chkUserReload.Value = vbChecked Then
        Call SetEnabled(txtUserTime, True)
    Else
        Call SetEnabled(txtUserTime, False)
    End If
End Sub

Private Sub cmdOpen_Click()
    Dim FileName As String
    Dim FilePath As String
    Dim i As Long

    On Error Resume Next

    FilePath = gstrSave_SoundPath
    If cboSound.Text <> LoadString(245) And cboSound.Text <> LoadString(244) Then
        If ExtractPath(cboSound.Text) <> "" Then
            FilePath = ExtractPath(cboSound.Text)
        End If
    End If

    FileName = DialogFile(Me.hWnd, 1, _
                      LoadString(246), _
                      "", _
                      LoadString(247), _
                      FilePath, _
                      "")

    FileName = Trim(Replace(FileName, vbNullChar, " "))

    If FileName <> "" Then
        gstrSave_SoundPath = FilePath
        cboSound.Text = FileName
        Call SetSoundList(FileName)

        For i = 2 To cboSound.ListCount - 1
            If FileName = cboSound.List(i) Then
                Exit For
            End If
        Next i
        ' 如果没有则添加
        If i = cboSound.ListCount Then
            Call cboSound.AddItem(FileName)
        End If
    End If
End Sub

Private Sub cmdPlay_Click()
    On Error Resume Next

    If cmdPlay.Tag = "playing" Then
        Set cmdPlay.Picture = objSoundPlay
        tmrSound.Enabled = False
        Call SoundEffects.mmStop
        cmdPlay.Tag = SoundEffects.Status
        Exit Sub
    End If

    Select Case cboSound.Text
        Case LoadString(245)    ' 无
        Case LoadString(244)    ' 默认
            ' 播放默认声音
            Call PlaySoundEffects(lvwSound.SelectedItem.Index, lvwSound.SelectedItem.Tag)
        Case Else
            If FileExisted(cboSound.Text) Then
                Set cmdPlay.Picture = objSoundStop
                Me.MousePointer = vbHourglass
                Call SoundEffects.mmOpen(cboSound.Text)
                Call SoundEffects.mmPlay
                tmrSound.Enabled = True
                Me.MousePointer = vbDefault
                cmdPlay.Tag = SoundEffects.Status
            Else
                Call MessageBox(Me.hWnd, LoadString(162), vbCritical, LoadString(185))
            End If
    End Select
End Sub

Private Sub fltbtnAdd_Click(Button As Integer)
    Dim Temp As String
    Dim Temp1 As String
    Dim i As Long
    Dim BasePath As String
    Dim itmX As ListItem

    On Error Resume Next

    Temp = DialogFile(Me.hWnd, 2, _
                      LoadString(248), _
                      "", _
                      LoadString(249), _
                      gstrSave_PlayListPath, _
                      "")

    'Debug.Print GetCount(Temp, vbNullChar)
    'If mlngPlayListNumber + Number > MAX_PLAY_LIST Then
    '    Call MessageBox(Me.hwnd, "注意！播放列表最大只能容纳 " & CStr(MAX_PLAY_LIST) & " 首歌曲！", vbExclamation, "错误")
    '    Exit Sub
    'End If

    If Temp <> "" Then
        BasePath = GetInfo(Temp, 1, vbNullChar)
        gstrSave_PlayListPath = ExtractPath(BasePath)
        If GetInfo(Temp, 2, vbNullChar) = "" Then
            gstrSave_PlayListPath = ExtractPath(BasePath)
            If mlngPlayListNumber >= MAX_PLAY_LIST Then
                lblPlayListNumber.Caption = CStr(mlngPlayListNumber)
                Call MessageBox(Me.hWnd, "注意！播放列表最大只能容纳 " & CStr(MAX_PLAY_LIST) & " 首歌曲！", vbExclamation, LoadString(181))
                Exit Sub
            End If
            'Temp1 = GetInfo(Temp, 1, vbNullChar)
            Set itmX = lvwPlayList.ListItems.Add(, , ExtractName(BasePath))
            itmX.Tag = BasePath
            mlngPlayListNumber = mlngPlayListNumber + 1
        Else
            gstrSave_PlayListPath = BasePath
            For i = 2 To 255
                If GetInfo(Temp, i, vbNullChar) = "" Then Exit For
                If mlngPlayListNumber >= MAX_PLAY_LIST Then
                    Call MessageBox(Me.hWnd, "注意！播放列表最大只能容纳 " & CStr(MAX_PLAY_LIST) & " 首歌曲！", vbExclamation, LoadString(181))
                    Exit For
                End If
                Set itmX = lvwPlayList.ListItems.Add(, , GetInfo(Temp, i, vbNullChar))
                itmX.Tag = BasePath & "\" & GetInfo(Temp, i, vbNullChar)
                mlngPlayListNumber = mlngPlayListNumber + 1
            Next i
        End If
        lblPlayListNumber.Caption = CStr(mlngPlayListNumber)
    End If
    Call SetControlFocus(lvwPlayList)
End Sub

Private Sub fltbtnCancel_Click(Button As Integer)
    Call Unload(Me)
End Sub

Private Sub fltbtnDown_Click(Button As Integer)
    Dim Name As String
    Dim FullName As String

    On Error Resume Next

    If (lvwPlayList.SelectedItem Is Nothing) Then Exit Sub

    If lvwPlayList.SelectedItem.Index < lvwPlayList.ListItems.Count Then
        Name = lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index).Text
        FullName = lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index).Tag

        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index).Text = lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index + 1).Text
        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index).Tag = lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index + 1).Tag

        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index + 1).Text = Name
        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index + 1).Tag = FullName

        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index + 1).Selected = True
    End If
    Call lvwPlayList.SelectedItem.EnsureVisible
    Call SetControlFocus(lvwPlayList)
End Sub

Private Sub fltbtnEdit_Click(Button As Integer)
    Dim Temp As String

    On Error Resume Next

    If lvwPlayList.SelectedItem Is Nothing Then Exit Sub
    Temp = frmEditPlayList.ShowEx(lvwPlayList.SelectedItem.Tag)
    If Temp <> "" Then
        lvwPlayList.SelectedItem.Text = ExtractName(Temp)
        lvwPlayList.SelectedItem.Tag = Temp
    End If
    Call SetControlFocus(lvwPlayList)
End Sub

Private Sub fltbtnOK_Click(Button As Integer)
    Call fltbtnApply_Click(Button)
    Call Unload(Me)
End Sub

Private Sub fltbtnRemove_Click(Button As Integer)
    On Error Resume Next

    If (lvwPlayList.SelectedItem Is Nothing) Then Exit Sub

    Call lvwPlayList.ListItems.Remove(lvwPlayList.SelectedItem.Index)
    If Not (lvwPlayList.SelectedItem Is Nothing) Then
        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index).Selected = True
    End If
    mlngPlayListNumber = mlngPlayListNumber - 1
    lblPlayListNumber.Caption = CStr(mlngPlayListNumber)
    Call SetControlFocus(lvwPlayList)
End Sub

Private Sub fltbtnTest_Click(Button As Integer)
    On Error GoTo ErrorHandler

    fltbtnTest.Enabled = False
    Socks5Status = 1
    Call tcpSocks5.Connect(txtSocks5IP.Text, Int(txtSocks5Port.Text))

    Exit Sub

ErrorHandler:
    Call tcpSocks5.Close
    Call MessageBox(Me.hWnd, LoadString(163), vbExclamation, LoadString(186))
    fltbtnTest.Enabled = True
End Sub

Private Sub fltbtnUp_Click(Button As Integer)
    Dim Name As String
    Dim FullName As String

    On Error Resume Next

    If (lvwPlayList.SelectedItem Is Nothing) Then Exit Sub

    If lvwPlayList.SelectedItem.Index > 1 Then
        Name = lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index).Text
        FullName = lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index).Tag

        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index).Text = lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index - 1).Text
        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index).Tag = lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index - 1).Tag

        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index - 1).Text = Name
        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index - 1).Tag = FullName

        lvwPlayList.ListItems.Item(lvwPlayList.SelectedItem.Index - 1).Selected = True
    End If
    Call lvwPlayList.SelectedItem.EnsureVisible
    Call SetControlFocus(lvwPlayList)
End Sub

Private Sub imgcboFace_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
    End If
End Sub

Private Sub lvwPlayList_DblClick()
    Call fltbtnEdit_Click(vbLeftButton)
End Sub

Private Sub lvwPlayList_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyDelete Then
        Call fltbtnRemove_Click(vbLeftButton)
    End If

    If (Shift And vbCtrlMask) = vbCtrlMask Then
        If KeyCode = vbKeyUp Then
            Call fltbtnUp_Click(vbLeftButton)
        End If
        If KeyCode = vbKeyDown Then
            Call fltbtnDown_Click(vbLeftButton)
        End If
        If KeyCode = vbKeyA Then
            Call fltbtnAdd_Click(vbLeftButton)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub lvwSound_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Tag = DEFAULT_SOUND Then
        cboSound.Text = LoadString(244)
    ElseIf Item.Tag = "" Then
        cboSound.Text = LoadString(245)
    Else
        cboSound.Text = Item.Tag
    End If
End Sub

Private Sub tcpSocks5_Close()
    Call tcpSocks5.Close
    fltbtnTest.Enabled = True
End Sub

Private Sub tcpSocks5_Connect()
    Dim Mem(3) As Byte

    On Error Resume Next

    Mem(0) = 5
    Mem(1) = 2
    Mem(2) = 0
    Mem(3) = 2

    Socks5Status = 2
    Call tcpSocks5.SendData(Mem())
End Sub

Private Sub tcpSocks5_DataArrival(ByVal bytesTotal As Long)
    Dim Mem(9) As Byte
    Dim Socks5Data As String
    Dim i As Integer

    On Error Resume Next

    Call tcpSocks5.GetData(Socks5Data, vbString, bytesTotal)
    If Socks5Status = 2 Then
        ' 需要验证，则发送验证信息
        If AscW(Mid(Socks5Data, 2, 1)) = 2 Then
            Dim Anthreq As String

            Anthreq = Anthreq & Chr(5)
            Anthreq = Anthreq & Chr(Len(txtSocks5Username.Text))
            Anthreq = Anthreq & txtSocks5Username.Text
            Anthreq = Anthreq & Chr(Len(txtSocks5Password.Text))
            Anthreq = Anthreq & txtSocks5Password.Text

            Socks5Status = 3
            Call tcpSocks5.SendData(Anthreq)
            Exit Sub
        End If

        Mem(0) = 5
        Mem(1) = 1
        Mem(2) = 0
        Mem(3) = 1
        Mem(4) = 127
        Mem(5) = 0
        Mem(6) = 0
        Mem(7) = 1
        Mem(8) = (&HFF00 And Int(txtSocks5Port.Text)) \ &H100
        Mem(9) = &HFF And Int(txtSocks5Port.Text)

        Socks5Status = 4
        Call tcpSocks5.SendData(Mem())

    ElseIf Socks5Status = 3 Then
        If AscW(Mid(Socks5Data, 2, 1)) <> 0 Then
            Call MessageBox(Me.hWnd, LoadString(164), vbExclamation, LoadString(186))
            Call tcpSocks5.Close
            fltbtnTest.Enabled = True
            Exit Sub
        End If

        Mem(0) = 5
        Mem(1) = 1
        Mem(2) = 0
        Mem(3) = 1
        Mem(4) = 127
        Mem(5) = 0
        Mem(6) = 0
        Mem(7) = 1
        Mem(8) = (&HFF00 And Int(txtSocks5Port.Text)) \ &H100
        Mem(9) = &HFF And Int(txtSocks5Port.Text)

        Socks5Status = 4
        Call tcpSocks5.SendData(Mem())
    ElseIf Socks5Status = 4 Then
        If AscW(Mid(Socks5Data, 2, 1)) <> 0 Then
            Call MessageBox(Me.hWnd, LoadString(163), vbExclamation, LoadString(186))
        Else
            Call MessageBox(Me.hWnd, LoadString(166), vbInformation, LoadString(186))
        End If
        Call tcpSocks5.Close
        fltbtnTest.Enabled = True
    End If
End Sub

Private Sub tcpSocks5_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Number <> sckSuccess Then
        Call tcpSocks5.Close
        Call MessageBox(Me.hWnd, LoadString(163), vbExclamation, LoadString(186))
        fltbtnTest.Enabled = True
    End If
End Sub

Private Sub tmrSound_Timer()
    If SoundEffects.Status = "stopped" Then
        tmrSound.Enabled = False
        cmdPlay.Tag = SoundEffects.Status
        Set cmdPlay.Picture = objSoundPlay
    End If
End Sub

Private Sub treOption_Collapse(ByVal Node As MSComctlLib.Node)
    Call treOption_NodeClick(Node)
End Sub

Private Sub treOption_Expand(ByVal Node As MSComctlLib.Node)
    Call treOption_NodeClick(Node)
End Sub

Private Sub treOption_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim i As Long

    On Error Resume Next

    For i = 0 To fraOption.Count - 1
        fraOption(i).Visible = False
    Next i
    Call DisplayItem(Node)
    Call fraOption(Node.Index - 1).Move(2625, 990, 4815, 3180)
    fraOption(Node.Index - 1).Caption = Node.Text
    fraOption(Node.Index - 1).Visible = True
End Sub

Private Sub DisplayTitle(ByVal Name As String, ByVal Tips As String)
    Call Gradient(picTitle, 0, 0, 255)
    picTitle.CurrentX = 100
    picTitle.CurrentY = 80
    picTitle.FontBold = True
    picTitle.Print Name
    picTitle.CurrentX = 250
    picTitle.CurrentY = 320
    picTitle.FontBold = False
    picTitle.Print Tips
End Sub

Private Sub DisplayItem(ByVal Item As Node)
    Call DisplayTitle(Item.FullPath, Item.Tag)
End Sub

Private Sub SetSoundList(ByVal SoundText As String)
    On Error Resume Next

    Select Case SoundText
        Case LoadString(245)    ' 无
            If lvwSound.SelectedItem.Tag <> "" Then
                lvwSound.SelectedItem.Tag = ""
                lvwSound.SelectedItem.SubItems(1) = ""
                lvwSound.SelectedItem.SmallIcon = 0
            End If
        Case LoadString(244)    ' 默认
            If lvwSound.SelectedItem.Tag <> DEFAULT_SOUND Then
                lvwSound.SelectedItem.Tag = DEFAULT_SOUND
                lvwSound.SelectedItem.SubItems(1) = SoundText
                lvwSound.SelectedItem.SmallIcon = 1
            End If
        Case Else
            If lvwSound.SelectedItem.Tag <> SoundText Then
                If FileExisted(SoundText) Then
                    lvwSound.SelectedItem.Tag = SoundText
                    lvwSound.SelectedItem.SubItems(1) = ExtractName(SoundText)
                    lvwSound.SelectedItem.SmallIcon = 1
                Else
                    Call MessageBox(Me.hWnd, LoadString(162), vbCritical, LoadString(185))
                End If
            End If
    End Select
End Sub

Private Sub SetEnabled(ByRef objControl As Object, ByVal Status As Boolean)
    On Error Resume Next

    objControl.Enabled = Status
    If Status Then
        objControl.BackColor = vbWindowBackground
    Else
        objControl.BackColor = vbButtonFace
    End If
End Sub

Private Sub txtHttpIP_GotFocus()
    Call AutoSelectText(txtHttpIP)
End Sub

Private Sub txtHttpPort_GotFocus()
    Call AutoSelectText(txtHttpPort)
End Sub

Private Sub txtServerURL_GotFocus()
    Call AutoSelectText(txtServerURL)
End Sub

Private Sub txtSocks5IP_GotFocus()
    Call AutoSelectText(txtSocks5IP)
End Sub

Private Sub txtSocks5Password_GotFocus()
    Call AutoSelectText(txtSocks5Password)
End Sub

Private Sub txtSocks5Port_GotFocus()
    Call AutoSelectText(txtSocks5Port)
End Sub

Private Sub txtSocks5Username_GotFocus()
    Call AutoSelectText(txtSocks5Username)
End Sub

Private Sub txtTableTime_GotFocus()
    Call AutoSelectText(txtTableTime)
End Sub

Private Sub txtUserTime_GotFocus()
    Call AutoSelectText(txtUserTime)
End Sub
