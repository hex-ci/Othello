VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmRegister 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户注册向导 -- 欢迎"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "frmRegister.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin Othello.FlatButton fltbtnBack 
      Height          =   375
      Left            =   1320
      TabIndex        =   60
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "< 上一步(&B)"
      MousePointer    =   99
      Style           =   2
      Enabled         =   0   'False
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
      ForeColor       =   -2147483631
   End
   Begin VB.TextBox txtRegisterBack 
      Height          =   285
      Left            =   285
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   59
      Text            =   "frmRegister.frx":0442
      Top             =   3885
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3990
      Index           =   1
      Left            =   360
      TabIndex        =   53
      Tag             =   "用户注册向导 -- 用户协议  "
      Top             =   -3825
      Visible         =   0   'False
      Width           =   4200
      Begin VB.OptionButton optDisaccord 
         Caption         =   "我不同意(&D)"
         Height          =   300
         Left            =   2265
         TabIndex        =   58
         Top             =   3645
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optAgree 
         Caption         =   "我同意(&A)"
         Height          =   300
         Left            =   420
         TabIndex        =   57
         Top             =   3645
         Width           =   1350
      End
      Begin VB.TextBox txtRegister 
         Height          =   2400
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   1110
         Width           =   3945
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "请仔细阅读下面的游戏规则与说明，您必须同意以下内容，方可继续注册！"
         Height          =   345
         Left            =   180
         TabIndex        =   56
         Top             =   615
         Width           =   3780
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "游戏规则与注册说明。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   225
         TabIndex        =   55
         Top             =   225
         Width           =   2835
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3795
      Index           =   5
      Left            =   6705
      TabIndex        =   47
      Tag             =   "用户注册向导 -- 注册  "
      Top             =   4500
      Visible         =   0   'False
      Width           =   4020
      Begin VB.Frame fraStatus 
         Caption         =   "状态信息"
         Height          =   1500
         Left            =   225
         TabIndex        =   49
         Top             =   765
         Width           =   3390
         Begin VB.Label lblStatus 
            Caption         =   "正在提交信息，请稍候......"
            Height          =   915
            Left            =   360
            TabIndex        =   50
            Top             =   405
            Width           =   2760
         End
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "说明:"
         Height          =   180
         Left            =   270
         TabIndex        =   52
         Top             =   2655
         Width           =   450
      End
      Begin VB.Label Label30 
         Caption         =   "    如果程序长时间没有反应或提示出错，请返回到开始，检查错误，并重新注册！"
         Height          =   420
         Left            =   270
         TabIndex        =   51
         Top             =   2925
         Width           =   3480
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "提交用户资料并注册！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   315
         TabIndex        =   48
         Top             =   225
         Width           =   1965
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3930
      Index           =   4
      Left            =   6750
      TabIndex        =   43
      Tag             =   "用户注册向导 -- 检查信息  "
      Top             =   195
      Visible         =   0   'False
      Width           =   4290
      Begin VB.TextBox txtInfo 
         BackColor       =   &H8000000F&
         Height          =   1905
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   945
         Width           =   4020
      End
      Begin VB.Label Label28 
         Caption         =   "请再次检查您输入的信息，如果确认无误，并且已经连接到了 Internet，请单击“下一步”按钮。"
         Height          =   405
         Left            =   135
         TabIndex        =   46
         Top             =   3105
         Width           =   3960
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "用户资料:"
         Height          =   180
         Left            =   135
         TabIndex        =   45
         Top             =   675
         Width           =   810
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "检查并确认用户输入的资料。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   225
         TabIndex        =   44
         Top             =   225
         Width           =   2550
      End
   End
   Begin InetCtlsObjects.Inet ietRegister 
      Left            =   675
      Top             =   4275
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
      RequestTimeout  =   30
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3840
      Index           =   3
      Left            =   -3825
      TabIndex        =   38
      Tag             =   "用户注册向导 -- 其它信息  "
      Top             =   4365
      Visible         =   0   'False
      Width           =   4380
      Begin VB.Frame fraArea 
         Caption         =   "所在地区"
         Height          =   960
         Left            =   135
         TabIndex        =   41
         Top             =   2115
         Width           =   4065
         Begin VB.TextBox txtCity 
            Height          =   285
            Left            =   2880
            MaxLength       =   20
            TabIndex        =   21
            Top             =   540
            Width           =   1050
         End
         Begin VB.ComboBox cboState 
            Height          =   300
            ItemData        =   "frmRegister.frx":0906
            Left            =   1665
            List            =   "frmRegister.frx":0989
            TabIndex        =   19
            Top             =   540
            Width           =   1185
         End
         Begin VB.ComboBox cboCountry 
            Height          =   300
            ItemData        =   "frmRegister.frx":0AD3
            Left            =   135
            List            =   "frmRegister.frx":0ADA
            TabIndex        =   17
            Top             =   540
            Width           =   1500
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "城市(&I):"
            Height          =   180
            Left            =   2925
            TabIndex        =   20
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "省份(&E):"
            Height          =   180
            Left            =   1710
            TabIndex        =   18
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "国家/地区(&U):"
            Height          =   180
            Left            =   180
            TabIndex        =   16
            Top             =   270
            Width           =   1170
         End
      End
      Begin VB.Frame fraInfo 
         Caption         =   "其它资料"
         Height          =   1365
         Left            =   135
         TabIndex        =   40
         Top             =   585
         Width           =   4065
         Begin VB.TextBox txtAge 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   270
            Left            =   2250
            MaxLength       =   3
            TabIndex        =   15
            Top             =   945
            Width           =   1455
         End
         Begin VB.ComboBox cboSex 
            Height          =   300
            ItemData        =   "frmRegister.frx":0AEE
            Left            =   2250
            List            =   "frmRegister.frx":0AF8
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   570
            Width           =   825
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   2250
            MaxLength       =   15
            TabIndex        =   11
            Top             =   225
            Width           =   1455
         End
         Begin MSComctlLib.ImageCombo imgcboFace 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   225
            TabIndex        =   9
            Top             =   585
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            ForeColor       =   -2147483633
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "年龄(&G):"
            Height          =   180
            Left            =   1440
            TabIndex        =   14
            Top             =   990
            Width           =   720
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "性别(&S):"
            Height          =   180
            Left            =   1440
            TabIndex        =   12
            Top             =   630
            Width           =   720
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "昵称(&A):"
            Height          =   180
            Left            =   1440
            TabIndex        =   10
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "我的肖像(&F):"
            Height          =   180
            Left            =   180
            TabIndex        =   8
            Top             =   270
            Width           =   1080
         End
      End
      Begin VB.Label Label25 
         Caption         =   "说明: 以上内容为非必填项目，可根据个人情况       选择填写。"
         Height          =   375
         Left            =   135
         TabIndex        =   42
         Top             =   3285
         Width           =   3840
      End
      Begin VB.Label Label17 
         Caption         =   "收集用户其它信息。"
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
         Left            =   225
         TabIndex        =   39
         Top             =   225
         Width           =   1950
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3930
      Index           =   2
      Left            =   -4065
      TabIndex        =   24
      Tag             =   "用户注册向导 -- 用户信息  "
      Top             =   165
      Visible         =   0   'False
      Width           =   4200
      Begin VB.Frame fraUserInfo 
         Caption         =   "基本资料"
         Height          =   1905
         Left            =   135
         TabIndex        =   33
         Top             =   585
         Width           =   3840
         Begin VB.TextBox txtEmail 
            Height          =   285
            Left            =   1305
            TabIndex        =   7
            ToolTipText     =   "联系 E-mail，请正确填写。"
            Top             =   1395
            Width           =   2355
         End
         Begin VB.TextBox txtPassword1 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   1305
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   1035
            Width           =   2355
         End
         Begin VB.TextBox txtPassword 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   1305
            PasswordChar    =   "*"
            TabIndex        =   3
            ToolTipText     =   "登陆用的密码。"
            Top             =   675
            Width           =   2355
         End
         Begin VB.TextBox txtUserName 
            Height          =   270
            Left            =   1305
            TabIndex        =   1
            ToolTipText     =   "您的用户名，以后用此登陆。"
            Top             =   315
            Width           =   2355
         End
         Begin VB.Label Label13 
            Caption         =   "&E-mail 地址:"
            Height          =   240
            Left            =   135
            TabIndex        =   6
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "校验密码(&S):"
            Height          =   240
            Left            =   135
            TabIndex        =   4
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label Label11 
            Caption         =   "密码(&P):"
            Height          =   240
            Left            =   135
            TabIndex        =   2
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label10 
            Caption         =   "用户名(&U):"
            Height          =   240
            Left            =   135
            TabIndex        =   0
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.Label Label16 
         Caption         =   "2. 密码 5-15 位，同样不能包含空格等特殊字符。"
         Height          =   375
         Left            =   135
         TabIndex        =   37
         Top             =   3375
         Width           =   3750
      End
      Begin VB.Label Label15 
         Caption         =   "1. 用户名 3-15 位，其中不能包含空格等特殊字符，建议使用字母和数字。"
         Height          =   375
         Left            =   135
         TabIndex        =   36
         Top             =   2970
         Width           =   3750
      End
      Begin VB.Label Label14 
         Caption         =   "说明:"
         Height          =   240
         Left            =   135
         TabIndex        =   35
         Top             =   2700
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "收集用户最基本的信息。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   34
         Top             =   225
         Width           =   2175
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3750
      Index           =   0
      Left            =   2400
      TabIndex        =   23
      Tag             =   "用户注册向导 -- 欢迎  "
      Top             =   -15
      Visible         =   0   'False
      Width           =   4155
      Begin VB.Label Label1 
         Caption         =   "欢迎您使用用户注册向导！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   225
         TabIndex        =   32
         Top             =   225
         Width           =   2400
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "本向导将帮助您快速准确地创建一个属于您自己的账户！"
         Height          =   420
         Left            =   180
         TabIndex        =   31
         Top             =   630
         Width           =   3420
      End
      Begin VB.Label Label3 
         Caption         =   "简要步骤及介绍:"
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   1305
         Width           =   3480
      End
      Begin VB.Label Label4 
         Caption         =   "1. 输入用户相关基本信息。"
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   1665
         Width           =   3660
      End
      Begin VB.Label Label5 
         Caption         =   "2. 帮助用户输入其它信息，包括头像、昵称、    性别、年龄等内容。"
         Height          =   375
         Left            =   180
         TabIndex        =   28
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "3. 提交所有用户信息到服务器。"
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   2370
         Width           =   3750
      End
      Begin VB.Label Label7 
         Caption         =   "4. 完成用户注册。"
         Height          =   255
         Left            =   180
         TabIndex        =   26
         Top             =   2655
         Width           =   3645
      End
      Begin VB.Label Label8 
         Caption         =   "单击“下一步”按钮继续进行。"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   3195
         Width           =   2535
      End
   End
   Begin Othello.FlatButton fltbtnNext 
      Default         =   -1  'True
      Height          =   375
      Left            =   2535
      TabIndex        =   61
      Top             =   4440
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      forecolor       =   0
      font            =   "frmRegister.frx":0B04
      style           =   2
      caption         =   "下一步(&N) >"
      enablehot       =   -1
   End
   Begin Othello.FlatButton fltbtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   62
      Top             =   4440
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      forecolor       =   0
      font            =   "frmRegister.frx":0B28
      style           =   2
      caption         =   "取消(&C)"
      enablehot       =   -1
   End
   Begin Othello.FlatButton fltbtnHelp 
      Height          =   375
      Left            =   5400
      TabIndex        =   63
      Top             =   4440
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      forecolor       =   -2147483631
      enabled         =   0
      font            =   "frmRegister.frx":0B4C
      style           =   2
      caption         =   "帮助"
      enablehot       =   -1
   End
   Begin VB.Image imgWizard 
      BorderStyle     =   1  'Fixed Single
      Height          =   3615
      Left            =   270
      Top             =   225
      Width           =   2115
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   255
      X2              =   6715
      Y1              =   4215
      Y2              =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   255
      X2              =   6715
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Dim TotalStep As Long
Dim Step As Long

Dim mblnServerOK As Boolean

Private Sub fltbtnBack_Click(Button As Integer)
    Select Case Step
        Case 2
            Step = Step - 1
            Call Step1(2, Step)
        Case 3
            Step = Step - 1
            Call Step2(3, Step)
        Case 4
            Step = Step - 1
            Call Step3(4, Step)
        Case 5
            Step = Step - 1
            Call Step4(5, Step, False)
        Case 6
            Step = Step - 1
            Call Step5(6, Step, False)
    End Select
End Sub

Private Sub fltbtnCancel_Click(Button As Integer)
    Call Unload(Me)
End Sub

Private Sub fltbtnNext_Click(Button As Integer)
    On Error Resume Next

    Select Case Step
        Case 1
            Step = Step + 1
            Call Step2(1, Step)
        Case 2
            Step = Step + 1
            Call Step3(2, Step)
        Case 3
            Step = Step + 1
            Call Step4(3, Step, True)
        Case 4
            Step = Step + 1
            Call Step5(4, Step, True)
        Case 5
            Step = Step + 1
            Call Step6(5, Step)
        Case 6
            Step = Step + 1

            Dim Temp As String
            Temp = txtUserName.Text
            Call Unload(Me)
            If Not gblnLogin Then
                frmLogin.cboUserName.Text = Temp
                Call frmLogin.Show(vbModal, MainForm)
            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Long

    On Error Resume Next

    TotalStep = 6
    Step = 1

    Set imgWizard.Picture = frmResource.imgResWizard.Picture
    Set imgcboFace.ImageList = MainForm.ilsFace
    For i = 1 To Val(MainForm.ilsFace.Tag)
        imgcboFace.ComboItems.Add.Image = i
    Next i
    imgcboFace.ComboItems.Item(1).Selected = True

    'Call MoveWindow(cboSex.hwnd, cboSex.Left, cboSex.Top, _
        cboSex.Width, 50, 1)

    txtRegister.Text = txtRegisterBack.Text

    Call DispWizard(0)
End Sub

Private Sub Step1(ByVal Hide As Long, ByVal Disp As Long)
    fraWizard(Hide - 1).Visible = False

    Call DispWizard(Disp - 1)
    fltbtnBack.Enabled = False
    fltbtnNext.Enabled = True
    Call SetControlFocus(fltbtnNext)
End Sub

Private Sub Step2(ByVal Hide As Long, ByVal Disp As Long)
    fraWizard(Hide - 1).Visible = False
    
    Call DispWizard(Disp - 1)
    fltbtnBack.Enabled = True
    fltbtnNext.Enabled = optAgree.Value
    Call SetControlFocus(txtRegister)
End Sub

Private Sub Step3(ByVal Hide As Long, ByVal Disp As Long)
    fraWizard(Hide - 1).Visible = False
    
    Call DispWizard(Disp - 1)
    fltbtnBack.Enabled = True
    Call SetControlFocus(txtUserName)
End Sub

Private Sub Step4(ByVal Hide As Long, ByVal Disp As Long, ByVal Check As Boolean)
    If Check Then
        If Not CheckBaseInfo() Then
            Step = Step - 1
            Exit Sub
        End If
    End If
    fraWizard(Hide - 1).Visible = False
    
    Call DispWizard(Disp - 1)
    Set imgWizard.Picture = frmResource.imgResWizard.Picture
    Call SetControlFocus(txtName)
End Sub

Private Sub Step5(ByVal Hide As Long, ByVal Disp As Long, ByVal Check As Boolean)
    On Error Resume Next

    If Check Then
        If Not CheckInfo() Then
            Step = Step - 1
            Exit Sub
        End If
    End If
    'ietRegister.Cancel
    fraWizard(Hide - 1).Visible = False

    txtInfo.Text = LoadString(222)
    txtInfo.Text = txtInfo.Text & LoadString(223) & txtUserName.Text & vbCrLf
    txtInfo.Text = txtInfo.Text & LoadString(224) & vbCrLf
    txtInfo.Text = txtInfo.Text & LoadString(225) & txtEmail.Text & vbCrLf & vbCrLf

    txtInfo.Text = txtInfo.Text & LoadString(226)
    txtInfo.Text = txtInfo.Text & LoadString(227)
    If txtName.Text = "" Then
        txtInfo.Text = txtInfo.Text & LoadString(204) & vbCrLf
    Else
        txtInfo.Text = txtInfo.Text & txtName.Text & vbCrLf
    End If

    txtInfo.Text = txtInfo.Text & LoadString(228)
    If cboSex.ListIndex = -1 Then
        txtInfo.Text = txtInfo.Text & LoadString(204) & vbCrLf
    Else
        txtInfo.Text = txtInfo.Text & cboSex.Text & vbCrLf
    End If

    txtInfo.Text = txtInfo.Text & LoadString(229)
    If txtAge.Text = "" Then
        txtInfo.Text = txtInfo.Text & LoadString(204) & vbCrLf
    Else
        txtInfo.Text = txtInfo.Text & CStr(Int(Val(txtAge.Text))) & vbCrLf
    End If

    txtInfo.Text = txtInfo.Text & LoadString(230)
    If cboCountry.Text = "" Then
        txtInfo.Text = txtInfo.Text & LoadString(204) & vbCrLf
    Else
        txtInfo.Text = txtInfo.Text & cboCountry.Text & vbCrLf
    End If

    txtInfo.Text = txtInfo.Text & LoadString(231)
    If cboState.Text = "" Then
        txtInfo.Text = txtInfo.Text & LoadString(204) & vbCrLf
    Else
        txtInfo.Text = txtInfo.Text & cboState.Text & vbCrLf
    End If

    txtInfo.Text = txtInfo.Text & LoadString(232)
    If txtCity.Text = "" Then
        txtInfo.Text = txtInfo.Text & LoadString(204)
    Else
        txtInfo.Text = txtInfo.Text & txtCity.Text
    End If

    fltbtnNext.Enabled = True
    Set imgWizard.Picture = frmResource.imgResFinished.Picture
    Call DispWizard(Disp - 1)
    Call SetControlFocus(txtInfo)
End Sub

Private Sub Step6(ByVal Hide As Long, ByVal Disp As Long)
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String
    Dim strID As String

    On Error Resume Next

    fraWizard(Hide - 1).Visible = False
    fltbtnBack.Enabled = False
    fltbtnNext.Enabled = False

    lblStatus.Caption = LoadString(233)
    Call DispWizard(Disp - 1)
    Call SetControlFocus(fltbtnCancel)

    ' 首先获得注册许可(获得安全检查数据)
    strUrl = gstrSave_ServerUrl & SERVER_ACTION_GET
    If ServerCommand(ietRegister, mblnServerOK, strUrl, strStatus, strData, True) Then
        If strStatus = STATUS_OK Then
            strID = GetRecord(strData, 1)
            gstrSecurity1 = GetRecord(strData, 2)
            gstrSecurity2 = GetRecord(strData, 3)
        End If
    End If

    ' 提交信息
    strUrl = gstrSave_ServerUrl & SERVER_ACTION_REGISTER & "?username=" & ToUrlString(txtUserName.Text) & _
                                  "&password=" & MD5(txtPassword.Text) & _
                                  "&email=" & ToUrlString(txtEmail.Text) & _
                                  "&face=" & CStr(imgcboFace.SelectedItem.Index) & _
                                  "&name=" & ToUrlString(txtName.Text) & _
                                  "&country=" & ToUrlString(cboCountry.Text) & _
                                  "&state=" & ToUrlString(cboState.Text) & _
                                  "&city=" & ToUrlString(txtCity.Text)

    If cboSex.Text = "" Then
        strUrl = strUrl & "&sex=0"
    ElseIf cboSex.Text = LoadString(209) Then
        strUrl = strUrl & "&sex=" & CStr(SEX_MAN)
    Else
        strUrl = strUrl & "&sex=" & CStr(SEX_WOMAN)
    End If

    If txtAge.Text = "" Then
        strUrl = strUrl & "&age=0"
    Else
        strUrl = strUrl & "&age=" & CStr(Int(Val(txtAge.Text)))
    End If

    strUrl = strUrl & "&id=" & strID & "&" & MakeServerPassword() & "&" & MakeVersion()

    fltbtnCancel.Enabled = False

    If ServerCommand(ietRegister, mblnServerOK, strUrl, strStatus, strData, True) Then
        fltbtnCancel.Enabled = True
        Select Case strStatus
            Case STATUS_OK
                lblStatus.Caption = LoadString(234)
                fltbtnNext.Caption = LoadString(112)
                fltbtnBack.Enabled = False
                fltbtnCancel.Enabled = False
                fltbtnNext.Enabled = True
                Call SetControlFocus(fltbtnNext)
            Case STATUS_ERROR
                lblStatus.Caption = strData
                fltbtnBack.Enabled = True
                Call SetControlFocus(fltbtnBack)
            Case Else
                lblStatus.Caption = LoadString(235) & LoadString(101)
                fltbtnBack.Enabled = True
                Call SetControlFocus(fltbtnBack)
        End Select
    Else
        fltbtnCancel.Enabled = True
        lblStatus.Caption = LoadString(235) & LoadString(101)
        fltbtnBack.Enabled = True
        Call SetControlFocus(fltbtnBack)
    End If
End Sub

Private Sub ietRegister_StateChanged(ByVal State As Integer)
    'Dim strMess As String '消息变量。
    Select Case State
        Case icResponseReceived
            mblnServerOK = True
        Case icError  '11
            mblnServerOK = False
            '得到错误文本。
            'strMess = "ErrorCode: " & ietRegister.ResponseCode & " : " & ietRegister.ResponseInfo
    End Select
End Sub

Private Sub DispWizard(ByVal Num As Long)
    On Error Resume Next

    With fraWizard(Num)
        .Left = 2385
        .Top = 0
        .Visible = True
        Me.Caption = .Tag & GetStep(Step)
    End With
End Sub
Private Function CheckBaseInfo() As Boolean
    On Error Resume Next

    If StrLen(txtUserName.Text) < 3 Or StrLen(txtUserName.Text) > 15 Or Not CheckString(txtUserName.Text) Then
        CheckBaseInfo = False
        Call MessageBox(hWnd, LoadString(102) & LoadString(106), vbCritical, LoadString(181))
        Call SetControlFocus(txtUserName)
        Call SendKeys("{Home}+{End}")
        Exit Function
    End If
    
    If StrLen(txtPassword.Text) < 5 Or StrLen(txtPassword.Text) > 15 Or Not CheckString(txtPassword.Text) Then
        CheckBaseInfo = False
        Call MessageBox(hWnd, LoadString(103) & LoadString(106), vbCritical, LoadString(181))
        Call SetControlFocus(txtPassword)
        Call SendKeys("{Home}+{End}")
        Exit Function
    End If
    
    If txtPassword1.Text <> txtPassword.Text Then
        CheckBaseInfo = False
        Call MessageBox(hWnd, LoadString(104), vbCritical, LoadString(181))
        txtPassword.Text = ""
        txtPassword1.Text = ""
        Call SetControlFocus(txtPassword)
        Exit Function
    End If
    
    If StrLen(txtEmail.Text) < 3 Or StrLen(txtEmail.Text) > 30 Or InStr(1, txtEmail.Text, "@") < 2 Or Not CheckString(txtEmail.Text) Then
        CheckBaseInfo = False
        Call MessageBox(hWnd, LoadString(105) & LoadString(106), vbCritical, LoadString(181))
        Call SetControlFocus(txtEmail)
        Call SendKeys("{Home}+{End}")
        Exit Function
    End If

    CheckBaseInfo = True
End Function

Private Function CheckInfo() As Boolean
    On Error Resume Next

    If Not CheckString(txtName.Text) Then
        CheckInfo = False
        Call MessageBox(hWnd, LoadString(107), vbCritical, LoadString(181))
        Call SetControlFocus(txtName)
        Call SendKeys("{Home}+{End}")
        Exit Function
    End If

    If txtAge.Text <> "" Then
        If Val(txtAge.Text) < 1 Or Val(txtAge.Text) > 99 Or Not CheckString(txtAge.Text) Then
            CheckInfo = False
            Call MessageBox(hWnd, LoadString(108), vbCritical, LoadString(181))
            Call SetControlFocus(txtAge)
            Call SendKeys("{Home}+{End}")
            Exit Function
        End If
    End If
    If StrLen(cboCountry.Text) > 20 Or Not CheckString(cboCountry.Text) Then
        CheckInfo = False
        Call MessageBox(hWnd, LoadString(109), vbCritical, LoadString(181))
        Call SetControlFocus(cboCountry)
        Call SendKeys("{Home}+{End}")
        Exit Function
    End If

    If StrLen(cboState.Text) > 20 Or Not CheckString(cboState.Text) Then
        CheckInfo = False
        Call MessageBox(hWnd, LoadString(110), vbCritical, LoadString(181))
        Call SetControlFocus(cboState)
        Call SendKeys("{Home}+{End}")
        Exit Function
    End If

    If Not CheckString(txtCity.Text) Then
        CheckInfo = False
        Call MessageBox(hWnd, LoadString(111), vbCritical, LoadString(181))
        Call SetControlFocus(txtCity)
        Call SendKeys("{Home}+{End}")
        Exit Function
    End If

    CheckInfo = True
End Function

Private Sub optAgree_Click()
    fltbtnNext.Enabled = True
End Sub

Private Sub optDisaccord_Click()
    fltbtnNext.Enabled = False
End Sub

Private Sub txtAge_GotFocus()
    Call AutoSelectText(txtAge)
End Sub
Private Sub txtCity_GotFocus()
    Call AutoSelectText(txtCity)
End Sub
Private Sub txtName_GotFocus()
    Call AutoSelectText(txtName)
End Sub
Private Sub txtUserName_GotFocus()
    Call AutoSelectText(txtUserName)
End Sub
Private Sub txtPassword_GotFocus()
    Call AutoSelectText(txtPassword)
End Sub
Private Sub txtPassword1_GotFocus()
    Call AutoSelectText(txtPassword1)
End Sub
Private Sub txtEmail_GotFocus()
    Call AutoSelectText(txtEmail)
End Sub

Private Function GetStep(ByVal Step As Long) As String
    GetStep = "(" & CStr(TotalStep) & " 步骤之第 " & CStr(Step) & " 步)"
End Function
