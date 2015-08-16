VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUserInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户信息"
   ClientHeight    =   3660
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5010
   Icon            =   "frmUserInfo.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5010
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   240
      TabIndex        =   38
      Top             =   495
      Width           =   795
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   120
         Picture         =   "frmUserInfo.frx":0E42
         Top             =   405
         Width           =   480
      End
   End
   Begin InetCtlsObjects.Inet ietView 
      Left            =   4935
      Top             =   2415
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
      RequestTimeout  =   30
   End
   Begin Othello.FlatButton fltbtnClose 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   2925
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   661
      Caption         =   "关闭"
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
   Begin Othello.FlatButton fltbtnReload 
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2460
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   661
      Caption         =   "刷新"
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
   Begin VB.Frame fraTabs 
      BorderStyle     =   0  'None
      Height          =   3045
      Index           =   1
      Left            =   975
      TabIndex        =   28
      Top             =   3540
      Visible         =   0   'False
      Width           =   3780
      Begin VB.Frame fraInfo 
         Caption         =   "其它信息"
         Height          =   2850
         Left            =   0
         TabIndex        =   29
         Top             =   90
         Width           =   3600
         Begin VB.TextBox txtAge 
            BackColor       =   &H8000000F&
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   10
            Top             =   1665
            Width           =   735
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   8
            Top             =   585
            Width           =   1785
         End
         Begin VB.TextBox txtSex 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2655
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1665
            Width           =   735
         End
         Begin VB.TextBox txtCity 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   2325
            Width           =   1560
         End
         Begin VB.TextBox txtState 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   195
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   2325
            Width           =   1530
         End
         Begin VB.TextBox txtCountry 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   195
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1665
            Width           =   1545
         End
         Begin VB.PictureBox picFace 
            AutoRedraw      =   -1  'True
            Height          =   630
            Left            =   420
            ScaleHeight     =   570
            ScaleWidth      =   645
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   585
            Width           =   705
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "年龄:"
            Height          =   180
            Left            =   1845
            TabIndex        =   37
            Top             =   1425
            Width           =   450
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "性别:"
            Height          =   180
            Left            =   2670
            TabIndex        =   36
            Top             =   1425
            Width           =   450
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "昵称:"
            Height          =   180
            Left            =   1680
            TabIndex        =   35
            Top             =   315
            Width           =   450
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "头像:"
            Height          =   180
            Left            =   225
            TabIndex        =   34
            Top             =   285
            Width           =   450
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "国家/地区:"
            Height          =   180
            Left            =   210
            TabIndex        =   33
            Top             =   1425
            Width           =   900
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "省份:"
            Height          =   180
            Left            =   210
            TabIndex        =   32
            Top             =   2100
            Width           =   450
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "城市:"
            Height          =   180
            Left            =   1845
            TabIndex        =   31
            Top             =   2100
            Width           =   450
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            X1              =   1365
            X2              =   3435
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   1365
            X2              =   3435
            Y1              =   1170
            Y2              =   1170
         End
      End
   End
   Begin VB.Frame fraTabs 
      BorderStyle     =   0  'None
      Height          =   3045
      Index           =   0
      Left            =   1080
      TabIndex        =   17
      Top             =   420
      Visible         =   0   'False
      Width           =   3780
      Begin VB.Frame fraStatus 
         Caption         =   "比赛信息"
         Height          =   1440
         Left            =   0
         TabIndex        =   21
         Top             =   1500
         Width           =   3600
         Begin VB.TextBox txtWin 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   225
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   570
            Width           =   765
         End
         Begin VB.TextBox txtLose 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   570
            Width           =   765
         End
         Begin VB.TextBox txtDraw 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   570
            Width           =   765
         End
         Begin VB.TextBox txtWinP 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   2655
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   570
            Width           =   765
         End
         Begin VB.TextBox txtScore 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   975
            Width           =   765
         End
         Begin VB.TextBox txtDisconnectP 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   2655
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   975
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "获胜:"
            Height          =   180
            Left            =   240
            TabIndex        =   27
            Top             =   330
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "失败:"
            Height          =   180
            Left            =   1050
            TabIndex        =   26
            Top             =   330
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "平局:"
            Height          =   180
            Left            =   1860
            TabIndex        =   25
            Top             =   330
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "胜率:"
            Height          =   180
            Left            =   2670
            TabIndex        =   24
            Top             =   330
            Width           =   450
         End
         Begin VB.Label Label7 
            Caption         =   "总积分:"
            Height          =   180
            Left            =   285
            TabIndex        =   23
            Top             =   1020
            Width           =   645
         End
         Begin VB.Label Label8 
            Caption         =   "断线率:"
            Height          =   180
            Left            =   1935
            TabIndex        =   22
            Top             =   1020
            Width           =   645
         End
      End
      Begin VB.Frame fraBaseInfo 
         Caption         =   "基本信息"
         Height          =   1320
         Left            =   0
         TabIndex        =   18
         Top             =   90
         Width           =   3600
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   330
            Width           =   2505
         End
         Begin VB.TextBox txtEmail 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   780
            Width           =   2505
         End
         Begin VB.Label Label1 
            Caption         =   "用户名:"
            Height          =   180
            Left            =   225
            TabIndex        =   20
            Top             =   390
            Width           =   630
         End
         Begin VB.Label Label2 
            Caption         =   "E-mail:"
            Height          =   180
            Left            =   225
            TabIndex        =   19
            Top             =   840
            Width           =   630
         End
      End
   End
   Begin MSComctlLib.TabStrip tabInfo 
      Height          =   3420
      Left            =   135
      TabIndex        =   16
      Top             =   90
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   6033
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本资料"
            Key             =   "BaseInfo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "详细资料"
            Key             =   "Info"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrUserName As String

Dim mblnServerOK As Boolean

Dim FormVisible As Boolean

Public Sub ShowEx(ByVal UserName As String, ByVal Name As String)
    On Error Resume Next

    mstrUserName = UserName
    FormVisible = True
    Call Show(vbModeless)
    If Me.WindowState <> vbNormal Then
        Me.WindowState = vbNormal
    End If

    'Me.Caption = LoadString(208) & GetDisplayName(UserName, Name)

    Call Clear
    Call Me.Refresh
    Call ReloadInfo(UserName)
End Sub

Public Sub FormMinimize()
    If Me.Visible Then
        Call Me.Hide
    End If
End Sub

Public Sub FormNormal()
    If FormVisible Then
        Call Me.Show(vbModeless)
    End If
End Sub

Private Function ReloadInfo(ByVal UserName As String) As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    On Error Resume Next

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_USER_VIEW & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&name=" & ToUrlString(UserName) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If ServerCommand(ietView, mblnServerOK, strUrl, strStatus, strData) Then
        If strStatus = STATUS_OK Then
            Call SetInfo(strData)
            ReloadInfo = True
            Exit Function
        End If
    End If

    ReloadInfo = False
End Function

Private Sub SetInfo(ByVal strData As String)
    Dim TempUserInfo As tagUserInfo
    Dim PlayTimes As Long

    On Error Resume Next

    Call LoadUserInfo(TempUserInfo, strData)
    TempUserInfo.UserName = mstrUserName

    PlayTimes = TempUserInfo.Win + TempUserInfo.Draw + TempUserInfo.Lose

    Me.Caption = LoadString(208) & GetDisplayName(TempUserInfo.UserName, TempUserInfo.Name)

    ' 第一部分
    txtUserName.Text = TempUserInfo.UserName
    txtEmail.Text = TempUserInfo.Email
    txtWin.Text = CStr(TempUserInfo.Win)
    txtDraw.Text = CStr(TempUserInfo.Draw)
    txtLose.Text = CStr(TempUserInfo.Lose)
    txtScore.Text = CStr(TempUserInfo.Score)
    If PlayTimes < 1 Then
        txtWinP.Text = "0%"
    Else
        txtWinP.Text = Format(TempUserInfo.Win / PlayTimes, "0%")
    End If
    If gblnGameStart Then
        If TempUserInfo.GameTimes < 2 Then
            txtDisconnectP.Text = "0%"
        Else
            txtDisconnectP.Text = Format((TempUserInfo.GameTimes - 1 - PlayTimes) / (TempUserInfo.GameTimes - 1), "0%")
        End If
    Else
        If TempUserInfo.GameTimes < 1 Then
            txtDisconnectP.Text = "0%"
        Else
            txtDisconnectP.Text = Format((TempUserInfo.GameTimes - PlayTimes) / TempUserInfo.GameTimes, "0%")
        End If
    End If

    ' 第二部分
    'Call picFace.PaintPicture(MainForm.GetFace(TempUserInfo.Face), GetTwipX(6), GetTwipY(3))
    Call picFace.Cls
    If TempUserInfo.Face > CLng(MainForm.ilsFace.Tag) Then
        Call picFace.PaintPicture(frmResource.imgResDefaultFace.Picture, GetTwipX(5), GetTwipY(3))
    Else
        Call MainForm.ilsFace.ListImages(TempUserInfo.Face).Draw(picFace.hDC, GetTwipX(5), GetTwipY(3), imlTransparent)
    End If
    Call picFace.Refresh

    txtName.Text = TempUserInfo.Name
    If TempUserInfo.Sex = SEX_MAN Then
        txtSex.Text = LoadString(209)
    ElseIf TempUserInfo.Sex = SEX_WOMAN Then
        txtSex.Text = LoadString(210)
    Else
        txtSex.Text = "-"
    End If
    If TempUserInfo.Age = 0 Then
        txtAge.Text = "-"
    Else
        txtAge.Text = CStr(TempUserInfo.Age)
    End If
    txtCountry.Text = TempUserInfo.Country
    txtState.Text = TempUserInfo.State
    txtCity.Text = TempUserInfo.City
End Sub

Private Sub fltbtnClose_Click(Button As Integer)
    FormVisible = False
    Call Me.Hide
End Sub

Private Sub fltbtnReload_Click(Button As Integer)
    Call ReloadInfo(mstrUserName)
End Sub

Private Sub Form_Load()
    On Error Resume Next

    Call Me.Move(gptsSave_ViewUserInfo.X, gptsSave_ViewUserInfo.Y)
    Call fraTabs(0).Move(1080, tabInfo.ClientTop)
    fraTabs(0).Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        FormVisible = False
        Call Me.Hide
    End If
End Sub

Private Sub ietView_StateChanged(ByVal State As Integer)
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

Private Sub tabInfo_Click()
    On Error Resume Next

    fraTabs(0).Visible = False
    fraTabs(1).Visible = False
    Call fraTabs(tabInfo.SelectedItem.Index - 1).Move(1080, tabInfo.ClientTop)
    fraTabs(tabInfo.SelectedItem.Index - 1).Visible = True
End Sub

Private Sub Clear()
    Dim i As Long
    Dim j As Object

    On Error Resume Next

    For i = 0 To Me.Controls.Count - 1
        Set j = Me.Controls(i)
        If TypeName(j) = "TextBox" Then
            j.Text = ""
        ElseIf TypeName(j) = "PictureBox" Then
            Call j.Cls
        End If
    Next i
End Sub
