VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmEditInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改资料"
   ClientHeight    =   3660
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5010
   Icon            =   "frmEditInfo.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5010
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   225
      TabIndex        =   33
      Top             =   465
      Width           =   795
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   135
         Picture         =   "frmEditInfo.frx":0E42
         Top             =   315
         Width           =   480
      End
   End
   Begin Othello.FlatButton fltbtnEdit 
      Height          =   375
      Left            =   225
      TabIndex        =   12
      Top             =   2040
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   661
      Caption         =   "修改"
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
      Left            =   225
      TabIndex        =   13
      Top             =   2490
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
   Begin InetCtlsObjects.Inet ietEdit 
      Left            =   4935
      Top             =   2445
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
      RequestTimeout  =   30
   End
   Begin VB.Frame fraTabs 
      BorderStyle     =   0  'None
      Height          =   3045
      Index           =   1
      Left            =   1020
      TabIndex        =   15
      Top             =   3555
      Visible         =   0   'False
      Width           =   3780
      Begin VB.Frame fraInfo 
         Caption         =   "其它信息"
         Height          =   2850
         Left            =   0
         TabIndex        =   16
         Top             =   90
         Width           =   3600
         Begin VB.ComboBox cboCountry 
            Height          =   300
            ItemData        =   "frmEditInfo.frx":170C
            Left            =   195
            List            =   "frmEditInfo.frx":1713
            TabIndex        =   6
            Top             =   1665
            Width           =   1545
         End
         Begin VB.ComboBox cboState 
            Height          =   300
            ItemData        =   "frmEditInfo.frx":1727
            Left            =   195
            List            =   "frmEditInfo.frx":17AA
            TabIndex        =   9
            Top             =   2325
            Width           =   1545
         End
         Begin VB.ComboBox cboSex 
            Height          =   300
            ItemData        =   "frmEditInfo.frx":18F4
            Left            =   2655
            List            =   "frmEditInfo.frx":18FE
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1665
            Width           =   735
         End
         Begin VB.TextBox txtCity 
            Height          =   315
            Left            =   1830
            MaxLength       =   20
            TabIndex        =   10
            Top             =   2325
            Width           =   1560
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   5
            Top             =   585
            Width           =   1785
         End
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
            Height          =   300
            Left            =   1830
            MaxLength       =   3
            TabIndex        =   7
            Top             =   1665
            Width           =   735
         End
         Begin MSComctlLib.ImageCombo imgcboFace 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   360
            TabIndex        =   4
            Top             =   585
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            ForeColor       =   -2147483633
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   1365
            X2              =   3435
            Y1              =   1170
            Y2              =   1170
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            X1              =   1365
            X2              =   3435
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "城市:"
            Height          =   180
            Left            =   1845
            TabIndex        =   23
            Top             =   2100
            Width           =   450
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "省份:"
            Height          =   180
            Left            =   210
            TabIndex        =   22
            Top             =   2100
            Width           =   450
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "国家/地区:"
            Height          =   180
            Left            =   210
            TabIndex        =   21
            Top             =   1425
            Width           =   900
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "头像:"
            Height          =   180
            Left            =   225
            TabIndex        =   20
            Top             =   285
            Width           =   450
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "昵称:"
            Height          =   180
            Left            =   1680
            TabIndex        =   19
            Top             =   315
            Width           =   450
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "性别:"
            Height          =   180
            Left            =   2670
            TabIndex        =   18
            Top             =   1425
            Width           =   450
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "年龄:"
            Height          =   180
            Left            =   1845
            TabIndex        =   17
            Top             =   1425
            Width           =   450
         End
      End
   End
   Begin VB.Frame fraTabs 
      BorderStyle     =   0  'None
      Height          =   3045
      Index           =   0
      Left            =   1080
      TabIndex        =   24
      Top             =   420
      Visible         =   0   'False
      Width           =   3780
      Begin VB.Frame fraBaseInfo 
         Caption         =   "基本信息"
         Height          =   1320
         Left            =   0
         TabIndex        =   29
         Top             =   90
         Width           =   3600
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   915
            TabIndex        =   0
            Top             =   780
            Width           =   2505
         End
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   330
            Width           =   2505
         End
         Begin VB.Label Label2 
            Caption         =   "E-mail:"
            Height          =   180
            Left            =   225
            TabIndex        =   31
            Top             =   840
            Width           =   630
         End
         Begin VB.Label Label1 
            Caption         =   "用户名:"
            Height          =   180
            Left            =   225
            TabIndex        =   30
            Top             =   390
            Width           =   630
         End
      End
      Begin VB.Frame fraPassword 
         Caption         =   "修改密码"
         Height          =   1440
         Left            =   0
         TabIndex        =   25
         Top             =   1500
         Width           =   3600
         Begin VB.TextBox txtNewPassword2 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   1080
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   1005
            Width           =   2100
         End
         Begin VB.TextBox txtNewPassword1 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   1080
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   645
            Width           =   2100
         End
         Begin VB.TextBox txtOldPassword 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   1080
            PasswordChar    =   "*"
            TabIndex        =   1
            Top             =   285
            Width           =   2100
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "校  验:"
            Height          =   180
            Left            =   360
            TabIndex        =   28
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "新密码:"
            Height          =   180
            Left            =   360
            TabIndex        =   27
            Top             =   690
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "旧密码:"
            Height          =   180
            Left            =   360
            TabIndex        =   26
            Top             =   330
            Width           =   630
         End
      End
   End
   Begin Othello.FlatButton fltbtnClose 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   225
      TabIndex        =   14
      Top             =   2940
      Width           =   780
      _extentx        =   1376
      _extenty        =   661
      forecolor       =   0
      mousepointer    =   99
      font            =   "frmEditInfo.frx":190A
      style           =   2
      caption         =   "关闭"
      enablehot       =   -1
   End
   Begin MSComctlLib.TabStrip tabInfo 
      Height          =   3420
      Left            =   135
      TabIndex        =   32
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
Attribute VB_Name = "frmEditInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnServerOK As Boolean
Dim FormVisible As Boolean

Public Sub ShowEx()
    On Error Resume Next

    FormVisible = True
    Call Show(vbModeless)
    If Me.WindowState <> vbNormal Then
        Me.WindowState = vbNormal
    End If

    'Me.Caption = LoadString(252) & GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name)

    Call Me.Refresh
    Call SetInfo
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

Private Sub fltbtnClose_Click(Button As Integer)
    FormVisible = False
    Call Me.Hide
End Sub

Private Sub fltbtnEdit_Click(Button As Integer)
    On Error Resume Next

    If Not CheckBaseInfo() Then Exit Sub
    If Not CheckInfo() Then Exit Sub
    If SaveInfo() Then
        If gblnConnect And (Not (imgcboFace.SelectedItem Is Nothing)) Then
            Call MainForm.SendCommand(CMD_InfoChanged & CStr(imgcboFace.SelectedItem.Index) & "|" & txtName.Text)
        End If
        Me.Caption = LoadString(252) & GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name)
        If Not (imgcboFace.SelectedItem Is Nothing) Then
            Call MainForm.RefreshMyFace(imgcboFace.SelectedItem.Index)
        End If
        Call MainForm.RefreshMyName

        Call frmOnline.ReloadOnline
        Call frmTable.ReloadTable(True)
        Call MessageBox(hWnd, LoadString(169), vbInformation, LoadString(180))
    Else
        Call MessageBox(hWnd, LoadString(170), vbExclamation, LoadString(181))
    End If

    txtOldPassword.Text = ""
    txtNewPassword1.Text = ""
    txtNewPassword2.Text = ""
End Sub

Private Sub fltbtnReload_Click(Button As Integer)
    Call ReloadInfo
End Sub

Private Sub Form_Load()
    Dim i As Long

    On Error Resume Next

    Call Me.Move(gptsSave_EditUserInfo.X, gptsSave_EditUserInfo.Y)

    Set imgcboFace.ImageList = MainForm.ilsFace
    For i = 1 To Val(MainForm.ilsFace.Tag)
        imgcboFace.ComboItems.Add.Image = i
    Next i
    'imgcboFace.ComboItems.Item(1).Selected = True

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

Private Sub ietEdit_StateChanged(ByVal State As Integer)
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

Private Sub imgcboFace_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
    End If
End Sub

Private Sub tabInfo_Click()
    On Error Resume Next

    fraTabs(0).Visible = False
    fraTabs(1).Visible = False
    Call fraTabs(tabInfo.SelectedItem.Index - 1).Move(1080, tabInfo.ClientTop)
    fraTabs(tabInfo.SelectedItem.Index - 1).Visible = True
End Sub

Private Sub SetInfo()
    On Error Resume Next

    Me.Caption = LoadString(252) & GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name)

    ' 第一部分
    txtUserName.Text = gMyUserInfo.UserName
    txtEmail.Text = gMyUserInfo.Email

    ' 第二部分
    If gMyUserInfo.Face <= CLng(MainForm.ilsFace.Tag) Then
        imgcboFace.ComboItems.Item(gMyUserInfo.Face).Selected = True
    End If
    txtName.Text = gMyUserInfo.Name

    If gMyUserInfo.Sex > 0 And gMyUserInfo.Sex < 3 Then
        cboSex.ListIndex = gMyUserInfo.Sex - 1
    Else
        cboSex.ListIndex = -1
    End If

    If gMyUserInfo.Age = 0 Then
        txtAge.Text = ""
    Else
        txtAge.Text = CStr(gMyUserInfo.Age)
    End If

    cboCountry = gMyUserInfo.Country
    cboState.Text = gMyUserInfo.State
    txtCity.Text = gMyUserInfo.City
End Sub

Private Function CheckBaseInfo() As Boolean
    On Error Resume Next

    If StrLen(txtEmail.Text) < 3 Or StrLen(txtEmail.Text) > 30 Or InStr(1, txtEmail.Text, "@") < 2 Or Not CheckString(txtEmail.Text) Then
        CheckBaseInfo = False
        Call MessageBox(hWnd, LoadString(105) & LoadString(106), vbCritical, LoadString(181))
        Call SetControlFocus(txtEmail)
        Call SendKeys("{Home}+{End}")
        Exit Function
    End If

    If txtOldPassword.Text <> "" Then
        If txtOldPassword.Text <> gMyUserInfo.Password Then
            CheckBaseInfo = False
            Call MessageBox(hWnd, LoadString(171) & LoadString(113), vbCritical, LoadString(181))
            Call SetControlFocus(txtOldPassword)
            Call SendKeys("{Home}+{End}")
            Exit Function
        End If

        If StrLen(txtNewPassword1.Text) < 5 Or StrLen(txtNewPassword1.Text) > 15 Or Not CheckString(txtNewPassword1.Text) Then
            CheckBaseInfo = False
            Call MessageBox(hWnd, LoadString(103) & LoadString(113), vbCritical, LoadString(181))
            Call SetControlFocus(txtNewPassword1)
            Call SendKeys("{Home}+{End}")
            Exit Function
        End If
        If txtNewPassword2.Text <> txtNewPassword1.Text Then
            CheckBaseInfo = False
            Call MessageBox(hWnd, LoadString(104), vbCritical, LoadString(181))
            txtNewPassword1.Text = ""
            txtNewPassword2.Text = ""
            Call SetControlFocus(txtNewPassword1)
            Exit Function
        End If
    Else
        txtNewPassword1.Text = ""
        txtNewPassword2.Text = ""
        'CheckBaseInfo = False
        'Call MessageBox(hWnd, LoadString(171) & LoadString(113), vbCritical, LoadString(181))
        'Call SetControlFocus(txtOldPassword)
        'Call SendKeys("{Home}+{End}")
        'Exit Function
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

Private Function SaveInfo() As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    On Error Resume Next

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_USER_EDIT & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName)

    If txtNewPassword1.Text <> "" Then strUrl = strUrl & "&password=" & MD5(txtNewPassword1.Text)

    If Not (imgcboFace.SelectedItem Is Nothing) Then
        strUrl = strUrl & "&face=" & CStr(imgcboFace.SelectedItem.Index)
    End If
    strUrl = strUrl & "&name=" & ToUrlString(txtName.Text) & _
                      "&country=" & ToUrlString(cboCountry.Text) & _
                      "&state=" & ToUrlString(cboState.Text) & _
                      "&city=" & ToUrlString(txtCity.Text)

    If txtOldPassword.Text <> "" Then
        strUrl = strUrl & "&oldpassword=" & MD5(txtOldPassword.Text)
    Else
        strUrl = strUrl & "&oldpassword=" & MD5(gMyUserInfo.Password)
    End If

    If txtEmail.Text <> gMyUserInfo.Email Then
        strUrl = strUrl & "&email=" & ToUrlString(txtEmail.Text)
    End If

    If cboSex.ListIndex = -1 Then
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

    strUrl = strUrl & "&" & MakeServerPassword() & "&" & MakeVersion()

    Me.MousePointer = vbHourglass
    Me.Caption = LoadString(257)
    If ServerCommand(ietEdit, mblnServerOK, strUrl, strStatus, strData) Then
        Me.MousePointer = vbDefault
        If strStatus = STATUS_OK Then
            If txtNewPassword1.Text <> "" Then gMyUserInfo.Password = txtNewPassword1.Text
            gMyUserInfo.Email = txtEmail.Text
            gMyUserInfo.Face = imgcboFace.SelectedItem.Index
            gMyUserInfo.Name = txtName.Text
            gMyUserInfo.Country = cboCountry.Text
            gMyUserInfo.State = cboState.Text
            gMyUserInfo.City = txtCity.Text

            If cboSex.List(cboSex.ListIndex) = "" Then
                gMyUserInfo.Sex = 0
            ElseIf cboSex.List(cboSex.ListIndex) = LoadString(209) Then
                gMyUserInfo.Sex = SEX_MAN
            Else
                gMyUserInfo.Sex = SEX_WOMAN
            End If

            gMyUserInfo.Age = Val(txtAge.Text)

            SaveInfo = True
            Exit Function
        End If
    End If

    Me.MousePointer = vbDefault
    Me.Caption = LoadString(252) & GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name)
    SaveInfo = False
End Function

Private Function ReloadInfo() As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    On Error Resume Next

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_USER_VIEW & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&name=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If ServerCommand(ietEdit, mblnServerOK, strUrl, strStatus, strData) Then
        If strStatus = STATUS_OK Then
            Call LoadUserInfo(gMyUserInfo, strData)
            Call SetInfo
            ReloadInfo = True
            Exit Function
        End If
    End If

    ReloadInfo = False
End Function
