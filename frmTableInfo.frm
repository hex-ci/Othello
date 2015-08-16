VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmTableInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "棋局信息"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmTableInfo.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5535
   Begin InetCtlsObjects.Inet ietTableInfo 
      Left            =   4770
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
      RequestTimeout  =   30
   End
   Begin VB.TextBox txtTime 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2820
      Width           =   2940
   End
   Begin VB.ComboBox cboType 
      Height          =   300
      ItemData        =   "frmTableInfo.frx":030A
      Left            =   1245
      List            =   "frmTableInfo.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1815
      Width           =   2940
   End
   Begin VB.CheckBox chkUpLevel 
      Caption         =   "晋级游戏(&U)"
      Height          =   315
      Left            =   1245
      TabIndex        =   12
      Top             =   3300
      Width           =   1380
   End
   Begin VB.ComboBox cboGameTimer 
      Height          =   300
      ItemData        =   "frmTableInfo.frx":0324
      Left            =   1245
      List            =   "frmTableInfo.frx":033A
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2325
      Width           =   2940
   End
   Begin VB.TextBox txtVisitor 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1305
      Width           =   2940
   End
   Begin VB.TextBox txtCreator 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   780
      Width           =   2940
   End
   Begin VB.TextBox txtTableName 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   270
      Width           =   2940
   End
   Begin Othello.FlatButton fltbtnClose 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4335
      TabIndex        =   15
      Top             =   1125
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   609
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
      Height          =   345
      Left            =   4335
      TabIndex        =   13
      Top             =   270
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   609
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
   Begin Othello.FlatButton fltbtnEdit 
      Height          =   345
      Left            =   4335
      TabIndex        =   14
      Top             =   690
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   609
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "时间(&M):"
      Height          =   180
      Left            =   270
      TabIndex        =   10
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "类型(&Y):"
      Height          =   180
      Left            =   270
      TabIndex        =   6
      Top             =   1860
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "计时器(&T):"
      Height          =   180
      Left            =   270
      TabIndex        =   8
      Top             =   2370
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "访问者(&V):"
      Height          =   180
      Left            =   270
      TabIndex        =   4
      Top             =   1350
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "创建人(&C):"
      Height          =   180
      Left            =   270
      TabIndex        =   2
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "棋局名(&N):"
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   330
      Width           =   900
   End
End
Attribute VB_Name = "frmTableInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnServerOK As Boolean

Dim mstrTableName As String

Dim FormVisible As Boolean

Public Sub ShowEx(ByVal TableName As String, ByVal IsEdit As Boolean)
    On Error Resume Next

    mstrTableName = TableName
    FormVisible = True
    Call Show(vbModeless)

    If Me.WindowState <> vbNormal Then
        Me.WindowState = vbNormal
    End If

    If IsEdit Then
        Me.Caption = LoadString(253) & TableName
        cboType.Enabled = True
        cboGameTimer.Enabled = True
        chkUpLevel.Enabled = True
        fltbtnEdit.Enabled = True
    Else
        Me.Caption = LoadString(254) & TableName
        cboType.Enabled = False
        cboGameTimer.Enabled = False
        chkUpLevel.Enabled = False
        fltbtnEdit.Enabled = False
    End If

    Call Clear
    Call Me.Refresh
    Call ReloadInfo(TableName)
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

Private Function ReloadInfo(ByVal TableName As String) As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_TABLE_VIEW & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&name=" & ToUrlString(TableName) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If ServerCommand(ietTableInfo, mblnServerOK, strUrl, strStatus, strData) Then
        If strStatus = STATUS_OK Then
            Call SetInfo(strData)
            ReloadInfo = True
            Exit Function
        End If
    End If

    ReloadInfo = False
End Function

Private Sub SetInfo(ByVal strData As String)
    Dim TempTableInfo As tagTableInfo

    On Error Resume Next

    Call LoadTableInfo(TempTableInfo, strData)

    txtTableName.Text = TempTableInfo.TableName
    txtCreator.Text = GetDisplayName(TempTableInfo.Creator, TempTableInfo.CreatorName)
    txtVisitor.Text = GetDisplayName(TempTableInfo.Visitor, TempTableInfo.VisitorName)
    txtTime.Text = Format(TempTableInfo.LastTime, "yyyy年m月d日 hh:mm:ss")

    cboType.ListIndex = TempTableInfo.TableType - 1
    cboGameTimer.ListIndex = TempTableInfo.Timer
    chkUpLevel.Value = Abs(TempTableInfo.UpLevel)
End Sub

Private Function SaveInfo() As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_TABLE_EDIT & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&creator=" & ToUrlString(gMainTableInfo.Creator) & _
                                  "&nickname=" & ToUrlString(gMyUserInfo.Name) & _
                                  "&name=" & ToUrlString(txtTableName.Text) & _
                                  "&type=" & CStr(cboType.ItemData(cboType.ListIndex)) & _
                                  "&timer=" & CStr(cboGameTimer.ItemData(cboGameTimer.ListIndex)) & _
                                  "&level=" & CStr(chkUpLevel.Value) & _
                                  "&port=" & CStr(glngSave_GamePort) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    Me.MousePointer = vbHourglass
    Me.Caption = LoadString(258)
    If ServerCommand(ietTableInfo, mblnServerOK, strUrl, strStatus, strData) Then
        Me.MousePointer = vbDefault
        If strStatus = STATUS_OK Then
            Call LoadTableInfo(gMainTableInfo, strData)
            Call SetInfo(strData)
            SaveInfo = True
            Exit Function
        End If
    End If
    Me.MousePointer = vbDefault
    Me.Caption = LoadString(253) & gMainTableInfo.TableName
    SaveInfo = False
End Function

Private Sub fltbtnClose_Click(Button As Integer)
    FormVisible = False
    Call Me.Hide
End Sub

Private Sub fltbtnEdit_Click(Button As Integer)
    On Error Resume Next

    If (Not gblnLogin) Or (Not gblnCreator) Or gblnGameStart Then
        Call Me.Hide
        Exit Sub
    End If

    If Len(txtTableName.Text) < 2 Or Len(txtTableName.Text) > 15 Or Not CheckString(txtTableName.Text) Then
        Call MessageBox(Me.hWnd, LoadString(172), vbCritical, LoadString(181))
        Call SetControlFocus(txtTableName)
        Call SendKeys("{Home}+{End}")
        Exit Sub
    End If

    If SaveInfo() Then
        Me.Caption = LoadString(253) & gMainTableInfo.TableName
        If gblnConnect Then
            Call MainForm.SendCommand(CMD_TableChanged & gMainTableInfo.TableType & "|" & gMainTableInfo.Timer & "|" & Abs(gMainTableInfo.UpLevel))
        End If
        ' 更新棋局显示
        If gMainTableInfo.Timer < 1 Then
            MainForm.lblTime.Caption = "--:--"
        Else
            MainForm.lblTime.Caption = GetTime(gMainTableInfo.Timer * 60)
        End If

        Call frmTable.ReloadTable(True)
        Call frmOnline.ReloadOnline
        Call MessageBox(hWnd, LoadString(173), vbInformation, LoadString(180))
    Else
        Call MessageBox(hWnd, LoadString(174), vbCritical, LoadString(181))
    End If
End Sub

Private Sub fltbtnReload_Click(Button As Integer)
    Call ReloadInfo(mstrTableName)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call Me.Move(gptsSave_TableInfo.X, gptsSave_TableInfo.Y)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        FormVisible = False
        Call Me.Hide
    End If
End Sub

Private Sub ietTableInfo_StateChanged(ByVal State As Integer)
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

Private Sub Clear()
    Dim i As Long
    Dim j As Object

    On Error Resume Next

    For i = 0 To Me.Controls.Count - 1
        Set j = Me.Controls(i)
        If TypeName(j) = "TextBox" Then
            j.Text = ""
        End If
    Next i
End Sub
