VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmPublicChat 
   AutoRedraw      =   -1  'True
   Caption         =   "公共聊天区"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5115
   Icon            =   "frmPublicChat.frx":0000
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5115
   StartUpPosition =   3  '窗口缺省
   Begin Othello.FlatButton fltbtnReload 
      Height          =   315
      Left            =   4335
      TabIndex        =   2
      Top             =   165
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
      Caption         =   "刷新"
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
   Begin VB.Timer tmrReload 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4545
      Top             =   1500
   End
   Begin InetCtlsObjects.Inet ietChat 
      Left            =   4455
      Top             =   645
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
      RequestTimeout  =   30
   End
   Begin VB.TextBox txtTalk 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   615
      Width           =   4155
   End
   Begin VB.ComboBox cboTalk 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1140
      TabIndex        =   0
      Top             =   165
      Width           =   2220
   End
   Begin Othello.FlatButton fltbtnSendTalk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3525
      TabIndex        =   1
      Top             =   165
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
      Caption         =   "发送"
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
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "聊天:"
      Height          =   180
      Left            =   495
      TabIndex        =   4
      Top             =   225
      Width           =   450
   End
End
Attribute VB_Name = "frmPublicChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngSecond As Long
Dim FormVisible As Boolean

Dim mblnServerOK As Boolean

Private Sub fltbtnReload_Click(Button As Integer)
    Call ReloadChat
    Call SetControlFocus(cboTalk)
End Sub

Private Sub fltbtnSendTalk_Click(Button As Integer)
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String
    Dim strTalk As String
    Dim strRecord As String
    Dim i As Long

    On Error Resume Next

    If Trim(cboTalk.Text) = "" Then
        Call SetControlFocus(cboTalk)
        Exit Sub
    End If
    If Len(cboTalk.Text) > 50 Or Not CheckString(cboTalk.Text) Then
        Call MessageBox(Me.hWnd, LoadString(176), vbExclamation, LoadString(177))
        Call SetControlFocus(cboTalk)
        Exit Sub
    End If

    strTalk = cboTalk.Text
    cboTalk.Text = ""

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_CHAT_SEND & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&name=" & ToUrlString(GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name)) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&text=" & ToUrlString(strTalk) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    tmrReload.Enabled = False
    mlngSecond = 0
    If ServerCommand(ietChat, mblnServerOK, strUrl, strStatus, strData) Then
        If strStatus = STATUS_OK Then
            txtTalk.Text = ""
            For i = 1 To GetFieldCount(strData)
                strRecord = GetField(strData, i)
                Call Chat(GetRecord(strRecord, 1), GetRecord(strRecord, 2), GetRecord(strRecord, 3))
            Next i

            For i = 0 To cboTalk.ListCount - 1
                If strTalk = cboTalk.List(i) Then
                    tmrReload.Enabled = True
                    Call SetControlFocus(cboTalk)
                    Exit Sub
                End If
            Next i

            If cboTalk.ListCount >= 20 Then
                For i = cboTalk.ListCount - 1 To 1 Step -1
                    cboTalk.List(i) = cboTalk.List(i - 1)
                Next i
                cboTalk.List(0) = strTalk
            Else
                Call cboTalk.AddItem(strTalk, 0)
            End If
        End If
    End If
    tmrReload.Enabled = True

    Call SetControlFocus(cboTalk)
End Sub

Private Sub Form_Activate()
    Call SetControlFocus(cboTalk)
End Sub

Private Sub Form_Load()
    On Error Resume Next

    Call Me.Move(gwifSave_PublicChatWindow.Left, gwifSave_PublicChatWindow.Top, gwifSave_PublicChatWindow.Width, gwifSave_PublicChatWindow.Height)
    Me.WindowState = glngSave_PublicChatWindowState
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next

    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Call Me.HideEx
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Call Me.HideEx
        Me.WindowState = vbNormal
        Exit Sub
    End If

    If Me.Width < LIMIT_WIDTH + 500 Or Me.Height < LIMIT_HEIGHT Then Exit Sub

    Call lblUserName.Move(230, 350)

    Call txtTalk.Move(220, 750, ScaleWidth - GetTwipX(31), ScaleHeight - GetTwipY(70))

    Call cboTalk.Move(lblUserName.Left + lblUserName.Width + GetTwipX(3), lblUserName.Top - GetTwipY(4), ScaleWidth - lblUserName.Width - GetTwipX(150))

    Call fltbtnSendTalk.Move(cboTalk.Left + cboTalk.Width + GetTwipX(8), lblUserName.Top - GetTwipY(4))
    Call fltbtnReload.Move(fltbtnSendTalk.Left + fltbtnSendTalk.Width + GetTwipX(8), lblUserName.Top - GetTwipY(4))
End Sub

Public Sub EnableChat(ByVal Name As String)
    lblUserName.Caption = Name & ":"
    cboTalk.Enabled = True
    fltbtnSendTalk.Enabled = True
    fltbtnReload.Enabled = True
    txtTalk.Enabled = True
    tmrReload.Enabled = True
    Call ReloadChat
End Sub

Public Sub DisableChat()
    lblUserName.Caption = LoadString(255)
    cboTalk.Enabled = False
    fltbtnSendTalk.Enabled = False
    fltbtnReload.Enabled = False
    txtTalk.Enabled = False
    tmrReload.Enabled = False
End Sub

Private Sub ietChat_StateChanged(ByVal State As Integer)
    On Error Resume Next

    Select Case State
        Case icResponseReceived
            mblnServerOK = True

        Case icError  '11
            mblnServerOK = False

        Case icResponseCompleted
            'mblnServerOK = True
            Dim strStatus As String
            Dim strData As String
            Dim i As Long
            Dim strRecord As String

            If GetServerExecute(ietChat, strStatus, strData) Then
                If strStatus = STATUS_OK Then
                    txtTalk.Text = ""
                    'Call SetControlFocus(cboTalk)
                    For i = 1 To GetFieldCount(strData)
                        strRecord = GetField(strData, i)
                        Call Chat(GetRecord(strRecord, 1), GetRecord(strRecord, 2), GetRecord(strRecord, 3))
                    Next i
                End If
            End If
    End Select
End Sub

Public Sub Chat(ByVal Who As String, ByVal Talk As String, ByVal ChatDate As String)
    Dim Temp As String

    Temp = Who & ": " & Talk & vbTab & "[" & Format(ChatDate, "yyyy-m-d hh:mm:ss") & "]"
    txtTalk.Text = Temp & vbCrLf & txtTalk.Text

    ' 存储聊天数据，以备保存。
    'Call lstChatData.AddItem(Temp)
End Sub

Private Function ReloadChat() As Boolean
    Dim strUrl As String

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_CHAT_GET & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    ReloadChat = ServerExecute(ietChat, strUrl)
End Function

Private Sub tmrReload_Timer()
    On Error Resume Next

    If Not gblnLogin Or Not Me.Visible Then
        tmrReload.Enabled = False
        mlngSecond = 0
        Exit Sub
    End If

    mlngSecond = mlngSecond + 1
    If mlngSecond > PUBLIC_CHAT_RELOAD_TIME Then
        mlngSecond = 0
        tmrReload.Enabled = False
        Call ReloadChat
        tmrReload.Enabled = True
    End If
End Sub

Private Sub lblUserName_Change()
    Call Form_Resize
End Sub

Public Sub ShowEx()
    On Error Resume Next

    FormVisible = True
    Call Me.Show(vbModeless)
    If gblnLogin Then Call ReloadChat
    tmrReload.Enabled = True
    mlngSecond = 0
End Sub

Public Sub HideEx()
    tmrReload.Enabled = False
    mlngSecond = 0
    FormVisible = False
    Call Me.Hide
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
