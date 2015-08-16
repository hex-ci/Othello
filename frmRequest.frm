VERSION 5.00
Begin VB.Form frmRequest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "请求"
   ClientHeight    =   2475
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4395
   Icon            =   "frmRequest.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4395
   StartUpPosition =   2  '屏幕中心
   Begin Othello.FlatButton fltbtnAgree 
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   2115
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      Caption         =   "同意请求"
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
   Begin Othello.FlatButton fltbtnSend 
      Default         =   -1  'True
      Height          =   375
      Left            =   2055
      TabIndex        =   1
      Top             =   1965
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      Caption         =   "送出请求"
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
   Begin VB.TextBox txtMessage 
      Height          =   300
      Left            =   105
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1410
      Width           =   4215
   End
   Begin VB.Frame fraPartner 
      Caption         =   "对方"
      Height          =   915
      Left            =   75
      TabIndex        =   8
      Top             =   45
      Width           =   4260
      Begin VB.PictureBox picFace 
         AutoRedraw      =   -1  'True
         Height          =   540
         Left            =   195
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   225
         Width           =   540
      End
      Begin Othello.FlatButton fltbtnView 
         Height          =   315
         Left            =   3180
         TabIndex        =   7
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         Caption         =   "查看"
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
      Begin VB.TextBox txtUserName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   510
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名字:"
         Height          =   180
         Left            =   930
         TabIndex        =   10
         Top             =   240
         Width           =   450
      End
   End
   Begin Othello.FlatButton fltbtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3210
      TabIndex        =   2
      Top             =   1965
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      Caption         =   "取消"
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
   Begin Othello.FlatButton fltbtnDisagree 
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   1710
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      Caption         =   "拒绝请求"
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "简短附言:"
      Height          =   180
      Left            =   165
      TabIndex        =   9
      Top             =   1140
      Width           =   810
   End
End
Attribute VB_Name = "frmRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrCommand As String
Dim mblnAgree As Boolean
Dim mblnSendRequest As Boolean

Dim FormVisible As Boolean

Public Sub ShowEx(ByVal UserName As String, ByVal Name As String, ByVal Face As Long, Optional ByVal Command As String = "", Optional ByVal Message As String = "", Optional ByVal SendRequest As Boolean = True)
    On Error Resume Next

    Call MainForm.ilsFace.ListImages(Face).Draw(picFace.hDC, 0, 0, imlTransparent)
    txtUserName.Tag = UserName
    txtUserName.Text = GetDisplayName(UserName, Name)
    If SendRequest Then
        mstrCommand = CMD_Request & Command
    Else
        mstrCommand = CMD_ReRequest & Command
    End If
    mblnSendRequest = SendRequest
    If Command = CMD_RequestCancelGame Then
        Me.Caption = LoadString(219) & GetDisplayName(UserName, Name)
    ElseIf Command = CMD_RequestDrawGame Then
        Me.Caption = LoadString(220) & GetDisplayName(UserName, Name)
    Else
        Me.Caption = LoadString(221) & GetDisplayName(UserName, Name)
    End If
    If SendRequest Then
        fltbtnSend.Default = True
        fltbtnSend.Visible = True
        fltbtnCancel.Cancel = True
        fltbtnCancel.Visible = True
    Else
        txtMessage.Text = Message
        txtMessage.Locked = True
        txtMessage.BackColor = vbButtonFace
        Call fltbtnAgree.Move(fltbtnSend.Left, fltbtnSend.Top)
        Call fltbtnDisagree.Move(fltbtnCancel.Left, fltbtnCancel.Top)
        fltbtnSend.Visible = False
        fltbtnSend.Default = False
        fltbtnCancel.Visible = False
        fltbtnCancel.Cancel = False
        fltbtnAgree.Default = True
        fltbtnAgree.Visible = True
        fltbtnDisagree.Cancel = True
        fltbtnDisagree.Visible = True
    End If

    FormVisible = True
    Call Me.Show(vbModeless, MainForm)
End Sub

Public Sub FormMinimize()
    If Me.Visible Then
        Call Me.Hide
    End If
End Sub

Public Sub FormNormal()
    If FormVisible Then
        Call Me.Show(vbModeless)
        Call KeepOnTop(Me.hWnd)
    End If
End Sub

Private Sub fltbtnAgree_Click(Button As Integer)
    On Error Resume Next

    Call Me.Hide
    If gblnGameStart Then
        mblnAgree = True
        Call MainForm.SendCommand(mstrCommand & CMD_Agree)
        Call MainForm.AgreeRequest(Mid(mstrCommand, 2, 1))
    End If
    Call Unload(Me)
End Sub

Private Sub fltbtnCancel_Click(Button As Integer)
    mblnAgree = False
    Call Unload(Me)
End Sub

Private Sub fltbtnDisagree_Click(Button As Integer)
    'If mstrCommand = CMD_RequestJoin And Not gblnGameStart Then
    'If gblnGameStart Then
        mblnAgree = False
        'Call MainForm.SendCommand(CMD_DisagreeCancelGame)
    'End If
    Call Unload(Me)
End Sub

Private Sub fltbtnSend_Click(Button As Integer)
    On Error Resume Next

    If gblnGameStart Then
        If mstrCommand <> "" Then
            Call MainForm.SendCommand(mstrCommand & txtMessage.Text)
        Else
            Call MessageBox(Me.hWnd, LoadString(161), vbExclamation, LoadString(181))
        End If
    End If
    Call Unload(Me)
End Sub

Private Sub fltbtnView_Click(Button As Integer)
    On Error Resume Next

    'fltbtnView.Enabled = False
    Call frmUserInfo.ShowEx(txtUserName.Tag, txtUserName.Text)
    'fltbtnView.Enabled = True
End Sub

Private Sub Form_Activate()
    Call SetControlFocus(txtMessage)
End Sub

Private Sub Form_Load()
    Call KeepOnTop(Me.hWnd)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next

    If (Not mblnSendRequest) And (Not mblnAgree) Then
        Call MainForm.SendCommand(mstrCommand & CMD_Disagree)
        Call MainForm.DisagreeRequest(Mid(mstrCommand, 2, 1))
    End If
    Call KillOnTop(Me.hWnd)
    FormVisible = False
End Sub
