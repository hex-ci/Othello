VERSION 5.00
Begin VB.Form frmProgress 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ControlBox      =   0   'False
   Icon            =   "frmProgress.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrProgress 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2250
      Top             =   675
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   465
      Stretch         =   -1  'True
      Top             =   195
      Width           =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   5040
      Y1              =   15
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   5025
      X2              =   5025
      Y1              =   0
      Y2              =   945
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   5040
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Label txtProgress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "......"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3885
      TabIndex        =   1
      Top             =   315
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "正在连接服务器，请稍候"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1050
      TabIndex        =   0
      Top             =   315
      Width           =   2805
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mCount As Long
Dim FormVisible As Boolean

Public Sub ShowEx()
    On Error Resume Next

    FormVisible = True
    Call Me.Show(vbModeless)
    Call KeepOnTop(Me.hWnd)
    mCount = 7
    txtProgress.Caption = "......"
    tmrProgress.Enabled = True
End Sub

Public Sub HideEx()
    On Error Resume Next

    tmrProgress.Enabled = False
    FormVisible = False
    Call KillOnTop(Me.hWnd)
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

Private Sub Form_Load()
    On Error Resume Next

    Set imgIcon = MainForm.Icon
End Sub

Private Sub tmrProgress_Timer()
    Dim i As Long

    On Error Resume Next

    If mCount > 6 Then
        mCount = 1
        txtProgress.Caption = ""
    End If

    txtProgress.Caption = txtProgress.Caption & "."

    mCount = mCount + 1
End Sub
