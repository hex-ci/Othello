VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于 黑白棋.Net"
   ClientHeight    =   3315
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   382
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin Othello.ScrollBox ScrollText 
      CausesValidation=   0   'False
      Height          =   1545
      Left            =   1140
      TabIndex        =   5
      Top             =   870
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   2725
      BackColor       =   0
      ForeColor       =   14737632
   End
   Begin VB.Timer TitleTimer 
      Interval        =   55
      Left            =   5070
      Top             =   570
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   1050
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   645
      ScaleWidth      =   2145
      TabIndex        =   4
      Top             =   60
      Width           =   2145
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   315
      Picture         =   "frmAbout.frx":1326
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   225
      Width           =   480
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "版本: 1.0"
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   3660
      TabIndex        =   2
      Top             =   660
      UseMnemonic     =   0   'False
      Width           =   810
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   " 点击这里继续 "
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   3750
      MouseIcon       =   "frmAbout.frx":146F
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "单击鼠标返回游戏！"
      Top             =   2430
      UseMnemonic     =   0   'False
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   7
      X2              =   376.933
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "声明: 本软件为免费软件，您可以任意分发其副本，但不得修改本软件及其相关文档！特此声明！"
      ForeColor       =   &H0080FFFF&
      Height          =   420
      Left            =   240
      TabIndex        =   1
      Top             =   2715
      UseMnemonic     =   0   'False
      Width           =   5220
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TitleX As Single
Dim X As Single

Public Sub ShowEx()
    On Error Resume Next

    lblVersion.Caption = LoadString(256) & App.Major & "." & App.Minor & "a Build " & App.Revision
    lblVersion.Visible = False
    Line1.Visible = False
    lblDisclaimer.Visible = False
    lblExit.Visible = False

    TitleX = picTitle.Left
    picTitle.Left = Me.ScaleWidth - 120
    X = picTitle.Left
    picTitle.Visible = True

    ScrollText.TextHeight = 210
    Call BackText

    Call Me.Move((Screen.Width - Width) \ 2, (Screen.Height * 0.9 - Height) \ 2)

    Call AniRotateShowFrm(Me.hWnd, 10)

    TitleTimer.Enabled = True
    Call Show(vbModal)
End Sub

Private Sub Form_Load()
    Set Me.Icon = MainForm.Icon
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call Unload(Me)
End Sub

Private Sub BackText()
    On Error Resume Next

    Call ScrollText.Cls
    ScrollText.Font.Name = "宋体"
    ScrollText.Font.Bold = True
    ScrollText.Font.Size = 9
    ScrollText.ForeColor = &HE0E0E0  '正常颜色
    Call ScrollText.TextOut("制作人员:" & vbCr)
    ScrollText.Font.Bold = False
    Call ScrollText.TextOut("          Hex" & vbCr & vbCr)
    'ScrollText.Font.Bold = True
    'Call ScrollText.TextOut("特别感谢:" & vbCr)
    'ScrollText.Font.Bold = False
    'Call ScrollText.TextOut("          心雨如歌" & vbCr & vbCr)
    'Call ScrollText.TextOut("          Kenny" & vbCr & vbCr)
    'Call ScrollText.TextOut("          我的朋友们等" & vbCr & vbCr)
    ScrollText.Font.Bold = True
    Call ScrollText.TextOut("版权所有(C) 2002-2014" & vbCr)
    Call ScrollText.TextOut(" 幸福家园BBS 保留所有权利" & vbCr & vbCr)
    Call ScrollText.TextOut("主页:" & vbCr)
    ScrollText.ForeColor = &HFF8080   ' 蓝色
    ScrollText.Font.Bold = False
    Call ScrollText.TextOut("      http://www.ourhf.com" & vbCr)
    ScrollText.ForeColor = &HE0E0E0  '正常颜色
    ScrollText.Font.Bold = True
    Call ScrollText.TextOut("E-mail:" & vbCr)
    ScrollText.ForeColor = &HFF8080  ' 蓝色
    ScrollText.Font.Bold = False
    Call ScrollText.TextOut("      webmaster@ourhf.com")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.Visible Then Call AniRotateUnloadFrm(Me, 10)
End Sub

Private Sub lblExit_Click()
    Call Unload(Me)
End Sub

Private Sub TitleTimer_Timer()
    On Error Resume Next

    'TitleTimer.Enabled = False
    X = X - 30
    If X <= TitleX Then
        TitleTimer.Enabled = False
        picTitle.Left = 70
        lblVersion.Visible = True
        Line1.Visible = True
        lblDisclaimer.Visible = True
        lblExit.Visible = True
        ScrollText.Scroll = True
        Exit Sub
    End If
    picTitle.Left = X
End Sub
