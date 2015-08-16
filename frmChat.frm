VERSION 5.00
Begin VB.Form frmChat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picResize 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3765
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   4035
      Width           =   240
   End
   Begin Othello.FlatButton fltbtnSendTalk 
      Default         =   -1  'True
      Height          =   315
      Left            =   2985
      TabIndex        =   2
      Top             =   720
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
      ForeColor       =   -2147483631
      Caption         =   "发送"
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
   End
   Begin VB.ComboBox cboTalk 
      Enabled         =   0   'False
      Height          =   300
      Left            =   660
      TabIndex        =   1
      Top             =   720
      Width           =   2220
   End
   Begin VB.TextBox txtTalk 
      Enabled         =   0   'False
      Height          =   2355
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1170
      Width           =   3510
   End
   Begin Othello.FlatButton fltbtnHide 
      Height          =   240
      Left            =   3015
      TabIndex        =   6
      Top             =   45
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BackColor       =   255
      ForeColor       =   0
      Caption         =   ""
      Style           =   1
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstChatData 
      Height          =   420
      ItemData        =   "frmChat.frx":000C
      Left            =   195
      List            =   "frmChat.frx":000E
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "--  聊天  --"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1155
      TabIndex        =   7
      Top             =   75
      Width           =   1095
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   8
      X1              =   315
      X2              =   1530
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   315
      X2              =   1530
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   315
      X2              =   1530
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   1620
      X2              =   2835
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   1620
      X2              =   2835
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   315
      X2              =   1530
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   315
      X2              =   1530
      Y1              =   2745
      Y2              =   2745
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   1620
      X2              =   2835
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   1620
      X2              =   2835
      Y1              =   2745
      Y2              =   2745
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "聊天:"
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   780
      Width           =   450
   End
   Begin VB.Shape shpTitle 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   375
      Top             =   30
      Width           =   3210
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DockStyle As eDockStyle
Public DockPosition As Single

Public DockingLevel As Single
Public Parent As Form

Dim mMouseX As Single
Dim mMouseY As Single
Dim FormMoving As Boolean
Dim FormResizing As Boolean
Dim FormVisible As Boolean

Public Sub ShowEx(ByRef Frm As Form)
    On Error Resume Next

    Set Parent = Frm
    FormVisible = True
    Call FormMove(True)
    Call Me.Show(vbModeless)
End Sub

Public Sub HideEx()
    FormVisible = False
    MainForm.fltbtnChat.Value = False
    Call Me.Hide
End Sub

'Private Sub cboTalk_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        Call fltbtnSendTalk_Click(vbLeftButton)
'    End If
'End Sub

Private Sub fltbtnSendTalk_Click(Button As Integer)
    On Error Resume Next

    If cboTalk.Text = "" Then
        Call SetControlFocus(cboTalk)
        Exit Sub
    End If
    If Len(cboTalk.Text) > 50 Then
        Call MessageBox(Me.hWnd, LoadString(175), vbExclamation, LoadString(177))
        Call SetControlFocus(cboTalk)
        Exit Sub
    End If

    Dim i As Long
    Dim strTalk As String

    strTalk = cboTalk.Text
    cboTalk.Text = ""
    Call MainForm.SendTalk(strTalk)
    Call Chat(GetDisplayName(gMyUserInfo.UserName, gMyUserInfo.Name), strTalk)

    For i = 0 To cboTalk.ListCount - 1
        If strTalk = cboTalk.List(i) Then
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

    Call SetControlFocus(cboTalk)
End Sub

Private Sub Form_Activate()
    shpTitle.BackColor = CLR_ACTIVATE
    Call SetControlFocus(cboTalk)
End Sub

Private Sub Form_Deactivate()
    shpTitle.BackColor = CLR_DEACTIVATE
End Sub

Private Sub Form_Load()
    On Error Resume Next

    Set picResize.Picture = frmResource.imgResResizer.Picture
    Set fltbtnHide.Picture = frmResource.imgResHide1.Picture
    Set fltbtnHide.DownPicture = frmResource.imgResHide2.Picture

    DockingLevel = GetTwipX(10)

    FormMoving = False
    FormResizing = False

    linBorder(0).BorderColor = vbWhite
    linBorder(1).BorderColor = vbWhite
    linBorder(2).BorderColor = vbWhite
    linBorder(3).BorderColor = vbWhite

    linBorder(4).BorderColor = RGB(210, 220, 210)
    linBorder(5).BorderColor = RGB(180, 180, 180)
    linBorder(6).BorderColor = RGB(180, 180, 180)
    linBorder(7).BorderColor = RGB(210, 220, 210)

    linBorder(8).BorderColor = vbWhite

    If gwifSave_ChatWindow.DockStyle = dsDockNone Then
        Call Me.Move(gwifSave_ChatWindow.Left, gwifSave_ChatWindow.Top, gwifSave_ChatWindow.Width, gwifSave_ChatWindow.Height)
    Else
        DockPosition = gwifSave_ChatWindow.DockPosition
        Me.Width = gwifSave_ChatWindow.Width
        Me.Height = gwifSave_ChatWindow.Height
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbLeftButton) = vbLeftButton And FormMoving Then
        ' 停靠到左边
        If Abs(Left + Width - (Parent.Left + gsngBorderX)) < DockingLevel And (Top + Height > Parent.Top + gsngCaptionHeight + gsngBorderY And Top < Parent.Top + gsngCaptionHeight + gsngBorderY + WindowHeight) Then
            Call Me.Move(Parent.Left + gsngBorderX - Width, Y - mMouseY + Me.Top)
            DockStyle = dsDockLeft
            DockPosition = Me.Top - Parent.Top
            If Abs(X - mMouseX) > DockingLevel Then
                GoTo FreeMove
            End If
            Exit Sub
        End If
        ' 停靠到右边
        If Abs(Me.Left - (Parent.Left + gsngBorderX + WindowWidth + GetTwipX(1))) < DockingLevel And (Top + Height > Parent.Top + gsngCaptionHeight + gsngBorderY And Top < Parent.Top + gsngCaptionHeight + gsngBorderY + WindowHeight) Then
            Call Me.Move(Parent.Left + gsngBorderX + WindowWidth + GetTwipX(1), Y - mMouseY + Me.Top)
            DockStyle = dsDockRight
            DockPosition = Me.Top - Parent.Top
            If Abs(X - mMouseX) > DockingLevel Then
                GoTo FreeMove
            End If
            Exit Sub
        End If
        ' 停靠到上边
        If Abs(Top + Height - (Parent.Top + gsngCaptionHeight + gsngBorderY)) < DockingLevel And (Left + Width > Parent.Left + gsngBorderX And Left < Parent.Left + gsngBorderX + WindowWidth) Then
            Call Me.Move(X - mMouseX + Left, Parent.Top + gsngCaptionHeight + gsngBorderY - Height)
            DockStyle = dsDockTop
            DockPosition = Me.Left - Parent.Left
            If Abs(Y - mMouseY) > DockingLevel Then
                GoTo FreeMove
            End If
            Exit Sub
        End If
        ' 停靠到下面
        If Abs(Me.Top - (Parent.Top + gsngCaptionHeight + gsngBorderY + WindowHeight + GetTwipY(1))) < DockingLevel And (Left + Width > Parent.Left + gsngBorderX And Left < Parent.Left + gsngBorderX + WindowWidth) Then
            Call Me.Move(X - mMouseX + Left, Parent.Top + gsngCaptionHeight + gsngBorderY + WindowHeight + GetTwipY(1))
            DockStyle = dsDockBottom
            DockPosition = Me.Left - Parent.Left
            If Abs(Y - mMouseY) > DockingLevel Then
                GoTo FreeMove
            End If
            Exit Sub
        End If
        
        ' 自由移动
FreeMove:
        Call Me.Move(X - mMouseX + Me.Left, Y - mMouseY + Me.Top)
        DockStyle = dsDockNone
        'Debug.Print x / Screen.TwipsPerPixelX, mMouseX, Me.left
    Else
        If X < lblTitle.Left Or X > lblTitle.Left + lblTitle.Width Or Y < lblTitle.Top Or Y > lblTitle.Top + lblTitle.Height Then
            Call ReleaseCapture
        End If
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetControlFocus(Me)
    If (Button And vbLeftButton) = vbLeftButton Then
        Call SetCapture(Me.hWnd)
        mMouseX = X: mMouseY = Y
        FormMoving = True
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbLeftButton) = vbLeftButton And FormMoving Then
        Call ReleaseCapture
        FormMoving = False
    End If
End Sub

Private Sub Form_Resize()
    'Call ControlPosition(txtTalk, 200, 1000, 0.8, 0.7)
    'Call ResizeControl(fltbtnSendTalk, Me)
    'Call ResizeControl(txtTalk, Me)

    Call LinePosition(linBorder(0), 0, 0, ScaleWidth, 0)
    Call LinePosition(linBorder(1), ScaleWidth - GetTwipX(1), 0, 0, ScaleHeight)
    Call LinePosition(linBorder(2), 0, ScaleHeight - GetTwipY(1), ScaleWidth, 0)
    Call LinePosition(linBorder(3), 0, 0, 0, ScaleHeight)

    Call LinePosition(linBorder(4), GetTwipX(1), GetTwipY(1), ScaleWidth - GetTwipX(2), 0)
    Call LinePosition(linBorder(5), ScaleWidth - GetTwipX(2), GetTwipY(1), 0, ScaleHeight - GetTwipY(2))
    Call LinePosition(linBorder(6), GetTwipX(1), ScaleHeight - GetTwipY(2), ScaleWidth - GetTwipX(2), 0)
    Call LinePosition(linBorder(7), GetTwipX(1), GetTwipY(1), 0, ScaleHeight - GetTwipY(2))

    Call LinePosition(linBorder(8), GetTwipX(2), shpTitle.Height + GetTwipY(1), ScaleWidth - GetTwipX(4), 0)

    Call shpTitle.Move(GetTwipX(2), GetTwipY(2), ScaleWidth - GetTwipX(3))
    Call lblTitle.Move((shpTitle.Width - lblTitle.Width - fltbtnHide.Width - GetTwipX(2)) \ 2)
    Call fltbtnHide.Move(shpTitle.Width - fltbtnHide.Width - GetTwipX(1))

    Call lblUserName.Move(230, 550)

    Call txtTalk.Move(220, 900, ScaleWidth - GetTwipX(31), ScaleHeight - GetTwipY(79))

    Call cboTalk.Move(lblUserName.Left + lblUserName.Width + GetTwipX(3), lblUserName.Top - GetTwipY(4), ScaleWidth - lblUserName.Width - GetTwipX(90))

    Call fltbtnSendTalk.Move(cboTalk.Left + cboTalk.Width + GetTwipX(8), lblUserName.Top - GetTwipY(4))

    Call picResize.Move(Me.ScaleWidth - picResize.Width - GetTwipX(3), Me.ScaleHeight - picResize.Height - GetTwipY(3))
End Sub

Public Sub FormMove(Optional ByVal AlwaysMove As Boolean = False)
    If Me.Visible Or AlwaysMove Then
        Select Case DockStyle
            Case dsDockLeft
                Call Me.Move(Parent.Left + gsngBorderX - Width, Parent.Top + DockPosition)

            Case dsDockRight
                Call Me.Move(Parent.Left + gsngBorderX + WindowWidth + GetTwipX(1), Parent.Top + DockPosition)

            Case dsDockTop
                Call Me.Move(Parent.Left + DockPosition, Parent.Top + gsngCaptionHeight + gsngBorderY - Height)

            Case dsDockBottom
                Call Me.Move(Parent.Left + DockPosition, Parent.Top + gsngCaptionHeight + gsngBorderY + WindowHeight + GetTwipY(1))

            Case dsDockNone
                FormMoving = True
                mMouseX = 0: mMouseY = 0
                Call Form_MouseMove(vbLeftButton, 0, 0, 0)
                FormMoving = False
        End Select
    End If
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

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, X + lblTitle.Left, Y + lblTitle.Top)
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetCapture(Me.hWnd)
End Sub

Private Sub fltbtnHide_Click(Button As Integer)
    Call Me.HideEx
End Sub

Private Sub lblUserName_Change()
    Call Form_Resize
End Sub

Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbLeftButton) = vbLeftButton Then
        Call SetCapture(picResize.hWnd)
        mMouseX = X: mMouseY = Y
        FormResizing = True
    End If
End Sub

Private Sub picResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbLeftButton) = vbLeftButton And FormResizing Then
        If X - mMouseX + Me.Width < LIMIT_WIDTH And Y - mMouseY + Me.Height < LIMIT_HEIGHT Then
            Call Me.Move(Me.Left, Me.Top, LIMIT_WIDTH, LIMIT_HEIGHT)
            Exit Sub
        End If
        If X - mMouseX + Me.Width < LIMIT_WIDTH Then
            Call Me.Move(Me.Left, Me.Top, LIMIT_WIDTH, Y - mMouseY + Me.Height)
            Exit Sub
        End If
        If Y - mMouseY + Me.Height < LIMIT_HEIGHT Then
            Call Me.Move(Me.Left, Me.Top, X - mMouseX + Me.Width, LIMIT_HEIGHT)
            Exit Sub
        End If
        Call Me.Move(Me.Left, Me.Top, X - mMouseX + Me.Width, Y - mMouseY + Me.Height)
    End If
End Sub

Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbLeftButton) = vbLeftButton And FormResizing Then
        Call ReleaseCapture
        mMouseX = X: mMouseY = Y: FormMoving = True
        Call Form_MouseMove(vbLeftButton, Shift, X, Y)
        FormMoving = False
        FormResizing = False
    End If
End Sub

Public Sub Chat(ByVal Who As String, ByVal Talk As String)
    Dim Temp As String

    Temp = Who & ": " & Talk
    txtTalk.Text = Temp & vbCrLf & txtTalk.Text

    ' 存储聊天数据，以备保存。
    'Call lstChatData.AddItem(Temp)
End Sub

'Public Sub SysMsg(Message As String)
'    txtTalk.Text = Message & txtTalk.Text
'End Sub

Public Sub EnableChat(ByVal Name As String, ByVal Start As Boolean)
    lblUserName.Caption = Name & ":"
    If Start Then
        cboTalk.Enabled = True
        fltbtnSendTalk.Enabled = True
        txtTalk.Enabled = True
    End If
End Sub

Public Sub DisableChat(ByVal Finish As Boolean)
    If Finish Then
        lblUserName.Caption = LoadString(255)
    End If
    cboTalk.Enabled = False
    fltbtnSendTalk.Enabled = False
    txtTalk.Enabled = False
End Sub
