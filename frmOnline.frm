VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmOnline 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   4995
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmOnline.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList ilsFace 
      Left            =   4080
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8421376
      _Version        =   393216
   End
   Begin VB.Timer tmrReload 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4170
      Top             =   3810
   End
   Begin Othello.FlatButton fltbtnReload 
      Height          =   360
      Left            =   1050
      TabIndex        =   5
      Top             =   345
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   635
      Caption         =   "刷新"
      MousePointer    =   99
      Style           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmOnline.frx":000C
      OverBorderColor =   16744576
      BorderColor     =   65535
      OverBackColor   =   12648384
      DownBackColor   =   16777152
      BackColor       =   16761024
      ForeColor       =   0
   End
   Begin Othello.FlatButton fltbtnView 
      Height          =   360
      Left            =   135
      TabIndex        =   4
      Top             =   345
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   635
      Caption         =   "查看"
      MousePointer    =   99
      Style           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LightBorderColor1=   16711680
      DarkBorderColor1=   16711680
      HotColor        =   16777215
      EnableHot       =   -1  'True
      Icon            =   "frmOnline.frx":148E
      OverBorderColor =   16777215
      DownBorderColor =   8454143
      BorderColor     =   65535
      OverBackColor   =   16711680
      DownBackColor   =   4194304
      BackColor       =   16761024
      ForeColor       =   0
   End
   Begin VB.PictureBox picResize 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4635
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   5100
      Width           =   240
   End
   Begin MSComctlLib.ListView lvwOnline 
      Height          =   3210
      Left            =   105
      TabIndex        =   0
      Top             =   1170
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   5662
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名字"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "断线率"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "积分"
         Object.Width           =   2540
      EndProperty
   End
   Begin InetCtlsObjects.Inet ietOnline 
      Left            =   4095
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
      RequestTimeout  =   30
   End
   Begin Othello.FlatButton fltbtnHide 
      Height          =   240
      Left            =   3585
      TabIndex        =   2
      Top             =   45
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Caption         =   ""
      MousePointer    =   99
      Style           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   255
      ForeColor       =   0
      AutoSize        =   -1  'True
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   1920
      X2              =   3135
      Y1              =   5145
      Y2              =   5145
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   1875
      X2              =   3090
      Y1              =   5010
      Y2              =   5010
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   435
      X2              =   1650
      Y1              =   5145
      Y2              =   5145
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   390
      X2              =   1605
      Y1              =   5010
      Y2              =   5010
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   1965
      X2              =   3180
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   1920
      X2              =   3135
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   480
      X2              =   1695
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   435
      X2              =   1650
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   8
      X1              =   390
      X2              =   1605
      Y1              =   4845
      Y2              =   4845
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "--  在线用户  --"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   915
      TabIndex        =   3
      Top             =   75
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Shape shpTitle 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   255
      Top             =   30
      Width           =   3735
   End
End
Attribute VB_Name = "frmOnline"
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

Dim mlngSecond As Long

'Dim mblnServerOK As Boolean

Public Sub ShowEx(ByRef Frm As Form)
    On Error Resume Next

    Set Parent = Frm
    FormVisible = True
    Call FormMove(True)
    Call Me.Show(vbModeless)
End Sub

Public Sub HideEx()
    On Error Resume Next

    FormVisible = False
    MainForm.fltbtnOnline.Value = False
    Call Me.Hide
End Sub

Private Sub fltbtnReload_Click(Button As Integer)
    If Not gblnLogin Then Exit Sub
    Call ReloadOnline
End Sub

Private Sub fltbtnView_Click(Button As Integer)
    On Error Resume Next

    If Not gblnLogin Then Exit Sub

    If Not (lvwOnline.SelectedItem Is Nothing) Then
        'fltbtnView.Enabled = False
        Call frmUserInfo.ShowEx(lvwOnline.SelectedItem.Tag, lvwOnline.SelectedItem.Text)
        'fltbtnView.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    shpTitle.BackColor = CLR_ACTIVATE
End Sub

Private Sub Form_Deactivate()
    shpTitle.BackColor = CLR_DEACTIVATE
End Sub

Private Sub Form_Load()
    Dim i As Long

    On Error Resume Next

    Set picResize.Picture = frmResource.imgResResizer.Picture
    Set fltbtnHide.Picture = frmResource.imgResHide1.Picture
    Set fltbtnHide.DownPicture = frmResource.imgResHide2.Picture
    Set lvwOnline.ColumnHeaderIcons = frmResource.ilsResSortIcon

    DockingLevel = GetTwipX(10)
    'FormMoving = False
    'FormResizing = False

    linBorder(0).BorderColor = vbWhite
    linBorder(1).BorderColor = vbWhite
    linBorder(2).BorderColor = vbWhite
    linBorder(3).BorderColor = vbWhite

    linBorder(4).BorderColor = RGB(210, 220, 210)
    linBorder(5).BorderColor = RGB(180, 180, 180)
    linBorder(6).BorderColor = RGB(180, 180, 180)
    linBorder(7).BorderColor = RGB(210, 220, 210)

    linBorder(8).BorderColor = vbWhite

    For i = 1 To MainForm.ilsFace.ListImages.Count
        Call ilsFace.ListImages.Add(, , MainForm.ilsFace.ListImages(i).Picture)
    Next i
    Set lvwOnline.SmallIcons = ilsFace

    If gwifSave_OnlineWindow.DockStyle = dsDockNone Then
        Call Me.Move(gwifSave_OnlineWindow.Left, gwifSave_OnlineWindow.Top, gwifSave_OnlineWindow.Width, gwifSave_OnlineWindow.Height)
    Else
        DockPosition = gwifSave_OnlineWindow.DockPosition
        Me.Width = gwifSave_OnlineWindow.Width
        Me.Height = gwifSave_OnlineWindow.Height
    End If

    For i = 1 To MAX_ONLINE_ITEM
        lvwOnline.ColumnHeaders(i).Width = gsngSave_OnlineItemWidth(i)
        Call ShowHeaderIcon(lvwOnline, i - 1, 0, 0)
    Next i

    If glngSave_OnlineSort > 0 Then
        With lvwOnline
            .Sorted = True
            .SortOrder = glngSave_OnlineSort - 1
            .SortKey = glngSave_OnlineSortKey
        End With
        Call ShowHeaderIcon(lvwOnline, lvwOnline.SortKey, lvwOnline.SortOrder, HDF_IMAGE)
    Else
        With lvwOnline
            .Sorted = False
            .SortOrder = lvwDescending
            .SortKey = 0
        End With
    End If
    If gblnSave_OnlineAutoReload Then
        tmrReload.Enabled = True
    End If
    'fltbtnView.ForeColor = vbWhite
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

    'Call lblToolbar.Move(210, 600)

    Call lvwOnline.Move(100, 800, ScaleWidth - GetTwipX(30), ScaleHeight - GetTwipY(80))

    'Call ColumnSize(lvwOnline, 1, 40)
    'Call ColumnSize(lvwOnline, 2, 30)
    'Call ColumnSize(lvwOnline, 3, 28)

    'Call fltbtnView.Move(lvwOnline.Left + lvwOnline.Width + GetTwipX(7), lvwOnline.Top)
    'Call fltbtnReload.Move(lvwOnline.Left + lvwOnline.Width + GetTwipX(7), fltbtnView.Top + fltbtnView.Height + GetTwipY(10))

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

Private Sub ietOnline_StateChanged(ByVal State As Integer)
    Dim strStatus As String
    Dim strData As String
    Dim lngSelected As Long
    Dim itmX As ListItem

    On Error Resume Next

    Select Case State
        Case icResponseCompleted
            If GetServerExecute(ietOnline, strStatus, strData) Then
                Select Case strStatus
                    Case STATUS_OK
                        If Not (lvwOnline.SelectedItem Is Nothing) Then
                            lngSelected = lvwOnline.SelectedItem.Index
                        End If
                        Call AddListView(strData)
                        If lngSelected > 0 And lngSelected <= lvwOnline.ListItems.Count Then
                            Set itmX = lvwOnline.ListItems.Item(lngSelected)
                        End If
                        If Not (itmX Is Nothing) Then
                            itmX.Selected = True
                        End If
                    Case STATUS_NONE
                        lvwOnline.ListItems.Clear
                End Select
            End If
            tmrReload.Enabled = True
    End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' 子程序: AddListView
'''
''' 描述:   把接收到的数据显示到列表中。
'''
''' 参数:   strInfo - 接收字符串
'''
''' 日期:   2002.6.17
'''
''' 作者:   赵畅
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddListView(ByVal strInfo As String)
    Dim strRecord As String
    Dim OldSorted As Boolean
    Dim i As Long
    Dim GameTimes As Long
    Dim PlayTimes As Long
    Dim itmX As ListItem

    On Error Resume Next

    OldSorted = lvwOnline.Sorted
    lvwOnline.Sorted = False
    Call lvwOnline.ListItems.Clear

    For i = 1 To GetFieldCount(strInfo)
        strRecord = GetField(strInfo, i)
        GameTimes = Val(GetRecord(strRecord, 4))
        PlayTimes = Val(GetRecord(strRecord, 5))
        Set itmX = lvwOnline.ListItems.Add(, , GetDisplayName(GetRecord(strRecord, 1), GetRecord(strRecord, 2)))
        itmX.SmallIcon = Val(GetRecord(strRecord, 3))
        itmX.Tag = GetRecord(strRecord, 1)
        ' 如果是自己，则显示为蓝色。
        If itmX.Tag = gMyUserInfo.UserName Then
            itmX.ForeColor = vbBlue
        End If
        If gblnGameStart Then
            If GameTimes < 2 Then
                itmX.SubItems(1) = "0%"
            Else
                itmX.SubItems(1) = Format((GameTimes - 1 - PlayTimes) / (GameTimes - 1), "0%")
            End If
        Else
            If GameTimes < 1 Then
                itmX.SubItems(1) = "0%"
            Else
                itmX.SubItems(1) = Format((GameTimes - PlayTimes) / GameTimes, "0%")
            End If
        End If
        itmX.SubItems(2) = GetRecord(strRecord, 6)
    Next i

    lvwOnline.Sorted = OldSorted
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

Private Sub lvwOnline_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim i As Long

    On Error Resume Next

    For i = 0 To lvwOnline.ColumnHeaders.Count - 1
        Call ShowHeaderIcon(lvwOnline, i, 0, 0)
    Next i

    If lvwOnline.SortKey = ColumnHeader.Index - 1 Then
        ' 交换升降序！
        lvwOnline.SortOrder = ToPartner(lvwOnline.SortOrder, 1)
    Else
        lvwOnline.SortOrder = lvwAscending
    End If

    Call ShowHeaderIcon(lvwOnline, ColumnHeader.Index - 1, lvwOnline.SortOrder, HDF_IMAGE)
    lvwOnline.SortKey = ColumnHeader.Index - 1

    lvwOnline.Sorted = True
End Sub

Private Sub lvwOnline_DblClick()
    Call fltbtnView_Click(vbLeftButton)
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
        If X - mMouseX + Me.Width < lblTitle.Width + fltbtnHide.Width + 100 And Y - mMouseY + Me.Height < LIMIT_HEIGHT Then
            Call Me.Move(Me.Left, Me.Top, lblTitle.Width + fltbtnHide.Width + 100, LIMIT_HEIGHT)
            Exit Sub
        End If
        If X - mMouseX + Me.Width < lblTitle.Width + fltbtnHide.Width + 100 Then
            Call Me.Move(Me.Left, Me.Top, lblTitle.Width + fltbtnHide.Width + 100, Y - mMouseY + Me.Height)
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

Public Function ReloadOnline() As Boolean
    Dim strUrl As String

    On Error Resume Next

    If Not gblnLogin Then
        Call lvwOnline.ListItems.Clear
        ReloadOnline = True
        Exit Function
    End If

    tmrReload.Enabled = False

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_ONLINE_GET & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    ReloadOnline = ServerExecute(ietOnline, strUrl)
    'Else
        'MsgBox "刷新在线用户错误！" & vbCr & vbCr & LoadString(101), vbCritical, "在线"
    'End If
    'tmrReload.Enabled = True

    'ReloadOnline = False
End Function

Private Sub tmrReload_Timer()
    On Error Resume Next

    If Not gblnLogin Or Not gblnSave_OnlineAutoReload Then
        tmrReload.Enabled = False
        mlngSecond = 0
        Exit Sub
    End If

    mlngSecond = mlngSecond + 1
    If mlngSecond > glngSave_OnlineAutoReloadTime Then
        mlngSecond = 0
        'tmrReload.Enabled = False
        If Me.Visible Then Call ReloadOnline
        'tmrReload.Enabled = True
    End If
End Sub
