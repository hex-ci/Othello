VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmTable 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   4740
   ControlBox      =   0   'False
   Icon            =   "frmTable.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin Othello.FlatButton fltbtnReload 
      Height          =   330
      Left            =   3675
      TabIndex        =   6
      Top             =   2520
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
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
   Begin Othello.FlatButton fltbtnJoin 
      Height          =   330
      Left            =   3675
      TabIndex        =   5
      Top             =   1170
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
      Caption         =   "加入"
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
   Begin VB.PictureBox picResize 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4410
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   5100
      Width           =   240
   End
   Begin VB.Timer tmrReload 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   4095
   End
   Begin InetCtlsObjects.Inet ietTable 
      Left            =   3885
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
      RequestTimeout  =   30
   End
   Begin MSComctlLib.ListView lvwTable 
      Height          =   3360
      Left            =   240
      TabIndex        =   0
      Top             =   1170
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   5927
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "TableName"
         Text            =   "棋局名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Creator"
         Text            =   "创建人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Visitor"
         Text            =   "访问者"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "CanJoin"
         Text            =   "加入"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "TableType"
         Text            =   "棋局类型"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Timer"
         Text            =   "计时器"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "UpLevel"
         Text            =   "晋级"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "TableTime"
         Text            =   "更新时间"
         Object.Width           =   2540
      EndProperty
   End
   Begin Othello.FlatButton fltbtnHide 
      Height          =   240
      Left            =   3000
      TabIndex        =   3
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
   Begin Othello.FlatButton fltbtnCreator 
      Height          =   330
      Left            =   3675
      TabIndex        =   7
      Top             =   2070
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
      Caption         =   "创建人"
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
   Begin Othello.FlatButton fltbtnView 
      Height          =   330
      Left            =   3675
      TabIndex        =   8
      Top             =   1620
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
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
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "--  棋局列表  --"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   825
      TabIndex        =   4
      Top             =   75
      Width           =   1455
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   8
      X1              =   585
      X2              =   1800
      Y1              =   4905
      Y2              =   4905
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   2025
      X2              =   3240
      Y1              =   5175
      Y2              =   5175
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   1980
      X2              =   3195
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   540
      X2              =   1755
      Y1              =   5175
      Y2              =   5175
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   495
      X2              =   1710
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   2070
      X2              =   3285
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   2025
      X2              =   3240
      Y1              =   4635
      Y2              =   4635
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   585
      X2              =   1800
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   540
      X2              =   1755
      Y1              =   4635
      Y2              =   4635
   End
   Begin VB.Label lblTips 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "请选择棋局:"
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   765
      Width           =   990
   End
   Begin VB.Shape shpTitle 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   480
      Top             =   30
      Width           =   2895
   End
End
Attribute VB_Name = "frmTable"
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

Dim mblnServerOK As Boolean

Private Sub fltbtnCreator_Click(Button As Integer)
    If Not gblnLogin Then Exit Sub

    If lvwTable.SelectedItem Is Nothing Then Exit Sub

    fltbtnCreator.Enabled = False
    Call frmUserInfo.ShowEx(lvwTable.SelectedItem.Tag, lvwTable.SelectedItem.SubItems(1))
    fltbtnCreator.Enabled = True
End Sub

Private Sub fltbtnView_Click(Button As Integer)
    ' 查看棋局详细信息。
    If Not gblnLogin Then Exit Sub

    If lvwTable.SelectedItem Is Nothing Then Exit Sub

    fltbtnView.Enabled = False
    Call frmTableInfo.ShowEx(lvwTable.SelectedItem.Text, False)
    fltbtnView.Enabled = True
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
    Set lvwTable.ColumnHeaderIcons = frmResource.ilsResSortIcon

    DockingLevel = GetTwipX(10)
    'FormMoving = False
    'FormResizing = False

    'ReloadTimer.Enabled = True
    'Set lvwTable.ColumnHeaderIcons = ilsIcon
    'Call ShowHeaderIcon(lvwTable, 1, 0, HDF_IMAGE)

    linBorder(0).BorderColor = vbWhite
    linBorder(1).BorderColor = vbWhite
    linBorder(2).BorderColor = vbWhite
    linBorder(3).BorderColor = vbWhite

    linBorder(4).BorderColor = RGB(210, 220, 210)
    linBorder(5).BorderColor = RGB(180, 180, 180)
    linBorder(6).BorderColor = RGB(180, 180, 180)
    linBorder(7).BorderColor = RGB(210, 220, 210)

    linBorder(8).BorderColor = vbWhite

    If gwifSave_TableWindow.DockStyle = dsDockNone Then
        Call Me.Move(gwifSave_TableWindow.Left, gwifSave_TableWindow.Top, gwifSave_TableWindow.Width, gwifSave_TableWindow.Height)
    Else
        DockPosition = gwifSave_TableWindow.DockPosition
        Me.Width = gwifSave_TableWindow.Width
        Me.Height = gwifSave_TableWindow.Height
    End If

    For i = 1 To MAX_TABLE_ITEM
        lvwTable.ColumnHeaders(i).Width = gsngSave_TableItemWidth(i)
        Call ShowHeaderIcon(lvwTable, i - 1, 0, 0)
    Next i

    If glngSave_TableSort > 0 Then
        With lvwTable
            .Sorted = True
            .SortOrder = glngSave_TableSort - 1
            .SortKey = glngSave_TableSortKey
        End With
        Call ShowHeaderIcon(lvwTable, lvwTable.SortKey, lvwTable.SortOrder, HDF_IMAGE)
    Else
        With lvwTable
            .Sorted = False
            .SortOrder = lvwDescending
            .SortKey = 0
        End With
    End If
    If gblnSave_TableAutoReload Then
        tmrReload.Enabled = True
    End If
End Sub

Public Sub ShowEx(ByRef Frm As Form)
    On Error Resume Next

    Set Parent = Frm
    FormVisible = True
    Call FormMove(True)
    'If (Not Started) And gblnLogin Then StartTimer.Enabled = True
    Call Me.Show(vbModeless)
End Sub

Public Sub HideEx()
    FormVisible = False
    MainForm.fltbtnTable.Value = False
    Call Me.Hide
End Sub

Private Sub fltbtnReload_Click(Button As Integer)
    'If Not fltbtnReload.Enabled Or Not Me.Visible Then Exit Sub
    If Not gblnLogin Then Exit Sub
    Call ReloadTable(True)
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

    Call lblTips.Move(210, 600)

    Call lvwTable.Move(200, 900, ScaleWidth - GetTwipX(90), ScaleHeight - GetTwipY(80))

    'Call ColumnSize(lvwTable, 1, 30)
    'Call ColumnSize(lvwTable, 2, 30)
    'Call ColumnSize(lvwTable, 3, 20)
    'Call ColumnSize(lvwTable, 4, 19)

    Call fltbtnJoin.Move(lvwTable.Left + lvwTable.Width + GetTwipX(7), lvwTable.Top)
    Call fltbtnView.Move(lvwTable.Left + lvwTable.Width + GetTwipX(7), fltbtnJoin.Top + fltbtnJoin.Height + GetTwipY(5))
    Call fltbtnCreator.Move(lvwTable.Left + lvwTable.Width + GetTwipX(7), fltbtnView.Top + fltbtnView.Height + GetTwipY(5))
    Call fltbtnReload.Move(lvwTable.Left + lvwTable.Width + GetTwipX(7), fltbtnCreator.Top + fltbtnCreator.Height + GetTwipY(5))

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

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    Call ietTable.Cancel
    tmrReload.Enabled = False
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, X + lblTitle.Left, Y + lblTitle.Top)
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetCapture(Me.hWnd)
End Sub

Private Sub ietTable_StateChanged(ByVal State As Integer)
    'Dim strMess As String '消息变量。
    On Error Resume Next

    Select Case State
        Case icResponseReceived
            mblnServerOK = True
        Case icError  '11
            mblnServerOK = False
            tmrReload.Enabled = True
            '得到错误文本。
            'strMess = "ErrorCode: " & ietRegister.ResponseCode & " : " & ietRegister.ResponseInfo
        Case icResponseCompleted
            mblnServerOK = True
            Dim strStatus As String
            Dim strData As String

            If GetServerExecute(ietTable, strStatus, strData) Then
                Select Case strStatus
                    Case STATUS_OK
                        Dim lngSelected As Long
                        Dim itmX As ListItem
                        If Not (lvwTable.SelectedItem Is Nothing) Then
                            lngSelected = lvwTable.SelectedItem.Index
                        End If
                        'MsgBox GetRecord(GetField(msg, 2), 1), vbExclamation, "正确"
                        Call AddListView(strData)
                        If lngSelected > 0 And lngSelected <= lvwTable.ListItems.Count Then
                            Set itmX = lvwTable.ListItems.Item(lngSelected)
                        End If
                        If Not (itmX Is Nothing) Then
                            itmX.Selected = True
                        End If
                    Case STATUS_NONE
                        Call lvwTable.ListItems.Clear
                End Select
            End If
            tmrReload.Enabled = True
    End Select
End Sub

Private Sub fltbtnJoin_Click(Button As Integer)
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    On Error Resume Next

    If (Not gblnLogin) Or gblnCreator Or gblnConnect Or (lvwTable.SelectedItem Is Nothing) Then Exit Sub

    fltbtnJoin.Enabled = False
    If Not ReloadTable(False) Then Exit Sub


    strUrl = gstrSave_ServerUrl & SERVER_ACTION_TABLE_VIEW & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&name=" & ToUrlString(lvwTable.SelectedItem.Text) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If ServerCommand(ietTable, mblnServerOK, strUrl, strStatus, strData) Then
        Select Case strStatus
            Case STATUS_OK
                'lvwTable_GotFocus
                Dim TableInfo As tagTableInfo
                Call LoadTableInfo(TableInfo, strData)
                If TableInfo.Visitor = "" Then
                    ' 临时存储棋局名及类型
                    gMainTableInfo.TableName = TableInfo.TableName
                    gMainTableInfo.TableType = TableInfo.TableType
                    Call SetControlFocus(MainForm)

                    If Len(TableInfo.LANIP) > 5 Then
                        ' 试所有IP。
                        'MsgBox GetRecord(strData, 13) & ";" & TableInfo.IP
                        Call MainForm.ReadyTryJoin(TableInfo.LANIP, TableInfo.ip, TableInfo.Port)
                    Else
                        Call MainForm.ReadyJoinTable(TableInfo.ip, TableInfo.Port)
                    End If
                Else
                    Call MessageBox(Me.hWnd, LoadString(160), vbExclamation, LoadString(181))
                End If
                'MsgBox GetRecord(GetField(msg, 2), 1), vbExclamation, "正确"
                'Call AddListView(Msg)
            Case STATUS_ERROR
                Call MessageBox(Me.hWnd, LoadString(143) & strData, vbExclamation, LoadString(181))
            'Case STATUS_NONE
                'lvwTable.ListItems.Clear
                'Call MessageBox(Me.hWnd, strData, vbExclamation, "加入棋局")
            Case Else
                Call MessageBox(Me.hWnd, LoadString(144), vbExclamation, LoadString(181))
        End Select
    'Else
        'Call MessageBox(hWnd, "加入棋局列表错误！" & vbCr & vbCr & LoadString(101), vbCritical, "棋局")
    End If
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
    'Dim strRecord As String
    Dim OldSorted As Boolean
    Dim TempTableInfo As tagTableInfo
    Dim itmX As ListItem
    Dim i As Long

    On Error Resume Next

    OldSorted = lvwTable.Sorted
    lvwTable.Sorted = False
    Call lvwTable.ListItems.Clear

    For i = 1 To GetFieldCount(strInfo)
        Call LoadTableInfo(TempTableInfo, GetField(strInfo, i))
        Set itmX = lvwTable.ListItems.Add(, , TempTableInfo.TableName)
        ' 把自己创建的或加入的棋局显示为蓝色。
        If itmX.Text = gMainTableInfo.TableName Then
            itmX.ForeColor = vbBlue
        End If

        With itmX
            .Tag = TempTableInfo.Creator
            .SubItems(1) = GetDisplayName(TempTableInfo.Creator, TempTableInfo.CreatorName)
            .SubItems(2) = GetDisplayName(TempTableInfo.Visitor, TempTableInfo.VisitorName)
            .SubItems(3) = IIf(TempTableInfo.Visitor = "", LoadString(211), LoadString(212))
            If TempTableInfo.TableType = TABLE_PUBLIC Then
                .SubItems(4) = LoadString(213)
            Else
                .SubItems(4) = LoadString(214)
            End If
            If TempTableInfo.Timer > 0 Then
                .SubItems(5) = CStr(TempTableInfo.Timer) & LoadString(215)
            Else
                .SubItems(5) = LoadString(216)
            End If
            .SubItems(6) = IIf(TempTableInfo.UpLevel, LoadString(217), LoadString(218))
            .SubItems(7) = Format(TempTableInfo.LastTime, "yyyy年m月d日 hh:mm:ss")
        End With
    Next i

    lvwTable.Sorted = OldSorted
End Sub

Private Sub lvwTable_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim i As Long

    On Error Resume Next

    For i = 0 To lvwTable.ColumnHeaders.Count - 1
        Call ShowHeaderIcon(lvwTable, i, 0, 0)
    Next i

    If lvwTable.SortKey = ColumnHeader.Index - 1 Then
        ' 交换升降序！
        lvwTable.SortOrder = ToPartner(lvwTable.SortOrder, 1)
    Else
        lvwTable.SortOrder = lvwAscending
    End If

    Call ShowHeaderIcon(lvwTable, ColumnHeader.Index - 1, lvwTable.SortOrder, HDF_IMAGE)
    lvwTable.SortKey = ColumnHeader.Index - 1

    lvwTable.Sorted = True
End Sub

Private Sub lvwTable_DblClick()
    If fltbtnJoin.Enabled Then
        Call fltbtnJoin_Click(vbLeftButton)
    End If
End Sub

Private Sub lvwTable_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.SubItems(2) = "" And (Not gblnCreator) And (Not gblnConnect) Then
        fltbtnJoin.Enabled = True
    Else
        fltbtnJoin.Enabled = False
    End If
End Sub

Private Sub fltbtnHide_Click(Button As Integer)
    Call Me.HideEx
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

Public Function ReloadTable(ByVal Async As Boolean) As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    On Error Resume Next

    If Not gblnLogin Then
        Call lvwTable.ListItems.Clear
        ReloadTable = True
        Exit Function
    End If

    tmrReload.Enabled = False

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_TABLE_GET & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If Async Then
        ReloadTable = ServerExecute(ietTable, strUrl)
    Else
        If ServerCommand(ietTable, mblnServerOK, strUrl, strStatus, strData, True) Then
            tmrReload.Enabled = True
            Select Case strStatus
                Case STATUS_OK
                    Dim lngSelected As Long
                    Dim itmX As ListItem
                    If Not (lvwTable.SelectedItem Is Nothing) Then
                        lngSelected = lvwTable.SelectedItem.Index
                    End If
                    'MsgBox GetRecord(GetField(msg, 2), 1), vbExclamation, "正确"
                    Call AddListView(strData)
                    If lngSelected > 0 And lngSelected <= lvwTable.ListItems.Count Then
                        Set itmX = lvwTable.ListItems.Item(lngSelected)
                    End If
                    If Not (itmX Is Nothing) Then
                        itmX.Selected = True
                    End If
                    ReloadTable = True
                    Exit Function
                'Case STATUS_ERROR
                    'Call MessageBox(Me.hwnd, strData, vbCritical, "棋局")
                Case STATUS_NONE
                    Call lvwTable.ListItems.Clear
                    'Call MessageBox(Me.hwnd, msg, vbCritical, "棋局")
                'Case Else
                    'Call MessageBox(Me.hwnd, strData, vbCritical, "棋局")
                    'Call MessageBox(Me.hwnd, "错误！请重试！", vbCritical, "棋局")
            End Select
        'Else
            'MsgBox "刷新棋局列表错误！" & vbCr & vbCr & LoadString(101), vbCritical, "棋局"
        End If
        tmrReload.Enabled = True

        ReloadTable = False
    End If
End Function

Public Function AutoJoin() As Boolean
    Dim strUrl As String
    Dim strStatus As String
    Dim strData As String

    On Error Resume Next

    If (Not ReloadTable(False)) Or lvwTable.ListItems.Count < 1 Then
        AutoJoin = False
        Exit Function
    End If

    strUrl = gstrSave_ServerUrl & SERVER_ACTION_TABLE_AUTOJOIN & _
                                  "?username=" & ToUrlString(gMyUserInfo.UserName) & _
                                  "&password=" & MD5(gMyUserInfo.Password) & _
                                  "&" & MakeServerPassword() & _
                                  "&" & MakeVersion()

    If ServerCommand(ietTable, mblnServerOK, strUrl, strStatus, strData, True) Then
        Select Case strStatus
            Case STATUS_OK
                AutoJoin = BestJoin(strData)
            Case Else
                AutoJoin = False
                Exit Function
        End Select
    Else
        AutoJoin = False
    End If
End Function

Private Function BestJoin(ByVal strInfo As String) As Boolean
    Dim TableInfo() As tagTableInfo
    Dim CreatorInfo() As tagUserInfo
    Dim Count As Long
    Dim i As Long
    Dim Temp As String
    Dim BestTableInfo As tagTableInfo
    Dim BestCreatorInfo As tagUserInfo

    On Error Resume Next

    Count = GetFieldCount(strInfo) \ 2

    If Count < 1 Then
        BestJoin = False
        Exit Function
    End If

    ReDim TableInfo(Count) As tagTableInfo
    ReDim CreatorInfo(Count) As tagUserInfo

    For i = 1 To GetFieldCount(strInfo) Step 2
        If GetField(strInfo, i) <> "" Then
            Call LoadTableInfo(TableInfo(i \ 2 + 1), GetField(strInfo, i))
        End If
        If GetField(strInfo, i + 1) <> "" Then
            Call LoadUserInfo(CreatorInfo(i \ 2 + 1), GetField(strInfo, 1 + 1))
        End If
    Next i

    If TableInfo(1).Visitor <> "" Or CreatorInfo(1).UserName = gMyUserInfo.UserName Then
        BestJoin = False
        Exit Function
    End If

    BestTableInfo = TableInfo(1)
    BestCreatorInfo = CreatorInfo(1)

    For i = 2 To Count
        If CreatorInfo(i).UserName <> gMyUserInfo.UserName _
           And TableInfo(i).Visitor = "" _
           And (TableInfo(i).LANIP = "" Or TableInfo(i).ip = gstrIP) _
           And Abs(CreatorInfo(i).Score - gMyUserInfo.Score) < Abs(BestCreatorInfo.Score - gMyUserInfo.Score) Then
                BestTableInfo = TableInfo(i)
                BestCreatorInfo = CreatorInfo(i)
        End If
    Next i

    ' 在棋局列表中，查找并定位棋局
    Set lvwTable.SelectedItem = lvwTable.FindItem(BestTableInfo.TableName, lvwText)

    ' 加入棋局！
    Call fltbtnJoin_Click(vbLeftButton)

    BestJoin = True
End Function

Private Sub tmrReload_Timer()
    On Error Resume Next

    If Not gblnLogin Or Not gblnSave_TableAutoReload Then
        tmrReload.Enabled = False
        mlngSecond = 0
        Exit Sub
    End If

    mlngSecond = mlngSecond + 1
    If mlngSecond > glngSave_TableAutoReloadTime Then
        mlngSecond = 0
        'tmrReload.Enabled = False
        If Me.Visible Then Call ReloadTable(True)
        'tmrReload.Enabled = True
    End If
End Sub
