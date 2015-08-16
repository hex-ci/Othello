VERSION 5.00
Begin VB.Form frmCreateTable 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "创建棋局"
   ClientHeight    =   3075
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5040
   Icon            =   "frmCreateTable.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin Othello.FlatButton fltbtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   2550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "取消(&E)"
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
   Begin Othello.FlatButton fltbtnCreate 
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "创建棋局(&C)"
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
   Begin VB.ComboBox cboGameTimer 
      Height          =   300
      ItemData        =   "frmCreateTable.frx":030A
      Left            =   1320
      List            =   "frmCreateTable.frx":0320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   3225
   End
   Begin VB.CheckBox chkUpLevel 
      Caption         =   "晋级游戏(&U)"
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1815
      Width           =   1380
   End
   Begin VB.ComboBox cboType 
      Height          =   300
      ItemData        =   "frmCreateTable.frx":0356
      Left            =   1320
      List            =   "frmCreateTable.frx":0360
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   3225
   End
   Begin VB.TextBox txtTableName 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   3225
   End
   Begin VB.Label Label3 
      Caption         =   "计时器(&T):"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1380
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "类型(&Y):"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   900
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "棋局名(&N):"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   420
      Width           =   915
   End
End
Attribute VB_Name = "frmCreateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fltbtnCancel_Click(Button As Integer)
    Call Unload(Me)
End Sub

Private Sub fltbtnCreate_Click(Button As Integer)
    On Error Resume Next

    If gblnOfflineMode Then
        gMainTableInfo.Timer = 0

        Call Unload(Me)
        ' 这里调用主窗体中的创建棋局函数（单机模式）
        Call MainForm.CreateOfflineTable
    Else
        If Len(txtTableName.Text) < 2 Or Len(txtTableName.Text) > 15 Or Not CheckString(txtTableName.Text) Then
            Call MessageBox(Me.hWnd, LoadString(172), vbCritical, LoadString(181))
            Call SetControlFocus(txtTableName)
            Call SendKeys("{Home}+{End}")
            Exit Sub
        End If
    
        gMainTableInfo.Creator = gMyUserInfo.UserName
        gMainTableInfo.CreatorName = gMyUserInfo.Name
        gMainTableInfo.Visitor = ""
        gMainTableInfo.VisitorName = ""
        gMainTableInfo.TableName = txtTableName.Text
        gMainTableInfo.TableType = cboType.ItemData(cboType.ListIndex)
        gMainTableInfo.Timer = cboGameTimer.ItemData(cboGameTimer.ListIndex)
        gMainTableInfo.UpLevel = CBool(chkUpLevel.Value)
    
        Call Unload(Me)
        Call MainForm.ReadyCreateTable
    End If
End Sub

Private Sub Form_Activate()
    Call SetControlFocus(txtTableName)
End Sub

Private Sub Form_Load()
    On Error Resume Next

    If glngSave_TableType < 0 Or glngSave_TableType > 1 Then glngSave_TableType = 0
    If glngSave_TableTimer < 0 Or glngSave_TableTimer > 5 Then glngSave_TableTimer = 0
    cboType.ListIndex = glngSave_TableType
    cboGameTimer.ListIndex = glngSave_TableTimer
    chkUpLevel.Value = glngSave_TableUpLevel

    If gblnOfflineMode Then
        txtTableName.Enabled = False
        cboType.Enabled = False
        cboGameTimer.Enabled = False
        cboGameTimer.ListIndex = 0
        chkUpLevel.Enabled = False
        Me.Caption = "创建棋局 -- 单机模式"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next

    glngSave_TableType = cboType.ListIndex
    glngSave_TableTimer = cboGameTimer.ListIndex
    glngSave_TableUpLevel = chkUpLevel.Value
End Sub
