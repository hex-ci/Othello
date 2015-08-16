VERSION 5.00
Begin VB.Form frmEditPlayList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "编辑播放列表项目"
   ClientHeight    =   1440
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6060
   Icon            =   "frmEditPlayList.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin Othello.FlatButton fltbtnCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4905
      TabIndex        =   3
      Top             =   1005
      Width           =   1005
      _extentx        =   1773
      _extenty        =   582
      caption         =   "取消"
      mousepointer    =   99
      style           =   2
      font            =   "frmEditPlayList.frx":000C
      enablehot       =   -1  'True
      forecolor       =   0
   End
   Begin Othello.FlatButton fltbtnOK 
      Default         =   -1  'True
      Height          =   330
      Left            =   3780
      TabIndex        =   2
      Top             =   1005
      Width           =   1005
      _extentx        =   1773
      _extenty        =   582
      caption         =   "确定"
      mousepointer    =   99
      style           =   2
      font            =   "frmEditPlayList.frx":0030
      enablehot       =   -1  'True
      forecolor       =   0
   End
   Begin VB.TextBox txtNew 
      Height          =   330
      Left            =   795
      TabIndex        =   1
      Top             =   540
      Width           =   5130
   End
   Begin VB.TextBox txtOld 
      BackColor       =   &H8000000F&
      Height          =   345
      Left            =   795
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   5130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "新项目:"
      Height          =   180
      Left            =   135
      TabIndex        =   5
      Top             =   600
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "旧项目:"
      Height          =   180
      Left            =   135
      TabIndex        =   4
      Top             =   255
      Width           =   630
   End
End
Attribute VB_Name = "frmEditPlayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrItem As String

Public Function ShowEx(ByVal Item As String) As String
    On Error Resume Next

    txtOld.Text = Item
    txtNew.Text = Item
    Call Show(vbModal, MainForm)
    ShowEx = mstrItem
End Function

Private Sub fltbtnCancel_Click(Button As Integer)
    mstrItem = ""
    Call Unload(Me)
End Sub

Private Sub fltbtnOK_Click(Button As Integer)
    mstrItem = txtNew.Text
    Call Unload(Me)
End Sub

Private Sub Form_Activate()
    Call SetControlFocus(txtNew)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEditPlayList = Nothing
End Sub

Private Sub txtNew_GotFocus()
    Call AutoSelectText(txtNew)
End Sub
