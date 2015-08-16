VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3900
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5250
   ControlBox      =   0   'False
   Enabled         =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmSplash.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   3900
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2310
      TabIndex        =   0
      Top             =   2310
      Width           =   120
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next

    lblVersion.Caption = App.Major & "." & App.Minor & " Build " & App.Revision
    Call Me.ZOrder(0)
    Call Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set frmSplash = Nothing
End Sub
