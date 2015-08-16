VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResource 
   BorderStyle     =   0  'None
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmResource.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ilsResSortIcon 
      Left            =   4815
      Top             =   5895
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResource.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResource.frx":0084
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgResResizer 
      Height          =   240
      Left            =   4485
      Picture         =   "frmResource.frx":00FE
      Top             =   5370
      Width           =   240
   End
   Begin VB.Image imgResHide2 
      Height          =   240
      Left            =   3975
      Picture         =   "frmResource.frx":01F2
      Top             =   5355
      Width           =   240
   End
   Begin VB.Image imgResHide1 
      Height          =   240
      Left            =   3465
      Picture         =   "frmResource.frx":0352
      Top             =   5355
      Width           =   240
   End
   Begin VB.Image imgResTips 
      Height          =   525
      Left            =   2775
      Picture         =   "frmResource.frx":04B8
      Top             =   5355
      Width           =   525
   End
   Begin VB.Image imgResSelectDown 
      Height          =   525
      Left            =   930
      Picture         =   "frmResource.frx":06AC
      Top             =   5970
      Width           =   525
   End
   Begin VB.Image imgResSelectIcon 
      Height          =   525
      Left            =   930
      Picture         =   "frmResource.frx":0738
      Top             =   5325
      Width           =   525
   End
   Begin VB.Image imgResSelBlackMan 
      Height          =   525
      Left            =   1515
      Picture         =   "frmResource.frx":07AE
      Top             =   5955
      Width           =   525
   End
   Begin VB.Image imgResDefaultFace 
      Height          =   480
      Left            =   2790
      Picture         =   "frmResource.frx":0B2A
      Top             =   6000
      Width           =   480
   End
   Begin VB.Image imgResSelWhiteMan 
      Height          =   525
      Left            =   2145
      Picture         =   "frmResource.frx":0CDA
      Top             =   5970
      Width           =   525
   End
   Begin VB.Image imgResWhiteMan 
      Height          =   525
      Left            =   2145
      Picture         =   "frmResource.frx":1063
      Top             =   5355
      Width           =   525
   End
   Begin VB.Image imgResBlackMan 
      Height          =   525
      Left            =   1515
      Picture         =   "frmResource.frx":13E6
      Top             =   5355
      Width           =   525
   End
   Begin VB.Image imgResFooter 
      Height          =   285
      Left            =   315
      Picture         =   "frmResource.frx":1763
      Top             =   6780
      Width           =   1650
   End
   Begin VB.Image imgResOpenFile 
      Height          =   240
      Left            =   3510
      Picture         =   "frmResource.frx":2060
      Top             =   5985
      Width           =   240
   End
   Begin VB.Image imgResSoundStop 
      Height          =   345
      Left            =   3915
      Picture         =   "frmResource.frx":20DB
      Top             =   5940
      Width           =   345
   End
   Begin VB.Image imgResSoundPlay 
      Height          =   345
      Left            =   4275
      Picture         =   "frmResource.frx":2130
      Top             =   5955
      Width           =   345
   End
   Begin VB.Image imgResLightYellow 
      Height          =   300
      Left            =   465
      Picture         =   "frmResource.frx":2184
      Top             =   6060
      Width           =   300
   End
   Begin VB.Image imgResWizard 
      Height          =   3555
      Left            =   8400
      Picture         =   "frmResource.frx":2372
      Top             =   3930
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image imgResFinished 
      Height          =   3555
      Left            =   6285
      Picture         =   "frmResource.frx":4740
      Top             =   3945
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image imgResLightOff 
      Height          =   300
      Left            =   465
      Picture         =   "frmResource.frx":6CC6
      Top             =   5685
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgResLightOn 
      Height          =   300
      Left            =   465
      Picture         =   "frmResource.frx":6EA8
      Top             =   5325
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMenuSide 
      Height          =   7200
      Left            =   5565
      Picture         =   "frmResource.frx":7097
      Top             =   285
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgResTitle1 
      Height          =   1575
      Left            =   6285
      Picture         =   "frmResource.frx":7C82
      Top             =   315
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Image imgResTitle 
      Height          =   1575
      Left            =   6315
      Picture         =   "frmResource.frx":9ABA
      Top             =   2100
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Image imgResChessBoard 
      Height          =   5100
      Left            =   150
      Picture         =   "frmResource.frx":BDB2
      Top             =   180
      Visible         =   0   'False
      Width           =   5100
   End
End
Attribute VB_Name = "frmResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next

    Set ChessBoard = imgResChessBoard.Picture
    Set GameTitle = imgResTitle.Picture
    Set NoFocusTitle = imgResTitle1.Picture
    Set objLightOn = imgResLightOn.Picture
    Set objLightOff = imgResLightOff.Picture
    Set objLightYellow = imgResLightYellow.Picture
    Set objSoundPlay = imgResSoundPlay.Picture
    Set objSoundStop = imgResSoundStop.Picture
    Set objOpenFile = imgResOpenFile.Picture

    Set BlackMan = imgResBlackMan.Picture
    Set WhiteMan = imgResWhiteMan.Picture
    Set SelBlackMan = imgResSelBlackMan.Picture
    Set SelWhiteMan = imgResSelWhiteMan.Picture
    Set TipsBitmap = imgResTips.Picture
    Set SelectIcon = imgResSelectIcon.Picture
    Set SelectDown = imgResSelectDown.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    Set ChessBoard = Nothing
    Set GameTitle = Nothing
    Set NoFocusTitle = Nothing
    Set objLightOn = Nothing
    Set objLightOff = Nothing
    Set objLightYellow = Nothing
    Set objSoundPlay = Nothing
    Set objSoundStop = Nothing
    Set objOpenFile = Nothing

    Set BlackMan = Nothing
    Set WhiteMan = Nothing
    Set SelBlackMan = Nothing
    Set SelWhiteMan = Nothing
    Set TipsBitmap = Nothing
    Set SelectIcon = Nothing
    Set SelectDown = Nothing
End Sub
