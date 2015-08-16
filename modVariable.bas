Attribute VB_Name = "modVariable"
Option Explicit

Public Type tagPointS
    X As Single
    Y As Single
End Type

Public Type tagPointI
    Col As Integer
    Row As Integer
End Type

Public Type tagWindowInfo
    Show As Boolean
    Center As Boolean
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    DockStyle As eDockStyle
    DockPosition As Single
End Type

Public Type tagUserInfo
    UserName As String
    Password As String
    UserClass As String
    Email As String
    Face As Long
    Name As String
    Sex As Long
    Age As Long
    Country As String
    State As String
    City As String
    Win As Long
    Lose As Long
    Draw As Long
    GameTimes As Long
    Score As Long
End Type

Public Type tagTableInfo
    TableName As String
    Creator As String       ' 用户名
    CreatorName As String   ' 昵称
    Visitor As String
    VisitorName As String
    TableType As Long
    Timer As Long
    UpLevel As Boolean
    LastTime As Date
    ip As String
    LANIP As String
    Port As Long
End Type

Public SoundBuffer() As Byte

' 应用程序路径
Public gstrAppPath As String

' 服务器端安全检查
Public gstrSecurity1 As String
Public gstrSecurity2 As String

Public gstrLocalPassword As String  ' *本地口令字
Public gstrAzDGPrivateKey As String ' *与服务器通信的密钥

Public TablePos As tagPointS     ' 棋盘图片显示坐标
Public Current As tagPointI     ' 当前棋盘位置
Public LastDown As tagPointI
Public LastMan As Byte

Public gsngCaptionHeight As Single
Public gsngBorderX As Single
Public gsngBorderY As Single

''' 图像对象
Public BlackMan As StdPicture
Public WhiteMan As StdPicture
Public SelBlackMan As StdPicture
Public SelWhiteMan As StdPicture
Public DefaultCursor As StdPicture
Public HandCursor As StdPicture
Public BlackCursor As StdPicture
Public WhiteCursor As StdPicture
Public HandUpCursor As StdPicture
Public HandDownCursor As StdPicture
Public TipsBitmap As StdPicture
Public ChessBoard As StdPicture
Public GameTitle As StdPicture
Public NoFocusTitle As StdPicture
Public SelectIcon As StdPicture
Public SelectDown As StdPicture
Public objLightOn As StdPicture
Public objLightOff As StdPicture
Public objLightYellow As StdPicture
Public objSoundPlay As StdPicture
Public objSoundStop As StdPicture
Public objOpenFile As StdPicture


' 全局配置变量，用于暂时保存配置数据
Public gblnSave_OfflineMode As Boolean
Public glngSave_Level As Long
Public glngSave_OfflineFace As Long

Public gstrSave_UserList(MAX_USER_LIST) As String
Public gblnSave_AutoLogin As Boolean
Public gblnSave_SavePassword As Boolean
Public gstrSave_UserName As String
Public gstrSave_Password As String

Public gwifSave_MainWindow As tagWindowInfo
Public gwifSave_OnlineWindow As tagWindowInfo
Public gwifSave_TableWindow As tagWindowInfo
Public gwifSave_ChatWindow As tagWindowInfo
Public gwifSave_PublicChatWindow As tagWindowInfo

Public gsngSave_TableItemWidth(MAX_TABLE_ITEM) As Single
Public glngSave_TableSort As Long
Public glngSave_TableSortKey As Long
Public glngSave_TableAutoReloadTime As Long
Public gblnSave_TableAutoReload As Boolean

Public gsngSave_OnlineItemWidth(MAX_ONLINE_ITEM) As Single
Public glngSave_OnlineSort As Long
Public glngSave_OnlineSortKey As Long
Public glngSave_OnlineAutoReloadTime As Long
Public gblnSave_OnlineAutoReload As Boolean

Public glngSave_PublicChatWindowState As Long

Public gstrSave_FacePath As String
Public glngSave_FaceNumber As Long

Public gblnSave_DownTip As Boolean

Public gstrSave_PlayListPath As String
Public glngSave_PlayListNumber As Long
Public glngSave_PlayListPosition As Long
Public gstrSave_PlayListName(MAX_PLAY_LIST) As String

Public gstrSave_SoundPath As String
Public gstrSave_SoundName(MAX_SOUND) As String
Public gstrSave_SoundValue(MAX_SOUND) As String

Public gstrSave_ServerUrl As String
Public glngSave_GamePort As Long

Public gblnSave_UseProxy As Boolean

Public gstrSave_HttpProxyIP As String
Public glngSave_HttpProxyPort As Long

Public gstrSave_Socks5ProxyIP As String
Public glngSave_Socks5ProxyPort As Long
Public gstrSave_Socks5Username As String
Public gstrSave_Socks5Password As String

' 保存创建棋局信息
Public glngSave_TableType As Long
Public glngSave_TableTimer As Long
Public glngSave_TableUpLevel As Long

' 保存最后使用的选项页
Public glngSave_OptionPage As Long

' 保存其它窗体的位置
Public gptsSave_ViewUserInfo As tagPointS
Public gptsSave_EditUserInfo As tagPointS
Public gptsSave_TableInfo As tagPointS

' 用户信息变量
Public gMyUserInfo As tagUserInfo
Public gYourUserInfo As tagUserInfo

' 棋局信息变量
Public gMainTableInfo As tagTableInfo

Public gblnSndCard As Boolean

Public gBackgroundMusic As Mmedia
Public gSoundEffects As Mmedia

Public WindowWidth As Single
Public WindowHeight As Single

' 连接、登陆等全局 Bool 变量
Public gblnConnect As Boolean
Public gblnLogin As Boolean
Public gblnCreator As Boolean
Public gblnGameStart As Boolean

Public glngRetryTimes As Long

Public gstrLocalIP As String
Public gstrIP As String

'Public glngReSendTimes As Long

Public gblnMenuDisplay As Boolean

' 单机模式
Public gblnOfflineMode As Boolean
