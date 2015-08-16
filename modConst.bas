Attribute VB_Name = "modConst"
Option Explicit


' 是否为调试模式
Public Const OTHELLO_DEBUG = False


' 窗体停靠样式
Public Enum eDockStyle
    dsDockNone = 0
    dsDockLeft = 1
    dsDockTop = 2
    dsDockRight = 3
    dsDockBottom = 4
End Enum

' 灯样式
Public Enum eLightStyle
    lsLightGreen = 0
    lsLightRed = 1
    lsLightYellow = 2
End Enum

' 鼠标指针样式
Public Enum ePointerStyle
    psDefault = 0
    psBlack = 1
    psWhite = 2
    psPointer = 3
    psHourglass = 4
End Enum

'Public Const WM_LBUTTONDOWN = &H201
'Public Const WM_LBUTTONUP = &H202

' 窗体消息常量
Public Const WM_USER = &H400
Public Const WM_THINKEND = WM_USER + &H2000
Public Const WM_ACTIVEWINDOW = WM_USER + &H1984

'Public Const SW_RESTORE = 9
Public Const SRCCOPY& = &HCC0020

Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1

Public Const SND_ASYNC = &H1     ' 异步播放
Public Const SND_NODEFAULT = &H2 ' 不使用缺省声音
Public Const SND_MEMORY = &H4    ' lpszSoundName 指向一个内存文件
Public Const SND_ALIAS = &H10000     ' name is a WIN.INI [sounds] entry

Public Const LVM_FIRST = &H1000
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const HDI_IMAGE = &H20
Public Const HDI_FORMAT = &H4
Public Const HDF_LEFT = 0
Public Const HDF_IMAGE = &H800
Public Const HDF_BITMAP_ON_RIGHT = &H1000
Public Const HDF_STRING = &H4000
Public Const HDM_FIRST = &H1200
Public Const HDM_SETITEM = (HDM_FIRST + 4)

Public Const LIMIT_WIDTH = 3200
Public Const LIMIT_HEIGHT = 2700

Public Const T_NONE = 0
Public Const T_BLACK = 1
Public Const T_WHITE = 2

' 版本
Public Const OTHELLO_VERSION = 1

' 主页地址
Public Const HAPPY_FAMILY_BBS = "http://bbs.ourhf.com"
Public Const HAPPY_FAMILY_BBS_MAIL = "webmaster@ourhf.com"

' 热键
Public Const HOTKEY_ID = 1000
Public Const MOD_CONTROL = &H2
Public Const KEY_HOTKEY = 192

' 总游戏人数，等于人数减 1
Public Const PLY_NUMBER = 2 - 1
' 代表自己
Public Const PLY_ME = 0
' 代表对方
Public Const PLY_YOU = PLY_ME + 1

' 最大用户列表的用户数
Public Const MAX_USER_LIST = 20

' 最大棋局项目数
Public Const MAX_TABLE_ITEM = 8
' 最大在线项目数
Public Const MAX_ONLINE_ITEM = 3

Public Const MAX_PLAY_LIST = 50

Public Const GAME_LOSE = "0"
Public Const GAME_DRAW = "1"
Public Const GAME_WIN = "2"

' 每秒的长度，标准为 18
Public Const PER_SECOND = 15

' 公共聊天区刷新间隔（秒）
Public Const PUBLIC_CHAT_RELOAD_TIME = 10

Public Const CLR_ACTIVATE = vbRed
Public Const CLR_DEACTIVATE = &H808080
Public Const CLR_TABLE_NORMAL = &H78DDFF
Public Const CLR_SELECT_MENU = &HD26900    ' RGB(0, 105, 210)

Public Const STATUS_OK = "0"
Public Const STATUS_ERROR = "1"
Public Const STATUS_NONE = "2"
Public Const STATUS_BUSY = "3"

Public Const TABLE_PUBLIC = 1
Public Const TABLE_LIMIT = 2

Public Const SEX_MAN = 2
Public Const SEX_WOMAN = 1

Public Const STY_NORMAL_MAN = 0
Public Const STY_SELECT_MAN = 1

Public Const MAX_SOUND = 12
Public Const DEFAULT_SOUND = "."
Public Const RES_DEFAULT_SOUND = 117

Public Const SOUND_LOGIN = 1
Public Const SOUND_LOGOUT = 2
Public Const SOUND_JOIN_TABLE = 3
Public Const SOUND_EXIT_TABLE = 4
Public Const SOUND_GAME_START = 5
Public Const SOUND_CHAT = 6
Public Const SOUND_DOWN_MAN = 7
Public Const SOUND_DOWN_ERROR = 8
Public Const SOUND_NOT_DOWN = 9
Public Const SOUND_GAME_WIN = 10
Public Const SOUND_GAME_LOSE = 11
Public Const SOUND_GAME_DRAW = 12


' 服务器程序路径
Public Const SERVER_APPLICATION_PATH = "/othello/"
Public Const SERVER_ACTION_GET = SERVER_APPLICATION_PATH & "security/get"
Public Const SERVER_ACTION_ONLINE_GET = SERVER_APPLICATION_PATH & "online/get"
Public Const SERVER_ACTION_REGISTER = SERVER_APPLICATION_PATH & "user/register"
Public Const SERVER_ACTION_LOGIN = SERVER_APPLICATION_PATH & "user/login"
Public Const SERVER_ACTION_LOGOUT = SERVER_APPLICATION_PATH & "user/logout"
Public Const SERVER_ACTION_USER_VIEW = SERVER_APPLICATION_PATH & "user/view"
Public Const SERVER_ACTION_USER_EDIT = SERVER_APPLICATION_PATH & "user/edit"
Public Const SERVER_ACTION_TABLE_VIEW = SERVER_APPLICATION_PATH & "table/view"
Public Const SERVER_ACTION_TABLE_GET = SERVER_APPLICATION_PATH & "table/get"
Public Const SERVER_ACTION_TABLE_AUTOJOIN = SERVER_APPLICATION_PATH & "table/autojoin"
Public Const SERVER_ACTION_TABLE_EDIT = SERVER_APPLICATION_PATH & "table/edit"
Public Const SERVER_ACTION_TABLE_CREATE = SERVER_APPLICATION_PATH & "table/create"
Public Const SERVER_ACTION_TABLE_JOIN = SERVER_APPLICATION_PATH & "table/join"
Public Const SERVER_ACTION_TABLE_EXIT = SERVER_APPLICATION_PATH & "table/quit"
Public Const SERVER_ACTION_TABLE_REMOVE = SERVER_APPLICATION_PATH & "table/remove"
Public Const SERVER_ACTION_GAME_START = SERVER_APPLICATION_PATH & "game/start"
Public Const SERVER_ACTION_GAME_CANCEL = SERVER_APPLICATION_PATH & "game/cancel"
Public Const SERVER_ACTION_GAME_OVER = SERVER_APPLICATION_PATH & "game/over"
Public Const SERVER_ACTION_CHAT_SEND = SERVER_APPLICATION_PATH & "chat/send"
Public Const SERVER_ACTION_CHAT_GET = SERVER_APPLICATION_PATH & "chat/get"


' 全局配置变量的默认值
Public Const DEFAULT_gOfflineMode = False
Public Const DEFAULT_gLevel = 5
Public Const DEFAULT_gOfflineFace = 1

Public Const DEFAULT_gDownTip = True

Public Const DEFAULT_gFaceNumber = 100
'Public Const DEFAULT_gFacePath = "/Images/"
Public Const DEFAULT_gServerUrl = "www.ourhf.com"

Public Const DEFAULT_gPlayListNumber = 0

Public Const DEFAULT_gMainWindowCenter = True
Public Const DEFAULT_gMainWindowLeft = 0
Public Const DEFAULT_gMainWindowTop = 0

Public Const DEFAULT_gUseProxy = False

' 在线用户窗口
Public Const DEFAULT_gOnlineWindowLeft = 0
Public Const DEFAULT_gOnlineWindowTop = 0
Public Const DEFAULT_gOnlineWindowWidth = LIMIT_WIDTH
Public Const DEFAULT_gOnlineWindowHeight = LIMIT_HEIGHT
Public Const DEFAULT_gOnlineWindowShow = True
Public Const DEFAULT_gOnlineWindowDockStyle = dsDockNone
Public Const DEFAULT_gOnlineWindowDockPosition = 0
Public Const DEFAULT_gOnlineSort = 0
Public Const DEFAULT_gOnlineSortKey = 0
Public Const DEFAULT_gOnlineItemWidth = 1440
Public Const DEFAULT_gOnlineAutoReload = False
Public Const DEFAULT_gOnlineAutoReloadTime = 10

Public Const DEFAULT_gTableWindowLeft = 0
Public Const DEFAULT_gTableWindowTop = 0
Public Const DEFAULT_gTableWindowWidth = LIMIT_WIDTH
Public Const DEFAULT_gTableWindowHeight = LIMIT_HEIGHT
Public Const DEFAULT_gTableWindowShow = True
Public Const DEFAULT_gTableWindowDockStyle = dsDockNone
Public Const DEFAULT_gTableWindowDockPosition = 0
Public Const DEFAULT_gTableSort = 0
Public Const DEFAULT_gTableSortKey = 0
Public Const DEFAULT_gTableItemWidth = 1440
Public Const DEFAULT_gTableAutoReload = False
Public Const DEFAULT_gTableAutoReloadTime = 10

Public Const DEFAULT_gChatWindowLeft = 0
Public Const DEFAULT_gChatWindowTop = 0
Public Const DEFAULT_gChatWindowWidth = LIMIT_WIDTH
Public Const DEFAULT_gChatWindowHeight = LIMIT_WIDTH
Public Const DEFAULT_gChatWindowShow = True
Public Const DEFAULT_gChatWindowDockStyle = dsDockNone
Public Const DEFAULT_gChatWindowDockPosition = 0

Public Const DEFAULT_gPublicChatWindowLeft = 0
Public Const DEFAULT_gPublicChatWindowTop = 0
Public Const DEFAULT_gPublicChatWindowWidth = LIMIT_WIDTH
Public Const DEFAULT_gPublicChatWindowHeight = LIMIT_WIDTH
Public Const DEFAULT_gPublicChatWindowShow = True
Public Const DEFAULT_gPublicChatWindowState = vbNormal

Public Const DEFAULT_gTableType = 0
Public Const DEFAULT_gTableTimer = 3
Public Const DEFAULT_gTableUpLevel = 1

Public Const DEFAULT_gOptionPage = 1

Public Const DEFAULT_gViewUserInfoLeft = 0
Public Const DEFAULT_gViewUserInfoTop = 0
Public Const DEFAULT_gEditUserInfoLeft = 0
Public Const DEFAULT_gEditUserInfoTop = 0
Public Const DEFAULT_gTableInfoLeft = 0
Public Const DEFAULT_gTableInfoTop = 0

Public Const CMD_Connected = "0"
Public Const CMD_DownChessMan = "1"
Public Const CMD_NoDown = "2"
Public Const CMD_GameStart = "3"
Public Const CMD_SitDown = "4"
Public Const CMD_GetSitDown = "5"
Public Const CMD_NoneSitDown = "6"
Public Const CMD_Talk = "7"
Public Const CMD_AgainStart = "8"
Public Const CMD_OK = "9"
Public Const CMD_Request = "A"
Public Const CMD_ReRequest = "B"
Public Const CMD_Agree = "C"
Public Const CMD_Disagree = "D"
Public Const CMD_RequestJoin = "E"
Public Const CMD_RequestCancelGame = "F"
Public Const CMD_RequestExitTable = "G"
Public Const CMD_RequestExitGame = "H"
Public Const CMD_RequestLogout = "I"
Public Const CMD_RequestDrawGame = "J"
Public Const CMD_CancelGame = "M"
Public Const CMD_GameOver = "N"
Public Const CMD_TimeOver = "O"
Public Const CMD_GameReadyStart = "P"
Public Const CMD_InfoChanged = "Q"
Public Const CMD_TableChanged = "R"
