Attribute VB_Name = "mCoolMenu"
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
''  mCoolMenu Module v1.3
''
''  Copyright Olivier Martin 2000
''
''  martin.olivier@bigfoot.com
''
''  Code based on Paul Dilascia's work from the
''  Microsoft System Journal January 1998
''  Visit Paul's page at www.dilascia.com
''
''  This module allows an application to show
''  icons in menus just like the VB IDE and
''  MS Office applications.  The link between
''  the menus and an ImageList is the image tag.
''  The test forms show all the possibilities.
''
''  People who contributed with suggestions :
''  Pietro Cecchi : help callback
''  Kayl Magnus : separator font; sub menu help bug;
''                main bar icon bug
''  Nabil AbuSharane : Right to left property
''  Kenneth (aka Maverick) : MDI window lists bug
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwndParent As Long, pt As POINTAPI) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal ByPosition As Long, ByRef lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hDC As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hinst As Long, ByVal lpszName As Any, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function MenuItemFromPoint Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal ptScreen As Double) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
'Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
'Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long


'Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal diIgnore As Long) As Long
'Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
'Private Declare Function ImageList_GetImageInfo Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, IMAGEINFO As Any) As Long

'Used by CreateBrushIndirect
Private Type LOGBRUSH
  lbStyle As Long
  lbColor As Long
  lbHatch As Long
End Type

Private Const SC_MOVE = &HF010&

'LOGBRUSH constants
Private Const BS_SOLID = 0
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_HATCHED = 2
Private Const BS_PATTERN = 3
Private Const BS_INDEXED = 4
Private Const BS_DIBPATTERN = 5
Private Const BS_DIBPATTERNPT = 6
Private Const BS_PATTERN8X8 = 7
Private Const BS_DIBPATTERN8X8 = 8

'LoadImage constants
Private Const IMAGE_BITMAP = 0&
Private Const IMAGE_ICON = 1&
Private Const IMAGE_CURSOR = 2&

'LoadImage constants
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000

'LoadImage constants
Private Const OBM_LFARROWI = 32734
Private Const OBM_RGARROWI = 32735
Private Const OBM_DNARROWI = 32736
Private Const OBM_UPARROWI = 32737
Private Const OBM_COMBO = 32738
Private Const OBM_MNARROW = 32739
Private Const OBM_LFARROWD = 32740
Private Const OBM_RGARROWD = 32741
Private Const OBM_DNARROWD = 32742
Private Const OBM_UPARROWD = 32743
Private Const OBM_RESTORED = 32744
Private Const OBM_ZOOMD = 32745
Private Const OBM_REDUCED = 32746
Private Const OBM_RESTORE = 32747
Private Const OBM_ZOOM = 32748
Private Const OBM_REDUCE = 32749
Private Const OBM_LFARROW = 32750
Private Const OBM_RGARROW = 32751
Private Const OBM_DNARROW = 32752
Private Const OBM_UPARROW = 32753
Private Const OBM_CLOSE = 32754
Private Const OBM_OLD_RESTORE = 32755
Private Const OBM_OLD_ZOOM = 32756
Private Const OBM_OLD_REDUCE = 32757
Private Const OBM_BTNCORNERS = 32758
Private Const OBM_CHECKBOXES = 32759
Private Const OBM_CHECK = 32760
Private Const OBM_BTSIZE = 32761
Private Const OBM_OLD_LFARROW = 32762
Private Const OBM_OLD_RGARROW = 32763
Private Const OBM_OLD_DNARROW = 32764
Private Const OBM_OLD_UPARROW = 32765
Private Const OBM_SIZE = 32766
Private Const OBM_OLD_CLOSE = 32767

' GetSystemMetrics() constants
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CYCAPTION = 4
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CXDLGFRAME = 7
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYVTHUMB = 9
Private Const SM_CXHTHUMB = 10
Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXCURSOR = 13
Private Const SM_CYCURSOR = 14
Private Const SM_CYMENU = 15
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYKANJIWINDOW = 18
Private Const SM_MOUSEPRESENT = 19
Private Const SM_CYVSCROLL = 20
Private Const SM_CXHSCROLL = 21
Private Const SM_DEBUG = 22
Private Const SM_SWAPBUTTON = 23
Private Const SM_RESERVED1 = 24
Private Const SM_RESERVED2 = 25
Private Const SM_RESERVED3 = 26
Private Const SM_RESERVED4 = 27
Private Const SM_CXMIN = 28
Private Const SM_CYMIN = 29
Private Const SM_CXSIZE = 30
Private Const SM_CYSIZE = 31
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
Private Const SM_CXMINTRACK = 34
Private Const SM_CYMINTRACK = 35
Private Const SM_CXDOUBLECLK = 36
Private Const SM_CYDOUBLECLK = 37
Private Const SM_CXICONSPACING = 38
Private Const SM_CYICONSPACING = 39
Private Const SM_MENUDROPALIGNMENT = 40
Private Const SM_PENWINDOWS = 41
Private Const SM_DBCSENABLED = 42
Private Const SM_CMOUSEBUTTONS = 43

Private Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Private Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Private Const SM_CXSIZEFRAME = SM_CXFRAME
Private Const SM_CYSIZEFRAME = SM_CYFRAME

Private Const SM_SECURE = 44
Private Const SM_CXEDGE = 45
Private Const SM_CYEDGE = 46
Private Const SM_CXMINSPACING = 47
Private Const SM_CYMINSPACING = 48
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
Private Const SM_CYSMCAPTION = 51
Private Const SM_CXSMSIZE = 52
Private Const SM_CYSMSIZE = 53
Private Const SM_CXMENUSIZE = 54
Private Const SM_CYMENUSIZE = 55
Private Const SM_ARRANGE = 56
Private Const SM_CXMINIMIZED = 57
Private Const SM_CYMINIMIZED = 58
Private Const SM_CXMAXTRACK = 59
Private Const SM_CYMAXTRACK = 60
Private Const SM_CXMAXIMIZED = 61
Private Const SM_CYMAXIMIZED = 62
Private Const SM_NETWORK = 63
Private Const SM_CLEANBOOT = 67
Private Const SM_CXDRAG = 68
Private Const SM_CYDRAG = 69
Private Const SM_SHOWSOUNDS = 70
Private Const SM_CXMENUCHECK = 71  'Use instead of GetMenuCheckMarkDimensions()!
Private Const SM_CYMENUCHECK = 72
Private Const SM_SLOWMACHINE = 73
Private Const SM_MIDEASTENABLED = 74

' Return values for ExcludeClipRect
Private Const NULLREGION = 1
Private Const SIMPLEREGION = 2
Private Const COMPLEXREGION = 3

' Hatch constants for CreateHatchBrush
Private Const HS_HORIZONTAL = 0
Private Const HS_VERTICAL = 1
Private Const HS_FDIAGONAL = 2
Private Const HS_BDIAGONAL = 3
Private Const HS_CROSS = 4
Private Const HS_DIAGCROSS = 5
Private Const HS_FDIAGONAL1 = 6
Private Const HS_BDIAGONAL1 = 7
Private Const HS_SOLID = 8
Private Const HS_DENSE1 = 9
Private Const HS_DENSE2 = 10
Private Const HS_DENSE3 = 11
Private Const HS_DENSE4 = 12
Private Const HS_DENSE5 = 13
Private Const HS_DENSE6 = 14
Private Const HS_DENSE7 = 15
Private Const HS_DENSE8 = 16
Private Const HS_NOSHADE = 17
Private Const HS_HALFTONE = 18
Private Const HS_SOLIDCLR = 19
Private Const HS_DITHEREDCLR = 20
Private Const HS_SOLIDTEXTCLR = 21
Private Const HS_DITHEREDTEXTCLR = 22
Private Const HS_SOLIDBKCLR = 23
Private Const HS_DITHEREDBKCLR = 24
Private Const HS_API_MAX = 25

' Image List draw constants
Private Const ILD_NORMAL = &H0
Private Const ILD_TRANSPARENT = &H1
Private Const ILD_MASK = &H10
Private Const ILD_IMAGE = &H20

'' Image type for DrawState
Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4

' ' State type for DrawState
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000

' SysColor constants *some could be wrong in the code*
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_ADJ_MAX = 100
Private Const COLOR_ADJ_MIN = -100
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22
Private Const COLOR_MENU = 4
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_WINDOWTEXT = 8

' Owner draw actions
Private Const ODA_DRAWENTIRE = &H1
Private Const ODA_SELECT = &H2
Private Const ODA_FOCUS = &H4

' Owner draw state
Private Const ODS_SELECTED = &H1
Private Const ODS_GRAYED = &H2
Private Const ODS_DISABLED = &H4
Private Const ODS_CHECKED = &H8
Private Const ODS_FOCUS = &H10
Private Const ODS_DEFAULT = &H20
Private Const ODS_COMBOBOXEDIT = &H1000

'required for font API functions
Private Const LF_FACESIZE = 32
Private Const SYMBOL_CHARSET = 2

Private Const LOGPIXELSY = 90

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700

Private Const DEFAULT_CHARSET = 1
Private Const GB2312_CHARSET = 134

'for subclassing
Private Const GWL_WNDPROC = -4

'for BitBlt
Private Const NOTSRCERASE = &H1100A6
Private Const NOTSRCCOPY = &H330008
Private Const SRCERASE = &H440328
Private Const SRCINVERT = &H660046
Private Const SRCAND = &H8800C6
Private Const MERGEPAINT = &HBB0226
Private Const MERGECOPY = &HC000CA
Private Const SRCCOPY = &HCC0020
Private Const SRCPAINT = &HEE0086
Private Const PATPAINT = &HFB0A09

Private Const BLACKNESS = &H42
Private Const DSTINVERT = &H550009
Private Const PATINVERT = &H5A0049
Private Const PATCOPY = &HF00021
Private Const WHITENESS = &HFF0062

Private Const MAGICROP = &HB8074A

' Background Modes
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2

' DrawText constants
Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_INTERNAL = &H1000
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

Private Const ODT_MENU = 1

Private Const MNC_IGNORE = 0
Private Const MNC_CLOSE = 1
Private Const MNC_EXECUTE = 2
Private Const MNC_SELECT = 3

' Menu Item Info Mask constants
Private Const MIIM_STATE = &H1&
Private Const MIIM_ID = &H2
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_CHECKMARKS = &H8
Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20
Private Const MIIM_STRING = &H40
Private Const MIIM_BITMAP = &H80
Private Const MIIM_FTYPE = &H100

Private Const MF_INSERT = &H0
Private Const MF_CHANGE = &H80
Private Const MF_APPEND = &H100
Private Const MF_DELETE = &H200
Private Const MF_REMOVE = &H1000

Private Const MF_BYCOMMAND = &H0
Private Const MF_BYPOSITION = &H400

Private Const MF_SEPARATOR = &H800

Private Const MF_ENABLED = &H0
Private Const MF_GRAYED = &H1
Private Const MF_DISABLED = &H2

Private Const MF_UNCHECKED = &H0
Private Const MF_CHECKED = &H8
Private Const MF_USECHECKBITMAPS = &H200

Private Const MF_STRING = &H0
Private Const MF_BITMAP = &H4
Private Const MF_OWNERDRAW = &H100

Private Const MF_POPUP = &H10
Private Const MF_MENUBARBREAK = &H20
Private Const MF_MENUBREAK = &H40

Private Const MF_UNHILITE = &H0
Private Const MF_HILITE = &H80

Private Const MF_DEFAULT = &H1000
Private Const MF_SYSMENU = &H2000
Private Const MF_HELP = &H4000
Private Const MF_RIGHTJUSTIFY = &H4000

Private Const MF_MOUSESELECT = &H8000
Private Const MF_END = &H80                     ' ' Obsolete -- only used by old RES files

Private Const MFT_STRING = MF_STRING
Private Const MFT_BITMAP = MF_BITMAP
Private Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Private Const MFT_MENUBREAK = MF_MENUBREAK
Private Const MFT_OWNERDRAW = MF_OWNERDRAW
Private Const MFT_RADIOCHECK = &H200
Private Const MFT_SEPARATOR = MF_SEPARATOR
Private Const MFT_RIGHTORDER = &H2000
Private Const MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY

Private Const MFS_GRAYED = &H3
Private Const MFS_DISABLED = MFS_GRAYED
Private Const MFS_CHECKED = MF_CHECKED
Private Const MFS_HILITE = MF_HILITE
Private Const MFS_ENABLED = MF_ENABLED
Private Const MFS_UNCHECKED = MF_UNCHECKED
Private Const MFS_UNHILITE = MF_UNHILITE
Private Const MFS_DEFAULT = MF_DEFAULT
'Private Const MFS_MASK = &H108B
'Private Const MFS_HOTTRACKDRAWN = &H10000000
'Private Const MFS_CACHEDBMP = &H20000000
'Private Const MFS_BOTTOMGAPDROP = &H40000000
'Private Const MFS_TOPGAPDROP = &H80000000
'Private Const MFS_GAPDROP = &HC0000000

' Menu item drawing constants
Private Const CXGAP = 0           ' num pixels between button and text
Private Const CXTEXTMARGIN = 2    ' num pixels after hilite to start text
Private Const CXBUTTONMARGIN = 2  ' num pixels wider button is than bitmap
Private Const CYBUTTONMARGIN = 2  ' ditto for height

' 3D border styles
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

' Border flags
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BF_DIAGONAL = &H10

' For diagonal lines, the BF_RECT flags specify the end point of the
' vector bounded by the rectangle parameter.
Private Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

Private Const BF_MIDDLE = &H800         ' Fill in the middle
Private Const BF_SOFT = &H1000          ' For softer buttons
Private Const BF_ADJUST = &H2000        ' Calculate the space left over
Private Const BF_FLAT = &H4000          ' For flat rather than 3D borders
Private Const BF_MONO = &H8000          ' For monochrome borders

' Window messages
Private Const WM_HOTKEY = &H312
Private Const WM_ACTIVATE = &H6
'Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_COMMAND = &H111
Private Const WM_CLOSE = &H10
Private Const WM_DRAWITEM = &H2B
Private Const WM_GETFONT = &H31
Private Const WM_MEASUREITEM = &H2C
Private Const WM_NCHITTEST = &H84
Private Const WM_MENUSELECT = &H11F
Private Const WM_MENUCHAR = &H120
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_INITMENU = &H116
Private Const WM_WININICHANGE = &H1A
Private Const WM_SETCURSOR = &H20
Private Const WM_SETTINGCHANGE = WM_WININICHANGE
Private Const WM_CANCELMODE = &H1F
Private Const WM_MDISETMENU = &H230
Private Const WM_MDIREFRESHMENU = &H234
Private Const WM_MOVE = &H3

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion(1 To 128) As Byte      '  Maintenance string for PSS usage
End Type

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0

Private Type MEASUREITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemWidth As Long
  itemHeight As Long
  ItemData As Long
End Type

Private Type DRAWITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemAction As Long
  itemState As Long
  hwndItem As Long
  hDC As Long
  rcItem As RECT
  ItemData As Long
End Type

Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As Long
  cch As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte '0=false; 255=true
  lfUnderline As Byte '0=f; 255=t
  lfStrikeOut As Byte '0=f; 255=t
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
End Type

Private Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type

Private Type IMAGEINFO
  hbmImage As Long
  hbmMask As Long
  Unused1 As Long
  Unused2 As Long
  rcImage As RECT
End Type

' 我加的，菜单边宽度
Private Const SIDE_WIDTH = 21
' 我加的，菜单边图像高度
Private Const MENUSIDE_HEIGHT = 510
Private Const MENUSIDE_TOP = -98

' Bitmap objects for quick redrawing
Private m_bmpChecked As Long, m_bmpRadioed As Long

Private m_MarlettFont As Long 'Font used to draw Window items
Private m_iBitmapWidth As Integer 'width of menu bitmaps (square)
'Private m_SideBitmapWidth As Long
' 我加的，缩进的菜单数量
'Private m_SideMenuNumber As Long
' 我加的，菜单左面的图像
'Private m_SideMenuBitmap As StdPicture

Private pmds As CMyItemDatas 'the collection of pmd
Private WndCol As Collection 'the collection of WndCoolMenu

Private Sub ConvertMenu(hWnd As Long, hMenu As Long, nIndex As Long, bSysMenu As Boolean, bShowButtons As Boolean, Optional Permanent As Boolean = False)
'Based on Paul DiLascia's
'Converts submenus into OwnerDraw

  On Error GoTo ErrorHandle
  
  Dim i As Long, k As Byte
  Dim Info As MENUITEMINFO

  Dim dwItemData As Long
  Dim pmd As CMyItemData
          
  Dim Text As String
  Dim ByteBuffer() As Byte
  
  ' Get the number of menu items
  Dim nItem As Long
  nItem = GetMenuItemCount(hMenu&)

  'Debug.Print nItem
  'On GetmenuItemCount error, exit
  If nItem = -1 Then Exit Sub
  'Debug.Print nItem

  For i& = 0 To nItem& - 1
  
    'Create and initialize a byte array
    ReDim ByteBuffer(0 To 200) As Byte
    For k = 0 To 200
      ByteBuffer(k) = 0
    Next k
    
    'information to retreive with GetMenuItemInfo
    Info.fMask = MIIM_DATA Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU
    
    Info.dwTypeData = VarPtr(ByteBuffer(0))
    Info.cch = UBound(ByteBuffer)
    Info.cbSize = LenB(Info) 'size in byte of structure
    
    Call GetMenuItemInfo(hMenu&, i&, MF_BYPOSITION, Info)

    dwItemData& = Info.dwItemData
  
    If bSysMenu And (Info.wID >= &HF000) Then _
      GoTo NextGoto 'not touching

    Info.fMask = 0& 'reset mask value
    
    If bShowButtons Then
      'showing buttons. if not, no OwnerDraw needed
      
      If Not CBool(Info.fType And MFT_OWNERDRAW) Then
        
        'Convert if not OWNERDRAW
        Info.fType = Info.fType Or MFT_OWNERDRAW
        Info.fMask = Info.fMask Or MIIM_TYPE
        
        If dwItemData& = 0& Then
          ' no reference; create one
          ' Paul used a pointer in original code
          '   but this works fine
          
          Info.dwItemData = CLng(pmds.Count + 1)
          Info.fMask = Info.fMask Or MIIM_DATA
          
          Set pmd = pmds.Add(CStr(Info.dwItemData))
        
          Text$ = Left(StrConv(ByteBuffer, vbUnicode), Info.cch)
          pmd.sMenuText = Text$

          Dim iBreakPos As Integer
          iBreakPos% = InStr(Text$, "|")

          If iBreakPos% Then
            
            Dim iBreak2Pos As Integer
            iBreak2Pos% = InStr(Right(Text$, Len(Text$) - iBreakPos%), "|")
            
            Dim HelpText As String
            Dim iHelpLen As Integer
            HelpText$ = Mid(Text$, iBreakPos% + 1, iBreak2Pos% - 1)
            iHelpLen% = Len(HelpText$)
            
            pmd.sMenuHelp = HelpText$
            pmd.sMenuText = Right(Text$, Len(Text$) - (iBreakPos% + iBreak2Pos%))

          Else
            pmd.sMenuText = Text$
          End If

          Dim cFirstChar As String * 1
          cFirstChar$ = Left(Text$, 1)
          
          If cFirstChar$ = "-" Then
            Info.fType = Info.fType Or MF_SEPARATOR
            If pmd.sMenuHelp = "" Then _
              pmd.sMenuText = Right(Text$, Len(Text$) - 1)
          End If
          
          pmd.bAsMark = (cFirstChar$ = "*") Or (cFirstChar$ = "#")
          If pmd.bAsMark Then
            pmd.bAsCheck = (cFirstChar$ = "#")
          
            If pmd.sMenuHelp = "" Then _
              pmd.sMenuText = Right(Text$, Len(Text$) - 1)
          Else
          
            'The "Window List" bug was found by Kenneth. Thank you
            If Mid(Text$, 3, 1) = " " Then
              If InStr("123456789", Mid(Text$, 2, 1)) > 0 And cFirstChar$ = "&" Then
                pmd.bAsMark = True
                pmd.bAsCheck = False
              End If
            End If
            
          End If
          
          'get image index
          If Permanent Then
            'main bar icon bug identified at conception
            'but since Kayl Magnus had it, I fixed it
            pmd.iButton = -1
          Else
            pmd.iButton = GetButtonIndex(hWnd&, pmd.sMenuText)
          End If
          
          pmd.fType = Info.fType
          
          pmd.bTrueSub = (Info.hSubMenu <> 0&) And (Not Permanent)
          
        Else
          'A reference exists
          Set pmd = pmds(CStr(dwItemData&))
        End If
        
        pmd.bMainMenu = Permanent ' it's a main menu
        
        
      End If 'Changed to OWNERDRAW
    
      If Not Permanent Then _
        Call WndCol(CStr(hWnd&)).AddMenuHead(hMenu)
    
    Else
      'No buttons
      
      If Info.fType And MFT_OWNERDRAW Then
        
        Info.fType = Info.fType And (Not MFT_OWNERDRAW)
        Info.fMask = Info.fMask Or MIIM_TYPE
        
        Set pmd = pmds(CStr(dwItemData&))
        
        Dim cLeadChar As String
        cLeadChar$ = ""
        If pmd.bAsMark Then
          If pmd.bAsCheck Then
            cLeadChar = "#"
          Else
            cLeadChar = "*"
          End If
        End If
        
        If pmd.fType And MFT_SEPARATOR Then
          cLeadChar$ = "-"
          Info.fType = Info.fType And (Not MFT_SEPARATOR)
        End If
        
        If pmd.sMenuHelp <> "" Then _
          pmd.sMenuHelp = "|" + pmd.sMenuHelp + "|"
          
        Text$ = cLeadChar$ + pmd.sMenuHelp + pmd.sMenuText
        
        Info.cch = BSTRtoLPSTR(Text$, ByteBuffer, Info.dwTypeData)
        
      End If
      
      If dwItemData <> 0& Then
        'remove reference
        Info.dwItemData = 0&
        Info.fMask = Info.fMask Or MIIM_DATA
        pmds.Remove CStr(dwItemData&) 'by key
      End If
      
    End If
    
    ' make changes if any
    If Info.fMask Then _
      Call SetMenuItemInfo(hMenu&, i&, MF_BYPOSITION, Info)
    
NextGoto:
  Next i&

  Exit Sub

ErrorHandle:
  Debug.Print Err.Number; Err.Description; " ConvertMenu"
  Err.Clear
End Sub

Private Sub OnInitMenuPopup(hWnd As Long, hMenu As Long, nIndex As Long, bSysMenu As Boolean)
'Based on Paul DiLascia's
'Bridge to ConvertMenu(ON) when in menu loop
  
  WndCol(CStr(hWnd&)).MainPopedIndex = -2 ' Deselect main menu item

  Call ConvertMenu(hWnd&, hMenu&, nIndex&, bSysMenu, True, False)
End Sub

Private Function OnMenuChar(nChar As Long, nFlags As Long, hMenu As Long) As Long
'Based on Paul DiLascia's
'Local character accelerator:
' the underlined character by the ampersand ("&")
  
  Dim i As Long
  Dim nItem As Long
  Dim dwItemData As Long
  
  Dim Info As MENUITEMINFO
  
  Dim Count As Integer: Count% = 0
  Dim iCurrent As Integer
  
  ReDim ItemMatch(0 To 0) As Integer
  
  nItem& = GetMenuItemCount(hMenu&)
  
  For i& = 0 To nItem& - 1
    
    Info.cbSize = LenB(Info)
    Info.fMask = MIIM_DATA Or MIIM_TYPE Or MIIM_STATE
    
    Call GetMenuItemInfo(hMenu&, i&, MF_BYPOSITION, Info)
    
    dwItemData& = Info.dwItemData
    
    If (Info.fType And MFT_OWNERDRAW) And dwItemData& <> 0 Then
      Dim Text As String
      Dim iAmpersand As Integer
      
      Text$ = pmds(CStr(dwItemData&)).sMenuText
      iAmpersand% = InStr(Text$, "&")
      
      If (iAmpersand% > 0) And (UCase(Chr(nChar&)) _
          = UCase(Mid(Text$, iAmpersand% + 1, 1))) Then
        
        If Count > UBound(ItemMatch) Then _
          ReDim Preserve ItemMatch(0 To Count%)
        
        'Build an array of matching elements
        ItemMatch(Count%) = i&
        Count% = Count% + 1
      
      End If
        
    End If
    
    'Identify the selected menu item
    If Info.fState And MFS_HILITE Then _
      iCurrent% = i&
  
  Next i&
  Count% = Count% - 1 'back
  
  If Count% = -1 Then 'no match
    OnMenuChar = 0&
    Exit Function
  End If
  
  Dim bMainMenu As Boolean
  bMainMenu = pmds(CStr(dwItemData&)).bMainMenu
  
  If Count% = 0 Then '1 match
      OnMenuChar = MakeLong(ItemMatch(0), MNC_EXECUTE)
    Exit Function
  End If
    
  Dim iSelect As Integer 'multiple matches
  For i& = 0 To Count%
    If ItemMatch(i&) = iCurrent% Then
      iSelect% = i&
      Exit For
    End If
  Next i&
  
  OnMenuChar = MakeLong(ItemMatch(iSelect%), MNC_SELECT)
End Function

Private Sub DrawMenuText(hWnd As Long, hDC As Long, rc As RECT, Text As String, Color As Long, Optional bLeftAlign As Boolean = True, Optional bRightToLeft As Boolean = False)
'Based on Paul DiLascia'
'Draw menu text; added main menu text and RightToLeft property

  On Error GoTo ErrHandler

  Dim LeftStr As String
  Dim RightStr As String
  Dim iTabPos As Integer
  
  Dim OldFont As Long

  LeftStr$ = Text$
  iTabPos = InStr(LeftStr$, Chr(9)) ' 9 = tab
  
  If iTabPos > 0 Then 'for the global accelerator (Ctrl+?)
    RightStr$ = Right$(LeftStr$, Len(LeftStr$) - iTabPos)
    LeftStr$ = Left$(LeftStr$, iTabPos - 1)
  End If

  Call SetTextColor(hDC&, Color&)
 
  OldFont& = SelectObject(hDC&, GetMenuFont(hWnd&))
  'Call DrawText(hdc&, LeftStr$, Len(LeftStr$), rc, IIf(bLeftAlign, IIf(bRightToLeft, DT_RIGHT, DT_LEFT), DT_CENTER) Or DT_VCENTER Or DT_SINGLELINE)
  'Call FillRectEx(hDC&, rc, vbWhite)
  Call DrawText(hDC&, LeftStr$, -1, rc, IIf(bLeftAlign, IIf(bRightToLeft, DT_RIGHT, DT_LEFT), DT_CENTER) Or DT_VCENTER Or DT_SINGLELINE)

  If iTabPos > 0 Then
    'Call DrawText(hdc&, RightStr$, Len(RightStr$), rc, IIf(bRightToLeft, DT_LEFT, DT_RIGHT) Or DT_VCENTER Or DT_SINGLELINE)
    Call DrawText(hDC&, RightStr$, -1, rc, IIf(bRightToLeft, DT_LEFT, DT_RIGHT) Or DT_VCENTER Or DT_SINGLELINE)
  End If

  Call SelectObject(hDC&, OldFont&)

  Exit Sub

ErrHandler:
  Debug.Print Err.Number; Err.Description; " DrawMenuText"
  Err.Clear
End Sub

Private Function OnDrawItem(hWnd As Long, ByRef dsPtr As Long, Optional bOverMain As Boolean = False) As Boolean
'Based on Paul DiLascia's
'Draw menu items

  On Error GoTo ErrHandler
  
  Dim lpds As DRAWITEMSTRUCT
  
  Call CopyMemory(lpds, ByVal dsPtr&, Len(lpds))
  
  Dim rt As RECT
  Dim rtItem As RECT
  Dim rtText As RECT
  Dim rtButn As RECT
  Dim rtIcon As RECT
  Dim rtHighlight As RECT
   
  Dim WndObj As WndCoolMenu
  Set WndObj = WndCol(CStr(hWnd&))
  
  Dim dwItemData As Long
  dwItemData& = lpds.ItemData
  
  If (dwItemData& = 0&) Or (lpds.CtlType <> ODT_MENU) Or (dwItemData& > pmds.Count) Then
    OnDrawItem = False
    Exit Function
  End If
  
  Dim pmd As CMyItemData
  Set pmd = pmds(CStr(dwItemData&))
  
  Dim hDC As Long
  hDC& = lpds.hDC
  LSet rtItem = lpds.rcItem

  'Debug.Print "Data:"; pmds.Count
  ' 我们可以先简单的这样写：
  If dwItemData <= WndObj.MenuSideNumber And Not (WndObj.MenuSideImage Is Nothing) Then
    rtItem.Left = rtItem.Left + WndObj.MenuSideWidth
  End If

  'rtItem.Bottom = rtItem.Top + 19
  Call FillRectEx(hDC&, rtItem, vbWhite)

  If pmd.fType And MFT_SEPARATOR Then
  
    LSet rt = rtItem
    LSet rtText = rtItem
    
    Dim SepMargin As Integer
    SepMargin = 5
    rt.Left = rt.Left + SepMargin
    rt.Right = rt.Right - SepMargin

    rt.Top = rt.Top + ((rt.Bottom - rt.Top) \ 2) - 1
    'Debug.Print rt.Bottom - rt.Top,
    Call DrawEdge(hDC&, rt, EDGE_ETCHED, BF_TOP)
    
    If pmd.sMenuText <> "" Then
      Dim OldFont As Long
      OldFont& = SelectObject(hDC&, GetMenuFontSep(hWnd&))
      
      rtText = OffsetRect(rtText, 1, 1)
      Call SetBkMode(hDC&, OPAQUE)
      Call SetTextColor(hDC&, GetSysColor(COLOR_BTNLIGHT))
      'Call SetTextColor(hdc&, vbBlue)
      Call DrawText(hDC&, " " + pmd.sMenuText + " ", 2 + Len(pmd.sMenuText), rtText, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
      
      rtText = OffsetRect(rtText, -1, -1)
      Call SetBkMode(hDC&, TRANSPARENT)
      Call SetTextColor(hDC&, GetSysColor(COLOR_BTNSHADOW))
      Call DrawText(hDC&, " " + pmd.sMenuText + " ", 2 + Len(pmd.sMenuText), rtText, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)

      Call SelectObject(hDC&, OldFont&)
    End If
  
  ElseIf Left(pmd.sMenuText, 1) = "!" Then
  
    ' !["s"("tile", "bottom", "top", "center")];["w"width];["n"name]
  
    Dim SideBitmap As Long
    'Dim sBitmap As String
    'sBitmap = "c:\win95\bureau\smart.bmp" + Chr(0)
    'sBitmap = "e:\hand.bmp" + Chr(0)
    'SideBitmap& = LoadImage(0&, sBitmap, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE)
    If Not (WndObj.MenuSideImage Is Nothing) Then
        'Dim lngMenuHeight As Long
        'Dim i As Long

        'lngMenuHeight = 0
'        For i = 1 To 21
'            'If pmds.Item(i).fType And MFT_SEPARATOR Then
'
'            'End If
'            lngMenuHeight = lngMenuHeight + GetSystemMetrics(SM_CYMENUSIZE)
'        Next i
        'Debug.Print "Height: "; lngMenuHeight
        'Debug.Print lngMenuHeight + (pmds.Count - 1) * GetSystemMetrics(SM_CYMENUSIZE)
        'lngMenuHeight = lngMenuHeight + (pmds.Count - 1) * GetSystemMetrics(SM_CYMENUSIZE) + 50

        SideBitmap = WndObj.MenuSideImage.Handle
        Dim hmemDC As Long
        hmemDC& = CreateCompatibleDC(hDC&)
    
        Dim hOldBitmap As Long
        hOldBitmap = SelectObject(hmemDC&, SideBitmap&)

        'Call BitBlt(hdc&, 0, 0, 32, 300, hmemDC&, 0, 0, SRCCOPY)
        'Debug.Print lpds.rcItem.Left
        'Debug.Print (WndObj.MenuSideNumber - 1); lngMenuHight
        'Call BitBlt(hDC&, 0, lngMenuHeight - GetPixelY(WndObj.MenuSideImage.Height), SIDE_WIDTH - 1, GetPixelY(WndObj.MenuSideImage.Height), hmemDC&, 0, 0, SRCCOPY)
        'Call BitBlt(hDC&, 0, MENUSIDE_HEIGHT - GetPixelY(WndObj.MenuSideImage.Height), SIDE_WIDTH - 1, GetPixelY(WndObj.MenuSideImage.Height), hmemDC&, 0, 0, SRCCOPY)
        'Call BitBlt(hDC&, 0, lngMenuHeight - MENUSIDE_HEIGHT, SIDE_WIDTH - 1, MENUSIDE_HEIGHT, hmemDC&, 0, 0, SRCCOPY)
        Call BitBlt(hDC&, 0, MENUSIDE_TOP, SIDE_WIDTH - 1, MENUSIDE_HEIGHT, hmemDC&, 0, 0, SRCCOPY)

        Call SelectObject(hmemDC&, hOldBitmap)
        'Call DeleteObject(SideBitmap&)
    
        Call DeleteDC(hmemDC&)
    Else
        WndObj.MenuSideWidth = 0
        WndObj.MenuSideNumber = 0
    End If

  Else
  
    Dim bDisabled As Boolean
    Dim bSelected As Boolean
    Dim bChecked As Boolean
    Dim bHaveButn As Boolean
  
    bDisabled = lpds.itemState And ODS_GRAYED
    bSelected = lpds.itemState And ODS_SELECTED
    bChecked = lpds.itemState And ODS_CHECKED
    bHaveButn = False
    
    Dim iButton As Integer
    iButton = pmd.iButton
    
    LSet rtButn = rtItem
    
    If WndObj.RightToLeft Then
      rtButn.Left = rtButn.Right - (m_iBitmapWidth + CXBUTTONMARGIN)
    Else
      rtButn.Right = rtButn.Left + m_iBitmapWidth + CXBUTTONMARGIN
    End If
    
    If iButton >= 0 Then
      bHaveButn = True
      
      rtIcon.Left = rtButn.Left + (CXBUTTONMARGIN \ 2)
      rtIcon.Right = rtIcon.Left + m_iBitmapWidth
      rtIcon.Top = rtButn.Top + ((rtButn.Bottom - rtButn.Top) - m_iBitmapWidth) \ 2
      rtIcon.Bottom = rtIcon.Top + m_iBitmapWidth
      'Debug.Print rtIcon.Bottom - rtIcon.Top,
      If Not bDisabled Then
        'Call FillRectEx(hDC&, rtButn, GetSysColor(IIf(bChecked And (Not bSelected), COLOR_BTNLIGHT, COLOR_MENU)))

        If bSelected Or bChecked Then
            Call FillRectEx(hDC&, rtButn, WndObj.SelectColor)
            'Call DrawEdge(hDC&, rtButn, IIf(bChecked, BDR_SUNKENOUTER, BDR_RAISEDINNER), BF_RECT)
            'Call ImageList_Draw(WndObj.ilHandle.hImageList, iButton%, hDC&, rtIcon.Left + 1, rtIcon.Top + 1, ILD_TRANSPARENT)
            Call WndObj.ilHandle.ListImages(iButton + 1).Draw(hDC, GetTwipX(rtIcon.Left + 1), GetTwipY(rtIcon.Top + 1), imlTransparent)
        Else
            'Call FillRectEx(hDC&, rtButn, vbRed)
            'Call ImageList_Draw(WndObj.ilHandle.hImageList, iButton, hDC&, rtIcon.Left, rtIcon.Top, ILD_TRANSPARENT)

            Call WndObj.ilHandle.ListImages(iButton + 1).Draw(hDC, GetTwipX(rtIcon.Left), GetTwipY(rtIcon.Top), imlTransparent)
            'Debug.Print ImageList_Draw(WndObj.ilHandle.hImageList, iButton, hDC&, rtIcon.Left, rtIcon.Top, ILD_NORMAL),
        End If

        'Call ImageList_Draw(WndObj.ilHandle.hImageList, iButton%, hDC&, rtIcon.Left, rtIcon.Top, ILD_TRANSPARENT)
      Else
        
        'Dim hIcon As Long
        
        'hIcon& = ImageList_GetIcon(WndObj.ilHandle.hImageList, iButton%, 0&)

'        Call DrawState(hdc&, 0&, 0&, hIcon&, 0&, rtIcon.Left, rtIcon.Top, rtIcon.Left + m_iBitmapWidth%, rtIcon.Top + m_iBitmapWidth%, DST_ICON Or DSS_DISABLED)
        'Paul DiLascia's funtion is quicker than DrawState;
        ' so he says; colors are a big plus anyway
        Call DrawEmbossed(hDC&, WndObj.ilHandle, iButton, rtIcon, WndObj.ColorEmbossed)
      End If
    Else
      
      Dim Info As MENUITEMINFO
      Info.cbSize = LenB(Info)
      Info.fMask = MIIM_CHECKMARKS
      Call GetMenuItemInfo(lpds.hwndItem, lpds.itemID, MF_BYCOMMAND, Info)
      
      'check marks and radio buttons
'** doesn't handle "info.hbmpUnchecked" yet
      If bChecked Or CBool(Info.hbmpUnchecked) Or (pmd.bAsMark And WndObj.ComplexChecks) Then
        bHaveButn = Draw3DMark(hWnd&, hDC&, rtButn, bChecked, bSelected, bDisabled, IIf(bChecked, Info.hbmpChecked, Info.hbmpUnchecked), pmd.bAsCheck)
      End If
    End If
    
    Dim iButnWidth As Integer
    iButnWidth% = m_iBitmapWidth% + CXBUTTONMARGIN
    
    Dim dwColorBG As Long
    dwColorBG = IIf(bSelected And WndObj.FullSelect, WndObj.SelectColor&, GetSysColor(COLOR_MENU))
    
    LSet rtText = rtItem
    
    If pmd.bMainMenu Then _
      Call FillRectEx(hDC&, rtItem, GetSysColor(COLOR_MENU))
   
    If (bSelected Or (lpds.itemAction = ODA_SELECT)) And (Not bDisabled) Then
      LSet rtHighlight = rtItem
      If bHaveButn Then
        If WndObj.RightToLeft Then
          rtHighlight.Right = rtItem.Right - (iButnWidth% + CXGAP)
        Else
          rtHighlight.Left = rtItem.Left + iButnWidth% + CXGAP
        End If
      End If
      
      If pmd.bMainMenu And bSelected Then
        
        rtText = OffsetRect(rtText, 2, 1)
        Call DrawEdge(hDC&, rtHighlight, BDR_SUNKENOUTER, BF_RECT)
        
      Else
        Call FillRectEx(hDC&, rtHighlight, dwColorBG&)
        'Call DrawEdge(hDC&, rtHighlight, BDR_RAISEDOUTER, BF_RECT)
      
      End If
    End If
    
    If Not pmd.bMainMenu Then
      If WndObj.RightToLeft Then
        rtText.Right = rtItem.Right - (iButnWidth% + CXGAP + CXTEXTMARGIN)
        rtText.Left = rtItem.Left + iButnWidth%
      Else
        rtText.Left = rtItem.Left + iButnWidth% + CXGAP + CXTEXTMARGIN
        rtText.Right = rtItem.Right - iButnWidth%
      End If
    End If
    
    Call SetBkMode(hDC&, TRANSPARENT)
    
    Dim dwSelTextColor As Long
'** This is where I most change text color according to selection brightness
    dwSelTextColor& = GetSysColor(COLOR_HIGHLIGHTTEXT)
    
    Dim dwColorText As Long
    dwColorText& = IIf(bDisabled, GetSysColor(COLOR_GRAYTEXT), _
                  IIf(bSelected And (Not pmd.bMainMenu), _
                  IIf(WndObj.FullSelect, dwSelTextColor&, WndObj.SelectColor&), _
                  IIf(WndObj.ForeColor& = 0&, GetSysColor(COLOR_MENUTEXT), WndObj.ForeColor&)))

    Dim TextOffset As Integer
    'TextOffset = IIf(WndObj.RightToLeft, -1, 1)
    TextOffset = 1


    ' 这是我加的！
    If bSelected Then
      ' 这里可以修改选定的文字颜色！
      If Not bDisabled Then dwColorText = vbWhite
    Else
      If Not bDisabled Then Call FillRectEx(hDC&, rtHighlight, vbWhite) 'GetSysColor(COLOR_MENU))
    End If


    If bDisabled Then _
      Call DrawMenuText(hWnd&, hDC&, OffsetRect(rtText, TextOffset, TextOffset), pmds(CStr(dwItemData)).sMenuText, GetSysColor(COLOR_BTNHIGHLIGHT), True, WndObj.RightToLeft)
    
    Call DrawMenuText(hWnd&, hDC&, rtText, pmd.sMenuText, dwColorText&, Not pmd.bMainMenu, WndObj.RightToLeft)
  End If
  
'Draws the arrows of submenus
  If pmd.bTrueSub Then
    Dim rtArrow As RECT
    LSet rtArrow = rtItem
    
    If WndObj.RightToLeft Then
      rtArrow.Left = rtArrow.Left + CXTEXTMARGIN
    Else
      rtArrow.Right = rtArrow.Right - CXTEXTMARGIN
    End If
    
    rtArrow.Top = rtArrow.Top + CXTEXTMARGIN
    
    Call PrintGlyph(hDC&, IIf(WndObj.RightToLeft, "3", "4"), dwColorText&, rtArrow, IIf(WndObj.RightToLeft, DT_LEFT, DT_RIGHT) Or DT_TOP Or DT_SINGLELINE)

    Call ExcludeClipRect(hDC&, rtItem.Left, rtItem.Top, rtItem.Right, rtItem.Bottom)
    
  End If
  
  Call CopyMemory(ByVal dsPtr&, lpds, Len(lpds))
  
  Set WndObj = Nothing
  
  OnDrawItem = True
  
  Exit Function
  
ErrHandler:
  Debug.Print Err.Number; Err.Description; " OnDrawItem"
  Err.Clear
  
End Function


Private Function OnMeasureItem(hWnd As Long, ByRef miPtr As Long) As Boolean
'Based on Paul DiLascia's
'Mesure menu items

  On Error GoTo ErrHandler

  Dim lpms As MEASUREITEMSTRUCT
  
  Call CopyMemory(lpms, ByVal miPtr, Len(lpms))

  Dim WndObj As WndCoolMenu
  Set WndObj = WndCol(CStr(hWnd&))

  Dim dwItemData As Long
  dwItemData& = lpms.ItemData
  
  If (dwItemData& = 0&) Or (lpms.CtlType <> ODT_MENU) Then
    OnMeasureItem = False
    Exit Function
  End If
  
  Dim pmd As CMyItemData
  Set pmd = pmds(CStr(dwItemData&))
  
  Dim iCYMENU As Integer
  'iCYMENU% = GetSystemMetrics(SM_CYMENU)
  iCYMENU% = 19
  
  Dim rc As RECT
  Dim rcHeight As Integer
  Dim OldFont As Long
  Dim hWndDC As Long

  'm_SideBitmapWidth = 0
  If pmd.fType And MFT_SEPARATOR Then
    hWndDC& = GetDC(hWnd&)
    OldFont& = SelectObject(hWndDC&, GetMenuFont(hWnd&))

    rcHeight = DrawText(hWndDC&, "A", 1&, rc, DT_SINGLELINE Or DT_CALCRECT) + 1
    lpms.itemHeight = IIf(iCYMENU% \ 2 > rcHeight, iCYMENU% \ 2, rcHeight)
    lpms.itemWidth = 0

    Call SelectObject(hWndDC&, OldFont&)
    Call ReleaseDC(hWnd&, hWndDC&)

  ElseIf Left(pmd.sMenuText, 1) = "!" Then
    'WndObj
    WndObj.MenuSideNumber = pmds.Count
    WndObj.MenuSideWidth = SIDE_WIDTH
    lpms.itemHeight = 0
    lpms.itemWidth = 0
    
  Else
    
    hWndDC& = GetDC(hWnd&)
    OldFont& = SelectObject(hWndDC&, GetMenuFont(hWnd&))
    
    'Call DrawText(hWndDC&, pmd.sMenuText, Len(pmd.sMenuText), rc, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT) 'Or DT_VCENTER
    Call DrawText(hWndDC&, pmd.sMenuText, -1, rc, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT) 'Or DT_VCENTER
    Call SelectObject(hWndDC&, OldFont&)
    Call ReleaseDC(hWnd&, hWndDC&)
    
    rcHeight = rc.Bottom - rc.Top
    
    lpms.itemHeight = IIf(rcHeight > iCYMENU%, rcHeight, iCYMENU%)
    
    Dim itemWidth As Long
    itemWidth& = (rc.Right - rc.Left)
    
    If Not pmd.bMainMenu Then
      
      itemWidth& = itemWidth& + (CXTEXTMARGIN * 2) + CXGAP + (m_iBitmapWidth% + CXBUTTONMARGIN) * 2
      itemWidth& = itemWidth& - (GetSystemMetrics(SM_CXMENUCHECK) - 1)
      'Paul explains it better than me; Windows(OS) stuff (the minus that is)
    End If

  ' 我们可以先简单的这样写：
    lpms.itemWidth = itemWidth&
    If Not (WndObj.MenuSideImage Is Nothing) And dwItemData <= WndObj.MenuSideNumber Then
        lpms.itemWidth = lpms.itemWidth + WndObj.MenuSideWidth
    End If

  End If

  'Debug.Print dwItemData; ":"; lpms.itemHeight

  Call CopyMemory(ByVal miPtr, lpms, Len(lpms))

  OnMeasureItem = True

  Exit Function

ErrHandler:
  Debug.Print Err.Number; Err.Description; " OnMeasureItem"
  Err.Clear
End Function

Public Function GetMenuFont(hWnd As Long, Optional bForceReset As Boolean = False) As Long
' This function is from me
' returns the created menu font or the existing one

  On Error GoTo ErrHandler

  Dim WndObj As WndCoolMenu
  Set WndObj = WndCol(CStr(hWnd&))
  
  If (WndObj.MenuFont = 0) Or bForceReset Then
    
    Dim sText As String
    Dim TextLen As Long
    Dim tLF As LOGFONT
    Dim tm As TEXTMETRIC
    
    
    If WndObj.FontName = "" Then
      sText$ = Space$(255)
      TextLen& = Len(sText$)
      
      'Window's DC
      Dim hWndDC As Long: hWndDC& = GetDC(hWnd&)
      
      'Font Name
      TextLen& = GetTextFace(hWndDC&, TextLen&, sText$)
      WndObj.FontName = Left$(sText$, TextLen&)
      
      'Form's fore color
      If WndObj.ForeColor = 0& Then _
        WndObj.ForeColor = GetTextColor(hWndDC&)
      
      Call GetTextMetrics(hWndDC&, tm)
      Call ReleaseDC(hWnd&, hWndDC&)
      
      tLF.lfCharSet = tm.tmCharSet
      tLF.lfHeight = tm.tmHeight
      tLF.lfWeight = tm.tmWeight

    Else
      'If FontName specified, use it + defined size
      tLF.lfWeight = FW_NORMAL
      
      Dim hDC As Long: hDC& = GetWindowDC(hWnd&)
      tLF.lfHeight = -MulDiv(WndObj.FontSize&, GetDeviceCaps(hDC&, LOGPIXELSY), 72)
      
      Call ReleaseDC(hWnd&, hDC&)
    End If

    tLF.lfFaceName = WndObj.FontName$ + Chr(0)
    WndObj.MenuFont& = CreateFontIndirect(tLF)

  End If

  GetMenuFont& = WndObj.MenuFont&

  Set WndObj = Nothing

  Exit Function

ErrHandler:
  Debug.Print Err.Number; Err.Description; " GetMenuFont"
  Err.Clear
End Function

Private Function GetMenuFontSep(hWnd As Long) As Long
' This function is from me
' returns the created menu font for the separator or return the existing one
  
  Dim WndObj As WndCoolMenu
  Set WndObj = WndCol(CStr(hWnd&))
  
  If WndObj.MenuFontSep& = 0& Then
    
    Dim tLF As LOGFONT

    tLF.lfFaceName = "宋体" + Chr(0)

    tLF.lfCharSet = GB2312_CHARSET
    tLF.lfHeight = 12
    tLF.lfWeight = FW_NORMAL

    WndObj.MenuFontSep& = CreateFontIndirect(tLF)
  End If
  
  GetMenuFontSep& = WndObj.MenuFontSep&
  Set WndObj = Nothing
End Function

Public Function Install(wndHandle As Long, Optional ilHandle As Object, Optional SideImage As StdPicture) As Boolean
'Install CoolMenu on the specified window handle
  
  On Error GoTo ErrHandler

  m_iBitmapWidth% = 16
  'm_SideBitmapWidth = 0
  
  If wndHandle <> 0 Then

    ' 删除系统菜单中的“移动”菜单项
    Call DeleteMenu(GetSystemMenu(wndHandle, False), SC_MOVE, MF_BYCOMMAND)

    If WndCol Is Nothing Then
      Set WndCol = New Collection
      Set pmds = New CMyItemDatas
    End If

    Dim NewWnd As WndCoolMenu

    Set NewWnd = New WndCoolMenu

    NewWnd.hWnd = wndHandle&
    NewWnd.PrevProc = GetWindowLong(wndHandle&, GWL_WNDPROC)

    NewWnd.SelectColor = GetSysColor(COLOR_HIGHLIGHT)

    Call SetWindowLong(wndHandle&, GWL_WNDPROC, AddressOf WindowProc)

    If Not (ilHandle Is Nothing) Then _
      Set NewWnd.ilHandle = ilHandle

    ' 我加的，图像及宽度
    If Not (SideImage Is Nothing) Then
        Set NewWnd.MenuSideImage = SideImage
        NewWnd.MenuSideWidth = SIDE_WIDTH
    End If

    NewWnd.SCMainMenu = False

    Call WndCol.Add(NewWnd, CStr(wndHandle&))

    Set NewWnd = Nothing

    'Main menu permanent subclassing
    Call ConvertMenu(wndHandle&, GetMenu(wndHandle&), 0&, False, True, True)

  End If

  Install = True

  Exit Function

ErrHandler:
  Debug.Print Err.Number; Err.Description; " Install"
  Install = False
  Err.Clear
End Function

Public Function Uninstall(wndHandle As Long) As Boolean
'Unintall sublassing on a window by handle

  On Error GoTo ErrHandler

  If (wndHandle <> 0) And (Not (WndCol Is Nothing)) Then

    Call SetWindowLong(wndHandle&, GWL_WNDPROC, WndCol(CStr(wndHandle&)).PrevProc)

    Call WndCol.Remove(CStr(wndHandle&))

    If WndCol.Count = 0 Then
      Set WndCol = Nothing
      Call DeleteObject(m_MarlettFont&)
      Call DeleteObject(m_bmpChecked&)
      Call DeleteObject(m_bmpRadioed)
      Set pmds = Nothing
    End If

    Uninstall = True
  End If

  Exit Function

ErrHandler:
  Debug.Print Err.Number; Err.Description; " Uninstall"
  Uninstall = False
  Err.Clear
End Function

Private Sub FillRectEx(hDC As Long, rc As RECT, Color As Long)
'Also based on Paul DiLascia's
'a good idea to simplify the calls to FillRect
  Dim hOldBrush As Long
  Dim hNewBrush As Long
  
  hNewBrush& = CreateSolidBrush(Color&)
  Call FillRect(hDC&, rc, hNewBrush&)
  Call DeleteObject(hNewBrush&)
End Sub

Private Function OffsetRect(InRect As RECT, ByVal xOffset As Long, ByVal yOffset As Long) As RECT
'don't ask
  OffsetRect.Left = InRect.Left + xOffset&
  OffsetRect.Right = InRect.Right + xOffset&
  OffsetRect.Top = InRect.Top + yOffset&
  OffsetRect.Bottom = InRect.Bottom + yOffset&

End Function

Private Sub OnMenuSelect(hWnd As Long, nItemID As Integer, nFlags As Integer, hSysMenu As Long)
' If click = the end of a menu loop: convert back from OWNERDRAWNED
  
  On Error GoTo ErrHandler
  
  Dim WndObj As WndCoolMenu
  Set WndObj = WndCol(CStr(hWnd&))
  
  Dim Info As MENUITEMINFO

  Info.cbSize = LenB(Info)
  Info.fMask = MIIM_DATA Or MIIM_STATE Or MIIM_TYPE Or MIIM_ID
  
  Call GetMenuItemInfo(GetMenu(hWnd&), nItemID, MF_BYCOMMAND, Info)

  If (hSysMenu = 0&) And (nFlags = &HFFFF) Then
    Dim i As Integer
    For i% = 0 To WndObj.CountMenuHeads
      Call ConvertMenu(hWnd&, WndObj.GetMenuHead(i%), 0&, False, False)
    Next i%
    WndObj.MainPopedIndex = -1
  End If

  Exit Sub

ErrHandler:
  Debug.Print Err.Number; Err.Description; " OnMenuSelect"
  Err.Clear
End Sub

Private Function GetButtonIndex(hWnd As Long, sMenuText As String) As Integer
'This function is from me

' Get the index of the image in the ImageList
' based on the menu caption (including ampersand)
' The function looks at the image tag and tries to
' match the smallest left most part of both strings

'This method is very useful in vb because the first
'objective is the quick user interface
'BUT it causes 2 "major" problems:
' 1- you have to load 2 identical images if you want
'    to use an image for 2 different menu items
' 2- if you have 2 or more identical captions, all those
'    items get the same image

  On Error GoTo ErrHandler

  Dim sTagText As String
  Dim iTagTextLen As Integer
  Dim iTextLen As Integer
  Dim m_ilHandle As Object
  
  Dim i As Integer
  
  GetButtonIndex% = -1
  
  Set m_ilHandle = WndCol(CStr(hWnd&)).ilHandle
  
  If Not (m_ilHandle Is Nothing) Then
    'ImageList object available (hope it's an ImageList)
    
    For i% = 1 To m_ilHandle.ListImages.Count
      
      iTextLen% = Len(sMenuText$)
      
      'Get the smallest length
      sTagText$ = m_ilHandle.ListImages(i%).Tag
      If sTagText$ = "" Then GoTo Continue

      iTagTextLen% = Len(sTagText$)
      If iTagTextLen% < iTextLen% Then iTextLen% = iTagTextLen%
      
      'Get the smallest match
      If UCase(Left(sMenuText$, iTextLen%)) = UCase(Left(sTagText$, iTextLen%)) Then
        
        GetButtonIndex = i% - 1 'starts at 0
        Exit Function
      End If

Continue:
    Next i%
  End If

  Set m_ilHandle = Nothing

  Exit Function

ErrHandler:
  Debug.Print Err.Number; Err.Description; " GetButtonIndex"
  Err.Clear
End Function

Private Function BSTRtoLPSTR(sBSTR As String, b() As Byte, lpsz As Long) As Long
'Convert a vb string to a byte array for GetMenuItemInfo and GetMenuFont
'This function is from Steven Roman who wrote a book
' which I didn't read (sorry). It's called
' "Win32 API Programming with Visual Basic", 534 pages
' ISDN : 1-56592-631-5; about $40 (got the info where I read the part)

'Chapter 6 was fortunatly made available on the net by
' the author at the right time

'Input: a non empty BSTR string
'Input: byte array b()
'Output: Fills byte array b() with ANSI har string
'Output: Fills lpsz with pointer to b() array
'Returns byte count, not including terminating null
'BSTR not affected
  
  Dim cBytes As Long
  'Get the number of bytes in the string
  cBytes = LenB(sBSTR)
  
  'Redim the array to hold it + 2 for Unicode null
  ReDim b(1 To cBytes + 2) As Byte
  
  Dim sABSTR As String
  'Set sABSTR to ASCII equivalent
  sABSTR = StrConv(sBSTR, vbFromUnicode)
  
  'Get a long pointer to the string
  lpsz = StrPtr(sABSTR)
  
  CopyMemory b(1), ByVal lpsz, cBytes + 2
  
  lpsz = VarPtr(b(1))
  
  BSTRtoLPSTR = cBytes

End Function

Private Sub DrawEmbossed(hDC As Long, ilHandle As Object, iButnIndex As Integer, rt As RECT, bInColor As Boolean)
'This function is almost word to word copy(well...translation)
'of Paul DiLascia's DrawEmbossed function.
'You can find most of the original C++ code in comment

  On Error GoTo ErrHandler

  'Dim Info As IMAGEINFO
  Dim rcImage As RECT

  'Call ImageList_GetImageInfo(ilHandle.hImageList, iButnIndex%, Info)

  Dim cx As Integer, cy As Integer

  'LSet rcImage = Info.rcImage
  rcImage.Right = GetPixelX(ilHandle.ListImages(iButnIndex + 1).Picture.Width)
  rcImage.Bottom = GetPixelY(ilHandle.ListImages(iButnIndex + 1).Picture.Height)
  cx% = rcImage.Right - rcImage.Left
  cy% = rcImage.Bottom - rcImage.Top

'  // create memory dc
  Dim hmemDC As Long
  hmemDC& = CreateCompatibleDC(hDC&)

'  // create mono or color bitmap
  Dim hBitmap As Long
  If bInColor Then
    hBitmap& = CreateCompatibleBitmap(hDC&, cx%, cy%)
  Else
    hBitmap& = CreateBitmap(cx%, cy%, 1, 1, vbNull)
'    bm.CreateBitmap(cx, cy, 1, 1, NULL);
  End If

'  // draw image into memory DC--fill BG white first
  Dim hOldBitmap As Long
  hOldBitmap = SelectObject(hmemDC&, hBitmap&)

'  memdc.PatBlt(0, 0, cx, cy, WHITENESS);
  Call PatBlt(hmemDC&, 0, 0, cx%, cy%, WHITENESS)
'  il.Draw(&memdc, i, CPoint(0,0), ILD_TRANSPARENT);
  'Call ImageList_Draw(ilHandle&, iButnIndex%, hmemDC&, 0, 0, ILD_TRANSPARENT)
  Call ilHandle.ListImages(iButnIndex + 1).Draw(hmemDC, 0, 0, imlTransparent)

'  // This seems to be required. Why, I don't know. ???
'  COLORREF colorOldBG = dc.SetBkColor(CWHITE);
  Dim hOldBackColor As Long
  hOldBackColor& = SetBkColor(hDC&, RGB(255, 255, 255))

'  // Draw using hilite offset by (1,1), then shadow
'  CBrush brShadow(GetSysColor(COLOR_3DSHADOW));
'  CBrush brHilite(GetSysColor(COLOR_3DHIGHLIGHT));
  Dim hbrShadow As Long, hbrHilite As Long
  hbrShadow& = CreateSolidBrush(GetSysColor(COLOR_BTNSHADOW))
  hbrHilite& = CreateSolidBrush(GetSysColor(COLOR_BTNHIGHLIGHT))

'  CBrush* phOldBrush = dc.SelectObject(&brHilite);
  Dim hOldBrush As Long
  hOldBrush& = SelectObject(hDC&, hbrHilite&)

'  dc.BitBlt(p.x+1, p.y+1, cx, cy, &memdc, 0, 0, MAGICROP);
  Call BitBlt(hDC&, rt.Left + 1, rt.Top + 1, cx%, cy%, hmemDC&, 0, 0, MAGICROP)
  Call SelectObject(hDC&, hbrShadow&)

'  dc.BitBlt(p.x, p.y, cx, cy, &memdc, 0, 0, MAGICROP);
  Call BitBlt(hDC&, rt.Left, rt.Top, cx%, cy%, hmemDC&, 0, 0, MAGICROP)
  
'  dc.SelectObject(phOldBrush);
  Call SelectObject(hDC&, hOldBrush&)

'  dc.SetBkColor(colorOldBG);         // restore
  Call SetBkColor(hDC&, hOldBackColor&)

'  memdc.SelectObject(pOldBitmap);    // ...
  Call SelectObject(hmemDC&, hOldBitmap&)

  Call DeleteObject(hOldBrush&)
  Call DeleteObject(hbrHilite&)
  Call DeleteObject(hbrShadow&)
  Call DeleteObject(hOldBackColor&)
  Call DeleteObject(hOldBitmap&)
  Call DeleteObject(hBitmap&)

  Call DeleteDC(hmemDC&)

  Exit Sub

ErrHandler:
  Debug.Print "DrawEmbossed Error!"
End Sub

Private Function Draw3DMark(hWnd As Long, hDC As Long, rc As RECT, bCheck As Boolean, bSelected As Boolean, bDisabled As Boolean, hBmp As Long, bDrawCheck As Boolean) As Boolean
'This horror is from me.  I wanted it to be clean and efficient
'turns out it's ugly and slow (but it works)

  On Error GoTo hError
  
  Dim WndObj As WndCoolMenu
  Set WndObj = WndCol(CStr(hWnd&))
  
  Dim cx As Integer, cy As Integer
  cx% = rc.Right - rc.Left
  cy% = rc.Bottom - rc.Top
  
  If Not CBool(hBmp) Then
    
    Dim hmemDC As Long
    Dim hBmpTemp As Long
    Dim hOldBmp As Long
    Dim hBrush As Long
    Dim lbInfo As LOGBRUSH
    Dim hOldBrush As Long
    
    If WndObj.ComplexChecks Then
      
      hmemDC& = CreateCompatibleDC(hDC&)
      
      Dim rcHighLigth As RECT
      LSet rcHighLigth = rc
      rcHighLigth.Right = rcHighLigth.Right + 1
      rcHighLigth.Left = rcHighLigth.Left - 1
      
      Call FillRectEx(hDC&, rcHighLigth, IIf(bSelected And (Not bDisabled) And WndObj.FullSelect, WndObj.SelectColor&, GetSysColor(COLOR_MENU)))
      
      If m_bmpChecked = 0& Then
        m_bmpChecked& = LoadImage(0&, CLng(OBM_CHECKBOXES), IMAGE_BITMAP, 0&, 0&, LR_DEFAULTCOLOR)
        m_bmpRadioed& = LoadImage(0&, CLng(OBM_BTNCORNERS), IMAGE_BITMAP, 0&, 0&, LR_MONOCHROME)
      End If
    
      'Using a Hollow brush.  I DON'T know why
      'it's needed. The effect was working on the form's
      'background and not in the menu.  I had to extract
      'the form's brush to find out it's properties and
      'it was a hollow brush.  So I'm using one in the menu.
      lbInfo.lbStyle = BS_HOLLOW
'      lbInfo.lbColor =
'      lbInfo.lbHatch =
      hBrush& = CreateBrushIndirect(lbInfo)
      
      hOldBrush& = SelectObject(hDC&, hBrush&)
      
      'The x and y are used to get the right square from
      'the image which contains all sorts of checks.
      Dim X As Long: X = 0
      Dim Y As Long: Y = 0
      If bCheck Then X = X + 13
      If bDisabled Then X = X + 26
      Y = 0
      
      Dim hOldBackColor As Long
      hOldBackColor& = SetBkColor(hDC&, RGB(255, 255, 255))
      
      If bDrawCheck Then
        hOldBmp& = SelectObject(hmemDC&, m_bmpChecked&)
'        Call BitBlt(hdc&, rc.Left + 3, rc.Top + 3, 13&, 13&, hmemDC&, 0&, 0&, WHITENESS)
'        Call PatBlt(hdc&, rc.Left + 3, rc.Top + 3, 13&, 13&, WHITENESS)
        Call BitBlt(hDC&, rc.Left + 3, rc.Top + 3, 13&, 13&, hmemDC&, X&, Y&, SRCCOPY)
      Else
        Y = 13
        'm_bmpRadioed contains the image mask of the round box
        hOldBmp& = SelectObject(hmemDC&, m_bmpRadioed&)
        Call BitBlt(hDC&, rc.Left + 3, rc.Top + 3, 13&, 13&, hmemDC&, 0&, 0&, MERGEPAINT)
        Call SelectObject(hmemDC&, m_bmpChecked&)
        Call BitBlt(hDC&, rc.Left + 3, rc.Top + 3, 13&, 13&, hmemDC&, X&, Y&, SRCAND)
      End If

''You can draw the entire image on the form with these four lines
'      Dim wdc As Long: wdc& = GetDC(hWnd&)
'      Call SelectObject(hmemDC&, m_bmpChecked&)
'      Call BitBlt(wdc&, 2, 2, 52, 37&, hmemDC&, x&, 13, SRCcopy)
'      Call ReleaseDC(hWnd&, wdc&)
      
      Call SetBkColor(hDC&, hOldBackColor&)
      
      Call SelectObject(hmemDC&, hOldBmp&)
      Call DeleteObject(hOldBmp&)
      
      Call SelectObject(hDC&, hOldBrush)
      Call DeleteObject(hBrush&)
    
      Call DeleteDC(hmemDC&)

    Else
      
      If bSelected Then
        Call FillRectEx(hDC&, rc, GetSysColor(COLOR_MENU))
      Else
        Dim i As Integer
        Dim BitArray(0 To 3) As Long
        For i = 0 To 3
          BitArray(i) = MakeLong(170, 85) '&HAA &H55 = 10101010
                                          '            01010101
        Next i
        
        Dim hPat As Long
        hPat& = CreateBitmap(8&, 8&, 1&, 1&, BitArray(0))
  
  
        Dim hPatBrush As Long
        hPatBrush& = CreatePatternBrush(hPat&)
        
        hOldBrush& = SelectObject(hDC&, hPatBrush&)
  
        Call SetBkColor(hDC&, GetSysColor(COLOR_MENU))
        Call SetTextColor(hDC&, GetSysColor(COLOR_BTNHIGHLIGHT))
        
        Call PatBlt(hDC&, rc.Left, rc.Top, cx%, cy%, PATCOPY)
  
        Call SelectObject(hDC&, hOldBrush&)
        
        Call DeleteObject(hPatBrush&)
        Call DeleteObject(hOldBrush&)
      End If
      
      If bDisabled Then
        
        Call PrintGlyph(hDC&, IIf(bDrawCheck, "a", "h"), GetSysColor(COLOR_BTNHIGHLIGHT), OffsetRect(rc, 1, 1), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
      
        Call PrintGlyph(hDC&, IIf(bDrawCheck, "a", "h"), GetSysColor(COLOR_GRAYTEXT), rc, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
      Else
        Call PrintGlyph(hDC&, IIf(bDrawCheck, "a", "h"), WndObj.ForeColor, rc, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
      End If
        
      Call DrawEdge(hDC&, rc, BDR_SUNKENOUTER, BF_RECT)
    End If
  Else
      
      'Bitmap argument is valid
  
  End If

  Draw3DMark = True

  Set WndObj = Nothing
  Exit Function

hError:
  Debug.Print Err.Number; Err.Description; " ( Draw3DMark )"
  Err.Clear
End Function

Private Function OnDrawMainMenu(hWnd As Long, lParam As Long, MousePosition As Long) As Long
'This draws IE4  (and W98) hover effect over main menu items
'I'm pretty pround of this one because I used the API
'MenuItemFromPoint by tricking it big time.
  
  On Error GoTo NOPOP
  
  Dim hDC As Long
  
  Dim WndObj As WndCoolMenu
  Set WndObj = WndCol(CStr(hWnd&))
  
  'A menu is selected and popped down.
  'A bug in the previous release cause the item to
  'popup even when menu was opened
  If WndObj.MainPopedIndex = -2 Then
    Set WndObj = Nothing
    Exit Function
  End If
  
  'Get the main menu handle from the window handle
  Dim hMenu As Long:  hMenu& = GetMenu(hWnd&)
  
  'This block is for MDI apps
  'the work area doesn't send the message I used so I use two msg
  If MousePosition <> 5 And MousePosition > 0 Then GoTo NOPOP
  If MousePosition = 5 Then
    Set WndObj = Nothing
    Exit Function
  End If

  
  Dim dwPapi As Double
  Dim Papi As POINTAPI
  
  'Get the position of the hit from lParam&
  Papi.X = LoWord(lParam&)
  Papi.Y = HiWord(lParam&)
  
  'This is the work around for the problem of the ByVal user-defined
  'structures.  If your try passing the POINTAPI byval to any "...FromPoint"
  'function, vb won't permit it; and removing the ByVal keyword
  'generates a Bad DLL calling error.  It works if you give the
  'function an 8 bytes variable instead of an 8 bytes structure (POINTAPI in this case)
  Call CopyMemory(dwPapi, Papi, LenB(Papi))
  
  Dim MenuHitIndex As Long
  'Get the hit, -1 if none
  MenuHitIndex& = MenuItemFromPoint(hWnd&, hMenu&, dwPapi)

  'No poped item, erase the old one if exists and exit
  If MenuHitIndex& = -1 Then GoTo NOPOP
  
  Dim PopedIndex As Long
  PopedIndex& = WndObj.MainPopedIndex

  'If the old and new pop are the same, don't redraw
  If MenuHitIndex& = PopedIndex& Then
    Set WndObj = Nothing
    Exit Function
  End If
  
  Dim Info As MENUITEMINFO
  Info.cbSize = LenB(Info)
  Info.fMask = MIIM_TYPE
  Call GetMenuItemInfo(hMenu&, MenuHitIndex&, MF_BYPOSITION, Info)

  If Info.fType And (Not MFT_OWNERDRAW) Then GoTo NOPOP
      
'Debug.Print "OnDrawMainMenu; hWnd&="; hWnd&; " hMenu&="; hMenu&; " MenuHit="; MenuHitIndex&
  
  'Erase the old pop if exists and return
  If PopedIndex& <> -1 Then GoSub DRAWFLAT 'yes gosub can still be useful
      
  
  Dim MenuRect As RECT
  'Get the new pop rect
  Call GetMenuItemRect(hWnd&, hMenu&, MenuHitIndex&, MenuRect)
      
  'Set the old pop reference with the new pop
  WndObj.MainPopedIndex = MenuHitIndex&

  'Draw on the sreen DC
  hDC& = GetDC(0&) 'Get the screen DC
  Call DrawEdge(hDC&, MenuRect, BDR_RAISEDINNER, BF_RECT)
  Call ReleaseDC(hWnd&, hDC&)
      
  OnDrawMainMenu = True
  
  Set WndObj = Nothing
  Exit Function
  
NOPOP:
  'Draw flat if there's a poped item
  If WndObj.MainPopedIndex > -1 Then
    GoSub DRAWFLAT
    WndObj.MainPopedIndex = -1
  End If
  
  Set WndObj = Nothing
  Exit Function
  
DRAWFLAT: 'Old style Sub
  'Erase old hit with flat edges
  
  Dim OldPopedRect As RECT
  'Get the old hit's rect
  Call GetMenuItemRect(hWnd&, hMenu&, CLng(WndObj.MainPopedIndex), OldPopedRect)
  
  'Draw on the sreen DC
  hDC& = GetDC(0&) 'Get the screen DC
  Call DrawEdge(hDC&, OldPopedRect, BDR_RAISEDINNER, BF_RECT Or BF_FLAT)
  Call ReleaseDC(hWnd&, hDC&)
    
Return

End Function

Private Sub PrintGlyph(hDC As Long, Glyph As String, Color As Long, rt As RECT, ByVal wFormat As Long)
'This function draws Window elements using the Marlett font
'it's hidden in the font directory but you can view it and it
'contains interesting elements.  Combining them create checks
'and radio buttons. Now I use the LoadImage API but previous
'version of PrintGlyph printed multiple layers of Marlett characters

  'Create the Marlett font if it doesn't exist already
  If m_MarlettFont& = 0& Then
    Dim tLF As LOGFONT
  
    tLF.lfFaceName = "Marlett" + Chr(0)
    tLF.lfCharSet = SYMBOL_CHARSET
    tLF.lfHeight = 13 'the value could be changed in relation to the real MenuFont to draw proportional boxes
'    tLF.lfWeight = 500
'    tLF.lfWidth = 31

    m_MarlettFont& = CreateFontIndirect(tLF)
  End If

  'write text with transparent background
  Call SetBkMode(hDC&, TRANSPARENT)
    
  Dim hOldFont As Long
  
  'Select the font for the device context
  hOldFont& = SelectObject(hDC&, m_MarlettFont&)
  
  'select the color for the glyph
  Call SetTextColor(hDC&, Color&)
  
  Call DrawText(hDC&, Glyph, 1, rt, wFormat&)

End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  On Error GoTo ErrorHandle

  Select Case Msg&

    Case WM_MOVE:
        If MainForm.WindowState = vbNormal Then
            Call frmOnline.FormMove(False)
            Call frmTable.FormMove(False)
            Call frmChat.FormMove(False)
        End If

    Case WM_ACTIVATE:
        If wParam = 0 Then
            Call MainForm.LostFocus
        Else
            Call MainForm.GetFocus
        End If

    Case WM_ACTIVEWINDOW:
        'Call ReleaseCapture
        'Call AppActivate(App.Title)
        Call KeepOnTop(hWnd)
        Call KillOnTop(hWnd)
        Call MainForm.BossKey(True)

    Case WM_HOTKEY
        If (wParam = HOTKEY_ID) And (LoWord(lParam) = MOD_CONTROL) And (HiWord(lParam) = KEY_HOTKEY) Then
            Call MainForm.BossKey
        End If

    Case WM_THINKEND
        Call MainForm.ComputerDownMan(wParam, lParam)

    Case WM_SETCURSOR:
'        Debug.Print "WM_SETCURSOR " & LoWord(lParam&); HiWord(lParam&); LoWord(wParam&); HiWord(wParam&)
'            If Not m_IsWin98 Then Call OnDrawMainMenu(hWnd&, 0&, LoWord(lParam&))
            Call OnDrawMainMenu(hWnd&, 0&, LoWord(lParam&))
    
    Case WM_NCHITTEST:
'        Debug.Print "WM_NCHITTEST " & LoWord(lParam&); HiWord(lParam&); LoWord(wParam&); HiWord(wParam&)
'            If Not m_IsWin98 Then Call OnDrawMainMenu(hWnd&, lParam&, -1)
            Call OnDrawMainMenu(hWnd&, lParam&, -1)
    
    Case WM_NCMOUSEMOVE:
'        Debug.Print "WM_NCMOUSEMOVE " & LoWord(lParam&); HiWord(lParam&); LoWord(wParam&); HiWord(wParam&)
    
    Case WM_MEASUREITEM:
'        Debug.Print "WM_MEASUREITEM " & LoWord(lParam&); HiWord(lParam&); LoWord(wParam&); HiWord(wParam&)
      
            If OnMeasureItem(hWnd&, lParam&) Then
              WindowProc = True
              Exit Function
            End If
            
    Case WM_DRAWITEM:
'        Debug.Print "WM_DRAWITEM " & LoWord(lParam&); HiWord(lParam&); LoWord(wParam&); HiWord(wParam&)
            
            If OnDrawItem(hWnd&, lParam&) Then
              WindowProc = True
              Exit Function
            End If
    
    Case WM_INITMENUPOPUP:
'        Debug.Print "WM_INITMENUPOPUP " & LoWord(lParam&); HiWord(lParam&); wParam&
            'm_SideBitmapWidth = 0
            'm_SideBitmapWidth = SIDE_WIDTH
            
            Call CallWindowProc(WndCol(CStr(hWnd&)).PrevProc, ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
            Call OnInitMenuPopup(hWnd&, wParam&, LoWord(lParam&), CBool(HiWord(lParam&)))
            WindowProc = 0&
            Exit Function
            
    Case WM_MENUCHAR:
            
            Dim result As Long
            result = OnMenuChar(LoWord(wParam&), HiWord(wParam&), lParam&)

            If result <> 0 Then
              WindowProc = result
              Exit Function
            End If
            
    Case WM_MENUSELECT:
'        Debug.Print "WM_MENUSELECT " & LoWord(lParam&); HiWord(lParam&); LoWord(wParam&); HiWord(wParam&)
            Call OnMenuSelect(hWnd&, LoWord(wParam&), HiWord(wParam&), lParam&)
      
    Case WM_WINDOWPOSCHANGED:
'        Debug.Print "WM_WINDOWPOSCHANGED " & LoWord(lParam&); HiWord(lParam&); LoWord(wParam&); HiWord(wParam&)
            'For MDI apps (child) main menu bar "refresh"
            'Dim Oldh As Long
            'Oldh& = WndCol(CStr(hWnd&)).SCMainMenu
            'If (Oldh& <> 0&) And (Oldh& <> -1&) And (GetMenu(hWnd&) <> Oldh&) Then
            '  WndCol(CStr(hWnd&)).SCMainMenu = True
            '  Call ConvertMenu(hWnd&, GetMenu(hWnd&), 0&, False, True, True)
            'End If

'    Case Else:
'    If msg& <> &H121 And msg& <> &H200 Then Debug.Print Hex(msg&) + "  " & LoWord(lParam&); HiWord(lParam&); LoWord(wParam&); HiWord(wParam&)
  End Select

Continue:
  WindowProc& = CallWindowProc(WndCol(CStr(hWnd&)).PrevProc, hWnd&, Msg&, wParam&, lParam&)
  Exit Function

ErrorHandle:
  Debug.Print Err.Number; Err.Description; " WindowProc"
  Err.Clear
End Function

Private Function HiWord(LongIn As Long) As Integer
  Call CopyMemory(HiWord, ByVal VarPtr(LongIn) + 2, 2)
End Function

Private Function LoWord(LongIn As Long) As Integer
  Call CopyMemory(LoWord, LongIn, 2)
End Function

Private Function HiByte(WordIn As Integer) As Byte
  Call CopyMemory(HiByte, ByVal VarPtr(WordIn) + 1, 2)
End Function

Private Function LoByte(WordIn As Integer) As Byte
  Call CopyMemory(LoByte, WordIn, 2)
End Function

Private Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
'Creates a Long value using Low and High integers
'Useful when converting code from C++

  MakeLong = CLng(LoWord)
  Call CopyMemory(ByVal VarPtr(MakeLong) + 2, HiWord, 2)
  
End Function

Private Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Integer
'Creates an integer value using Low and High bytes
'Useful when converting code from C++
  
  MakeWord = CInt(LoByte)
  Call CopyMemory(ByVal VarPtr(MakeWord) + 1, HiByte, 1)
End Function

Public Function ColorEmbossed(hWnd As Long, Optional Value As Variant) As Boolean
'Set/Get the "" property
  On Error Resume Next
  
  If IsMissing(Value) Then
    ColorEmbossed = WndCol(CStr(hWnd&)).ColorEmbossed
  Else
    WndCol(CStr(hWnd&)).ColorEmbossed = Value
    ColorEmbossed = Value
  End If
End Function

Public Function ComplexChecks(hWnd As Long, Optional Value As Variant) As Boolean
'Set/Get the "" property
  On Error Resume Next
  
  If IsMissing(Value) Then
    ComplexChecks = WndCol(CStr(hWnd&)).ComplexChecks
  Else
    WndCol(CStr(hWnd&)).ComplexChecks = Value
    ComplexChecks = Value
    
  End If
  
End Function

Public Function SelectColor(hWnd As Long, Optional Value As Variant) As Long
'Set/Get the "" property
  On Error Resume Next
  
  If IsMissing(Value) Then
    SelectColor = WndCol(CStr(hWnd&)).SelectColor
  Else
    WndCol(CStr(hWnd&)).SelectColor = Value
    SelectColor = Value
  End If
End Function

Public Function RightToLeft(hWnd As Long, Optional Value As Variant) As Boolean
'Set/Get the "" property
  On Error Resume Next
  
  If IsMissing(Value) Then
    RightToLeft = WndCol(CStr(hWnd&)).RightToLeft
  Else
    WndCol(CStr(hWnd&)).RightToLeft = Value
    RightToLeft = Value
  End If
End Function

Public Function FullSelect(hWnd As Long, Optional Value As Variant) As Boolean
'Set/Get the "" property
  On Error Resume Next

  If IsMissing(Value) Then
    FullSelect = WndCol(CStr(hWnd&)).FullSelect
  Else
    WndCol(CStr(hWnd&)).FullSelect = Value
    FullSelect = Value
  End If
End Function

Public Function FontSize(hWnd As Long, Optional Value As Variant) As Long
'Set/Get the "" property
  On Error Resume Next
  
  If IsMissing(Value) Then
    FontSize = WndCol(CStr(hWnd&)).FontSize
  Else
    WndCol(CStr(hWnd&)).FontSize = Value
    Call DrawMenuBar(hWnd&)
    FontSize = Value
  End If
End Function

Public Function ForeColor(hWnd As Long, Optional Value As Variant) As Long
'Set/Get the "" property
  On Error Resume Next
  
  If IsMissing(Value) Then
    ForeColor = WndCol(CStr(hWnd&)).ForeColor
  Else
    WndCol(CStr(hWnd&)).ForeColor = Value
    Call DrawMenuBar(hWnd&)
    ForeColor = Value
  End If
End Function

Public Function FontName(hWnd As Long, Optional Value As Variant) As String
'Set/Get the "" property
  On Error Resume Next
  
  If IsMissing(Value) Then
    FontName = WndCol(CStr(hWnd&)).FontName
  Else
    WndCol(CStr(hWnd&)).FontName = Value
    Call DrawMenuBar(hWnd&)
    FontName = Value
  End If
End Function

Public Sub MDIChildMenu(hWnd As Long)
' This sub is used in the Form_Load event of MDI childs
' that want their main menu bar to be subclassed.
' The menu convertion occurs in the WindowProc function
' on WindowPosChanged message; the best solution I could find
'
' NOTE : mother and childs all use the same ImageList located
' on the MDI mother form
  
  On Error Resume Next

'There seems to be a window between the MDI mother
'  form and MDI childs
  Dim ParentWnd As Long: ParentWnd& = GetParent(GetParent(hWnd&))
  
  Dim WndObj As WndCoolMenu
  Set WndObj = WndCol(CStr(ParentWnd&))
  
  'If the object exist, sets the current Main menu handle
  'for later comparaison
  If Not (WndObj Is Nothing) Then
    If WndObj.SCMainMenu Then WndObj.SCMainMenu = GetMenu(ParentWnd&)
  End If
  
  Set WndObj = Nothing
End Sub

Public Function MakeRop4(fore As RasterOpConstants, back As RasterOpConstants) As Long
'  (DWORD)((((back) << 8) & 0xFF000000) | (fore))
  MakeRop4 = MakeLong(0, MakeWord(0, LoByte(LoWord(back)))) Or fore
End Function
