VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WndCoolMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
''  WndCoolMenu Class v1.3
''
''  Copyright Olivier Martin 2000
''
''  martin.olivier@bigfoot.com
''
''  This class is used to keep information on
''  each of the subclassed windows.
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'local variable(s) to hold property value(s)
Private mvarPrevProc As Long 'local copy
Private mvarhWnd As Long 'local copy
Private mvarSCMainMenu As Long
Private mvarMenuHeads() As Long
Private mvarilHandle As Object

Private m_MenuFont As Long 'Font used in menus
Private m_MenuFontSep As Long 'Font used on separators
Private m_sFontName As String 'Font face name (no null)
Private m_FontSize As Long 'Font size
Private m_ForeColor As Long 'Color used for the unselected menu items
Private m_bColorEmbossed As Boolean 'draws disabled images in colore
Private m_bRightToLeft As Boolean 'draws text from right to left
Private m_FullSelect As Boolean 'Selection is a bar or just text
Private m_SelectColor As Long 'Color of the selection bar or text
Private m_bComplexRadioAndCheck As Boolean 'Draw complex radio and check boxes

Private iMainPopedIndex As Integer 'last poped main menu element
'保持属性值的局部变量
Private mvarMenuSideImage As StdPicture '局部复制
Private mvarMenuSideWidth As Long '局部复制
Private mvarMenuSideNumber As Long '局部复制

Public Property Let MenuSideNumber(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.MenuSideNumber = 5
    mvarMenuSideNumber = vData
End Property

Public Property Get MenuSideNumber() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.MenuSideNumber
    MenuSideNumber = mvarMenuSideNumber
End Property

Public Property Let MenuSideWidth(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.MenuSideWidth = 5
    mvarMenuSideWidth = vData
End Property

Public Property Get MenuSideWidth() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.MenuSideWidth
    MenuSideWidth = mvarMenuSideWidth
End Property

Public Property Set MenuSideImage(ByVal vData As StdPicture)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.MenuSideImage = Form1
    Set mvarMenuSideImage = vData
End Property

Public Property Get MenuSideImage() As StdPicture
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.MenuSideImage
    Set MenuSideImage = mvarMenuSideImage
End Property

Public Property Let hWnd(ByVal vData As Long)
    mvarhWnd = vData
End Property

Public Property Get hWnd() As Long
    hWnd = mvarhWnd
End Property

Public Property Let PrevProc(ByVal vData As Long)
    mvarPrevProc = vData
End Property

Public Property Get PrevProc() As Long
    PrevProc = mvarPrevProc
End Property

Public Function CountMenuHeads() As Long
    CountMenuHeads = UBound(mvarMenuHeads)
End Function

Public Sub AddMenuHead(ByVal hMenu As Long)
    Dim i As Integer
    Dim Bound As Integer

    On Error Resume Next

    Bound% = UBound(mvarMenuHeads)
    For i% = 0 To Bound%
      If mvarMenuHeads(i%) = hMenu Then Exit Sub
    Next i%

    ReDim Preserve mvarMenuHeads(0 To Bound% + 1) As Long

    mvarMenuHeads(Bound%) = hMenu
End Sub

Public Function GetMenuHead(ByVal HeadIndex As Long) As Long
    GetMenuHead = 0
    
    If UBound(mvarMenuHeads) >= HeadIndex Then _
      GetMenuHead = mvarMenuHeads(HeadIndex)

End Function

Public Property Set ilHandle(ByVal vData As Object)
    
    Set mvarilHandle = vData
End Property

Public Property Get ilHandle() As Object
    
    Set ilHandle = mvarilHandle
End Property

Private Sub Class_Initialize()
    On Error Resume Next

    ReDim mvarMenuHeads(0 To 0) As Long

    Set mvarilHandle = Nothing
    mvarSCMainMenu = 0&

    iMainPopedIndex = -1

    m_FontSize& = 9& 'Default value (8&)
    m_FullSelect = True 'Default value (True)

    m_bRightToLeft = False
    m_bComplexRadioAndCheck = True 'Default value (True)
    m_bColorEmbossed = True 'Default value (True)
End Sub

Private Sub Class_Terminate()
    On Error Resume Next

    ' delete the fonts
    Call DeleteObject(m_MenuFont&)
    Call DeleteObject(m_MenuFontSep&)

    Set mvarilHandle = Nothing
End Sub


Public Property Let ComplexChecks(ByVal vData As Boolean)
  m_bComplexRadioAndCheck = vData
    
End Property

Public Property Get ComplexChecks() As Boolean
    ComplexChecks = m_bComplexRadioAndCheck
End Property

Public Property Let ColorEmbossed(ByVal vData As Boolean)
    m_bColorEmbossed = vData
End Property

Public Property Get ColorEmbossed() As Boolean
    ColorEmbossed = m_bColorEmbossed
End Property

Public Property Let RightToLeft(ByVal vData As Boolean)
    m_bRightToLeft = vData
End Property

Public Property Get RightToLeft() As Boolean
    RightToLeft = m_bRightToLeft
End Property

Public Property Let SelectColor(ByVal vData As Long)
    m_SelectColor& = vData&
End Property

Public Property Get SelectColor() As Long
    SelectColor& = m_SelectColor&
End Property

Public Property Let FullSelect(ByVal vData As Boolean)
    m_FullSelect = vData
End Property

Public Property Get FullSelect() As Boolean
    FullSelect = m_FullSelect
End Property

Public Property Let FontSize(ByVal vData As Long)
    m_FontSize = vData
    Call GetMenuFont(mvarhWnd, True)
End Property

Public Property Get FontSize() As Long
    FontSize = m_FontSize
End Property

Public Property Let ForeColor(ByVal vData As Long)
    m_ForeColor& = vData
End Property

Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor&
End Property

Public Property Let FontName(ByVal vData As String)
    m_sFontName = vData
    Call GetMenuFont(mvarhWnd, True)
End Property

Public Property Get FontName() As String
    FontName = m_sFontName
End Property

Public Property Let MainPopedIndex(ByVal vData As String)
    iMainPopedIndex = vData
End Property

Public Property Get MainPopedIndex() As String
    MainPopedIndex = iMainPopedIndex
End Property

Public Property Let MenuFont(ByVal vData As Long)
    m_MenuFont = vData
End Property

Public Property Get MenuFont() As Long
    MenuFont = m_MenuFont
End Property

Public Property Let MenuFontSep(ByVal vData As Long)
    m_MenuFontSep = vData
End Property

Public Property Get MenuFontSep() As Long
    MenuFontSep = m_MenuFontSep
End Property

Public Property Let SCMainMenu(ByVal vData As Long)
    mvarSCMainMenu = vData
End Property

Public Property Get SCMainMenu() As Long
    SCMainMenu = mvarSCMainMenu
End Property
