VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMyItemData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Existe un pour chaque élément de menu OWNERDRAW"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
''  CMyItemData Class v1.3
''
''  Copyright Olivier Martin 2000
''
''  martin.olivier@bigfoot.com
''
''  This class will contain information about
''  a converted menu item.  It as the following
''  properties:
''
''  bAsCheck  : If the mark is a check
''  bAsMark   : If the item as a mark (no pic)
''  bMainMenu : If it's a main menu bar item
''  fType     : The type of menu item (flags)
''  iButton   : The index of the image list
''  sMenuText : The text of the menu item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private mvarsMenuText As String 'local copy
Private mvarfType As Long 'local copy
Private mvariButton As Integer 'local copy
Private mvarbMainMenu As Boolean 'local copy
Private mvarbAsCheck As Boolean 'local copy
Private mvarbAsMark As Boolean 'local copy
Private mvarsMenuHelp As String 'local copy
Private mvarbTrueSub As Boolean 'local copy

Public Property Let bTrueSub(ByVal vData As Boolean)
    mvarbTrueSub = vData
End Property

Public Property Get bTrueSub() As Boolean
    bTrueSub = mvarbTrueSub
End Property

Public Property Let sMenuHelp(ByVal vData As String)
    mvarsMenuHelp = vData
End Property

Public Property Get sMenuHelp() As String
    sMenuHelp = mvarsMenuHelp
End Property

Public Property Let bAsMark(ByVal vData As Boolean)
    mvarbAsMark = vData
End Property

Public Property Get bAsMark() As Boolean
    bAsMark = mvarbAsMark
End Property

Public Property Let bAsCheck(ByVal vData As Boolean)
    mvarbAsCheck = vData
End Property

Public Property Get bAsCheck() As Boolean
    bAsCheck = mvarbAsCheck
End Property

Public Property Let bMainMenu(ByVal vData As Boolean)
    mvarbMainMenu = vData
End Property

Public Property Get bMainMenu() As Boolean
    bMainMenu = mvarbMainMenu
End Property

Public Property Let iButton(ByVal vData As Integer)
Attribute iButton.VB_Description = "Indexe de l'image"
    mvariButton = vData
End Property

Public Property Get iButton() As Integer
    iButton = mvariButton
End Property

Public Property Let fType(ByVal vData As Long)
Attribute fType.VB_Description = "Type d'élément (item type) original"
    mvarfType = vData
End Property

Public Property Get fType() As Long
    fType = mvarfType
End Property

Public Property Let sMenuText(ByVal vData As String)
Attribute sMenuText.VB_Description = "Le texte du menu"
    mvarsMenuText = vData
End Property

Public Property Get sMenuText() As String
    sMenuText = mvarsMenuText
End Property
