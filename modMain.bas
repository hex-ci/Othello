Attribute VB_Name = "MainModule"
'***************************************************************************
'*
'* 模块名:  黑白棋.Net 主模块
'* 作者:    赵畅
'* 日期:    2002.6.17
'*
'* 描述:    程序的主模块。
'*
'*
'* 子程序:
'*   Main
'*   Quit           - 退出程序
'*   LoadSetting    - 装载 INI 设置
'*   SaveSetting    - 保存 INI 设置
'*
'***************************************************************************

Option Explicit

Public Declare Function Think Lib "Othello.dll" (ByVal hWnd As Long, ByRef Board As Long, ByVal Level As Long, ByVal Chess As Long, ByVal Step As Long) As Long
Public Declare Sub StopThink Lib "Othello.dll" ()
'Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' 子程序: Main
'''
''' 描述:   程序的主函数，程序从这里开始。
'''
''' 参数:   没有
'''
''' 日期:   2002.6.17
'''
''' 作者:   赵畅
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Main()
    Dim i As Long

    On Error Resume Next

    ' 判断程序是否已经执行
    Call HaveSelf

    Screen.MousePointer = vbHourglass

    Call frmSplash.Show(vbModeless)
    Call frmSplash.Refresh

    ' 变量初始化开始
    TablePos.X = 0: TablePos.Y = 0
    Current.Col = -1: Current.Row = -1
    LastDown.Col = -1: LastDown.Row = -1
    'NormalColor = RGB(255, 221, 120)   ' 正常颜色
    'FocusColor = RGB(255, 232, 162)  ' 加亮颜色
    glngSave_GamePort = 19833
    gstrLocalPassword = "AukdEf364Fg985erFG"
    ' 服务器通信密钥，需要保密，也许需要定期更换
    gstrAzDGPrivateKey = "KAUgjby73%82#476khf#$g)(gemneit#$^&$GFEd"
    gstrAppPath = App.Path

    glngRetryTimes = 2
    gstrLocalIP = GetLocalIP("")

    Call LoadSetting

    For i = 1 To MAX_SOUND
        gstrSave_SoundName(i) = GetRecord(LoadString(116), i)
    Next i

    ' 从资源窗体的 Image 控件中装载图片。
    Call Load(frmResource)

    Set DefaultCursor = LoadResPicture("Default", vbResCursor)
    Set HandCursor = LoadResPicture("Hand", vbResCursor)
    Set BlackCursor = LoadResPicture("Black", vbResCursor)
    Set WhiteCursor = LoadResPicture("White", vbResCursor)
    Set HandUpCursor = LoadResPicture("HandUp", vbResCursor)
    Set HandDownCursor = LoadResPicture("HandDown", vbResCursor)

    ' 变量初始化结束

    gblnSndCard = DetectSoundCard()

    ' 设置帮助文件
    If FileExisted(gstrAppPath & "\Othello.chm") Then
        App.HelpFile = gstrAppPath & "\Othello.chm"
    End If

    Call MainForm.Show(vbModeless)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' 子程序: Quit
'''
''' 描述:   退出处理函数。
'''
''' 参数:   Force  - 是否强制退出。
'''
''' 日期:   2002.6.17
'''
''' 作者:   赵畅
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Quit()
    Dim i As Long
    Dim j As Object

    On Error Resume Next

    'If gblnOfflineMode Then Call StopThink

    Call SaveSetting

    'Call Unload(frmResource)
    'Set frmResource = Nothing

    For i = Forms.Count - 1 To 0 Step -1
        'If Forms(i).Name <> "MainForm" Then
        Set j = Forms(i)
        Call Unload(j)
        Set j = Nothing
        'End If
    Next i

    If Not (gBackgroundMusic Is Nothing) Then
        Call gBackgroundMusic.mmStop
        Call gBackgroundMusic.mmClose
        Set gBackgroundMusic = Nothing
    End If
    If Not (gSoundEffects Is Nothing) Then
        Call gSoundEffects.mmStop
        Call gSoundEffects.mmClose
        Set gSoundEffects = Nothing
    End If

    Set DefaultCursor = Nothing
    Set HandCursor = Nothing
    Set BlackCursor = Nothing
    Set WhiteCursor = Nothing
    Set HandUpCursor = Nothing
    Set HandDownCursor = Nothing

    Close

    'Call Unload(MainForm)
    'If Force Or (Forms.Count = 0) Then Close
    'If Force Or (Forms.Count > 0) Then End
End Sub

' 装载所有设置到全局配置变量
Public Sub LoadSetting()
    Dim Ini As clsIniFile
    Dim i As Long
    Dim j As Long
    Dim Temp As String

    On Error Resume Next

    Set Ini = New clsIniFile
    Call Ini.SetFileName(gstrAppPath)

    gwifSave_MainWindow.Center = Ini.ReadBoolean("Main", "Center", DEFAULT_gMainWindowCenter)
    gwifSave_MainWindow.Left = Ini.ReadSingle("Main", "Left", DEFAULT_gMainWindowLeft)
    gwifSave_MainWindow.Top = Ini.ReadSingle("Main", "Top", DEFAULT_gMainWindowTop)
    If gwifSave_MainWindow.Left + 9800 < 0 _
       Or gwifSave_MainWindow.Left > Screen.Width _
       Or gwifSave_MainWindow.Top + 6800 < 0 _
       Or gwifSave_MainWindow.Top > Screen.Height Then
        gwifSave_MainWindow.Center = True
    End If

    ' 在线用户窗口
    gwifSave_OnlineWindow.Left = Ini.ReadSingle("Online", "Left", DEFAULT_gOnlineWindowLeft)
    gwifSave_OnlineWindow.Top = Ini.ReadSingle("Online", "Top", DEFAULT_gOnlineWindowTop)
    gwifSave_OnlineWindow.Width = Ini.ReadSingle("Online", "Width", DEFAULT_gOnlineWindowWidth)
    gwifSave_OnlineWindow.Height = Ini.ReadSingle("Online", "Height", DEFAULT_gOnlineWindowHeight)
    gwifSave_OnlineWindow.Show = Ini.ReadBoolean("Online", "Show", DEFAULT_gOnlineWindowShow)
    gwifSave_OnlineWindow.DockStyle = Ini.ReadLong("Online", "DockStyle", DEFAULT_gOnlineWindowDockStyle)
    glngSave_OnlineSort = Ini.ReadLong("Online", "Sort", DEFAULT_gOnlineSort)
    glngSave_OnlineSortKey = Ini.ReadLong("Online", "SortKey", DEFAULT_gOnlineSortKey)
    gblnSave_OnlineAutoReload = Ini.ReadBoolean("Online", "AutoReload", DEFAULT_gOnlineAutoReload)
    glngSave_OnlineAutoReloadTime = Ini.ReadLong("Online", "ReloadTime", DEFAULT_gOnlineAutoReloadTime)
    If gwifSave_OnlineWindow.DockStyle <> dsDockNone Then
        gwifSave_OnlineWindow.DockPosition = Ini.ReadSingle("Online", "DockPosition", DEFAULT_gOnlineWindowDockPosition)
    End If
    For i = 1 To MAX_ONLINE_ITEM
        gsngSave_OnlineItemWidth(i) = Ini.ReadSingle("Online", "ItemWidth" & CStr(i), DEFAULT_gOnlineItemWidth)
    Next i

    ' 棋局列表窗口
    gwifSave_TableWindow.Left = Ini.ReadSingle("Table", "Left", DEFAULT_gTableWindowLeft)
    gwifSave_TableWindow.Top = Ini.ReadSingle("Table", "Top", DEFAULT_gTableWindowTop)
    gwifSave_TableWindow.Width = Ini.ReadSingle("Table", "Width", DEFAULT_gTableWindowWidth)
    gwifSave_TableWindow.Height = Ini.ReadSingle("Table", "Height", DEFAULT_gTableWindowHeight)
    gwifSave_TableWindow.Show = Ini.ReadBoolean("Table", "Show", DEFAULT_gTableWindowShow)
    gwifSave_TableWindow.DockStyle = Ini.ReadLong("Table", "DockStyle", DEFAULT_gTableWindowDockStyle)
    glngSave_TableSort = Ini.ReadLong("Table", "Sort", DEFAULT_gTableSort)
    glngSave_TableSortKey = Ini.ReadLong("Table", "SortKey", DEFAULT_gTableSortKey)
    gblnSave_TableAutoReload = Ini.ReadBoolean("Table", "AutoReload", DEFAULT_gTableAutoReload)
    glngSave_TableAutoReloadTime = Ini.ReadLong("Table", "ReloadTime", DEFAULT_gTableAutoReloadTime)
    If gwifSave_TableWindow.DockStyle <> dsDockNone Then
        gwifSave_TableWindow.DockPosition = Ini.ReadSingle("Table", "DockPosition", DEFAULT_gTableWindowDockPosition)
    End If
    For i = 1 To MAX_TABLE_ITEM
        gsngSave_TableItemWidth(i) = Ini.ReadSingle("Table", "ItemWidth" & CStr(i), DEFAULT_gTableItemWidth)
    Next i

    ' 聊天窗口
    gwifSave_ChatWindow.Left = Ini.ReadSingle("Chat", "Left", DEFAULT_gChatWindowLeft)
    gwifSave_ChatWindow.Top = Ini.ReadSingle("Chat", "Top", DEFAULT_gChatWindowTop)
    gwifSave_ChatWindow.Width = Ini.ReadSingle("Chat", "Width", DEFAULT_gChatWindowWidth)
    gwifSave_ChatWindow.Height = Ini.ReadSingle("Chat", "Height", DEFAULT_gChatWindowHeight)
    gwifSave_ChatWindow.Show = Ini.ReadBoolean("Chat", "Show", DEFAULT_gChatWindowShow)
    gwifSave_ChatWindow.DockStyle = Ini.ReadLong("Chat", "DockStyle", DEFAULT_gChatWindowDockStyle)
    If gwifSave_ChatWindow.DockStyle <> dsDockNone Then
        gwifSave_ChatWindow.DockPosition = Ini.ReadSingle("Chat", "DockPosition", DEFAULT_gChatWindowDockPosition)
    End If

    ' 公共聊天区
    gwifSave_PublicChatWindow.Left = Ini.ReadSingle("PublicChat", "Left", DEFAULT_gPublicChatWindowLeft)
    gwifSave_PublicChatWindow.Top = Ini.ReadSingle("PublicChat", "Top", DEFAULT_gPublicChatWindowTop)
    gwifSave_PublicChatWindow.Width = Ini.ReadSingle("PublicChat", "Width", DEFAULT_gPublicChatWindowWidth)
    gwifSave_PublicChatWindow.Height = Ini.ReadSingle("PublicChat", "Height", DEFAULT_gPublicChatWindowHeight)
    gwifSave_PublicChatWindow.Show = Ini.ReadBoolean("PublicChat", "Show", DEFAULT_gPublicChatWindowShow)
    glngSave_PublicChatWindowState = Ini.ReadLong("PublicChat", "Window", DEFAULT_gPublicChatWindowState)

    ' 其它窗口
    gptsSave_ViewUserInfo.X = Ini.ReadSingle("ViewUserInfo", "Left", DEFAULT_gViewUserInfoLeft)
    gptsSave_ViewUserInfo.Y = Ini.ReadSingle("ViewUserInfo", "Top", DEFAULT_gViewUserInfoTop)
    gptsSave_EditUserInfo.X = Ini.ReadSingle("EditUserInfo", "Left", DEFAULT_gEditUserInfoLeft)
    gptsSave_EditUserInfo.Y = Ini.ReadSingle("EditUserInfo", "Top", DEFAULT_gEditUserInfoTop)
    gptsSave_TableInfo.X = Ini.ReadSingle("TableInfo", "Left", DEFAULT_gTableInfoLeft)
    gptsSave_TableInfo.Y = Ini.ReadSingle("TableInfo", "Top", DEFAULT_gTableInfoTop)

    gblnSave_DownTip = Ini.ReadBoolean("Options", "DownTip", DEFAULT_gDownTip)
    glngSave_FaceNumber = Ini.ReadLong("Options", "Images", DEFAULT_gFaceNumber)

    gstrSave_ServerUrl = Ini.ReadString("Options", "Server", DEFAULT_gServerUrl)
    gstrSave_ServerUrl = ParseURL(gstrSave_ServerUrl, False)

    ' 读取播放列表
    gstrSave_PlayListPath = Ini.ReadString("PlayList", "Path", "")
    If gstrSave_PlayListPath = "" Then gstrSave_PlayListPath = gstrAppPath
    glngSave_PlayListNumber = Ini.ReadLong("PlayList", "Number", DEFAULT_gPlayListNumber)
    glngSave_PlayListPosition = Ini.ReadLong("PlayList", "Position", 0)
    If glngSave_PlayListNumber > MAX_PLAY_LIST Then glngSave_PlayListNumber = MAX_PLAY_LIST
    For i = 1 To glngSave_PlayListNumber
        gstrSave_PlayListName(i) = Ini.ReadString("PlayList", "Music" & CStr(i), "")
    Next i
    For i = 1 To glngSave_PlayListNumber - 1
        For j = i + 1 To glngSave_PlayListNumber
            If gstrSave_PlayListName(i) = "" And gstrSave_PlayListName(j) <> "" Then
                Call Swap(gstrSave_PlayListName(i), gstrSave_PlayListName(j))
                Exit For
            End If
        Next j
    Next i
    For i = 1 To glngSave_PlayListNumber
        If gstrSave_PlayListName(i) = "" Then Exit For
    Next i
    glngSave_PlayListNumber = i - 1
    'ReDim Preserve gstrSave_PlayListName(glngSave_PlayListNumber) As String
    If glngSave_PlayListPosition < 0 Or glngSave_PlayListPosition > glngSave_PlayListNumber Then
        glngSave_PlayListPosition = 0
    End If

    ' 读取声音列表
    gstrSave_SoundPath = Ini.ReadString("Sound", "Path", "")
    If gstrSave_SoundPath = "" Then gstrSave_SoundPath = gstrAppPath
    For i = 1 To MAX_SOUND
        gstrSave_SoundValue(i) = Ini.ReadString("Sound", "Sound" & CStr(i), "")
    Next i

    ' 读取代理服务器信息
    gblnSave_UseProxy = Ini.ReadBoolean("Proxy", "Use", DEFAULT_gUseProxy)
    gstrSave_HttpProxyIP = Ini.ReadString("Proxy", "HttpIP", "")
    glngSave_HttpProxyPort = Ini.ReadLong("Proxy", "HttpPort", 0)
    gstrSave_Socks5ProxyIP = Ini.ReadString("Proxy", "Socks5IP", "")
    glngSave_Socks5ProxyPort = Ini.ReadLong("Proxy", "Socks5Port", 0)
    gstrSave_Socks5Username = Ini.ReadString("Proxy", "Socks5Username", "")
    gstrSave_Socks5Password = Ini.ReadString("Proxy", "Socks5Password", "")


    ' 读取用户列表数据到全局配置变量
    For i = 1 To MAX_USER_LIST
        If Ini.GetSetting("UserHistory", "User" + CStr(i), Temp) And Temp <> "" Then
            gstrSave_UserList(i) = Temp
        End If
    Next i

    gblnSave_AutoLogin = Ini.ReadBoolean("Options", "AutoLogin", gblnSave_AutoLogin)
    gblnSave_SavePassword = Ini.ReadBoolean("Options", "SavePassword", gblnSave_SavePassword)
    If Ini.GetSetting("Options", "UserName", Temp) And Temp <> "" Then
        gstrSave_UserName = Temp
    End If
    If Ini.GetSetting("Options", "Password", Temp) And Temp <> "" Then
        gstrSave_Password = Temp
    End If

    glngSave_TableType = Ini.ReadLong("Options", "TableType", DEFAULT_gTableType)
    glngSave_TableTimer = Ini.ReadLong("Options", "TableTimer", DEFAULT_gTableTimer)
    glngSave_TableUpLevel = Ini.ReadLong("Options", "TableUpLevel", DEFAULT_gTableUpLevel)

    glngSave_OptionPage = Ini.ReadLong("Options", "OptionPage", DEFAULT_gOptionPage)

    ' 单机模式
    gblnSave_OfflineMode = Ini.ReadBoolean("Options", "OfflineMode", DEFAULT_gOfflineMode)
    gblnOfflineMode = gblnSave_OfflineMode
    glngSave_Level = Ini.ReadLong("Options", "Level", DEFAULT_gLevel)
    glngSave_OfflineFace = Ini.ReadLong("Options", "Face", DEFAULT_gOfflineFace)

    ' 调试使用
    glngRetryTimes = Ini.ReadLong("Options", "RetryTimes", 2)

    Set Ini = Nothing
End Sub

' 保存所有设置到全局配置变量
Public Sub SaveSetting()
    Dim Ini As clsIniFile
    Dim i As Long

    On Error Resume Next

    Set Ini = New clsIniFile
    Call Ini.SetFileName(gstrAppPath)

    If gwifSave_MainWindow.Center And Abs(MainForm.Left - (Screen.Width - MainForm.Width) \ 2) < 100 And Abs(MainForm.Top - (Screen.Height * 0.9 - MainForm.Height) \ 2) < 100 Then
        Call Ini.WriteBoolean("Main", "Center", True)
    Else
        Call Ini.WriteBoolean("Main", "Center", False)
    End If
    Call Ini.WriteSingle("Main", "Left", MainForm.Left)
    Call Ini.WriteSingle("Main", "Top", MainForm.Top)

    ' 在线用户窗口
    Call Ini.WriteSingle("Online", "Left", frmOnline.Left)
    Call Ini.WriteSingle("Online", "Top", frmOnline.Top)
    Call Ini.WriteSingle("Online", "Width", frmOnline.Width)
    Call Ini.WriteSingle("Online", "Height", frmOnline.Height)
    Call Ini.WriteBoolean("Online", "Show", frmOnline.Visible)
    Call Ini.WriteInteger("Online", "DockStyle", frmOnline.DockStyle)
    Call Ini.WriteSingle("Online", "DockPosition", frmOnline.DockPosition)
    Call Ini.WriteBoolean("Online", "AutoReload", gblnSave_OnlineAutoReload)
    Call Ini.WriteLong("Online", "ReloadTime", glngSave_OnlineAutoReloadTime)
    If frmOnline.lvwOnline.Sorted Then
        Call Ini.WriteInteger("Online", "Sort", frmOnline.lvwOnline.SortOrder + 1)
        Call Ini.WriteInteger("Online", "SortKey", frmOnline.lvwOnline.SortKey)
    Else
        Call Ini.WriteInteger("Online", "Sort", 0)
        Call Ini.WriteInteger("Online", "SortKey", 0)
    End If
    For i = 1 To MAX_ONLINE_ITEM
        Call Ini.WriteSingle("Online", "ItemWidth" & CStr(i), frmOnline.lvwOnline.ColumnHeaders(i).Width)
    Next i

    ' 棋局列表窗口
    Call Ini.WriteSingle("Table", "Left", frmTable.Left)
    Call Ini.WriteSingle("Table", "Top", frmTable.Top)
    Call Ini.WriteSingle("Table", "Width", frmTable.Width)
    Call Ini.WriteSingle("Table", "Height", frmTable.Height)
    Call Ini.WriteBoolean("Table", "Show", frmTable.Visible)
    Call Ini.WriteInteger("Table", "DockStyle", frmTable.DockStyle)
    Call Ini.WriteSingle("Table", "DockPosition", frmTable.DockPosition)
    Call Ini.WriteBoolean("Table", "AutoReload", gblnSave_TableAutoReload)
    Call Ini.WriteLong("Table", "ReloadTime", glngSave_TableAutoReloadTime)
    If frmTable.lvwTable.Sorted Then
        Call Ini.WriteInteger("Table", "Sort", frmTable.lvwTable.SortOrder + 1)
        Call Ini.WriteInteger("Table", "SortKey", frmTable.lvwTable.SortKey)
    Else
        Call Ini.WriteInteger("Table", "Sort", 0)
        Call Ini.WriteInteger("Table", "SortKey", 0)
    End If
    For i = 1 To MAX_TABLE_ITEM
        Call Ini.WriteSingle("Table", "ItemWidth" & CStr(i), frmTable.lvwTable.ColumnHeaders(i).Width)
    Next i

    ' 聊天窗口
    Call Ini.WriteSingle("Chat", "Left", frmChat.Left)
    Call Ini.WriteSingle("Chat", "Top", frmChat.Top)
    Call Ini.WriteSingle("Chat", "Width", frmChat.Width)
    Call Ini.WriteSingle("Chat", "Height", frmChat.Height)
    Call Ini.WriteBoolean("Chat", "Show", frmChat.Visible)
    Call Ini.WriteInteger("Chat", "DockStyle", frmChat.DockStyle)
    Call Ini.WriteSingle("Chat", "DockPosition", frmChat.DockPosition)

    ' 公共聊天区
    If frmPublicChat.WindowState = vbNormal Then
        Call Ini.WriteSingle("PublicChat", "Left", frmPublicChat.Left)
        Call Ini.WriteSingle("PublicChat", "Top", frmPublicChat.Top)
        Call Ini.WriteSingle("PublicChat", "Width", frmPublicChat.Width)
        Call Ini.WriteSingle("PublicChat", "Height", frmPublicChat.Height)
    End If
    Call Ini.WriteBoolean("PublicChat", "Show", frmPublicChat.Visible)
    If frmPublicChat.WindowState = vbMinimized Then
        Call Ini.WriteInteger("PublicChat", "Window", vbNormal)
    Else
        Call Ini.WriteInteger("PublicChat", "Window", frmPublicChat.WindowState)
    End If

    ' 其它窗口
    Call Ini.WriteSingle("ViewUserInfo", "Left", frmUserInfo.Left)
    Call Ini.WriteSingle("ViewUserInfo", "Top", frmUserInfo.Top)
    Call Ini.WriteSingle("EditUserInfo", "Left", frmEditInfo.Left)
    Call Ini.WriteSingle("EditUserInfo", "Top", frmEditInfo.Top)
    Call Ini.WriteSingle("TableInfo", "Left", frmTableInfo.Left)
    Call Ini.WriteSingle("TableInfo", "Top", frmTableInfo.Top)

    Call Ini.WriteString("Options", "Server", ParseURL(gstrSave_ServerUrl, True))

    Call Ini.WriteBoolean("Options", "DownTip", gblnSave_DownTip)
    'Call Ini.SaveSetting("Options", "Images", CStr(glngSave_FaceNumber))

    ' 先删除 Play List 区域
    Call Ini.DeleteSection("PlayList")
    Call Ini.WriteString("PlayList", "Path", gstrSave_PlayListPath)
    Call Ini.WriteLong("PlayList", "Number", glngSave_PlayListNumber)
    Call Ini.WriteLong("PlayList", "Position", glngSave_PlayListPosition)
    For i = 1 To glngSave_PlayListNumber
        Call Ini.WriteString("PlayList", "Music" + CStr(i), gstrSave_PlayListName(i))
    Next i

    ' 保存声音列表
    Call Ini.WriteString("Sound", "Path", gstrSave_SoundPath)
    For i = 1 To MAX_SOUND
        Call Ini.WriteString("Sound", "Sound" + CStr(i), gstrSave_SoundValue(i))
    Next i

    ' 保存代理服务器信息
    Call Ini.WriteBoolean("Proxy", "Use", gblnSave_UseProxy)
    Call Ini.WriteString("Proxy", "HttpIP", gstrSave_HttpProxyIP)
    Call Ini.WriteLong("Proxy", "HttpPort", glngSave_HttpProxyPort)
    Call Ini.WriteString("Proxy", "Socks5IP", gstrSave_Socks5ProxyIP)
    Call Ini.WriteLong("Proxy", "Socks5Port", glngSave_Socks5ProxyPort)
    Call Ini.WriteString("Proxy", "Socks5Username", gstrSave_Socks5Username)
    Call Ini.WriteString("Proxy", "Socks5Password", gstrSave_Socks5Password)

    ' 保存全局配置变量中的数据
    For i = 1 To MAX_USER_LIST
        Call Ini.WriteString("UserHistory", "User" + CStr(i), gstrSave_UserList(i))
    Next i

    Call Ini.WriteBoolean("Options", "AutoLogin", gblnSave_AutoLogin)
    Call Ini.WriteBoolean("Options", "SavePassword", gblnSave_SavePassword)
    Call Ini.WriteString("Options", "UserName", gstrSave_UserName)
    Call Ini.WriteString("Options", "Password", gstrSave_Password)

    Call Ini.WriteLong("Options", "TableType", glngSave_TableType)
    Call Ini.WriteLong("Options", "TableTimer", glngSave_TableTimer)
    Call Ini.WriteLong("Options", "TableUpLevel", glngSave_TableUpLevel)

    Call Ini.WriteLong("Options", "OptionPage", glngSave_OptionPage)

    ' 单机模式
    Call Ini.WriteBoolean("Options", "OfflineMode", gblnSave_OfflineMode)
    Call Ini.WriteLong("Options", "Level", glngSave_Level)
    Call Ini.WriteLong("Options", "Face", glngSave_OfflineFace)

    Set Ini = Nothing
End Sub

Public Function MakeServerPassword() As String
    Dim i As Long
    Dim Style(5) As Long
    Dim Temp As String
    Dim strSecurity1 As String
    Dim strSecurity2 As String

    On Error Resume Next

    ' 存储“style=”
    Style(0) = 115
    Style(1) = 116
    Style(2) = 121
    Style(3) = 108
    Style(4) = 101
    Style(5) = 61

    For i = 0 To UBound(Style)
        Temp = Temp & Chr(Style(i))
    Next i

    strSecurity1 = AzDG_decrypt(gstrSecurity1)
    strSecurity2 = AzDG_decrypt(gstrSecurity2)

    MakeServerPassword = Temp & ToUrlString(AzDG_crypt(strSecurity1, strSecurity2))
End Function

Public Function MakeVersion() As String
    MakeVersion = "ver=" & CStr(OTHELLO_VERSION)
End Function

Private Sub HaveSelf()
    On Error Resume Next

    If App.PrevInstance Then
        Dim Temp As String
        Dim Handle As Long

        Temp = App.Title
        App.Title = ""  '如此才不会Avtivate到自己

        Handle = FindWindow(vbNullString, Temp)
        If Handle <> 0 Then
            Call PostMessage(Handle, WM_ACTIVEWINDOW, 0, 0)
        End If
        'Call AppActivate(Temp)      'Activate先前就已执行的程序
        End
    End If
End Sub
