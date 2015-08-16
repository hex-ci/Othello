Attribute VB_Name = "modInternet"
'***************************************************************************
'*
'* 模块名:  Internet 模块
'* 作者:    赵畅
'* 日期:    2002.9.23
'*
'* 描述:    设置/读取 IE 的联机或脱机状态
'*
'***************************************************************************

Option Explicit

Private Type INTERNET_CONNECTED_INFO
    dwConnectedState As Long
    dwFlags As Long
End Type

'Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer
Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Integer
'Private Declare Function FindFirstUrlCacheEntry Lib "wininet.dll" Alias "FindFirstUrlCacheEntryA" (ByVal lpszUrlSearchPattern As String, ByRef lpFirstCacheEntryInfo As Any, ByRef lpdwFirstCacheEntryInfoBufferSize As Long) As Long
'Private Declare Function FindNextUrlCacheEntry Lib "wininet.dll" Alias "FindNextUrlCacheEntryA" (ByVal hEnumHandle As Long, ByRef lpNextCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwNextCacheEntryInfoBufferSize As Long) As Integer
'Private Declare Function FindCloseUrlCache Lib "wininet.dll" Alias "FindCloseUrlCacheA" (ByVal hEnumHandle As Long) As Integer

Private Const INTERNET_OPTION_CONNECTED_STATE = 50
Private Const INTERNET_STATE_DISCONNECTED_BY_USER = &H10
Private Const ISO_FORCE_DISCONNECTED = &H1
Private Const INTERNET_STATE_CONNECTED = &H1
Private Const ERROR_NO_MORE_ITEMS = 259&


' Returns true if the global state is offline. Otherwise, false.
Public Function IsGlobalOffline(ByVal hInternet As Long) As Boolean
    Dim dwState As Long
    Dim dwSize As Long
    Dim fRet As Boolean

    dwSize = 4

    If InternetQueryOption(hInternet, INTERNET_OPTION_CONNECTED_STATE, dwState, dwSize) Then
        If dwState And INTERNET_STATE_DISCONNECTED_BY_USER Then
            fRet = True
        End If
    End If

    IsGlobalOffline = fRet
End Function

Public Sub SetGlobalOffline(ByVal hInternet As Long, ByVal fGoOffline As Boolean)
    Dim ci As INTERNET_CONNECTED_INFO

    'memset(&ci, 0, sizeof(ci));
    If fGoOffline Then
        ci.dwConnectedState = INTERNET_STATE_DISCONNECTED_BY_USER
        ci.dwFlags = ISO_FORCE_DISCONNECTED
    Else
        ci.dwConnectedState = INTERNET_STATE_CONNECTED
    End If

    Call InternetSetOption(hInternet, INTERNET_OPTION_CONNECTED_STATE, ci, LenB(ci))
End Sub

Public Function ServerCommand(ByRef objInetControl As Inet, ByRef blnInetState As Boolean, ByVal strUrl As String, ByRef strStatus As String, Optional ByRef strData As String = "", Optional ByVal blnQuiet As Boolean = False, Optional ByVal blnDisplayStatus As Boolean = False) As Boolean
    Dim strMessage As String
    Dim mbrMsgBox As VbMsgBoxResult
    Dim i As Long

    On Error GoTo ErrorHandler

Start:
    If objInetControl.StillExecuting Then
        strStatus = STATUS_BUSY
        strData = ""
        ServerCommand = False
        Exit Function
    End If

    If blnDisplayStatus Then
        Call CloseModal
        Call frmProgress.ShowEx
    End If

    objInetControl.AccessType = icUseDefault
    objInetControl.Proxy = ""
    If gblnSave_UseProxy And gstrSave_HttpProxyIP <> "" Then
        objInetControl.AccessType = icNamedProxy
        objInetControl.Proxy = gstrSave_HttpProxyIP & ":" & CStr(glngSave_HttpProxyPort)
    End If

    Do
        objInetControl.Parent.Enabled = False
        'Call objInetControl.Cancel
        blnInetState = True
        gblnMenuDisplay = True

        ' 如果在脱机状态则联机
        If IsGlobalOffline(objInetControl.hInternet) Then
            Call SetGlobalOffline(objInetControl.hInternet, False)
        End If

        For i = 1 To glngRetryTimes
            strMessage = ""
            strMessage = CStr(objInetControl.OpenURL(strUrl, icString))
            If blnInetState Or strMessage <> "" Then
                Call DeleteUrlCacheEntry(strUrl)
                Exit For
            End If
        Next i

        If blnQuiet Then Exit Do
        If Not blnInetState Then
            gblnMenuDisplay = False
            If blnDisplayStatus Then
                Call frmProgress.HideEx
            End If
            mbrMsgBox = MessageBox(objInetControl.Parent.hWnd, LoadString(101), vbRetryCancel Or vbCritical, LoadString(181))
            If mbrMsgBox <> vbRetry Then Exit Do
            If blnDisplayStatus Then
                Call CloseModal
                Call frmProgress.ShowEx
            End If
        End If
    Loop While Not blnInetState

    If blnDisplayStatus Then
        Call frmProgress.HideEx
    End If

    objInetControl.Parent.Enabled = True
    gblnMenuDisplay = False

    Debug.Print (strMessage)

    If blnInetState Then
        If Len(strMessage) > 0 Then
            strStatus = GetField(strMessage, 1)
        Else
            strStatus = ""
        End If

        If Len(strMessage) > 2 Then
            strData = Mid(strMessage, 3)
        Else
            strData = ""
        End If

        ServerCommand = True
    Else
        strStatus = STATUS_ERROR
        strData = ""
        ServerCommand = False
    End If

    Exit Function

ErrorHandler:
    If blnDisplayStatus Then
        Call frmProgress.HideEx
    End If
    If Not blnQuiet Then
        If MessageBox(objInetControl.Parent.hWnd, LoadString(101), vbRetryCancel Or vbCritical, LoadString(181)) = vbRetry Then
            GoTo Start
        End If
    End If
    objInetControl.Parent.Enabled = True
    gblnMenuDisplay = False
    strStatus = STATUS_ERROR
    strData = ""
    ServerCommand = False
End Function

Public Function ServerExecute(ByRef objInetControl As Inet, ByVal strUrl As String) As Boolean
    On Error GoTo ErrorHandler

    If objInetControl.StillExecuting Then
        ServerExecute = False
        Exit Function
    End If

    objInetControl.AccessType = icUseDefault
    objInetControl.Proxy = ""
    If gblnSave_UseProxy And gstrSave_HttpProxyIP <> "" Then
        objInetControl.AccessType = icNamedProxy
        objInetControl.Proxy = gstrSave_HttpProxyIP & ":" & CStr(glngSave_HttpProxyPort)
    End If

    'Call objInetControl.Cancel

    ' 如果在脱机状态则联机
    If IsGlobalOffline(objInetControl.hInternet) Then
        Call SetGlobalOffline(objInetControl.hInternet, False)
    End If

    objInetControl.Tag = strUrl
    Call objInetControl.Execute(strUrl, "GET")

    ServerExecute = True

    Exit Function

ErrorHandler:
    Call MessageBox(objInetControl.Parent.hWnd, LoadString(101), vbCritical, LoadString(181))
    ServerExecute = False
End Function

Public Function GetServerExecute(ByRef objInetControl As Inet, ByRef strStatus As String, Optional ByRef strData As String = "") As Boolean
    Dim Temp As String
    Dim strMessage As String

    On Error GoTo ErrorHandler

    '得到第一个大块。注意:指定 Byte 数组
    ' (icByteArray) 以获取二进制文件。
    Temp = objInetControl.GetChunk(1024, icString)

    Do While LenB(Temp) > 0
        strMessage = strMessage + Temp
        '得到下一大块。
        Temp = objInetControl.GetChunk(1024, icString)
    Loop

    Debug.Print (strMessage)

    If Len(strMessage) > 0 Then
        strStatus = GetField(strMessage, 1)
        GetServerExecute = True
    Else
        strStatus = ""
        GetServerExecute = False
    End If
    If Len(strMessage) > 2 Then
        strData = Mid(strMessage, 3)
    Else
        strData = ""
    End If

    Call DeleteUrlCacheEntry(objInetControl.Tag)

    Exit Function

ErrorHandler:
    Call DeleteUrlCacheEntry(objInetControl.Tag)
    strStatus = ""
    strData = ""
    GetServerExecute = False
End Function
