Attribute VB_Name = "modCommon"
Option Explicit

Private Const FLASHW_TRAY = &H2
Private Const FLASHW_TIMERNOFG = &HC

Private Type FLASHWINFO
    cbSize As Long
    hWnd As Long
    dwFlage As Long
    uCount As Long
    dwTimeout As Long
End Type

Private Type HD_ITEM
   mask        As Long
   cxy         As Long
   pszText     As String
   hbm         As Long
   cchTextMax  As Long
   fmt         As Long
   lParam      As Long
   iImage      As Long
   iOrder      As Long
End Type

Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function FlashWindowEx Lib "user32" (ByRef pfwi As FLASHWINFO) As Long

Public Function GetRecord(ByVal strInfo As String, ByVal Num As Long) As String
    GetRecord = GetInfo(strInfo, Num, "|")
End Function
Public Function GetRecordCount(ByVal strInfo As String) As Long
    GetRecordCount = GetCount(strInfo, "|")
End Function
Public Function GetField(ByVal strInfo As String, ByVal Num As Long) As String
    GetField = GetInfo(strInfo, Num, vbCr)
End Function
Public Function GetFieldCount(ByVal strInfo As String) As Long
    GetFieldCount = GetCount(strInfo, vbCr)
End Function

Public Function GetInfo(ByVal strInfo As String, ByVal Num As Long, ByVal Sign As String) As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Length As Long

    On Error GoTo ErrHandler

    If Num < 1 Then
        GetInfo = ""
        Exit Function
    End If

    Length = Len(strInfo)
    i = 1: j = 0: k = 0
    Do While i <= Length
        j = InStr(i, strInfo, Sign)
        If j > 0 Then
            k = k + 1
            If k = Num Then
                GetInfo = Mid(strInfo, i, j - i)
                Exit Do
            End If
            i = j + 1
        Else
            GetInfo = Mid(strInfo, i, Length - i + 1)
            Exit Do
        End If
    Loop

    If k + 1 < Num Then GetInfo = ""

    Exit Function

ErrHandler:
    GetInfo = ""
End Function

Public Function GetCount(ByVal strInfo As String, ByVal Sign As String) As Long
    Dim i As Long
    Dim j As Long
    Dim Number As Long
    Dim Length As Long

    On Error GoTo ErrHandler

    Length = Len(strInfo)
    i = 1: Number = 1
    Do While i <= Length
        j = InStr(i, strInfo, Sign)
        If j > 0 Then
            Number = Number + 1
        Else
            Exit Do
        End If
        i = j + 1
    Loop

    GetCount = Number

    Exit Function

ErrHandler:
    GetCount = 0
End Function

Public Function DetectSoundCard() As Boolean
    If waveOutGetNumDevs() > 0 Then
        DetectSoundCard = True
    Else
        DetectSoundCard = False
    End If
End Function

Public Sub BeginPlaySound(ByVal ResourceId As String)
    On Error Resume Next

    If gblnSndCard Then
        Call EndPlaySound
        SoundBuffer = LoadResData(ResourceId, "Wave")
        Call sndPlaySound(SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    End If
End Sub
Public Sub EndPlaySound()
    Call sndPlaySound(ByVal vbNullString, 0&)
End Sub

Public Function BoolToString(ByVal Value As Boolean) As String
    BoolToString = CStr(Abs(Value))
End Function

Public Function TrimPath(ByVal Path As String) As String
    On Error Resume Next

    If Right(Path, 1) = "\" Then
        TrimPath = Path
    Else
        TrimPath = Path & "\"
    End If
End Function

Public Function ExtractName(ByVal Name As String) As String
    On Error Resume Next

    ExtractName = Right(Name, Len(Name) - InStrRev(Name, "\"))
End Function

Public Function ExtractPath(ByVal Name As String) As String
    Dim Temp As String
    
    On Error Resume Next

    Temp = Left(Name, InStrRev(Name, "\"))
    If Temp <> "" Then
        ExtractPath = Left(Temp, Len(Temp) - 1)
        If ExtractPath Like "?:" Then ExtractPath = ExtractPath & "\"
    Else
        ExtractPath = Left(Name, InStrRev(Name, ":"))
    End If
End Function

Public Sub Swap(ByRef Var1 As Variant, ByRef Var2 As Variant)
    Dim Temp As Variant
    Temp = Var1
    Var1 = Var2
    Var2 = Temp
End Sub

Public Function ToUrlString(ByVal Text As String) As String
    Dim i As Long
    Dim Length As Long
    Dim Ret As String
    Dim Temp As String

    On Error GoTo ErrHandler

    Text = Trim(Text)
    Length = Len(Text)
    For i = 1 To Length
        Temp = CStr(Hex(Asc(Mid(Text, i, 1))))
        'Debug.Print i; ":"; temp,
        If Len(Temp) > 2 Then
            Ret = Ret & "%" & Left(Temp, 2)
            Ret = Ret & "%" & Right(Temp, 2)
        Else
            Ret = Ret & "%" & Temp
        End If
    Next i
    ToUrlString = Ret

    Exit Function

ErrHandler:
    ToUrlString = ""
End Function

Public Function CheckString(ByVal Text As String) As Boolean
    On Error Resume Next

    If InStr(1, Text, "&") > 0 _
       Or InStr(1, Text, "|") > 0 _
       Or InStr(1, Text, "'") > 0 _
       Or InStr(1, Text, Chr(34)) > 0 _
       Or InStr(1, Text, "[") > 0 _
       Or InStr(1, Text, "]") > 0 Then
        CheckString = False
    Else
        CheckString = True
    End If
End Function

Public Function FileExisted(ByVal FileName As String) As Boolean
    If Dir(FileName, vbArchive) = vbNullString Then
        FileExisted = False
    Else
        FileExisted = True
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' 子程序: ControlSize
'''
''' 描述:   比例缩放控件。
'''
''' 参数:   Controls      - 控件对象
'''         LeftPercent   - 左边距%
'''         TopPercent    - 上边距%
'''         WidthPercent  - 宽度%
'''         HeightPercent - 高度%
'''
''' 日期:   2002.6.17
'''
''' 作者:   赵畅
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ControlSize(ByRef Controls As Object, ByVal LeftPercent As Single, ByVal TopPercent As Single, ByVal WidthPercent As Single, ByVal HeightPercent As Single)
    On Error Resume Next

    Controls.Left = Controls.Parent.Width * LeftPercent
    Controls.Top = Controls.Parent.Height * TopPercent
    Controls.Width = Controls.Parent.Width * WidthPercent
    Controls.Height = Controls.Parent.Height * HeightPercent
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' 子程序: ControlPosition
'''
''' 描述:   比例缩放与定位控件。
'''
''' 参数:   Controls      - 控件对象
'''         Left          - 左边距
'''         Top           - 上边距
'''         WidthPercent  - 宽度%
'''         HeightPercent - 高度%
'''
''' 日期:   2002.6.17
'''
''' 作者:   赵畅
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ControlPosition(ByRef Controls As Object, ByVal Left As Single, ByVal Top As Single, ByVal WidthPercent As Single, ByVal HeightPercent As Single)
    On Error Resume Next

    If Left <> 0 Then Controls.Left = Left
    If Top <> 0 Then Controls.Top = Top
    If WidthPercent <> 0 Then Controls.Width = Controls.Parent.Width * WidthPercent
    If HeightPercent <> 0 Then Controls.Height = Controls.Parent.Height * HeightPercent
End Sub

Public Sub ColumnSize(ByRef Controls As Object, ByVal Index As Long, ByVal WidthPercent As Single)
    On Error Resume Next
    Controls.ColumnHeaders(Index).Width = Controls.Width * (WidthPercent / 100)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' 子程序: LinePosition
'''
''' 描述:   定位 Line 控件的位置。
'''
''' 参数:   LineControl - Line 控件对象
'''
''' 日期:   2002.6.17
'''
''' 作者:   赵畅
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LinePosition(ByRef LineControl As Line, ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single)
    On Error Resume Next

    With LineControl
        .X1 = Left
        .Y1 = Top
        .X2 = Left + Width
        .Y2 = Top + Height
    End With
End Sub

Public Function LoadUserInfo(ByRef UserInfo As tagUserInfo, ByVal Info As String) As Boolean
    On Error GoTo ErrorHandler

    UserInfo.Email = GetRecord(Info, 1)
    UserInfo.UserClass = GetRecord(Info, 2)
    UserInfo.Face = CLng(GetRecord(Info, 3))
    UserInfo.Name = GetRecord(Info, 4)
    UserInfo.Sex = CLng(GetRecord(Info, 5))
    UserInfo.Age = CLng(GetRecord(Info, 6))
    UserInfo.Country = GetRecord(Info, 7)
    UserInfo.State = GetRecord(Info, 8)
    UserInfo.City = GetRecord(Info, 9)
    UserInfo.Win = CLng(GetRecord(Info, 10))
    UserInfo.Lose = CLng(GetRecord(Info, 11))
    UserInfo.Draw = CLng(GetRecord(Info, 12))
    UserInfo.GameTimes = CLng(GetRecord(Info, 13))
    UserInfo.Score = CLng(GetRecord(Info, 14))

    LoadUserInfo = True

    Exit Function

ErrorHandler:
    LoadUserInfo = False
End Function

Public Function LoadTableInfo(ByRef TableInfo As tagTableInfo, ByVal Info As String) As Boolean
    On Error GoTo ErrorHandler

    TableInfo.TableName = GetRecord(Info, 1)
    TableInfo.Creator = GetRecord(Info, 2)
    TableInfo.CreatorName = GetRecord(Info, 3)
    TableInfo.Visitor = GetRecord(Info, 4)
    TableInfo.VisitorName = GetRecord(Info, 5)
    TableInfo.TableType = GetRecord(Info, 6)
    TableInfo.Timer = CLng(GetRecord(Info, 7))
    TableInfo.UpLevel = CBool(GetRecord(Info, 8))
    TableInfo.LastTime = CDate(GetRecord(Info, 9))
    TableInfo.ip = GetRecord(Info, 10)
    TableInfo.LANIP = GetRecord(Info, 11)
    TableInfo.Port = CLng(GetRecord(Info, 12))

    LoadTableInfo = True

    Exit Function

ErrorHandler:
    LoadTableInfo = False
End Function

' 缇 -> 像素
Public Function GetPixelX(ByVal Twips As Single) As Long
    On Error Resume Next
    GetPixelX = Twips \ Screen.TwipsPerPixelX
End Function
Public Function GetPixelY(ByVal Twips As Single) As Long
    On Error Resume Next
    GetPixelY = Twips \ Screen.TwipsPerPixelY
End Function

' 像素 -> 缇
Public Function GetTwipX(ByVal Pixels As Long) As Single
    GetTwipX = Screen.TwipsPerPixelX * Pixels
End Function
Public Function GetTwipY(ByVal Pixels As Long) As Single
    GetTwipY = Screen.TwipsPerPixelY * Pixels
End Function

Public Function StrLen(ByVal strText As String) As Long
    Dim i As Long
    Dim Length As Long
    Dim RealLen As Long

    On Error GoTo ErrHandler

    Length = Len(strText)
    For i = 1 To Length
        If ((Asc(Mid(strText, i, 1)) + 65536) And 65535) > 255 Then
        'If Len(CStr(Hex(Asc(Mid(strText, i, 1))))) > 2 Then
            RealLen = RealLen + 2
        Else
            RealLen = RealLen + 1
        End If
    Next i

    StrLen = RealLen

    Exit Function

ErrHandler:
    StrLen = 0
End Function

' 渐变色子程序
Public Sub Gradient(ByRef TheObject As Object, ByVal Redval&, ByVal Greenval&, ByVal Blueval&)
    Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$

    On Error Resume Next

    Step = (TheObject.Width / 63)
    FillTop = 0
    FillLeft = TheObject.Width - Step
    FillRight = FillLeft + Step
    FillBottom = TheObject.Height

    For Reps = 0 To 63

        TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF

        Redval = Redval - 4
        Greenval = Greenval - 4
        Blueval = Blueval - 4

        If Redval <= 0 Then Redval = 0
        If Greenval <= 0 Then Greenval = 0
        If Blueval <= 0 Then Blueval = 0

        FillLeft = FillLeft - Step
        FillRight = FillLeft + Step
    Next Reps
End Sub

Public Sub SetLabel(ByRef objControl As Label, ByVal strCaption As String, ByVal strToolTip As String, ByVal blnVisible As Boolean)
    On Error Resume Next

    With objControl
        .Caption = strCaption
        .ToolTipText = strToolTip
        .Visible = blnVisible
    End With
End Sub

Public Function ToPartner(ByVal Player As Long, Optional ByVal Max As Long = 2) As Integer
    ToPartner = Abs(Player - Max) + Max - 1
End Function

' 无错误设置控件焦点。
Public Sub SetControlFocus(ByRef objControl As Object)
    On Error Resume Next
    Call objControl.SetFocus
End Sub

Public Function LoadString(ByVal ID As Long) As String
    Dim strResource As String
    Dim strReturn As String
    Dim lngLength As Long
    Dim i As Long
    Dim j As Long

    On Error GoTo ErrHandler

    strResource = LoadResString(ID)
    lngLength = Len(strResource)

    For i = 1 To lngLength
        If Mid(strResource, i, 1) = "#" And lngLength - i > 1 Then
            If Mid(strResource, i + 1, 1) = "#" Then
                strReturn = strReturn & "#"
                i = i + 1
            Else
                j = Val("&H" & Mid(strResource, i + 1, 2))
                If j > 0 Then
                    strReturn = strReturn & Chr(j)
                    i = i + 2
                End If
            End If
        Else
            strReturn = strReturn & Mid(strResource, i, 1)
        End If
    Next i

    LoadString = strReturn

    Exit Function

ErrHandler:
    LoadString = ""
End Function

Public Sub CloseModal()
    On Error Resume Next

    If frmRegister.Visible Then Call frmRegister.Hide
    If frmEditPlayList.Visible Then Call frmEditPlayList.Hide
    If frmOption.Visible Then Call frmOption.Hide
    If frmAbout.Visible Then Call Unload(frmAbout)

    If frmCreateTable.Visible Then Call frmCreateTable.Hide
    If frmLogin.Visible Then Call frmLogin.Hide
End Sub

Public Sub ShowHeaderIcon(ByRef objListView As ListView, ByVal colNo As Long, ByVal imgIconNo As Long, ByVal showImage As Long)
    'Dim r As Long
    Dim hHeader As Long
    Dim HD As HD_ITEM

    On Error Resume Next

    'get a handle to the listview header component
    hHeader = SendMessageLong(objListView.hWnd, LVM_GETHEADER, 0, 0)
    'set up the required structure members
    With HD
      .mask = HDI_IMAGE Or HDI_FORMAT
      .fmt = HDF_LEFT Or HDF_STRING Or HDF_BITMAP_ON_RIGHT Or showImage
      .pszText = objListView.ColumnHeaders(colNo + 1).Text
      If showImage Then .iImage = imgIconNo
   End With     'modify the header
   Call SendMessageAny(hHeader, HDM_SETITEM, colNo, HD)
End Sub

Public Sub PlaySoundEffects(ByVal SoundNumber As Long, ByVal SoundValue As String)
    Dim Temp As String

    On Error Resume Next

    If SoundValue = "" Then Exit Sub

    If SoundValue = DEFAULT_SOUND Then
        Temp = GetRecord(LoadString(RES_DEFAULT_SOUND), SoundNumber)
        If Len(Temp) < 1 Then Exit Sub
        Call EndPlaySound
        If Left(Temp, 1) = "_" Then
            Call PlaySound(Right(Temp, Len(Temp) - 1), 0, SND_ASYNC Or SND_ALIAS)
        Else
            Call BeginPlaySound(Temp)
        End If
    Else
        Call gSoundEffects.mmStop
        Call gSoundEffects.mmOpen(SoundValue)
        Call gSoundEffects.mmPlay
    End If
End Sub

Public Sub AutoSelectText(ByRef TextControl As TextBox)
    TextControl.SelStart = 0
    TextControl.SelLength = Len(TextControl.Text)
End Sub

Public Function GetDisplayName(ByVal UserName As String, ByVal Name As String) As String
    GetDisplayName = IIf(Name = "", UserName, Name)
End Function

'Public Sub SetFormEnable(ByRef Form As Object, ByVal Enabled As Boolean)
'    Dim i As Long

'    On Error Resume Next

'    For i = 0 To Form.Controls.Count - 1
'        Form.Controls(i).Enabled = Enabled
'    Next i
'End Sub

Public Sub FlashWindow(ByVal hWnd As Long)
    Dim pfwi As FLASHWINFO

    On Error Resume Next

    pfwi.hWnd = hWnd
    pfwi.dwFlage = FLASHW_TRAY Or FLASHW_TIMERNOFG
    pfwi.uCount = 0
    pfwi.dwTimeout = 0
    pfwi.cbSize = Len(pfwi)

    Call FlashWindowEx(pfwi)
End Sub

Public Function GetTime(ByVal lngSecond As Long) As String
    GetTime = Format(Minute(TimeSerial(0, 0, lngSecond)), "0#") & ":" & Format(Second(TimeSerial(0, 0, lngSecond)), "0#")
End Function

Public Function ParseURL(ByVal URL As String, ByVal IsExtract As Boolean) As String
    Dim Temp As String
    Dim i As Long

    On Error GoTo ErrHandler

    If URL = "" Then Exit Function

    If IsExtract Then
        If LCase(Left(URL, 7)) = "http://" Then
            Temp = Mid(URL, 8)
        Else
            Temp = URL
        End If
        i = InStr(1, Temp, "/")
        If i > 0 Then
            Temp = Left(Temp, i - 1)
        End If
    Else
        If LCase(Left(URL, 7)) = "http://" Then
            Temp = URL
        Else
            Temp = "http://" & URL
        End If
        If Right(Temp, 1) = "/" And Len(Temp) > 0 Then
            Temp = Left(Temp, Len(Temp) - 1)
        End If
    End If

    ParseURL = Temp

    Exit Function

ErrHandler:
    ParseURL = ""
End Function
