VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Ini 文件类
' 作者: 赵畅
' 2003.3

Option Explicit

' --------
'  Public
' --------
'
' Property for file to read
Private File As String

' ---------
'  Private
' ---------
'
' API to read/write ini's
#If Win32 Then
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal FileName As String) As Integer
#Else
   Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal FileName As String) As Integer
#End If

Private Function SaveSetting(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Integer
    Dim retval As Integer

    SaveSetting = WritePrivateProfileString(Section, Key, Value, File)
End Function

Private Function TrimPath(ByVal Path As String) As String
    If Right(Path, 1) = "\" Then
        TrimPath = Path
    Else
        TrimPath = Path + "\"
    End If
End Function

Public Sub DeleteSection(ByVal Section As String)
    Dim retval As Integer

    retval = WritePrivateProfileString(Section, 0&, "", File)
End Sub

Public Function GetSetting(ByVal Section As String, ByVal KeyName As String, ByRef Value As String) As Boolean
    Dim retval As Integer
    Dim t As String * 255
    Dim Temp As Long

    On Error Resume Next

    ' Get the value
    retval = GetPrivateProfileString(Section, KeyName, "unknown value", t, Len(t), File)

    ' If there is one, return it
    If retval > 0 Then
        Temp = InStr(1, t, Chr(0), vbBinaryCompare)
        Value = Left(t, Temp - 1)
    Else
        'GetSetting = "unknown section or key"
        Value = ""
    End If
    If Value = "unknown value" Then
        GetSetting = False
    Else
        GetSetting = True
    End If
End Function

Public Function GetSection(ByVal Section As String, KeyArray() As String) As Integer
    Dim retval As Integer
    ' Allocate space for return value
    Dim t As String * 2500
    Dim lastpointer As Integer
    Dim nullpointer As Integer
    Dim ArrayCount As Integer
    Dim keystring As String

    ReDim KeyArray(0)

    On Error Resume Next

    ' Get the value
    retval = GetPrivateProfileString(Section, 0&, "", t, Len(t), File)

    ' If there is one, return it
    If retval > 0 Then
        '
        ' Separate the keys and store them in the array
        nullpointer = InStr(t, Chr$(0))
        lastpointer = 1
        Do While (nullpointer <> 0 And nullpointer > lastpointer + 1)
            '
            ' Extract key string
            keystring = Mid$(t, lastpointer, nullpointer - lastpointer)
            '
            ' Now add to array
            ArrayCount = ArrayCount + 1
            ReDim Preserve KeyArray(ArrayCount)
            KeyArray(ArrayCount) = keystring
            '
            ' Find next null
            lastpointer = nullpointer + 1
            nullpointer = InStr(nullpointer + 1, t, Chr$(0))
        Loop
    End If
    '
    ' Return the number of array elements
    GetSection = ArrayCount
End Function

Private Function GetValue(ByVal Section As String, ByVal Key As String, ByVal Default As String) As String
    Dim Temp As String

    If (Not GetSetting(Section, Key, Temp)) Or Temp = "" Then
        Call SaveSetting(Section, Key, Default)
        GetValue = Default
    Else
        GetValue = Temp
    End If
End Function

Public Sub SetFileName(ByVal Path As String)
    File = TrimPath(Path) + App.EXEName + ".INI"
End Sub

Public Function ReadString(ByVal Section As String, ByVal Key As String, ByVal Default As String) As String
    ReadString = GetValue(Section, Key, Default)
End Function

Public Function ReadBoolean(ByVal Section As String, ByVal Key As String, ByVal Default As Boolean) As Boolean
    ReadBoolean = CBool(GetValue(Section, Key, BoolToString(Default)))
End Function

Public Function ReadSingle(ByVal Section As String, ByVal Key As String, ByVal Default As Single) As Single
    ReadSingle = CSng(GetValue(Section, Key, CStr(Default)))
End Function

Public Function ReadInteger(ByVal Section As String, ByVal Key As String, ByVal Default As Integer) As Integer
    ReadInteger = CInt(GetValue(Section, Key, CStr(Default)))
End Function

Public Function ReadLong(ByVal Section As String, ByVal Key As String, ByVal Default As Long) As Long
    ReadLong = CLng(GetValue(Section, Key, CStr(Default)))
End Function

Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal Value As String)
    Call SaveSetting(Section, Key, Value)
End Sub

Public Sub WriteBoolean(ByVal Section As String, ByVal Key As String, ByVal Value As Boolean)
    Call SaveSetting(Section, Key, BoolToString(Value))
End Sub

Public Sub WriteSingle(ByVal Section As String, ByVal Key As String, ByVal Value As Single)
    Call SaveSetting(Section, Key, CStr(Value))
End Sub

Public Sub WriteInteger(ByVal Section As String, ByVal Key As String, ByVal Value As Integer)
    Call SaveSetting(Section, Key, CStr(Value))
End Sub

Public Sub WriteLong(ByVal Section As String, ByVal Key As String, ByVal Value As Long)
    Call SaveSetting(Section, Key, CStr(Value))
End Sub
