VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mmedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' 多媒体播放类
' 作者: 赵畅
' 2002.10.20

Option Explicit

Private sAlias As String        ' Used internally to give an alias name to
                          ' the multimedia resource

Private sFilename As String     ' Holds the filename internally
Private nLength As Single       ' Holds the length of the filename
                          ' internally
Private nPosition As Single     ' Holds the current position internally
Private sStatus As String       ' Holds the current status as a string
Private bWait As Boolean        ' Determines if VB should wait until play
                        ' is complete before returning.

'------------ API DECLARATIONS -------------
'note that this is all one code line:
Private Declare Function mciSendString Lib "winmm.dll" _
   Alias "mciSendStringA" (ByVal lpstrCommand As String, _
   ByVal lpstrReturnString As String, ByVal uReturnLength As Long, _
   ByVal hwndCallback As Long) As Long

Public Function mmOpen(ByVal sTheFile As String) As Long
    ' Declare a variable to hold the value returned by mciSendString
    Dim nReturn As Long
    Dim Pos As Long
    Dim Temp As Long

    ' Declare a string variable to hold the file type
    Dim sType As String

    On Error Resume Next

    ' Opens the specified multimedia file, and closes any
    ' other that may be open
    If sAlias <> "" Then
        Call mmClose
    End If
    
    ' Determine the type of file from the file extension
    Temp = InStr(1, sTheFile, ".")
    If Temp > 0 Then
        Pos = Len(sTheFile) - Temp
    Else
        Pos = 0
    End If
    Select Case UCase(Right(sTheFile, Pos))
       Case "WAV", "WAVE"
          sType = "WaveAudio"
       Case "AVI"
          sType = "AviVideo"
       Case "MID", "MIDI", "RMI"
          sType = "Sequencer"
       Case "MP3", "MP2", "MP1", "WMA"
          sType = "MPegVideo"
       Case Else
          sType = "MPegVideo"
          ' If the file extension is not known then exit the subroutine
    End Select
    'Randomize
    'sAlias = Right$(sTheFile, 3) & Minute(Now) & Second(Now) & Int(1000 * Rnd + 1)
    sAlias = Right(sTheFile, 3) & Minute(Now)

    ' At this point there is no file open, and we have determined the
    ' file type. Now would be a good time to open the new file.
    ' Note: if the name contains a space we have to enclose it in quotes
    'If InStr(sTheFile, " ") Then sTheFile = Chr(34) & sTheFile & Chr(34)
    'Debug.Print sTheFile, sAlias
    nReturn = mciSendString("Open " & Chr(34) & sTheFile & Chr(34) & " ALIAS " & sAlias _
            & " TYPE " & sType & " wait", "", 0, 0)

    mmOpen = nReturn
End Function

Public Sub mmClose()
    ' Closes the currently opened multimedia file

    ' Declare a variable to hold the return value from the mciSendString
    ' command
    Dim nReturn As Long

    ' If there is no file currently open then exit the subroutine
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Close " & sAlias, "", 0, 0)
    sAlias = ""
    sFilename = ""
    
End Sub

Public Sub mmPause()
    ' Pause playback of the file

    ' Declare a variable to hold the return value from the mciSendString
    ' command
    Dim nReturn As Long
    
    ' If there is no file currently open then exit the subroutine
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Pause " & sAlias, "", 0, 0)

End Sub

Public Function mmPlay() As Long
    ' Plays the currently open file, from the current position

    ' Declare a variable to hold the return value from the mciSendString
    ' command
    Dim nReturn As Long
    
    On Error Resume Next

    ' If there is no file currently open, then exit the routine
    If sAlias = "" Then
        mmPlay = -1
        Exit Function
    End If

    ' Now play the file
    If bWait Then
        nReturn = mciSendString("Play " & sAlias & " wait", "", 0, 0)
    Else
        nReturn = mciSendString("Play " & sAlias, "", 0, 0)
    End If

    mmPlay = nReturn
End Function

Public Sub mmStop()
    ' Stop using a file totally, be it playing or whatever

    ' Declare a variable to hold the return value from mciSendString
    Dim nReturn As Long
    
    ' If there is no file currently open then exit the subroutine
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Stop " & sAlias, "", 0, 0)
    
End Sub

Public Sub mmSeek(ByVal nPosition As Single)
    ' Seeks to a specific position within the file

    ' Declare a variable to hold the return value from the mciSendString
    ' function
    Dim nReturn As Long
    
    ' If there is no file currently open then exit the subroutine
    If sAlias = "" Then Exit Sub

    nReturn = mciSendString("Seek " & sAlias & " to " & nPosition, "", 0, 0)

End Sub

Property Get FileName() As String
' Routine to return a value when the programmer asks the
' object for the value of its Filename property
    FileName = sFilename
End Property

Property Let FileName(ByVal sTheFile As String)
' Routine to set the value of the filename property, should the programmer
' wish to do so. This implies that the programmer actually wants to open
' a file as well so control is passed to the mmOpen routine
   Call mmOpen(sTheFile)
End Property

Property Get Wait() As Boolean
' Routine to return the value of the object's wait property.
   Wait = bWait
End Property

Property Let Wait(ByVal bWaitValue As Boolean)
' Routine to set the value of the object's wait property
   bWait = bWaitValue
End Property

Property Get Length() As Single
   ' Routine to return the length of the currently opened multimedia file

   ' Declare a variable to hold the return value from the mciSendString
   Dim nReturn As Long, nLength As Integer

   ' Declare a string to hold the returned length from the mci Status call
   Dim sLength As String * 255
    
   On Error Resume Next

   ' If there is no file open then return 0
   If sAlias = "" Then
      Length = 0
      Exit Property
   End If

  nReturn = mciSendString("Status " & sAlias & " length", sLength, 255, 0)
  nLength = InStr(sLength, Chr$(0))
  Length = Val(Left$(sLength, nLength - 1))
End Property

Property Let Position(ByVal nPosition As Single)
' Sets the Position property effectively by seeking
    Call mmSeek(nPosition)
End Property

Property Get Position() As Single
   ' Returns the current position in the file
    
   ' Declare a variable to hold the return value from mciSendString
   Dim nReturn As Integer, nLength As Integer
    
   ' Declare a variable to hold the position returned
   ' by the mci Status position command
   Dim sPosition As String * 255

   On Error Resume Next

   ' If there is no file currently opened then exit the subroutine
   If sAlias = "" Then Exit Property
    
   ' Get the position and return
   nReturn = mciSendString("Status " & sAlias & " position", sPosition, 255, 0)
   nLength = InStr(sPosition, Chr$(0))
   Position = Val(Left$(sPosition, nLength - 1))

End Property

Property Get Status() As String
   ' Returns the playback/record status of the current file

   ' Declare a variable to hold the return value from mciSendString
   Dim nReturn As Integer, nLength As Integer

   ' Declare a variable to hold the return string from mciSendString
   Dim sStatus As String * 255

   On Error Resume Next

   ' If there is no file currently opened, then exit the subroutine
   If sAlias = "" Then Exit Property

   nReturn = mciSendString("Status " & sAlias & " mode", sStatus, 255, 0)

   nLength = InStr(sStatus, Chr$(0))
   Status = Left$(sStatus, nLength - 1)

End Property
