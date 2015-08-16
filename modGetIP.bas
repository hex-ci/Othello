Attribute VB_Name = "modGetIP"
Option Explicit

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
   hName As Long
   hAliases As Long
   hAddrType As Integer
   hLength As Integer
   hAddrList As Long
End Type

Private Type WSADATA
   wversion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADescription_Len) As Byte
   szSystemStatus(0 To WSASYS_Status_Len) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal HostName As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Private Function SocketsInitialize() As Boolean
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    'Dim sLowByte As String, sHighByte As String, sMsg As String

    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

    If iReturn <> 0 Then
        SocketsInitialize = False
        Exit Function
    End If

    If LoByte(WSAD.wversion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wversion) = WS_VERSION_MAJOR And HiByte(WSAD.wversion) < WS_VERSION_MINOR) Then
        SocketsInitialize = False
        Exit Function
    End If

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        SocketsInitialize = False
        Exit Function
    End If

    SocketsInitialize = True
End Function

Private Function SocketsCleanup() As Boolean
    Dim lReturn As Long

    lReturn = WSACleanup()

    If lReturn <> 0 Then
        SocketsCleanup = False
    Else
        SocketsCleanup = True
    End If
End Function

Public Function GetLocalIP(ByVal HostName As String) As String
    Dim Hostent_Addr As Long
    Dim Host As HOSTENT
    Dim HostIP_Addr As Long
    Dim Temp_IP_Address() As Byte
    Dim i As Integer
    Dim IP_Address As String

    On Error Resume Next

    GetLocalIP = ""
    If Not SocketsInitialize() Then Exit Function

    Hostent_Addr = gethostbyname(HostName)

    If Not SocketsCleanup() Then Exit Function

    If Hostent_Addr = 0 Then Exit Function

    Call RtlMoveMemory(Host, Hostent_Addr, LenB(Host))
    Call RtlMoveMemory(HostIP_Addr, Host.hAddrList, 4)

    ReDim Temp_IP_Address(1 To Host.hLength)
    Call RtlMoveMemory(Temp_IP_Address(1), HostIP_Addr, Host.hLength)

    For i = 1 To Host.hLength
       IP_Address = IP_Address & Temp_IP_Address(i) & "."
    Next
    IP_Address = Mid(IP_Address, 1, Len(IP_Address) - 1)

    GetLocalIP = IP_Address
End Function
