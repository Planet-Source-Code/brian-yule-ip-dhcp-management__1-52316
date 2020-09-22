Attribute VB_Name = "modWinsock"
Option Explicit

Public Const SOCKET_ERROR = -1

Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Public Enum WinsockVersion
    SOCKET_VERSION_11 = &H101
    SOCKET_VERSION_22 = &H202
End Enum

Public Const hostent_size = 16
Public Const sockaddr_size = 16

Public Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    imaxsockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Type SERVENT
    s_name    As Long
    s_aliases As Long
    s_port    As Integer
    s_proto   As Long
End Type
Public Enum AddressInterface
    INADDR_ANY = &H0
    INADDR_LOOPBACK = &H7F000001
    INADDR_BROADCAST = &HFFFFFFFF
    INADDR_NONE = &HFFFFFFFF
End Enum
Public Declare Function getservbyport Lib "ws2_32.dll" (ByVal port As Long, ByVal proto As Long) As Long
Public Declare Function getservbyname Lib "ws2_32.dll" (ByVal serv_name As String, ByVal proto As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long

Public Function GetPortNumber(lngNBO As Long) As Long
    GetPortNumber = ntohs(lngNBO)
End Function

Public Function GetPort(ByVal port As Long) As String
' Convert network port number to either the friendly port name or the application port number
    Call StartWinsock
    Dim portData As SERVENT
    Dim lpData As Long
    Dim lLen As Long
    Dim strOut As String

    lpData = getservbyport(port, 0&)
    If (lpData) Then
        CopyMemory portData, ByVal lpData, Len(portData)
        lLen = lstrlen(portData.s_name)
        strOut = Space$(lLen \ 2 + 1)
        CopyMemory ByVal StrPtr(strOut), ByVal portData.s_name, lLen
        strOut = StrConv(strOut, vbUnicode)
        GetPort = UCase(RTrim$(Replace$(strOut, Chr$(0), vbNullString)))
    Else
        GetPort = CStr(ntohs(port))
    End If
    Call EndWinsock
End Function

Public Function getascip(ByVal inn As Long) As String
' Convert a network address to a application IP address
    Call StartWinsock
1:  Dim lpStr&, nStr&, retString$
2:  getascip = "255.255.255.255"
3:  retString = String(32, 0)
4:  lpStr = inet_ntoa(inn)
5:  If lpStr <> 0 Then
6:      nStr = lstrlen(lpStr)
7:      If nStr > 32 Then nStr = 32
8:      CopyMemory ByVal retString, ByVal lpStr, nStr
9:      retString = Left(retString, nStr)
10:     getascip = retString
11: End If
    Call EndWinsock
End Function

Public Function GetHostByNameAlias(ByVal hostname As String) As Long
' Get the network long address from the application IP address
    Call StartWinsock
    GetHostByNameAlias = inet_addr(hostname)
    Call EndWinsock
End Function

Private Function StartWinsock() As Boolean
' Start Winsock
2:  Dim StartupData As WSADataType, RetVal As Long
3:  RetVal = WSAStartup(WinsockVersion.SOCKET_VERSION_22, StartupData)
End Function

Private Function EndWinsock() As Long
' Stop Winsock
2:  Dim RetVal As Long
3:  RetVal = WSACleanup
End Function

