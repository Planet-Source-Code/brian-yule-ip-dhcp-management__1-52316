VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetworkTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0

Private Const API_TRUE = 1
Private Const API_FALSE = 0

Public Enum MAXLEN
    MAXLEN_IFDESCR = 256
    MAXLEN_PHYSADDR = 8
End Enum

Public Enum IPFORWARD_TYPES
    other = 1
    invalid
    Local_Route
    Remote_Route
End Enum

Public Enum ARP_TYPES
    arpOther = 1
    arpInvalid
    arpDynamic
    arpStatic
End Enum

Private Const IPX_PROTOCOL_BASE = &H1FFFF

Public Enum PROTOCOLS
    IP_OTHER = 1
    IP_LOCAL = 2
    IP_NETMGMT = 3
    IP_ICMP = 4
    IP_EGP = 5
    IP_GGP = 6
    IP_HELLO = 7
    IP_RIP = 8
    IP_IS_IS = 9
    IP_ES_IS = 10
    IP_CISCO = 11
    IP_BBN = 12
    IP_OSPF = 13
    IP_BGP = 14
    IP_IDPR = 15
    
    IP_BOOTP = 9999
    IP_NT_AUTOSTATIC = 10002
    IP_NT_STATIC = 10006
    IP_NT_STATIC_NON_DOD = 10007
    
    IPX_PROTOCOL_RIP = IPX_PROTOCOL_BASE + 1
    IPX_PROTOCOL_SAP = IPX_PROTOCOL_BASE + 2
    IPX_PROTOCOL_NLSP = IPX_PROTOCOL_BASE + 3
End Enum

Private Const MAX_INTERFACE_NAME_LEN  As Long = 256

Public Type MIB_UDPROW
    dwLocalAddr As Long
    dwLocalPort As Long
End Type

Public Type MIB_TCPROW_EX
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
    dwProcessId As Long
End Type

Public Type MIB_UDPROW_EX
    dwLocalAddr As Long
    dwLocalPort As Long
    dwProcessId As Long
End Type

Public Type MIB_IFROW
    wszName(0 To (MAX_INTERFACE_NAME_LEN * 2) - 1) As Byte
    dwIndex              As Long
    dwType               As Long
    dwMtu                As Long
    dwSpeed              As Long
    dwPhysAddrLen        As Long
    bPhysAddr(0 To MAXLEN_PHYSADDR - 1) As Byte
    dwAdminStatus        As Long
    dwOperStatus         As Long
    dwLastChange         As Long
    dwInOctets           As Long
    dwInUcastPkts        As Long
    dwInNUcastPkts       As Long
    dwInDiscards         As Long
    dwInErrors           As Long
    dwInUnknownProtos    As Long
    dwOutOctets          As Long
    dwOutUcastPkts       As Long
    dwOutNUcastPkts      As Long
    dwOutDiscards        As Long
    dwOutErrors          As Long
    dwOutQLen            As Long
    dwDescrLen           As Long
    bDescr(0 To MAXLEN_IFDESCR - 1) As Byte
End Type

Public Enum MIB_TCP_STATE
    MIB_TCP_STATE_CLOSED = 1
    MIB_TCP_STATE_LISTEN
    MIB_TCP_STATE_SYN_SENT
    MIB_TCP_STATE_SYN_RCVD
    MIB_TCP_STATE_ESTAB
    MIB_TCP_STATE_FIN_WAIT1
    MIB_TCP_STATE_FIN_WAIT2
    MIB_TCP_STATE_CLOSE_WAIT
    MIB_TCP_STATE_CLOSING
    MIB_TCP_STATE_LAST_ACK
    MIB_TCP_STATE_TIME_WAIT
    MIB_TCP_STATE_DELETE_TCB
End Enum

Public Enum MIB_IF_OPER_STATUS
    MIB_IF_OPER_STATUS_NON_OPERATIONAL = 0
    MIB_IF_OPER_STATUS_UNREACHABLE = 1
    MIB_IF_OPER_STATUS_DISCONNECTED = 2
    MIB_IF_OPER_STATUS_CONNECTING = 3
    MIB_IF_OPER_STATUS_CONNECTED = 4
    MIB_IF_OPER_STATUS_OPERATIONAL = 5
End Enum

Public Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Public Type MIB_IPADDRROW
    dwAddr As Long
    dwIndex As Long
    dwMask As Long
    dwBCastAddr As Long
    dwReasmSize As Long
    unused1 As Integer
    unused2 As Integer
End Type

Public Type MIB_IPFORWARDROW
    dwForwardDest As Long
    dwForwardMask As Long
    dwForwardPolicy As Long
    dwForwardNextHop As Long
    dwForwardIfIndex As Long
    dwForwardType As Long
    dwForwardProto As Long
    dwForwardAge As Long
    dwForwardNextHopAS As Long
    dwForwardMetric1 As Long
    dwForwardMetric2 As Long
    dwForwardMetric3 As Long
    dwForwardMetric4 As Long
    dwForwardMetric5 As Long
End Type

Public Type MIB_IPNETROW
    dwIndex As Long
    dwPhysAddrLen As Long
    bPhysAddr(0 To MAXLEN_PHYSADDR - 1) As Byte
    dwAddr As Long
    dwType As Long
End Type

Private Const ADAPTER_ADDRESS_DELIMITER = "-"
Private Const STRING_ZERO = "0"

Public Property Get ExtendedProperties() As Boolean
    Dim hlib As Long
    ' Load the library(Open the file)
    hlib = LoadLibrary("IPhlpAPI.dll")
    If (hlib) Then ' If there is a handle to the library upon loading then
        ' If the function exists then
        If GetProcAddress(hlib, "AllocateAndGetTcpExTableFromStack") Then
            ' Return true because it does exist, therefore extended properties are available
            ExtendedProperties = True
        End If
    End If
End Property

Public Property Get TcpTableEx() As MIB_TCPROW_EX()
    Dim lngRet As Long
    Dim mTcpTable As MIB_TCPTABLE_EX
    Dim lPtr As Long
    Dim ArrSize As Long
    
    If ExtendedProperties Then ' If the extended properties are available then
        ' Return to the pointer the data for the TCP table
        lngRet = AllocateAndGetTcpExTableFromStack(lPtr, API_TRUE, GetProcessHeap, 0, 2)
        ' Copy from the number of entries from the pointer to the dwNumEntries property of the TCP Table struct.
        Call CopyMemory(mTcpTable, ByVal lPtr, 4)
            
        If mTcpTable.dwNumEntries > 0 Then ' If there is at least one entry then
            ' Resize the table to take the rows
            ReDim mTcpTable.table(0 To mTcpTable.dwNumEntries - 1) As MIB_TCPROW_EX
            ' Calculate the size of the data to copy
            ArrSize = mTcpTable.dwNumEntries * Len(mTcpTable.table(0))
            ' Copy the data from the pointer to the array of TCP table rows
            Call CopyMemory(mTcpTable.table(0), ByVal (lPtr + 4), ArrSize)
        End If
        ' Return the tables rows
        TcpTableEx = mTcpTable.table
    End If
End Property

Public Property Get TcpTable() As MIB_TCPROW()
    Dim lngRet As Long, tableSize As Long
    Dim mTcpTable As MIB_TCPTABLE
    Dim byteArr() As Byte
    Dim ArrSize As Long
    
    ' Get the TCP Table, this will fail because the buffer is too small, it will however return the buffer size required
    lngRet = GetTcpTable(ByVal 0&, tableSize, API_TRUE)

    'If the buffer size required is 0 then there are no entries
    If tableSize > 0 Then
        ' Redimention the array to the required buffer size
        ReDim byteArr(0 To tableSize - 1) As Byte
        ' Get the TCP Table, this will succeed because the correct buffer size is supplied.
        lngRet = GetTcpTable(ByVal VarPtr(byteArr(0)), tableSize, API_TRUE)
        ' Copy the number of entries to the MIB_TCPTABLE struct
        Call CopyMemory(ByVal VarPtr(mTcpTable.dwNumEntries), ByVal VarPtr(byteArr(0)), Len(mTcpTable.dwNumEntries))
        ' If there are at least one entry then
        If mTcpTable.dwNumEntries > 0 Then
            ' Resize the TCP table rows array
            ReDim mTcpTable.table(0 To mTcpTable.dwNumEntries - 1) As MIB_TCPROW
            ' Calculate the amount of data to be copied
            ArrSize = mTcpTable.dwNumEntries * Len(mTcpTable.table(0))
            ' Copy the data to the array of table rows
            Call CopyMemory(ByVal VarPtr(mTcpTable.table(0)), ByVal VarPtr(byteArr(4)), ArrSize)
        End If
        ' Erase the buffer
        Erase byteArr
    End If
    ' Return the tables rows
    TcpTable = mTcpTable.table
End Property

Public Function DeleteTCB(ByVal ItemId As Long) As Long
    ' This kills a connection on the TCP table row specified.
    ' This will not always succeed as we cannot kill connections, just some in different states.
    
    Dim tmpMIB_TCPRow() As MIB_TCPROW
    ' Get the rows in the TCP table
    tmpMIB_TCPRow = TcpTable
    ' Change the state of the table row with which is specified
    tmpMIB_TCPRow(ItemId).dwState = MIB_TCP_STATE.MIB_TCP_STATE_DELETE_TCB
    ' Set the rows information
    DeleteTCB = SetTcpEntry(tmpMIB_TCPRow(ItemId))
End Function

Public Property Get IpAddressTable() As MIB_IPADDRROW()
    Dim lngRet As Long, tableSize As Long
    Dim mIpAddrTable As MIB_IPADDRTABLE
    Dim byteArr() As Byte
    Dim ArrSize As Long
    
    lngRet = GetIpAddrTable(ByVal 0&, tableSize, API_TRUE)
    If tableSize > 0 Then
        ReDim byteArr(0 To tableSize - 1) As Byte
        lngRet = GetIpAddrTable(ByVal VarPtr(byteArr(0)), tableSize, API_TRUE)
        
        Call CopyMemory(ByVal VarPtr(mIpAddrTable.dwNumEntries), ByVal VarPtr(byteArr(0)), Len(mIpAddrTable.dwNumEntries))
        If mIpAddrTable.dwNumEntries > 0 Then
            ReDim mIpAddrTable.table(0 To mIpAddrTable.dwNumEntries - 1) As MIB_IPADDRROW
            ArrSize = mIpAddrTable.dwNumEntries * Len(mIpAddrTable.table(0))
            Call CopyMemory(ByVal VarPtr(mIpAddrTable.table(0)), ByVal VarPtr(byteArr(4)), ArrSize)
        End If
        Erase byteArr
    End If
    IpAddressTable = mIpAddrTable.table
End Property

Public Property Get IpForwardTable() As MIB_IPFORWARDROW()
    Dim lngRet As Long, tableSize As Long
    Dim mIpFwdTable As MIB_IPFORWARDTABLE
    Dim byteArr() As Byte
    Dim ArrSize As Long
    
    lngRet = GetIpForwardTable(ByVal 0&, tableSize, API_TRUE)
    If tableSize > 0 Then
        ReDim byteArr(0 To tableSize - 1) As Byte
        lngRet = GetIpForwardTable(ByVal VarPtr(byteArr(0)), tableSize, API_TRUE)
        
        Call CopyMemory(ByVal VarPtr(mIpFwdTable.dwNumEntries), ByVal VarPtr(byteArr(0)), Len(mIpFwdTable.dwNumEntries))
        If mIpFwdTable.dwNumEntries > 0 Then
            ReDim mIpFwdTable.table(0 To mIpFwdTable.dwNumEntries - 1) As MIB_IPFORWARDROW
            ArrSize = mIpFwdTable.dwNumEntries * Len(mIpFwdTable.table(0))
            Call CopyMemory(ByVal VarPtr(mIpFwdTable.table(0)), ByVal VarPtr(byteArr(4)), ArrSize)
        End If
        Erase byteArr
    End If
    IpForwardTable = mIpFwdTable.table
End Property

Public Property Get IpNetTable() As MIB_IPNETROW()
    Dim lngRet As Long, tableSize As Long
    Dim mIpNetTable As MIB_IPNETTABLE
    Dim byteArr() As Byte
    Dim ArrSize As Long
    
    lngRet = GetIpNetTable(ByVal 0&, tableSize, API_TRUE)
    If tableSize > 0 Then
        ReDim byteArr(0 To tableSize - 1) As Byte
        lngRet = GetIpNetTable(ByVal VarPtr(byteArr(0)), tableSize, API_TRUE)
        
        Call CopyMemory(ByVal VarPtr(mIpNetTable.dwNumEntries), ByVal VarPtr(byteArr(0)), Len(mIpNetTable.dwNumEntries))
        If mIpNetTable.dwNumEntries > 0 Then
            ReDim mIpNetTable.table(0 To mIpNetTable.dwNumEntries - 1) As MIB_IPNETROW
            ArrSize = mIpNetTable.dwNumEntries * Len(mIpNetTable.table(0))
            Call CopyMemory(ByVal VarPtr(mIpNetTable.table(0)), ByVal VarPtr(byteArr(4)), ArrSize)
        End If
        Erase byteArr
    End If
    IpNetTable = mIpNetTable.table
End Property

Public Property Get IfTable() As MIB_IFROW()
    Dim lngRet As Long, tableSize As Long
    Dim mIfTable As MIB_IFTABLE
    Dim byteArr() As Byte
    Dim ArrSize As Long
    
    lngRet = GetIfTable(ByVal 0&, tableSize, API_TRUE)
    If tableSize > 0 Then
        ReDim byteArr(0 To tableSize - 1) As Byte
        lngRet = GetIfTable(ByVal VarPtr(byteArr(0)), tableSize, API_TRUE)
        
        Call CopyMemory(ByVal VarPtr(mIfTable.dwNumEntries), ByVal VarPtr(byteArr(0)), Len(mIfTable.dwNumEntries))
        If mIfTable.dwNumEntries > 0 Then
            ReDim mIfTable.table(0 To mIfTable.dwNumEntries - 1) As MIB_IFROW
            ArrSize = mIfTable.dwNumEntries * Len(mIfTable.table(0))
            Call CopyMemory(ByVal VarPtr(mIfTable.table(0)), ByVal VarPtr(byteArr(4)), ArrSize)
        End If
        Erase byteArr
    End If
    IfTable = mIfTable.table
End Property

Public Property Get UdpTable() As MIB_UDPROW()
    Dim lngRet As Long, tableSize As Long
    Dim mUdpTable As MIB_UDPTABLE
    Dim byteArr() As Byte
    Dim ArrSize As Long
    
    lngRet = GetIfTable(ByVal 0&, tableSize, API_TRUE)
    If tableSize > 0 Then
        ReDim byteArr(0 To tableSize - 1) As Byte
        lngRet = GetIfTable(ByVal VarPtr(byteArr(0)), tableSize, API_TRUE)
        
        Call CopyMemory(ByVal VarPtr(mUdpTable.dwNumEntries), ByVal VarPtr(byteArr(0)), Len(mUdpTable.dwNumEntries))
        If mUdpTable.dwNumEntries > 0 Then
            ReDim mUdpTable.table(0 To mUdpTable.dwNumEntries - 1) As MIB_UDPROW
            ArrSize = mUdpTable.dwNumEntries * Len(mUdpTable.table(0))
            Call CopyMemory(ByVal VarPtr(mUdpTable.table(0)), ByVal VarPtr(byteArr(4)), ArrSize)
        End If
        Erase byteArr
    End If
    UdpTable = mUdpTable.table
End Property

Public Function ConvAddress(ByRef Addr() As Byte, ByVal Length As Long) As String
' This converts a byte array to a MAC address.
    Dim strdef As String, i As Integer

    For i = 0 To Length - 1
        If i = 0 Then
            strdef = strdef & Pad(Hex(Addr(i)))
        Else
            strdef = strdef & ADAPTER_ADDRESS_DELIMITER & Pad(Hex(Addr(i)))
        End If
    Next i
    ConvAddress = strdef
End Function

Private Function Pad(ByVal hexStr As String) As String
' This pads a character with a preceding 0 if it requires one
    If Len(hexStr) = 1 Then
        Pad = STRING_ZERO & hexStr
    ElseIf Len(hexStr) = 2 Then
        Pad = hexStr
    End If
End Function
