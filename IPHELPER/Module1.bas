Attribute VB_Name = "modIPHelp"
Option Explicit
Option Base 0

Public Const ERROR_SUCCESS       As Long = 0
Public Const ERROR_NOT_SUPPORTED As Long = 50
Public Const ERROR_BUFFER_OVERFLOW As Long = 111
Public Const ERROR_INSUFFICIENT_BUFFER = 122 '  dderror

Public Const MAX_ADAPTER_ADDRESS     As Long = 8
Public Const MAX_ADAPTER_DESCRIPTION As Long = 128
Public Const MAX_ADAPTER_NAME        As Long = 256
Public Const MAX_INTERFACE_NAME_LEN  As Long = 256
Public Const MAX_HOSTNAME_LEN    As Long = 128
Public Const MAX_DOMAIN_NAME_LEN As Long = 128
Public Const MAX_SCOPE_ID_LEN    As Long = 256
Public Const MAX_PATH = 260

Public Type IP_ADAPTER_INDEX_MAP
    Index As Long
    nname(0 To MAX_ADAPTER_DESCRIPTION - 1) As Byte
End Type

Public Type IP_INTERFACE_INFO
    NumAdapters As Long
    Adapter() As IP_ADAPTER_INDEX_MAP
End Type

Public Const MIB_TCP_MAXCONN_DYNAMIC = (-1&)


Public Enum MIB_IF_ADMIN_STATUS
    MIB_IF_ADMIN_STATUS_UP = 1
    MIB_IF_ADMIN_STATUS_DOWN = 2
    MIB_IF_ADMIN_STATUS_TESTING = 3
End Enum

Public Enum IP_PREFIX_ORIGIN
    IpPrefixOriginOther = 0
    IpPrefixOriginManual
    IpPrefixOriginWellKnown
    IpPrefixOriginDhcp
    IpPrefixOriginRouterAdvertisement
End Enum

Public Enum IF_OPER_STATUS
    IfOperStatusUp = 1
    IfOperStatusDown
    IfOperStatusTesting
    IfOperStatusUnknown
    IfOperStatusDormant
    IfOperStatusNotPresent
    IfOperStatusLowerLayerDown
End Enum

Public Enum IP_DAD_STATE
    IpDadStateInvalid = 0
    IpDadStateTentative
    IpDadStateDuplicate
    IpDadStateDeprecated
    IpDadStatePreferred
End Enum

Public Enum IP_SUFFIX_ORIGIN
    IpSuffixOriginOther = 0
    IpSuffixOriginManual
    IpSuffixOriginWellKnown
    IpSuffixOriginDhcp
    IpSuffixOriginLinkLayerAddress
    IpSuffixOriginRandom
End Enum

Public Enum SCOPE_LEVEL
    ScopeLevelInterface = 1
    ScopeLevelLink = 2
    ScopeLevelSubnet = 3
    ScopeLevelAdmin = 4
    ScopeLevelSite = 5
    ScopeLevelOrganization = 8
    ScopeLevelGlobal = 14
End Enum

Type IPADDR
    Octet(3) As Integer
End Type

Private Const MIB_IPROUTE_METRIC_UNUSED = -1

'#define MIB_IPPROTO_OTHER           1
'#define MIB_IPPROTO_LOCAL           2
'#define MIB_IPPROTO_NETMGMT         3
'#define MIB_IPPROTO_ICMP            4
'#define MIB_IPPROTO_EGP             5
'#define MIB_IPPROTO_GGP             6
'#define MIB_IPPROTO_HELLO           7
'#define MIB_IPPROTO_RIP             8
'#define MIB_IPPROTO_IS_IS           9
'#define MIB_IPPROTO_ES_IS           10
'#define MIB_IPPROTO_CISCO           11
'#define MIB_IPPROTO_BBN             12
'#define MIB_IPPROTO_OSPF            13
'#define MIB_IPPROTO_BGP             14

Public Type FIXED_INFO
   hostname(0 To (MAX_HOSTNAME_LEN + 3))      As Byte
   DomainName(0 To (MAX_DOMAIN_NAME_LEN + 3)) As Byte
   CurrentDnsServer   As Long ' Pointer to IP_ADDR_STRING
   DNSServerList      As IP_ADDR_STRING
   NodeType           As Long
   ScopeId(0 To (MAX_SCOPE_ID_LEN + 3))       As Byte
   EnableRouting      As Long
   EnableProxy        As Long
   EnableDns          As Long
End Type

Public Type MIB_TCPSTATS
   dwRtoAlgorithm   As Long  'time-out algorithm
   dwRtoMin         As Long  'minimum time-out
   dwRtoMax         As Long  'maximum time-out
   dwMaxConn        As Long  'maximum connections
   dwActiveOpens    As Long  'active opens
   dwPassiveOpens   As Long  'passive opens
   dwAttemptFails   As Long  'failed attempts
   dwEstabResets    As Long  'established connections reset
   dwCurrEstab      As Long  'established connections
   dwInSegs         As Long  'segments received
   dwOutSegs        As Long  'segment sent
   dwRetransSegs    As Long  'segments retransmitted
   dwInErrs         As Long  'incoming errors
   dwOutRsts        As Long  'outgoing resets
   dwNumConns       As Long  'cumulative connections
End Type

Type ARP_SEND_REPLY
    DestAddress As IPADDR
    SrcAddress As IPADDR
End Type

Type IP_ADAPTER_UNICAST_ADDRESS
    Length As Long
    Flags As Long
    dwNext     As Long
    Address As Long
    PrefixOrigin As IP_PREFIX_ORIGIN
    SuffixOrigin As IP_SUFFIX_ORIGIN
    DadState As IP_DAD_STATE
    ValidLifetime As Long
    PreferredLifetime As Long
    LeaseLifetime As Long
End Type

Type IP_ADAPTER_PREFIX
    Length As Long
    Flags As Long
    dwNext     As Long
    Address As Long
    PrefixLength As Long
End Type

Type IP_ADAPTER_ORDER_MAP
    NumAdapters As Long
    AdapterOrder(1) As Long
End Type

Type IP_ADAPTER_MULTICAST_ADDRESS
    Length As Long
    Flags As Long
    dwNext     As Long
    Address As Long
End Type

Type IP_ADAPTER_DNS_SERVER_ADDRESS
    Length As Long
    Reserved As Long
    dwNext     As Long
    Address As Long
End Type

Type IP_ADAPTER_ANYCAST_ADDRESS
    Length As Long
    Flags As Long
    dwNext     As Long
    Address As Long
End Type

Type IP_ADAPTER_ADDRESSES
    Length As Long
    IfIndex As Long
    dwNext     As Long
    AdapterName As Long
    FirstUnicastAddress As IP_ADAPTER_UNICAST_ADDRESS
    FirstAnycastAddress As IP_ADAPTER_ANYCAST_ADDRESS
    FirstMulticastAddress As IP_ADAPTER_MULTICAST_ADDRESS
    FirstDnsServerAddress As IP_ADAPTER_DNS_SERVER_ADDRESS
    DnsSuffix As Long
    Description As Long
    FriendlyName As Long
    PhysicalAddress(0 To MAX_ADAPTER_ADDRESS - 1) As Byte
    PhysicalAddressLength As Long
    Flags As Long
    Mtu As Long
    ifType As Long
    OperStatus As IF_OPER_STATUS
    Ipv6IfIndex As Long
    ZoneIndices(16) As Long
    FirstPrefix As IP_ADAPTER_PREFIX
End Type

Public Type IP_OPTION_INFORMATION ' The ip_option_information structure describes the options to be included in the header of an IP packet. The TTL, TOS, and Flags values are carried in specific fields in the header. The OptionsData bytes are carried in the options area following the standard IP header. With the exception of source route options, this data must be in the format to be transmitted on the wire as specified in RFC 791. A source route option should contain the full route - first hop thru final destination - in the route data. The first hop will be pulled out of the data and the option will be reformatted accordingly. Otherwise, the route option should be formatted as specified in RFC 791.
    Ttl         As Byte 'unsigned char       // Time To Live
    Tos         As Byte 'unsigned char       // Type Of Service
    Flags       As Byte 'unsigned char       // IP header flags
    OptionsSize As Byte 'unsigned char       // Size in bytes of options data
    OptionsData As Long 'unsigned char FAR * // Pointer to options data
End Type

Type ICMP_ECHO_REPLY
    Address As IPADDR
    status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    pData As Long
    Options As IP_OPTION_INFORMATION
End Type



Public Type GUID
    Data1          As Long
    Data2          As Integer
    Data3          As Integer
    Data4(0 To 7)  As String * 1
End Type

Type IP_INTERFACE_NAME_INFO
    Index As Long
    MediaType As Long
    ConnectionType As Byte
    AccessType As Byte
    DeviceGuid As GUID
    InterfaceGuid As GUID
End Type

Type IP_PER_ADAPTER_INFO
    AutoConfigEnabled As Integer
    AutoConfigActive As Integer
    CurrentDnsServer As IP_ADDR_STRING
    DNSServerList As IP_ADDR_STRING
End Type

Type IP_UNIDIRECTIONAL_ADAPTER_ADDRESS
    NumAdapters As Long
    Adapter(1) As IPADDR
End Type

Type TCP_RESERVE_PORT_RANGE
    UpperRange As Integer
    LowerRange As Integer
End Type



Public Type MIB_IPSTATS
    dwForwarding As Long
    dwDefaultTTL As Long
    dwInReceives As Long
    dwInHdrErrors As Long
    dwInAddrErrors As Long
    dwForwDatagrams As Long
    dwInUnknownProtos As Long
    dwInDiscards As Long
    dwInDelivers As Long
    dwOutRequests As Long
    dwRoutingDiscards As Long
    dwOutDiscards As Long
    dwOutNoRoutes As Long
    dwReasmTimeout As Long
    dwReasmReqds As Long
    dwReasmOks As Long
    dwReasmFails As Long
    dwFragOks As Long
    dwFragFails As Long
    dwFragCreates As Long
    dwNumIf As Long
    dwNumAddr As Long
    dwNumRoutes As Long
End Type

Public Type MIB_TCPTABLE
    dwNumEntries As Long
    table() As MIB_TCPROW
End Type

Public Type MIB_IPFORWARDTABLE
    dwNumEntries As Long
    table() As MIB_IPFORWARDROW
End Type

Public Type MIB_IPNETTABLE
    dwNumEntries As Long
    table() As MIB_IPNETROW
End Type

Public Type MIB_IFTABLE
    dwNumEntries As Long
    table() As MIB_IFROW
End Type

Public Type MIB_IPADDRTABLE
    dwNumEntries As Long
    table() As MIB_IPADDRROW
End Type

Public Type MIB_TCPTABLE_EX
    dwNumEntries As Long
    table() As MIB_TCPROW_EX
End Type

Public Type MIB_UDPTABLE_EX
    dwNumEntries  As Long
    table() As MIB_UDPROW_EX   ' ANY_SIZE : fake type
End Type

Public Type MIB_UDPSTATS
    dwInDatagrams As Long
    dwNoPorts As Long
    dwInErrors As Long
    dwOutDatagrams As Long
    dwNumAddrs As Long
End Type

Public Type MIB_UDPTABLE
    dwNumEntries As Long
    table() As MIB_UDPROW
End Type

Public Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Public Type MIBICMPSTATS
  dwMsgs As Long            ' number of messages
  dwErrors As Long          ' number of errors
  dwDestUnreachs As Long    ' destination unreachable messages
  dwTimeExcds As Long       ' time-to-live exceeded messages
  dwParmProbs As Long       ' parameter problem messages
  dwSrcQuenchs As Long      ' source quench messages
  dwRedirects As Long       ' redirection messages
  dwEchos As Long           ' echo requests
  dwEchoReps As Long        ' echo replies
  dwTimestamps As Long      ' timestamp requests
  dwTimestampReps As Long   ' timestamp replies
  dwAddrMasks As Long       ' address mask requests
  dwAddrMaskReps As Long    ' address mask replies
End Type

Public Type MIBICMPINFO
  icmpInStats As MIBICMPSTATS        ' stats for incoming messages
  icmpOutStats As MIBICMPSTATS       ' stats for outgoing messages
End Type

Public Declare Function AddIPAddress Lib "iphlpapi.dll" _
 (ByRef Address As IPADDR, ByRef IpMask As Long, ByVal IfIndex As Long, _
 ByRef NTEContext As Long, ByRef NTEInstance As Long) _
 As Long

Public Declare Function CreateIpForwardEntry Lib "iphlpapi.dll" _
 (ByRef pRoute As MIB_IPFORWARDROW) _
 As Long
 
Public Declare Function CreateIpNetEntry Lib "iphlpapi.dll" _
 (ByRef pArpEntry As MIB_IPNETROW) _
 As Long
 
Public Declare Function CreateProxyArpEntry Lib "iphlpapi.dll" _
 (ByVal dwAddress As Long, ByVal dwMask As Long, ByVal dwIfIndex As Long) _
 As Long
 
Public Declare Function DeleteIPAddress Lib "iphlpapi.dll" _
 (ByVal NTEContext As Long) _
 As Long
 
Public Declare Function DeleteIpForwardEntry Lib "iphlpapi.dll" _
 (ByRef pRoute As MIB_IPFORWARDROW) _
 As Long
 
Public Declare Function DeleteIpNetEntry Lib "iphlpapi.dll" _
 (ByRef pArpEntry As MIB_IPNETROW) _
 As Long
 
Public Declare Function DeleteProxyArpEntry Lib "iphlpapi.dll" _
 (ByVal dwAddress As Long, ByVal dwMask As Long, ByVal dwIfIndex As Long) _
 As Long
 
Public Declare Function EnableRouter Lib "iphlpapi.dll" _
 (ByRef pHandle As Long, ByRef pOverlapped As OVERLAPPED) _
 As Long
 
Public Declare Function FlushIpNetTable Lib "iphlpapi.dll" _
 (ByVal dwIfIndex As Long) _
 As Long
 
Public Declare Function GetNumberOfInterfaces Lib "iphlpapi.dll" _
 (ByRef pdwNumIf As Long) _
 As Long
 
Public Declare Function GetIcmpStatistics Lib "IPhlpAPI" _
  (ByRef pStats As MIBICMPINFO) As Long


Public Declare Function GetAdapterIndex Lib "iphlpapi.dll" _
 (ByVal AdapterName As String, ByRef IfIndex As Long) _
 As Long
 
Public Declare Function GetAdaptersInfo Lib "iphlpapi.dll" _
 (ByRef pAdapterInfo As IP_ADAPTER_INFO, ByRef pOutBufLen As Long) _
 As Long

Public Declare Function GetBestInterface Lib "iphlpapi.dll" _
 (ByRef dwDestAddr As IPADDR, ByRef pdwBestIfIndex As Long) _
 As Long
 
Public Declare Function GetBestRoute Lib "iphlpapi.dll" _
 (ByVal dwDestAddr As Long, ByVal dwSourceAddr As Long, ByRef pBestRoute As MIB_IPFORWARDROW) _
 As Long
 
Public Declare Function GetIfEntry Lib "iphlpapi.dll" _
 (ByRef pIfRow As MIB_IFROW) _
 As Long
 
Public Declare Function GetIfTable Lib "iphlpapi.dll" _
 (ByRef pIfTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) _
 As Long
 
Public Declare Function GetInterfaceInfo Lib "iphlpapi.dll" _
 (ByVal pIfTable As Long, ByVal dwOutBufLen As Long) _
 As Long
 
Public Declare Function GetIpAddrTable Lib "iphlpapi.dll" _
 (ByRef pIpAddrTable As Any, ByRef pdwSize As Long, _
 ByVal bOrder As Long) _
 As Long
 
Public Declare Function GetIpForwardTable Lib "iphlpapi.dll" _
 (ByRef pIpForwardTable As Any, ByRef pdwSize As Long, _
 ByVal bOrder As Long) _
 As Long
 
Public Declare Function GetIpNetTable Lib "iphlpapi.dll" _
 (ByRef pIpNetTable As Any, ByRef pdwSize As Long, _
 ByVal bOrder As Long) _
 As Long
 
Public Declare Function GetIpStatistics Lib "iphlpapi.dll" _
 (ByRef pStats As MIB_IPSTATS) _
 As Long
 
Public Declare Function GetNetworkParams Lib "iphlpapi.dll" _
 (ByRef pFixedInfo As Any, ByRef pOutBufLen As Long) _
 As Long
 
Public Declare Function GetPerAdapterInfo Lib "iphlpapi.dll" _
 (ByVal IfIndex As Long, ByRef pPerAdapterInfo As IP_PER_ADAPTER_INFO, _
 ByRef pOutBufLen As Long) _
 As Long
 
Public Declare Function GetTcpStatistics Lib "iphlpapi.dll" _
 (ByRef pStats As MIB_TCPSTATS) _
 As Long
 
Public Declare Function GetTcpTable Lib "iphlpapi.dll" _
 (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) _
 As Long

Public Declare Function GetUdpStatistics Lib "iphlpapi.dll" _
 (ByRef pStats As MIB_UDPSTATS) _
 As Long
 
 Public Declare Function GetProcessHeap Lib "kernel32" () As Long

Public Declare Function GetUdpTable Lib "iphlpapi.dll" _
 (ByRef pUdpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) _
 As Long

Public Declare Function AllocateAndGetTcpExTableFromStack Lib "IPhlpAPI" _
(lppTcpTable As Long, ByVal bOrder As Long, ByVal heap As Long, ByVal zero As Long, _
ByVal Flags As Long) As Long
Public Declare Function AllocateAndGetUdpExTableFromStack Lib "IPhlpAPI" _
(lppUdpTable As Long, ByVal bOrder As Long, ByVal heap As Long, ByVal zero As Long, _
ByVal Flags As Long) As Long
 
Public Declare Function GetUniDirectionalAdapterInfo Lib "iphlpapi.dll" _
 (ByRef pIPIfInfo As IP_UNIDIRECTIONAL_ADAPTER_ADDRESS, ByRef dwOutBufLen As Long) _
 As Long

Public Declare Function IpReleaseAddress Lib "iphlpapi.dll" _
 (ByRef AdapterInfo As IP_ADAPTER_INDEX_MAP) _
 As Long
 
Public Declare Function IpRenewAddress Lib "iphlpapi.dll" _
 (ByRef AdapterInfo As IP_ADAPTER_INDEX_MAP) _
 As Long
 
Public Declare Function NotifyAddrChange Lib "iphlpapi.dll" _
 (ByRef Handle As Long, ByRef OVERLAPPED As OVERLAPPED) _
 As Long
 
Public Declare Function NotifyRouteChange Lib "iphlpapi.dll" _
 (ByRef Handle As Long, ByRef OVERLAPPED As OVERLAPPED) _
 As Long
 
Public Declare Function SendARP Lib "iphlpapi.dll" _
 (ByVal DestIP As Long, ByVal SrcIP As Long, ByRef pMacAddr As Long, _
 ByRef PhyAddrLen As Long) _
 As Long
 
Public Declare Function SetIfEntry Lib "iphlpapi.dll" _
 (ByRef pIfRow As MIB_IFROW) _
 As Long
 
Public Declare Function SetIpForwardEntry Lib "iphlpapi.dll" _
 (ByRef pRoute As MIB_IPFORWARDROW) _
 As Long
 
Public Declare Function SetIpNetEntry Lib "iphlpapi.dll" _
 (ByRef pArpEntry As MIB_IPNETROW) _
 As Long
 
Public Declare Function SetIpStatistics Lib "iphlpapi.dll" _
 (ByRef pIpStats As MIB_IPSTATS) _
 As Long
 
Public Declare Function SetIpTTL Lib "iphlpapi.dll" _
 (ByVal nTTL As Long) _
 As Long
 
Public Declare Function SetTcpEntry Lib "iphlpapi.dll" _
 (ByRef pTcpRow As MIB_TCPROW) _
 As Long
 
Public Declare Function UnenableRouter Lib "iphlpapi.dll" _
 (ByRef pOverlapped As OVERLAPPED, ByRef lpdwEnableCount As Long) _
 As Long

Public Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (dest As Any, ByVal source As Long, ByVal Size As Long)

Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public Function TrimNull(Item As String)
' Trim any null characters from a piece of text
    If InStr(Item, Chr$(0)) Then
          TrimNull = Left$(Item, InStr(Item, Chr$(0)) - 1)
    Else: TrimNull = Item
    End If
End Function
