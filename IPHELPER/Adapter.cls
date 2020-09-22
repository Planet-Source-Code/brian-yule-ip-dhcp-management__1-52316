VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetworkAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const MAX_ADAPTER_ADDRESS     As Long = 8
Private Const MAX_ADAPTER_DESCRIPTION As Long = 128
Private Const MAX_ADAPTER_NAME        As Long = 256

Public Enum MIB_IF_TYPE
    MIB_IF_TYPE_OTHER = 1
    MIB_IF_TYPE_ETHERNET = 6
    MIB_IF_TYPE_TOKENRING = 9
    MIB_IF_TYPE_FDDI = 15
    MIB_IF_TYPE_PPP = 23
    MIB_IF_TYPE_LOOPBACK = 24
    MIB_IF_TYPE_SLIP = 28
End Enum

Public Type IP_ADDRESS_STRING
    IPADDR(0 To 15) As Byte
End Type

Public Type IP_MASK_STRING
    IpMask(0 To 15) As Byte
End Type

Public Type IP_ADDR_STRING
    dwNext     As Long
    IPAddress  As IP_ADDRESS_STRING
    IpMask     As IP_MASK_STRING
    dwContext  As Long
End Type

Public Type IP_ADAPTER_INFO
    dwNext                As Long
    ComboIndex            As Long
    sAdapterName(0 To (MAX_ADAPTER_NAME + 3))        As Byte
    sDescription(0 To (MAX_ADAPTER_DESCRIPTION + 3)) As Byte
    dwAddressLength       As Long
    Address(0 To (MAX_ADAPTER_ADDRESS - 1))       As Byte
    dwIndex               As Long
    uType                 As Long
    uDhcpEnabled          As Long
    CurrentIpAddress      As Long
    IPAddressList         As IP_ADDR_STRING
    GatewayList           As IP_ADDR_STRING
    DhcpServer            As IP_ADDR_STRING
    bHaveWins             As Long
    PrimaryWinsServer     As IP_ADDR_STRING
    SecondaryWinsServer   As IP_ADDR_STRING
    LeaseObtained         As Long
    LeaseExpires          As Long
End Type

Private Const DATE_SECONDS = "s"
Private Const STRING_ZERO = "0"
Private Const ADAPTER_ADDRESS_DELIMITER = "-"

Private LocalAdapter As IP_ADAPTER_INFO

Friend Property Let AdapterInfo(ByRef ipAdtInfo As IP_ADAPTER_INFO)
' Takes the IP_ADAPTER_INFO in
     LocalAdapter = ipAdtInfo
End Property

Public Property Get AutoConfigEnabled() As Boolean
' Returns the Auto Config Enabled property of the IP_ADAPTER_INFO struct taken in
    Dim RetVal As Long
    Dim ipPerAdapterInfo As IP_PER_ADAPTER_INFO
    RetVal = GetPerAdapterInfo(AdapterIndex, ipPerAdapterInfo, Len(ipPerAdapterInfo))
    
    AutoConfigEnabled = IIf(ipPerAdapterInfo.AutoConfigEnabled = 0, False, True)
End Property

Public Property Get AutoConfigActive() As Boolean
' Returns the Auto Config Active property of the IP_ADAPTER_INFO struct taken in
    Dim RetVal As Long
    Dim ipPerAdapterInfo As IP_PER_ADAPTER_INFO
    RetVal = GetPerAdapterInfo(AdapterIndex, ipPerAdapterInfo, Len(ipPerAdapterInfo))
    
    AutoConfigActive = IIf(ipPerAdapterInfo.AutoConfigActive = 0, False, True)
End Property

Public Property Get AdapterIndex() As Long
' Returns the AdapterIndex property of the IP_ADAPTER_INFO struct taken in
    AdapterIndex = LocalAdapter.dwIndex
End Property

Public Property Get LeaseObtained() As Date
' Returns the Lease Obtained property of the IP_ADAPTER_INFO struct taken in translated into a date
    If LocalAdapter.LeaseObtained > 0 Then
        LeaseObtained = DateAdd(DATE_SECONDS, LocalAdapter.LeaseObtained, #1/1/1970#)
    End If
End Property

Public Property Get LeaseExpires() As Date
' Returns the Lease Expires property of the IP_ADAPTER_INFO struct taken in translated into a date
    If LocalAdapter.LeaseExpires > 0 Then
        LeaseExpires = DateAdd(DATE_SECONDS, LocalAdapter.LeaseExpires, #1/1/1970#)
    End If
End Property

Public Property Get DHCP() As IPAddresses
' Returns the DHCP Server IP Addresses of the IP_ADAPTER_INFO struct taken in
    Dim addresses As New IPAddresses
    
    Do
        Call addresses.Add(New IPAddress, addresses.Count)
        Call addresses.Item(addresses.Count - 1).Load(LocalAdapter.DhcpServer)
        
        If LocalAdapter.DhcpServer.dwNext <> 0 Then
            Call CopyMemory(LocalAdapter.DhcpServer, ByVal LocalAdapter.DhcpServer.dwNext, Len(LocalAdapter.DhcpServer))
        Else
            Exit Do
        End If
    Loop

    Set DHCP = addresses
End Property

Public Property Get Gateway() As IPAddresses
' Returns the Gateway IP Addresses of the IP_ADAPTER_INFO struct taken in
    Dim addresses As New IPAddresses
    
    Do
        Call addresses.Add(New IPAddress, addresses.Count)
        Call addresses.Item(addresses.Count - 1).Load(LocalAdapter.GatewayList)
        
        If LocalAdapter.GatewayList.dwNext <> 0 Then
            Call CopyMemory(LocalAdapter.GatewayList, ByVal LocalAdapter.GatewayList.dwNext, Len(LocalAdapter.GatewayList))
        Else
            Exit Do
        End If
    Loop

    Set Gateway = addresses
End Property

Public Property Get PrimaryWins() As IPAddresses
' Returns the Primary Wins IP Addresses of the IP_ADAPTER_INFO struct taken in
    Dim addresses As New IPAddresses
    
    Do
        Call addresses.Add(New IPAddress, addresses.Count)
        Call addresses.Item(addresses.Count - 1).Load(LocalAdapter.PrimaryWinsServer)
        
        If LocalAdapter.PrimaryWinsServer.dwNext <> 0 Then
            Call CopyMemory(LocalAdapter.PrimaryWinsServer, ByVal LocalAdapter.PrimaryWinsServer.dwNext, Len(LocalAdapter.PrimaryWinsServer))
        Else
            Exit Do
        End If
    Loop

    Set PrimaryWins = addresses
End Property

Public Property Get SecondaryWins() As IPAddresses
' Returns the Secondary Wins IP Addresses of the IP_ADAPTER_INFO struct taken in
    Dim addresses As New IPAddresses
    
    Do
        Call addresses.Add(New IPAddress, addresses.Count)
        Call addresses.Item(addresses.Count - 1).Load(LocalAdapter.SecondaryWinsServer)
        
        If LocalAdapter.SecondaryWinsServer.dwNext <> 0 Then
            Call CopyMemory(LocalAdapter.SecondaryWinsServer, ByVal LocalAdapter.SecondaryWinsServer.dwNext, Len(LocalAdapter.SecondaryWinsServer))
        Else
            Exit Do
        End If
    Loop

    Set SecondaryWins = addresses
End Property

Public Property Get IPAddressList() As IPAddresses
' Returns the IP Address List of the IP_ADAPTER_INFO struct taken in
    Dim addresses As New IPAddresses
    
    Do
        Call addresses.Add(New IPAddress, addresses.Count)
        Call addresses.Item(addresses.Count - 1).Load(LocalAdapter.IPAddressList)
        
        If LocalAdapter.IPAddressList.dwNext <> 0 Then
            Call CopyMemory(LocalAdapter.IPAddressList, ByVal LocalAdapter.IPAddressList.dwNext, Len(LocalAdapter.IPAddressList))
        Else
            Exit Do
        End If
    Loop

    Set IPAddressList = addresses
End Property

Public Property Get AdapterName() As String
' Returns the Adapter Name of the IP_ADAPTER_INFO struct taken in translated from Unicode
    Dim strdef As String

    strdef = LocalAdapter.sAdapterName
    strdef = StrConv(strdef, vbUnicode)

    AdapterName = TrimNull(strdef)
End Property

Public Property Get AdapterDescription() As String
' Returns the Adapter Description of the IP_ADAPTER_INFO struct taken in translated from Unicode
    Dim strdef As String

    strdef = LocalAdapter.sDescription
    strdef = StrConv(strdef, vbUnicode)

    AdapterDescription = TrimNull(strdef)
End Property

Public Property Get Address() As String
' Returns the MAC Addresses of the IP_ADAPTER_INFO struct taken in
    Dim strdef As String, i As Integer

    For i = 0 To LocalAdapter.dwAddressLength - 1
        If i = 0 Then
            strdef = strdef & Pad(Hex(LocalAdapter.Address(i)))
        Else
            strdef = strdef & ADAPTER_ADDRESS_DELIMITER & Pad(Hex(LocalAdapter.Address(i)))
        End If
    Next i
    Address = strdef
End Property

Private Function Pad(hexStr As String) As String
    If Len(hexStr) = 1 Then
        Pad = STRING_ZERO & hexStr
    ElseIf Len(hexStr) = 2 Then
        Pad = hexStr
    End If
End Function

Public Property Get AdapterType() As MIB_IF_TYPE
' Returns the Adapter Type of the IP_ADAPTER_INFO struct taken in
    AdapterType = LocalAdapter.uType
End Property

Public Property Get HasWins() As Boolean
' Returns the Has Wins property of the IP_ADAPTER_INFO struct taken in
    HasWins = IIf(LocalAdapter.bHaveWins = 0, False, True)
End Property

Public Property Get DHCPEnabled() As Boolean
' Returns the DHCP Enabled property of the IP_ADAPTER_INFO struct taken in
    DHCPEnabled = IIf(LocalAdapter.uDhcpEnabled = 0, False, True)
End Property


