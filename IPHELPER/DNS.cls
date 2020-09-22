VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetworkParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum NodeTypes
    BROADCAST_NODETYPE = 1
    PEER_TO_PEER_NODETYPE = 2
    MIXED_NODETYPE = 4
    HYBRID_NODETYPE = 8
End Enum

Public Property Get DNSServerList() As IPAddresses
' Returns a IPAddresses class which is a collection of IPAddress Classes holding IP addresses for DNS Server's
    Dim fInfo As FIXED_INFO
    Dim RetVal As New IPAddresses

    'Because there can be multiple items we are going to use a Do Loop to enumerate the items
    fInfo = FixedInfo
    Do
        ' Add a new IPAddress object to the collection of IP Addresses
        Call RetVal.Add(New IPAddress, RetVal.Count)
        ' Load the newly created object with the first DNS Server IP Address
        Call RetVal.Item(RetVal.Count - 1).Load(fInfo.DNSServerList)
        
        ' If there is more Servers in the list then
        If fInfo.DNSServerList.dwNext <> 0 Then
            ' Overwrite the current DNSServerList with the next one
            Call CopyMemory(fInfo.DNSServerList, ByVal fInfo.DNSServerList.dwNext, Len(fInfo.DNSServerList))
        Else
            ' Exit the loop, no more DNS Servers in the list
            Exit Do
        End If
    Loop
    ' Return the IP Addresses
    Set DNSServerList = RetVal
End Property

Private Property Get FixedInfo() As FIXED_INFO
    Dim lngRet As Long
    Dim lngSize As Long
    Dim byteArr() As Byte
    
    ' Resize the array to a single byte so we can use it in the function below
    lngRet = GetNetworkParams(ByVal 0, lngSize)
    
    If lngRet = ERROR_BUFFER_OVERFLOW Then
        ' Resize the buffer to the correct size
        ReDim byteArr(0 To lngSize - 1) As Byte
        ' Call GetNetworkParams again with the correct buffer size
        lngRet = GetNetworkParams(byteArr(0), lngSize)
        ' Copy the data returned into the FIXED_INFO struct.
        Call CopyMemory(ByVal VarPtr(FixedInfo), ByVal VarPtr(byteArr(0)), Len(FixedInfo))
    End If
End Property

Public Property Get hostname() As String
    hostname = TrimNull(StrConv(FixedInfo.hostname, vbUnicode)) ' Trim the null characters and return the host name
End Property

Public Property Get NodeType() As NodeTypes
    NodeType = FixedInfo.NodeType
End Property

Public Property Get DomainName() As String
    DomainName = TrimNull(StrConv(FixedInfo.DomainName, vbUnicode))
End Property

Public Property Get RoutingEnabled() As Boolean
    RoutingEnabled = IIf(FixedInfo.EnableRouting = 0, False, True)
End Property

Public Property Get DNSEnabled() As Boolean
    DNSEnabled = IIf(FixedInfo.EnableDns = 0, False, True)
End Property

Public Property Get ARPProxyEnabled() As Boolean
    ARPProxyEnabled = IIf(FixedInfo.EnableProxy = 0, False, True)
End Property

Public Property Get DHCPScopeName() As String
    DHCPScopeName = TrimNull(StrConv(FixedInfo.ScopeId, vbUnicode))
End Property


