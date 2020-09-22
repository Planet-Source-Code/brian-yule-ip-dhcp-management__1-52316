Attribute VB_Name = "modFriendlyNames"
Option Explicit

Public Function GetOperStatusString(ByVal status As MIB_IF_OPER_STATUS) As String
' Gets the string representation of the If tables Operational status
    Select Case status
        Case MIB_IF_OPER_STATUS.MIB_IF_OPER_STATUS_CONNECTED: GetOperStatusString = LoadResString(270)
        Case MIB_IF_OPER_STATUS.MIB_IF_OPER_STATUS_CONNECTING: GetOperStatusString = LoadResString(271)
        Case MIB_IF_OPER_STATUS.MIB_IF_OPER_STATUS_DISCONNECTED: GetOperStatusString = LoadResString(272)
        Case MIB_IF_OPER_STATUS.MIB_IF_OPER_STATUS_NON_OPERATIONAL: GetOperStatusString = LoadResString(273)
        Case MIB_IF_OPER_STATUS.MIB_IF_OPER_STATUS_OPERATIONAL: GetOperStatusString = LoadResString(274)
        Case MIB_IF_OPER_STATUS.MIB_IF_OPER_STATUS_UNREACHABLE: GetOperStatusString = LoadResString(275)
    End Select
End Function

Public Function GetIfTypeString(ByVal ifType As MIB_IF_TYPE) As String
' Gets the string representation of the If tables type
    Select Case ifType
        Case MIB_IF_TYPE.MIB_IF_TYPE_OTHER: GetIfTypeString = LoadResString(306)
        Case MIB_IF_TYPE.MIB_IF_TYPE_ETHERNET: GetIfTypeString = LoadResString(307)
        Case MIB_IF_TYPE.MIB_IF_TYPE_TOKENRING: GetIfTypeString = LoadResString(308)
        Case MIB_IF_TYPE.MIB_IF_TYPE_FDDI: GetIfTypeString = LoadResString(309)
        Case MIB_IF_TYPE.MIB_IF_TYPE_PPP: GetIfTypeString = LoadResString(310)
        Case MIB_IF_TYPE.MIB_IF_TYPE_LOOPBACK: GetIfTypeString = LoadResString(311)
        Case MIB_IF_TYPE.MIB_IF_TYPE_SLIP: GetIfTypeString = LoadResString(312)
    End Select
End Function

Public Function GetARPTypeString(ByVal arpType As ARP_TYPES) As String
' Gets the string representation of the ARP type
    Select Case arpType
        Case ARP_TYPES.arpDynamic: GetARPTypeString = LoadResString(276)
        Case ARP_TYPES.arpInvalid: GetARPTypeString = LoadResString(277)
        Case ARP_TYPES.arpOther: GetARPTypeString = LoadResString(278)
        Case ARP_TYPES.arpStatic: GetARPTypeString = LoadResString(279)
    End Select
End Function

Public Function GetIPForwardTypeString(ByVal ipfType As IPFORWARD_TYPES) As String
' Gets the string representation of the IP Forward type
    Select Case ipfType
        Case IPFORWARD_TYPES.other: GetIPForwardTypeString = LoadResString(302)
        Case IPFORWARD_TYPES.invalid: GetIPForwardTypeString = LoadResString(303)
        Case IPFORWARD_TYPES.Local_Route: GetIPForwardTypeString = LoadResString(304)
        Case IPFORWARD_TYPES.Remote_Route: GetIPForwardTypeString = LoadResString(305)
    End Select
End Function

Public Function GetIPForwardProtocolString(ByVal ipfType As PROTOCOLS) As String
' Gets the string representation of the IP Forward Protocol
    Select Case ipfType
        Case PROTOCOLS.IP_BBN: GetIPForwardProtocolString = LoadResString(280)
        Case PROTOCOLS.IP_BGP: GetIPForwardProtocolString = LoadResString(281)
        Case PROTOCOLS.IP_BOOTP: GetIPForwardProtocolString = LoadResString(282)
        Case PROTOCOLS.IP_CISCO: GetIPForwardProtocolString = LoadResString(283)
        Case PROTOCOLS.IP_EGP: GetIPForwardProtocolString = LoadResString(284)
        Case PROTOCOLS.IP_ES_IS: GetIPForwardProtocolString = LoadResString(285)
        Case PROTOCOLS.IP_GGP: GetIPForwardProtocolString = LoadResString(286)
        Case PROTOCOLS.IP_HELLO: GetIPForwardProtocolString = LoadResString(287)
        Case PROTOCOLS.IP_ICMP: GetIPForwardProtocolString = LoadResString(288)
        Case PROTOCOLS.IP_IS_IS: GetIPForwardProtocolString = LoadResString(289)
        Case PROTOCOLS.IP_LOCAL: GetIPForwardProtocolString = LoadResString(290)
        Case PROTOCOLS.IP_NETMGMT: GetIPForwardProtocolString = LoadResString(291)
        Case PROTOCOLS.IP_NT_AUTOSTATIC: GetIPForwardProtocolString = LoadResString(292)
        Case PROTOCOLS.IP_OSPF: GetIPForwardProtocolString = LoadResString(293)
        Case PROTOCOLS.IP_OTHER: GetIPForwardProtocolString = LoadResString(294)
        Case PROTOCOLS.IP_RIP: GetIPForwardProtocolString = LoadResString(295)
        Case PROTOCOLS.IPX_PROTOCOL_NLSP: GetIPForwardProtocolString = LoadResString(296)
        Case PROTOCOLS.IPX_PROTOCOL_RIP: GetIPForwardProtocolString = LoadResString(297)
        Case PROTOCOLS.IPX_PROTOCOL_SAP: GetIPForwardProtocolString = LoadResString(298)
        Case PROTOCOLS.IP_NT_STATIC_NON_DOD: GetIPForwardProtocolString = LoadResString(299)
        Case PROTOCOLS.IP_NT_STATIC: GetIPForwardProtocolString = LoadResString(300)
        Case PROTOCOLS.IP_IDPR: GetIPForwardProtocolString = LoadResString(301)
    End Select
End Function

Public Function GetIPForwardProtocolValue(ByVal ipfType As String) As Long
' Gets the IP Forward Protocol from its string representation
    Select Case ipfType
        Case LoadResString(280): GetIPForwardProtocolValue = PROTOCOLS.IP_BBN
        Case LoadResString(281): GetIPForwardProtocolValue = PROTOCOLS.IP_BGP
        Case LoadResString(282): GetIPForwardProtocolValue = PROTOCOLS.IP_BOOTP
        Case LoadResString(283): GetIPForwardProtocolValue = PROTOCOLS.IP_CISCO
        Case LoadResString(284): GetIPForwardProtocolValue = PROTOCOLS.IP_EGP
        Case LoadResString(285): GetIPForwardProtocolValue = PROTOCOLS.IP_ES_IS
        Case LoadResString(286): GetIPForwardProtocolValue = PROTOCOLS.IP_GGP
        Case LoadResString(287): GetIPForwardProtocolValue = PROTOCOLS.IP_HELLO
        Case LoadResString(288): GetIPForwardProtocolValue = PROTOCOLS.IP_ICMP
        Case LoadResString(289): GetIPForwardProtocolValue = PROTOCOLS.IP_IS_IS
        Case LoadResString(290): GetIPForwardProtocolValue = PROTOCOLS.IP_LOCAL
        Case LoadResString(291): GetIPForwardProtocolValue = PROTOCOLS.IP_NETMGMT
        Case LoadResString(292): GetIPForwardProtocolValue = PROTOCOLS.IP_NT_AUTOSTATIC
        Case LoadResString(293): GetIPForwardProtocolValue = PROTOCOLS.IP_OSPF
        Case LoadResString(294): GetIPForwardProtocolValue = PROTOCOLS.IP_OTHER
        Case LoadResString(295): GetIPForwardProtocolValue = PROTOCOLS.IP_RIP
        Case LoadResString(296): GetIPForwardProtocolValue = PROTOCOLS.IPX_PROTOCOL_NLSP
        Case LoadResString(297): GetIPForwardProtocolValue = PROTOCOLS.IPX_PROTOCOL_RIP
        Case LoadResString(298): GetIPForwardProtocolValue = PROTOCOLS.IPX_PROTOCOL_SAP
        Case LoadResString(299): GetIPForwardProtocolValue = PROTOCOLS.IP_NT_STATIC_NON_DOD
        Case LoadResString(300): GetIPForwardProtocolValue = PROTOCOLS.IP_NT_STATIC
        Case LoadResString(301): GetIPForwardProtocolValue = PROTOCOLS.IP_IDPR
    End Select
End Function

Public Function GetState(ByVal lngState As MIB_TCP_STATE) As String
' Gets the string representation of the TCP socket State
    Select Case lngState
        Case MIB_TCP_STATE_CLOSED: GetState = LoadResString(258)
        Case MIB_TCP_STATE_LISTEN: GetState = LoadResString(259)
        Case MIB_TCP_STATE_SYN_SENT: GetState = LoadResString(260)
        Case MIB_TCP_STATE_SYN_RCVD: GetState = LoadResString(261)
        Case MIB_TCP_STATE_ESTAB: GetState = LoadResString(262)
        Case MIB_TCP_STATE_FIN_WAIT1: GetState = LoadResString(263)
        Case MIB_TCP_STATE_FIN_WAIT2: GetState = LoadResString(264)
        Case MIB_TCP_STATE_CLOSE_WAIT: GetState = LoadResString(265)
        Case MIB_TCP_STATE_CLOSING: GetState = LoadResString(266)
        Case MIB_TCP_STATE_LAST_ACK: GetState = LoadResString(267)
        Case MIB_TCP_STATE_TIME_WAIT: GetState = LoadResString(268)
        Case MIB_TCP_STATE_DELETE_TCB: GetState = LoadResString(269)
    End Select
End Function

Public Function GetNodeTypeString(NType As NodeTypes) As String
' Gets the string representation of the network node types
    Select Case NType
        Case NodeTypes.BROADCAST_NODETYPE: GetNodeTypeString = LoadResString(254)
        Case NodeTypes.PEER_TO_PEER_NODETYPE: GetNodeTypeString = LoadResString(255)
        Case NodeTypes.MIXED_NODETYPE: GetNodeTypeString = LoadResString(256)
        Case NodeTypes.HYBRID_NODETYPE: GetNodeTypeString = LoadResString(257)
    End Select
End Function

