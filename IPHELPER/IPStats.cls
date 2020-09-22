VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function GetStats() As MIB_IPSTATS
    Dim stats1 As MIB_IPSTATS, RetVal As Long
    ' Gets the IP Statistics
    RetVal = GetIpStatistics(stats1)
    ' Returns the IP Statistics struct
    GetStats = stats1
End Function

Public Property Get DefaultTTL() As Long
    DefaultTTL = GetStats.dwDefaultTTL
End Property

Public Property Get Forwarding() As Boolean
    Forwarding = IIf(GetStats.dwForwarding, True, False)
End Property

Public Property Get ForwDatagrams() As Long
    ForwDatagrams = GetStats.dwForwDatagrams
End Property

Public Property Get FragCreates() As Long
    FragCreates = GetStats.dwFragCreates
End Property

Public Property Get FragFails() As Long
    FragFails = GetStats.dwFragFails
End Property

Public Property Get FragOks() As Long
    FragOks = GetStats.dwFragOks
End Property

Public Property Get InAddrErrors() As Long
    InAddrErrors = GetStats.dwInAddrErrors
End Property

Public Property Get InDelivers() As Long
    InDelivers = GetStats.dwInDelivers
End Property

Public Property Get InDiscards() As Long
    InDiscards = GetStats.dwInDiscards
End Property

Public Property Get InHdrErrors() As Long
    InHdrErrors = GetStats.dwInHdrErrors
End Property

Public Property Get InReceives() As Long
    InReceives = GetStats.dwInReceives
End Property

Public Property Get InUnknownProtos() As Long
    InUnknownProtos = GetStats.dwInUnknownProtos
End Property

Public Property Get NumAddr() As Long
    NumAddr = GetStats.dwNumAddr
End Property

Public Property Get NumIf() As Long
    NumIf = GetStats.dwNumIf
End Property

Public Property Get NumRoutes() As Long
    NumRoutes = GetStats.dwNumRoutes
End Property

Public Property Get OutDiscards() As Long
    OutDiscards = GetStats.dwOutDiscards
End Property

Public Property Get OutNoRoutes() As Long
    OutNoRoutes = GetStats.dwOutNoRoutes
End Property

Public Property Get OutRequests() As Long
    OutRequests = GetStats.dwOutRequests
End Property

Public Property Get ReasmFails() As Long
    ReasmFails = GetStats.dwReasmFails
End Property

Public Property Get ReasmOks() As Long
    ReasmOks = GetStats.dwReasmOks
End Property

Public Property Get ReasmReqds() As Long
    ReasmReqds = GetStats.dwReasmReqds
End Property

Public Property Get ReasmTimeout() As Long
    ReasmTimeout = GetStats.dwReasmTimeout
End Property

Public Property Get RoutingDiscards() As Long
    RoutingDiscards = GetStats.dwRoutingDiscards
End Property
