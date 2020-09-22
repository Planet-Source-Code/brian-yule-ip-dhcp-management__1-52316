VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UDPStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get InDatagrams() As Long
' Return the in datagrams
    InDatagrams = GetStats.dwInDatagrams
End Property

Public Property Get InErrors() As Long
' Return the in errors
    InErrors = GetStats.dwInErrors
End Property

Public Property Get NoPorts() As Long
' Return the no of ports
    NoPorts = GetStats.dwNoPorts
End Property

Public Property Get NumAddrs() As Long
' Return the number of addresses
    NumAddrs = GetStats.dwNumAddrs
End Property

Public Property Get OutDatagrams() As Long
' Return the in number of out datagrams
    OutDatagrams = GetStats.dwOutDatagrams
End Property

Private Function GetStats() As MIB_UDPSTATS
    Dim stats1 As MIB_UDPSTATS, RetVal As Long
    ' Get the UDP statistics
    RetVal = GetUdpStatistics(stats1)
    ' Return the Statistics struct
    GetStats = stats1
End Function
