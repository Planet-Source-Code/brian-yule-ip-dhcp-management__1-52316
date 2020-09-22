VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCPStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum MIB_TCP_RTO
     MIB_TCP_RTO_OTHER = 1
     MIB_TCP_RTO_CONSTANT = 2
     MIB_TCP_RTO_RSRE = 3
     MIB_TCP_RTO_VANJ = 4
End Enum

Private Function GetStats() As MIB_TCPSTATS
    Dim stats1 As MIB_TCPSTATS, RetVal As Long
    ' Get the TCP Statistics
    RetVal = GetTcpStatistics(stats1)
    ' Return the TCP Statistics struct
    GetStats = stats1
End Function

Public Property Get retransmission_time_out_Algorithm() As MIB_TCP_RTO
    retransmission_time_out_Algorithm = GetStats.dwRtoAlgorithm
End Property

Public Property Get ActiveOpens() As Long
    ActiveOpens = GetStats.dwActiveOpens
End Property

Public Property Get AttemptFails() As Long
    AttemptFails = GetStats.dwAttemptFails
End Property

Public Property Get CurrEstab() As Long
    CurrEstab = GetStats.dwCurrEstab
End Property

Public Property Get EstabResets() As Long
    EstabResets = GetStats.dwEstabResets
End Property

Public Property Get InErrs() As Long
    InErrs = GetStats.dwInErrs
End Property

Public Property Get InSegs() As Long
    InSegs = GetStats.dwInSegs
End Property

Public Property Get MaxConn() As Long
    MaxConn = GetStats.dwMaxConn
End Property

Public Property Get NumConns() As Long
    NumConns = GetStats.dwNumConns
End Property

Public Property Get OutRsts() As Long
    OutRsts = GetStats.dwOutRsts
End Property

Public Property Get OutSegs() As Long
    OutSegs = GetStats.dwOutSegs
End Property

Public Property Get PassiveOpens() As Long
    PassiveOpens = GetStats.dwPassiveOpens
End Property

Public Property Get RetransSegs() As Long
    RetransSegs = GetStats.dwRetransSegs
End Property

Public Property Get RtoMax() As Long
    RtoMax = GetStats.dwRtoMax
End Property

Public Property Get RtoMin() As Long
    RtoMin = GetStats.dwRtoMin
End Property
