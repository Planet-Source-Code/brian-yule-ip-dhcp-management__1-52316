VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICMPStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function GetStats() As MIBICMPINFO
    Dim stats1 As MIBICMPINFO, RetVal As Long
    ' Get the ICMP Statistics
    RetVal = GetIcmpStatistics(stats1)
    ' Return the ICMP Information Struct
    GetStats = stats1
End Function

Public Property Get InAddrMaskReps() As Long
    InAddrMaskReps = GetStats.icmpInStats.dwAddrMaskReps
End Property

Public Property Get OutAddrMaskReps() As Long
    OutAddrMaskReps = GetStats.icmpOutStats.dwAddrMaskReps
End Property

Public Property Get InAddrMasks() As Long
    InAddrMasks = GetStats.icmpInStats.dwAddrMasks
End Property

Public Property Get OutAddrMasks() As Long
    OutAddrMasks = GetStats.icmpOutStats.dwAddrMasks
End Property

Public Property Get InDestUnreachs() As Long
    InDestUnreachs = GetStats.icmpInStats.dwDestUnreachs
End Property

Public Property Get OutDestUnreachs() As Long
    OutDestUnreachs = GetStats.icmpOutStats.dwDestUnreachs
End Property

Public Property Get InEchoReps() As Long
    InEchoReps = GetStats.icmpInStats.dwEchoReps
End Property

Public Property Get OutEchoReps() As Long
    OutEchoReps = GetStats.icmpOutStats.dwEchoReps
End Property

Public Property Get InEchos() As Long
    InEchos = GetStats.icmpInStats.dwEchos
End Property

Public Property Get OutEchos() As Long
    OutEchos = GetStats.icmpOutStats.dwEchos
End Property

Public Property Get InErrors() As Long
    InErrors = GetStats.icmpInStats.dwErrors
End Property

Public Property Get OutErrors() As Long
    OutErrors = GetStats.icmpOutStats.dwErrors
End Property

Public Property Get InMsgs() As Long
    InMsgs = GetStats.icmpInStats.dwMsgs
End Property

Public Property Get OutMsgs() As Long
    OutMsgs = GetStats.icmpOutStats.dwMsgs
End Property

Public Property Get InParmProbs() As Long
    InParmProbs = GetStats.icmpInStats.dwParmProbs
End Property

Public Property Get OutParmProbs() As Long
    OutParmProbs = GetStats.icmpOutStats.dwParmProbs
End Property

Public Property Get InRedirects() As Long
    InRedirects = GetStats.icmpInStats.dwRedirects
End Property

Public Property Get OutRedirects() As Long
    OutRedirects = GetStats.icmpOutStats.dwRedirects
End Property

Public Property Get InSrcQuenchs() As Long
    InSrcQuenchs = GetStats.icmpInStats.dwSrcQuenchs
End Property

Public Property Get OutSrcQuenchs() As Long
    OutSrcQuenchs = GetStats.icmpOutStats.dwSrcQuenchs
End Property

Public Property Get InTimeExcds() As Long
    InTimeExcds = GetStats.icmpInStats.dwTimeExcds
End Property

Public Property Get OutTimeExcds() As Long
    OutTimeExcds = GetStats.icmpOutStats.dwTimeExcds
End Property

Public Property Get InTimestampReps() As Long
    InTimestampReps = GetStats.icmpInStats.dwTimestampReps
End Property

Public Property Get OutTimestampReps() As Long
    OutTimestampReps = GetStats.icmpOutStats.dwTimestampReps
End Property

Public Property Get InTimestamps() As Long
    InTimestamps = GetStats.icmpInStats.dwTimestamps
End Property

Public Property Get OutTimestamps() As Long
    OutTimestamps = GetStats.icmpOutStats.dwTimestamps
End Property

