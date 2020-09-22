VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetworkInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private LocalAdapter As IP_ADAPTER_INDEX_MAP

Friend Property Let InterfaceInfo(ByRef ipAdtInfo As IP_ADAPTER_INDEX_MAP)
' Sets the IP_ADAPTER_INDEX_MAP struct
    LocalAdapter = ipAdtInfo
End Property

Public Property Get Name() As String
' Returns the name of the IP Adapter
    Dim tmpStr As String
    tmpStr = LocalAdapter.nname
    Name = TrimNull(tmpStr)
End Property

Public Property Get Index() As Long
' Returns the index of the IP Adapter
    Index = LocalAdapter.Index
End Property

Public Function Release() As Long
' Releases the DHCP address of the adapter
    Release = IpReleaseAddress(LocalAdapter)
End Function

Public Function Renew() As Long
' Renews the DHCP address of the adapter
    Renew = IpRenewAddress(LocalAdapter)
End Function
