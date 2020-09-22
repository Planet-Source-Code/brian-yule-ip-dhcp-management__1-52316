VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim localIPAddrStr As IP_ADDR_STRING

Friend Sub Load(ipAddrStr As IP_ADDR_STRING)
' Takes in the IP_ADDR_STRING Struct
    localIPAddrStr = ipAddrStr
End Sub

Public Property Get Address() As String
' Returns the string address from the IP_ADDR_STRING supplied
    Address = TrimNull(StrConv(localIPAddrStr.IPAddress.IPADDR, vbUnicode))
End Property

Public Property Get Mask() As String
' Returns the string Mask from the IP_ADDR_STRING supplied
    Mask = TrimNull(StrConv(localIPAddrStr.IpMask.IpMask, vbUnicode))
End Property

Public Property Get NTEContext() As Long
' Returns the string NTEContext from the IP_ADDR_STRING supplied
    NTEContext = localIPAddrStr.dwContext
End Property
