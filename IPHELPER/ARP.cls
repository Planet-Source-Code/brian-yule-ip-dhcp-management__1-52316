VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ARP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const STRING_ZERO = "0"
Private Const ADAPTER_ADDRESS_DELIMITER = "-"

Public Function Send(DestinationIP As String) As String
' Send an ARP Request and return a MAC Address
    Dim dest As Long, source As Long
    Dim MacAddr() As Byte
    Dim PhyAddrLen As Long
    Dim retPtr As Long
    
    PhyAddrLen = MAXLEN_PHYSADDR
    dest = GetHostByNameAlias(DestinationIP) ' Get the network long of the IP Address
    
    ' Send the arp request and if successfull then
    If SendARP(dest, 0&, retPtr, PhyAddrLen) = ERROR_SUCCESS Then
        'Resize the array to the size of the data in the pointer
        ReDim MacAddr(0 To PhyAddrLen - 1) As Byte
    
        ' Copy the data from the pointer to the byte array
        Call CopyMemory(ByVal VarPtr(MacAddr(0)), ByVal VarPtr(retPtr), ByVal PhyAddrLen)
        
        ' Return the convered byte array
        Send = ConvAddress(MacAddr, PhyAddrLen)
    End If
End Function

Private Function ConvAddress(ByRef Addr() As Byte, Length As Long) As String
' Converts a byte array to a physical mac address
    Dim strdef As String, i As Integer

    For i = 0 To Length - 1
        If i = 0 Then
            strdef = strdef & Pad(Hex(Addr(i)))
        Else
            strdef = strdef & ADAPTER_ADDRESS_DELIMITER & Pad(Hex(Addr(i)))
        End If
    Next i
    ConvAddress = strdef
End Function

Private Function Pad(hexStr As String) As String
' Pads a single character with a 0 to the left if it requires one
    If Len(hexStr) = 1 Then
        Pad = STRING_ZERO & hexStr
    ElseIf Len(hexStr) = 2 Then
        Pad = hexStr
    End If
End Function

