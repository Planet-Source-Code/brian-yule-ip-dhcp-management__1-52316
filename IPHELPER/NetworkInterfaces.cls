VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetworkInterfaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' This is a standard collection class but it includes the initializing of the collection when it is created!

Private mCol As New Collection

Friend Function Add(ByRef newAdapter As NetworkInterface, Optional ByVal sKey As String) As NetworkInterface
    If Len(sKey) = 0 Then mCol.Add newAdapter Else mCol.Add newAdapter, sKey
    Set Add = newAdapter
End Function

Public Property Get Item(vntIndexKey As Variant) As NetworkInterface
Attribute Item.VB_UserMemId = 0
    If Len(vntIndexKey) > 0 Then Set Item = mCol.Item(CStr(vntIndexKey))
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Private Sub Remove(vntIndexKey As Variant)
    If Len(vntIndexKey) > 0 Then mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Call Init
End Sub

Private Sub Init()
    Dim i As Integer
    Dim ipInterfaceinfo As IP_INTERFACE_INFO
    Dim byteArr() As Byte
    Dim tmpNetInt As NetworkInterface
    
    If NoOfAdapters > 0 Then ' If there is more than 1 adapter
        ' Resize the array to handle the number of adapters
        ReDim ipInterfaceinfo.Adapter(0 To NoOfAdapters - 1) As IP_ADAPTER_INDEX_MAP
        ' Get the byte Interface Information
        byteArr = byteIntInfo
        
        ' Copy the information to the array of adapters
        Call CopyMemory( _
            ByVal VarPtr(ipInterfaceinfo.Adapter(0)), _
            ByVal VarPtr(byteArr(4&)), _
            Len(ipInterfaceinfo.Adapter(0)) * NoOfAdapters _
        )
        
        ' For each adapter
        For i = 0 To NoOfAdapters - 1
            ' Add a new Network Interface object to the collection
            Call Add(New NetworkInterface, Trim(Str(mCol.Count)))
            ' Set the interface information of the new class
            Item(mCol.Count - 1).InterfaceInfo = ipInterfaceinfo.Adapter(i)
        Next i
    End If
End Sub

Public Property Get NoOfInterfaces() As Long
' Returns the number of interfaces
    Dim lngRet As Long
    lngRet = GetNumberOfInterfaces(NoOfInterfaces)
End Property

Public Property Get NoOfAdapters() As Long
' Returns the number of Adapters
    Dim byteArr() As Byte
    byteArr = byteIntInfo
    Call CopyMemory(ByVal VarPtr(NoOfAdapters), ByVal VarPtr(byteArr(0)), 4&)
End Property

Private Property Get byteIntInfo() As Byte()
' Returns the Network Interface Information Array in byte format
    Dim retSize As Long
    Dim byteArr() As Byte
    
    If GetInterfaceInfo(ByVal 0, VarPtr(retSize)) = modIPHelp.ERROR_INSUFFICIENT_BUFFER Then
        ' Resize byte array to hold all the data, retSize will have the size of the buffer required
        ReDim Preserve byteArr(0 To retSize - 1) As Byte
        ' Get the interface info again with the correct buffer size
        If GetInterfaceInfo(ByVal VarPtr(byteArr(0)), VarPtr(retSize)) = 0 Then
            ' Return the byte array information
            byteIntInfo = byteArr
        End If
    End If
End Property
