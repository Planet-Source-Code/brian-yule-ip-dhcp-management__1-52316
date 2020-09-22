VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetworkAdapters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Adapter"
Attribute VB_Ext_KEY = "Member0" ,"Adapter"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' This is a standard collection class but it includes the initializing of the collection when it is created!

Private mCol As New Collection

Friend Function Add(newAdapter As NetworkAdapter, Optional sKey As String) As NetworkAdapter
    If Len(sKey) = 0 Then mCol.Add newAdapter Else mCol.Add newAdapter, CStr(sKey)
    Set Add = newAdapter
End Function

Public Property Get Item(vntIndexKey As Variant) As NetworkAdapter
Attribute Item.VB_UserMemId = 0
    If Len(vntIndexKey) > 0 Then Set Item = mCol.Item(CStr(vntIndexKey))
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Private Sub Remove(vntIndexKey As Variant)
    If Len(vntIndexKey) > 0 Then mCol.Remove CStr(vntIndexKey)
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
    Dim lngRet As Long
    Dim ipAdtInfo() As IP_ADAPTER_INFO
    Dim outBufLen As Long
    Dim tmpAdapter  As NetworkAdapter
    Dim i As Integer

    ' Resize the array to at least a single object!
    ReDim ipAdtInfo(0) As IP_ADAPTER_INFO
    
    outBufLen = Len(ipAdtInfo(0))
    
    ' Gets Adapter Info, Usually returns Buffer is too small so it returns an error requiring a
    ' larger buffer and giving us the size of the required buffer
    lngRet = GetAdaptersInfo(ipAdtInfo(0), outBufLen)
    
    If lngRet = ERROR_BUFFER_OVERFLOW Then
        ' Resize the buffer to the size required
        ReDim ipAdtInfo(0 To (outBufLen / Len(ipAdtInfo(0))) - 1) As IP_ADAPTER_INFO
        ' Get the adapter info again, this time it will work because the buffer is the correct size
        lngRet = GetAdaptersInfo(ipAdtInfo(0), outBufLen)
    End If
    If lngRet = 0 Then ' If the function succeeds then
        ' For each Adapter in the array of adapters
        For i = 0 To UBound(ipAdtInfo)
            Set tmpAdapter = New NetworkAdapter ' Create a new network adapter class
            tmpAdapter.AdapterInfo = ipAdtInfo(i) ' Set the new classes adapter information
            Call mCol.Add(tmpAdapter, Trim(Str(mCol.Count))) ' Add the network adapter object to the collection
        Next i
    End If
End Sub
