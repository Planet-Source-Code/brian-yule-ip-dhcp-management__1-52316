VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"IPAddress"
Attribute VB_Ext_KEY = "Member0" ,"IPAddress"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Standard Object Collection Class

Private mCol As Collection

Friend Function Add(newAdapter As IPAddress, Optional sKey As String) As IPAddress
    If Len(sKey) = 0 Then mCol.Add newAdapter Else mCol.Add newAdapter, CStr(sKey)
    Set Add = newAdapter
End Function

Public Property Get Item(vntIndexKey As Variant) As IPAddress
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
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
