VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNetworkAdapters 
   Caption         =   "Network Adapters"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   2085
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   1440
   ScaleWidth      =   2085
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Menu mnuRelease 
      Caption         =   "&Release"
   End
   Begin VB.Menu mnuRenew 
      Caption         =   "Re&new"
   End
End
Attribute VB_Name = "frmNetworkAdapters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim y As NetworkAdapters
Dim currentMouseRow As Long

Private Sub SetColumnText(ByVal Column As Integer, ByRef colVal As Variant)
    ' Set the value of the grid column on the current row!
    MSFlexGrid1.Col = Column
    MSFlexGrid1.Text = colVal
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(321)
    Call Init ' Initialize the grid
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Resize the grid to the size of the form (roughly)!
        MSFlexGrid1.Width = Width - 105
        MSFlexGrid1.Height = Height - 465
    End If
End Sub

Private Sub Form_Activate()
    Call Init
End Sub


Private Sub Init()
    Dim i As Integer, j As Integer, tmpStr As String
    
    ' Create the network adapters class
    Set y = New NetworkAdapters
    
    With MSFlexGrid1
        ' Stop redrawing the grid
        .Redraw = False
        .Cols = 16
        
        .Rows = y.Count + 1
        
        .Row = 0
        
        ' Set the fixed column headers
        Call SetColumnText(0, LoadResString(199))
        Call SetColumnText(1, LoadResString(200))
        Call SetColumnText(2, LoadResString(201))
        Call SetColumnText(3, LoadResString(202))
        Call SetColumnText(4, LoadResString(203))
        Call SetColumnText(5, LoadResString(204))
        Call SetColumnText(6, LoadResString(205))
        Call SetColumnText(7, LoadResString(206))
        Call SetColumnText(8, LoadResString(207))
        Call SetColumnText(9, LoadResString(208))
        Call SetColumnText(10, LoadResString(209))
        Call SetColumnText(11, LoadResString(210))
        Call SetColumnText(12, LoadResString(211))
        Call SetColumnText(13, LoadResString(212))
        Call SetColumnText(14, LoadResString(213))
        Call SetColumnText(15, LoadResString(214))
    
        Dim tmpAdapter As NetworkAdapter
        
        ' For each network adapter
        For Each tmpAdapter In y
        
            tmpStr = vbNullString
            MSFlexGrid1.Row = i + 1
                
            ' Set the column data with the data from the network adapter!
            Call SetColumnText(0, tmpAdapter.AdapterDescription)
            Call SetColumnText(1, tmpAdapter.AdapterIndex)
            Call SetColumnText(2, tmpAdapter.AdapterName)
            Call SetColumnText(3, GetIfTypeString(tmpAdapter.AdapterType))
            Call SetColumnText(4, tmpAdapter.Address)
            Call SetColumnText(5, tmpAdapter.AutoConfigActive)
            Call SetColumnText(6, tmpAdapter.AutoConfigEnabled)
            Call SetColumnText(7, tmpAdapter.DHCP.Item(0).Address & Chr(vbKeySpace) & tmpAdapter.DHCP.Item(0).Mask)
            Call SetColumnText(8, tmpAdapter.DHCPEnabled)
            Call SetColumnText(9, tmpAdapter.Gateway.Item(0).Address & Chr(vbKeySpace) & tmpAdapter.Gateway.Item(0).Mask)
            Call SetColumnText(10, tmpAdapter.HasWins)
            
            ' The address list needs to be compiled into a string because it can be one address or many!
            For j = 0 To tmpAdapter.IPAddressList.Count - 1
                tmpStr = tmpStr & tmpAdapter.IPAddressList.Item(j).Address & Chr(vbKeySpace) & tmpAdapter.IPAddressList.Item(j).Mask & Chr(vbKeySpace)
            Next j
            
            Call SetColumnText(11, tmpStr)
            Call SetColumnText(12, tmpAdapter.LeaseExpires)
            Call SetColumnText(13, tmpAdapter.LeaseObtained)
            Call SetColumnText(14, tmpAdapter.PrimaryWins.Item(0).Address & Chr(vbKeySpace) & tmpAdapter.PrimaryWins.Item(0).Mask)
            Call SetColumnText(15, tmpAdapter.SecondaryWins.Item(0).Address & Chr(vbKeySpace) & tmpAdapter.SecondaryWins.Item(0).Mask)
            
            i = i + 1
        Next
        
        ' Auto size the grid columns
        modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
    
        ' Start redrawing the grid
        .Redraw = True
    End With
    
    ' Release the network adapters class
    Set y = Nothing
End Sub

Private Sub mnuRelease_Click()
    ' Create a network interfaces class
    Dim z As New NetworkInterfaces
    If currentMouseRow > 0 Then ' If the selected grid row is not fixed then
        Screen.MousePointer = vbHourglass
        Dim i As Integer, indexToFind As Long
        
        ' Get the index of the network adapter selected
        indexToFind = MSFlexGrid1.TextMatrix(currentMouseRow, 1)
        
        ' For each network interface
        For i = 0 To z.Count - 1
            ' Check if the index matches the one selected
            If z.Item(i).Index = indexToFind Then
                ' If a match is found then call its release method and if it succeeds then refresh the grid
                If z.Item(i).Release = 0 Then Call Init
            End If
        Next i
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnuRenew_Click()
    ' Create a network interfaces class
    Dim z As New NetworkInterfaces
    If currentMouseRow > 0 Then ' If the selected grid row is not fixed then
        Screen.MousePointer = vbHourglass
        Dim i As Integer, indexToFind As Long
        
        ' Get the index of the network adapter selected
        indexToFind = MSFlexGrid1.TextMatrix(currentMouseRow, 1)
        
        ' For each network interface
        For i = 0 To z.Count - 1
            ' Check if the index matches the one selected
            If z.Item(i).Index = indexToFind Then
                ' If a match is found then call its renew method and if it succeeds then refresh the grid
                If z.Item(i).Renew = 0 Then Call Init
            End If
        Next i
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub MSFlexGrid1_Click()
    ' Save the currently selected row
    currentMouseRow = MSFlexGrid1.MouseRow
End Sub
