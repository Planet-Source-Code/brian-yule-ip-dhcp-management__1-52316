VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNetworkParameters 
   Caption         =   "Network Parameters"
   ClientHeight    =   750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1725
   LinkTopic       =   "Form12"
   MDIChild        =   -1  'True
   ScaleHeight     =   750
   ScaleWidth      =   1725
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmNetworkParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x As NetworkParams

Private Sub SetColumnText(ByVal Column As Integer, ByRef colVal As Variant)
    ' Set the column value on the current row
    MSFlexGrid1.Col = Column
    MSFlexGrid1.Text = colVal
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(322)
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
    Dim tmpStr As String, i As Integer
    
    ' Create the network parameters class
    Set x = New NetworkParams
    
    With MSFlexGrid1
        ' Stop redrawing the grid
        .Redraw = False
        .Cols = 8
        .Rows = 2
        
        .Row = 0
        
        ' Set the fixed column headers
        Call SetColumnText(0, LoadResString(215))
        Call SetColumnText(1, LoadResString(216))
        Call SetColumnText(2, LoadResString(217))
        Call SetColumnText(3, LoadResString(218))
        Call SetColumnText(4, LoadResString(219))
        Call SetColumnText(5, LoadResString(220))
        Call SetColumnText(6, LoadResString(221))
        Call SetColumnText(7, LoadResString(222))
    
        .Row = 1
            
        ' Fill the grid with the network parameters
        Call SetColumnText(0, x.ARPProxyEnabled)
        Call SetColumnText(1, x.DHCPScopeName)
        Call SetColumnText(2, x.DNSEnabled)
        
        ' Because the DNS server list can be on or many we compile them all into one big string
        For i = 0 To x.DNSServerList.Count - 1
            tmpStr = tmpStr & x.DNSServerList.Item(i).Address & Chr(vbKeySpace) & x.DNSServerList.Item(i).Mask & Chr(vbKeySpace)
        Next i
        Call SetColumnText(3, tmpStr)
        
        Call SetColumnText(4, x.DomainName)
        Call SetColumnText(5, x.hostname)
        Call SetColumnText(6, GetNodeTypeString(x.NodeType))
        Call SetColumnText(7, x.RoutingEnabled)
        
        ' Auto size the grid columns
        modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
        
        ' Start redrawing the grid again
        .Redraw = True
    End With
    
    ' Release the network parameters class
    Set x = Nothing
End Sub
