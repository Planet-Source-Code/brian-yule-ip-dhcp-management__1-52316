VERSION 5.00
Begin VB.Form frmNewIpForwardEntry 
   BackColor       =   &H00808080&
   Caption         =   "New Ip Forward Entry"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form14"
   MDIChild        =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   10290
   Begin VB.ComboBox cboForwardType 
      Height          =   315
      Left            =   3480
      TabIndex        =   24
      Top             =   3000
      Width           =   6735
   End
   Begin VB.ComboBox cboProtocol 
      Height          =   315
      Left            =   3480
      TabIndex        =   26
      Top             =   2760
      Width           =   6735
   End
   Begin VB.TextBox txtForwardPolicy 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0"
      Top             =   2520
      Width           =   6735
   End
   Begin VB.TextBox txtForwardNextHopAS 
      Height          =   285
      Left            =   3480
      TabIndex        =   18
      Text            =   "0"
      Top             =   2280
      Width           =   6735
   End
   Begin VB.TextBox txtForwardNextHop 
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Top             =   2040
      Width           =   6735
   End
   Begin VB.TextBox txtForwardMetric5 
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Text            =   "-1"
      Top             =   1800
      Width           =   6735
   End
   Begin VB.TextBox txtForwardMetric4 
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      Text            =   "-1"
      Top             =   1560
      Width           =   6735
   End
   Begin VB.TextBox txtForwardMetric3 
      Height          =   285
      Left            =   3480
      TabIndex        =   14
      Text            =   "-1"
      Top             =   1320
      Width           =   6735
   End
   Begin VB.TextBox txtForwardMetric2 
      Height          =   285
      Left            =   3480
      TabIndex        =   13
      Text            =   "-1"
      Top             =   1080
      Width           =   6735
   End
   Begin VB.TextBox txtForwardMetric1 
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Text            =   "-1"
      Top             =   840
      Width           =   6735
   End
   Begin VB.TextBox txtForwardMask 
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Top             =   600
      Width           =   6735
   End
   Begin VB.ComboBox cboAdapterIndex 
      Height          =   315
      Left            =   3480
      TabIndex        =   25
      Top             =   360
      Width           =   6735
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H00808080&
      Caption         =   "Create"
      Height          =   495
      Left            =   7920
      TabIndex        =   19
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtForwardDest 
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label lblForwardType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label lblForwardProto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Protocol"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   585
   End
   Begin VB.Label lblForwardPolicy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Policy"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   420
   End
   Begin VB.Label lblForwardNextHopAS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Hop Autonomous System Number"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   2760
   End
   Begin VB.Label lblForwardNextHop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Hop IP Address"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1485
   End
   Begin VB.Label lblForwardMetric5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Metric 5"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label lblForwardMetric4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Metric 4"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label lblForwardMetric3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Metric 3"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label lblForwardMetric2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Metric 2"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   570
   End
   Begin VB.Label lblForwardMetric1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Metric 1"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   570
   End
   Begin VB.Label lblForwardMask 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mask"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   390
   End
   Begin VB.Label lblForwardIfIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interface"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   630
   End
   Begin VB.Label lblForwardDest 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destination IP Address"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1605
   End
End
Attribute VB_Name = "frmNewIpForwardEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim entry As MIB_IPFORWARDROW
Dim t As New NetworkTables

Private Sub cmdCreate_Click()
    Dim RetVal As Long
    Dim ifRow() As MIB_IFROW
    
    ' Get the destination IP address and convert it to a network address number
    entry.dwForwardDest = modWinsock.GetHostByNameAlias(Me.txtForwardDest.Text)
    
    ' Get the If table rows
    ifRow = t.IfTable
    
    ' Get the index of the currently selected adapter
    entry.dwForwardIfIndex = ifRow(cboAdapterIndex.ListIndex).dwIndex
    
    ' Get the destination Subnet Mask and convert it to a network address number
    entry.dwForwardMask = modWinsock.GetHostByNameAlias(Me.txtForwardMask.Text)
    
    ' Get the metrics!  These must be numberic or else there will be an exception
    entry.dwForwardMetric1 = Val(Me.txtForwardMetric1.Text)
    entry.dwForwardMetric2 = Val(Me.txtForwardMetric2.Text)
    entry.dwForwardMetric3 = Val(Me.txtForwardMetric3.Text)
    entry.dwForwardMetric4 = Val(Me.txtForwardMetric4.Text)
    entry.dwForwardMetric5 = Val(Me.txtForwardMetric5.Text)
    
    ' Get the Next Hop IP Address and conver it to a network address number
    entry.dwForwardNextHop = modWinsock.GetHostByNameAlias(Me.txtForwardNextHop.Text)
    
    ' Get the autonomous number of the next hop
    entry.dwForwardNextHopAS = Val(Me.txtForwardNextHopAS.Text)
    
    ' Get the forward policy number, normally in IP_TOS format
    entry.dwForwardPolicy = Val(Me.txtForwardPolicy.Text)
    
    ' Get the forward route IP Protocol
    entry.dwForwardProto = GetIPForwardProtocolValue(Me.cboProtocol.Text)
    
    ' Get the Type of forward route
    entry.dwForwardType = Me.cboForwardType.ListIndex + 1
    
    ' Create the entry in the IP Forward table
    RetVal = CreateIpForwardEntry(entry)
    If RetVal = 0 Then ' If the create succeeded then
        ' Unload the form, were done!
        Unload Me
    Else
        ' Inform the user that the items selected where not correct!
        MsgBox LoadResString(367) & Chr(vbKeySpace) & GetDLLMessage(RetVal)
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(323)
    
    ' Set the captions of the labels
    Me.lblForwardDest.Caption = LoadResString(351)
    Me.lblForwardIfIndex.Caption = LoadResString(352)
    Me.lblForwardMask.Caption = LoadResString(353)
    Me.lblForwardMetric1.Caption = LoadResString(354)
    Me.lblForwardMetric2.Caption = LoadResString(355)
    Me.lblForwardMetric3.Caption = LoadResString(356)
    Me.lblForwardMetric4.Caption = LoadResString(357)
    Me.lblForwardMetric5.Caption = LoadResString(358)
    Me.lblForwardNextHop.Caption = LoadResString(359)
    Me.lblForwardNextHopAS.Caption = LoadResString(360)
    Me.lblForwardPolicy.Caption = LoadResString(361)
    Me.lblForwardProto.Caption = LoadResString(362)
    Me.lblForwardType.Caption = LoadResString(363)

    '---------------------------------------------------------------------------------
    ' Load the IP Forward Types into the combo box
    Call cboForwardType.AddItem(GetIPForwardTypeString(IPFORWARD_TYPES.other))
    Call cboForwardType.AddItem(GetIPForwardTypeString(IPFORWARD_TYPES.invalid))
    Call cboForwardType.AddItem(GetIPForwardTypeString(IPFORWARD_TYPES.Local_Route))
    Call cboForwardType.AddItem(GetIPForwardTypeString(IPFORWARD_TYPES.Remote_Route))
    cboForwardType.ListIndex = 0 ' Select the first item
    
    ' Load the network adapters from the If table into the combo box
    Dim n() As MIB_IFROW, i As Integer
    n = t.IfTable
    For i = LBound(n) To UBound(n)
        cboAdapterIndex.AddItem (StrConv(n(i).bDescr, vbUnicode))
    Next i
    cboAdapterIndex.ListIndex = 0 ' Select the first item
    
    ' Load the IP Forward Protocols into the combo box
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_BBN))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_BGP))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_BOOTP))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_CISCO))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_EGP))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_ES_IS))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_GGP))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_HELLO))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_ICMP))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_IS_IS))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_LOCAL))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_NETMGMT))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_NT_AUTOSTATIC))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_NT_STATIC))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_NT_STATIC_NON_DOD))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_OSPF))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_OTHER))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_RIP))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IPX_PROTOCOL_NLSP))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IPX_PROTOCOL_RIP))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IPX_PROTOCOL_SAP))
    Call cboProtocol.AddItem(GetIPForwardProtocolString(PROTOCOLS.IP_IDPR))
    cboProtocol.ListIndex = 0 ' Select the first item
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Resize the grid to the size of the form (roughly)!
        Height = 4410
        Width = 10410
    End If
End Sub
