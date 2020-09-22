VERSION 5.00
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "IP Helper"
   ClientHeight    =   3780
   ClientLeft      =   3090
   ClientTop       =   4770
   ClientWidth     =   9435
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuStatistics 
      Caption         =   "&Statistics"
      Begin VB.Menu mnuIpStatistics 
         Caption         =   "&IP Statistics"
      End
      Begin VB.Menu mnuUdpStatistics 
         Caption         =   "&Udp Statistics"
      End
      Begin VB.Menu mnuTCPStatistics 
         Caption         =   "&Tcp Statistics"
      End
      Begin VB.Menu mnuICMPStatistics 
         Caption         =   "I&CMP Statistics"
      End
   End
   Begin VB.Menu mnuTables 
      Caption         =   "&Tables"
      Begin VB.Menu mnuIPAddressTable 
         Caption         =   "&IP Address Table"
      End
      Begin VB.Menu mnuTCPTable 
         Caption         =   "&TCP Table"
      End
      Begin VB.Menu mnuUDPTable 
         Caption         =   "&UDP Table"
      End
      Begin VB.Menu mnuIfTable 
         Caption         =   "I&F Table"
      End
      Begin VB.Menu mnuIPForwardTable 
         Caption         =   "I&P Forward Table"
      End
      Begin VB.Menu mnuIPNetTable 
         Caption         =   "IP &Net Table"
      End
   End
   Begin VB.Menu mnuInterfaceAdapters 
      Caption         =   "&Interface Adapters"
   End
   Begin VB.Menu mnuNetworkAdapters 
      Caption         =   "&Network Adapters"
   End
   Begin VB.Menu mnuNetworkParameters 
      Caption         =   "Network &Parameters"
   End
   Begin VB.Menu mnuOpenAll 
      Caption         =   "&Open All"
   End
   Begin VB.Menu mnuCloseAll 
      Caption         =   "&Close All"
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuTileHorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
' Sample of how easy it is to send an ARP request!
'    Dim h As New ARP
'    Debug.Print h.Send("192.168.1.33")

    Me.Caption = LoadResString(313)
    
    ' Set the captions of the menu items!
    Me.mnuArrangeIcons.Caption = LoadResString(328)
    Me.mnuCascade.Caption = LoadResString(329)
    Me.mnuCloseAll.Caption = LoadResString(330)
    Me.mnuExit.Caption = LoadResString(331)
    Me.mnuICMPStatistics.Caption = LoadResString(332)
    Me.mnuIfTable.Caption = LoadResString(333)
    Me.mnuInterfaceAdapters.Caption = LoadResString(334)
    Me.mnuIPAddressTable.Caption = LoadResString(335)
    Me.mnuIPForwardTable.Caption = LoadResString(336)
    Me.mnuIPNetTable.Caption = LoadResString(337)
    Me.mnuIpStatistics.Caption = LoadResString(338)
    Me.mnuNetworkAdapters.Caption = LoadResString(339)
    Me.mnuNetworkParameters.Caption = LoadResString(340)
    Me.mnuOpenAll.Caption = LoadResString(341)
    Me.mnuStatistics.Caption = LoadResString(342)
    Me.mnuTables.Caption = LoadResString(343)
    Me.mnuTCPStatistics.Caption = LoadResString(344)
    Me.mnuTCPTable.Caption = LoadResString(345)
    Me.mnuTileHorizontally.Caption = LoadResString(346)
    Me.mnuTileVertically.Caption = LoadResString(347)
    Me.mnuUdpStatistics.Caption = LoadResString(348)
    Me.mnuUDPTable.Caption = LoadResString(349)
    Me.mnuWindow.Caption = LoadResString(350)
End Sub

Private Sub mnuIfTable_Click()
    Load frmIFTable
    frmIFTable.Show
    frmIFTable.SetFocus
End Sub

Private Sub mnuArrangeIcons_Click()
    Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
    Arrange vbCascade
End Sub

Private Sub mnuCloseAll_Click()
    Dim x As Integer
    
    For x = (Forms.Count - 1) To 0 Step -1
        If Not TypeOf Forms(x) Is MDIForm Then
            Unload Forms(x)
        End If
    Next x
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuICMPStatistics_Click()
    Load frmICMPStatistics
    frmICMPStatistics.Show
    frmICMPStatistics.SetFocus
End Sub

Private Sub mnuInterfaceAdapters_Click()
    Load frmInterfaceAdapters
    frmInterfaceAdapters.Show
    frmInterfaceAdapters.SetFocus
End Sub

Private Sub mnuIPAddressTable_Click()
    Load frmIpAddressTable
    frmIpAddressTable.Show
    frmIpAddressTable.SetFocus
End Sub

Private Sub mnuIPForwardTable_Click()
    Load frmIPForwardTable
    frmIPForwardTable.Show
    frmIPForwardTable.SetFocus
End Sub

Private Sub mnuIPNetTable_Click()
    Load frmIPNetTable
    frmIPNetTable.Show
    frmIPNetTable.SetFocus
End Sub

Private Sub mnuIpStatistics_Click()
    Load frmIPStatistics
    frmIPStatistics.Show
    frmIPStatistics.SetFocus
End Sub

Private Sub mnuNetworkAdapters_Click()
    Load frmNetworkAdapters
    frmNetworkAdapters.Show
    frmNetworkAdapters.SetFocus
End Sub

Private Sub mnuNetworkParameters_Click()
    Load frmNetworkParameters
    frmNetworkParameters.Show
    frmNetworkParameters.SetFocus
End Sub

Private Sub mnuOpenAll_Click()
    Screen.MousePointer = vbHourglass
    
    Load frmIPStatistics
    Load frmUDPStatistics
    Load frmTCPStatistics
    Load frmIpAddressTable
    Load frmTCPTable
    Load frmUDPTable
    Load frmIFTable
    Load frmIPForwardTable
    Load frmIPNetTable
    Load frmInterfaceAdapters
    Load frmNetworkAdapters
    Load frmNetworkParameters
    Load frmICMPStatistics
    
    frmIPStatistics.Show
    frmUDPStatistics.Show
    frmTCPStatistics.Show
    frmIpAddressTable.Show
    frmTCPTable.Show
    frmUDPTable.Show
    frmIFTable.Show
    frmIPForwardTable.Show
    frmIPNetTable.Show
    frmInterfaceAdapters.Show
    frmNetworkAdapters.Show
    frmNetworkParameters.Show
    frmICMPStatistics.Show
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuTCPStatistics_Click()
    Load frmTCPStatistics
    frmTCPStatistics.Show
    frmTCPStatistics.SetFocus
End Sub

Private Sub mnuTCPTable_Click()
    Load frmTCPTable
    frmTCPTable.Show
    frmTCPTable.SetFocus
End Sub

Private Sub mnuTileHorizontally_Click()
    Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVertically_Click()
    Arrange vbTileVertical
End Sub

Private Sub mnuUdpStatistics_Click()
    Load frmUDPStatistics
    frmUDPStatistics.Show
    frmUDPStatistics.SetFocus
End Sub

Private Sub mnuUDPTable_Click()
    Load frmUDPTable
    frmUDPTable.Show
    frmUDPTable.SetFocus
End Sub
