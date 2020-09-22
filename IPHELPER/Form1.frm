VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIPStatistics 
   Caption         =   "IP Statistics"
   ClientHeight    =   840
   ClientLeft      =   1860
   ClientTop       =   2655
   ClientWidth     =   2610
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   840
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1508
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmIPStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim u As IPStats

Private Sub AddRow(ByVal RowHeader As String, ByRef RowText As Variant)
' Adds rows to the grid and places the information into the row created
    With MSFlexGrid1
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .Text = RowHeader
        .Col = 1
        .Text = RowText
    End With
End Sub

Private Sub Form_Activate()
    Call Init
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(320)
    Call Init ' Initialize the grid
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Resize the grid to the size of the form (roughly)!
        MSFlexGrid1.Width = Width - 105
        MSFlexGrid1.Height = Height - 465
    End If
End Sub

Private Sub Init()
    ' Create the IP statistics class
    Set u = New IPStats
    
    With MSFlexGrid1
        ' Stop redrawing the grid
        .Redraw = False
        
        .Cols = 2
        .Rows = 1
        .Col = 1
        .Row = 0
        .Text = LoadResString(175)
        
        ' Add the rows to the grid with the IP statistics data
        Call AddRow(LoadResString(176), u.DefaultTTL)
        Call AddRow(LoadResString(177), u.Forwarding)
        Call AddRow(LoadResString(178), u.ForwDatagrams)
        Call AddRow(LoadResString(179), u.FragCreates)
        Call AddRow(LoadResString(180), u.FragFails)
        Call AddRow(LoadResString(181), u.FragOks)
        Call AddRow(LoadResString(182), u.InAddrErrors)
        Call AddRow(LoadResString(183), u.InDelivers)
        Call AddRow(LoadResString(184), u.InDiscards)
        Call AddRow(LoadResString(185), u.InHdrErrors)
        Call AddRow(LoadResString(186), u.InReceives)
        Call AddRow(LoadResString(187), u.InUnknownProtos)
        Call AddRow(LoadResString(188), u.NumAddr)
        Call AddRow(LoadResString(189), u.NumIf)
        Call AddRow(LoadResString(190), u.NumRoutes)
        Call AddRow(LoadResString(191), u.OutDiscards)
        Call AddRow(LoadResString(192), u.OutNoRoutes)
        Call AddRow(LoadResString(193), u.OutRequests)
        Call AddRow(LoadResString(194), u.ReasmFails)
        Call AddRow(LoadResString(195), u.ReasmOks)
        Call AddRow(LoadResString(196), u.ReasmReqds)
        Call AddRow(LoadResString(197), u.ReasmTimeout)
        Call AddRow(LoadResString(198), u.RoutingDiscards)
        
        ' Auto size the grid
        modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
    
        ' Start redrawing the grid
        .Redraw = True
    End With
    
    ' Release the ip statistics class
    Set u = Nothing
End Sub
