VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTCPStatistics 
   Caption         =   "TCP Statistics"
   ClientHeight    =   705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1725
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   705
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
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmTCPStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim w As TCPStats

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
    Me.Caption = LoadResString(324)
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
    ' Create the TCP Statistics class
    Set w = New TCPStats
    
    With MSFlexGrid1
        ' Stop redrawing the grid
        .Redraw = False
        
        .Cols = 2
        .Rows = 1
        .Col = 1
        .Row = 0
        .Text = LoadResString(239)
        
        ' Add the rows of TCP statistics data to the grid
        Call AddRow(LoadResString(223), w.ActiveOpens)
        Call AddRow(LoadResString(224), w.AttemptFails)
        Call AddRow(LoadResString(225), w.CurrEstab)
        Call AddRow(LoadResString(226), w.EstabResets)
        Call AddRow(LoadResString(227), w.InErrs)
        Call AddRow(LoadResString(228), w.InSegs)
        Call AddRow(LoadResString(229), w.MaxConn)
        Call AddRow(LoadResString(230), w.NumConns)
        Call AddRow(LoadResString(231), w.OutRsts)
        Call AddRow(LoadResString(232), w.OutSegs)
        Call AddRow(LoadResString(233), w.PassiveOpens)
        Call AddRow(LoadResString(235), w.retransmission_time_out_Algorithm)
        Call AddRow(LoadResString(236), w.RetransSegs)
        Call AddRow(LoadResString(237), w.RtoMax)
        Call AddRow(LoadResString(238), w.RtoMin)
        
        ' Auto size the grid columns
        modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
    
        ' Start redrawing the grid again!
        .Redraw = True
    End With
    
    ' Release the TCP Statistics class
    Set w = Nothing
End Sub
