VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUDPStatistics 
   Caption         =   "UDP Statistics"
   ClientHeight    =   720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1725
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   720
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
Attribute VB_Name = "frmUDPStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim v As UDPStats

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
    Me.Caption = LoadResString(326)
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
    ' Create the UDP statistics class
    Set v = New UDPStats
    
    With MSFlexGrid1
        ' Stop redrawing the grid
        .Redraw = False
        
        .Cols = 2
        .Rows = 1
        .Col = 1
        .Row = 0
        .Text = LoadResString(246)
        
        ' Add rows of UDP Statistics data
        Call AddRow(LoadResString(247), v.InDatagrams)
        Call AddRow(LoadResString(248), v.InErrors)
        Call AddRow(LoadResString(249), v.NoPorts)
        Call AddRow(LoadResString(250), v.NumAddrs)
        Call AddRow(LoadResString(251), v.OutDatagrams)
        
        ' Auto size the grid columns
        modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
        
        ' Start redrawing the grid again
        .Redraw = True
    End With
    
    ' Release the UCP statistics class
    Set v = Nothing
End Sub
