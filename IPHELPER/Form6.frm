VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUDPTable 
   Caption         =   "UDP Table"
   ClientHeight    =   720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1725
   LinkTopic       =   "Form6"
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
      FixedCols       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmUDPTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim t As NetworkTables

Private Sub SetColumnText(ByVal Column As Integer, ByRef colVal As Variant)
    ' Set the column value of the current row
    MSFlexGrid1.Col = Column
    MSFlexGrid1.Text = colVal
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(327)
    Call Init ' Initialize the grid
End Sub

Private Sub Form_Activate()
    Call Init
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Resize the grid to the size of the form (roughly)!
        MSFlexGrid1.Width = Width - 105
        MSFlexGrid1.Height = Height - 465
    End If
End Sub

Private Sub Init()
    Dim AddRow() As MIB_UDPROW, i As Integer
    
    ' Create the Network Tables class
    Set t = New NetworkTables
    
    ' Get the UDP Table rows
    AddRow = t.UdpTable

    With MSFlexGrid1
        ' Stop redrawing the grid
        .Redraw = False
        .Cols = 2
        
        .Rows = UBound(AddRow) + 2
        
        .Row = 0
        
        ' Set the fixed column headers
        Call SetColumnText(0, LoadResString(252))
        Call SetColumnText(1, LoadResString(253))
        
        ' For each row of UDP info
        For i = 0 To UBound(AddRow)
    
            .Row = i + 1
            
            ' Fill the columns with the UDP info
            Call SetColumnText(0, modWinsock.getascip(AddRow(i).dwLocalAddr)) ' Conver the long network address to a string IP address
            Call SetColumnText(1, modWinsock.GetPort(AddRow(i).dwLocalPort)) ' Conver the network port number to an application port number
        Next i
        
        ' Auto size the grid columns
        modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
        
        ' Start redrawing the grid
        .Redraw = True
    End With
    
    ' Release the Network Tables class
    Set t = Nothing
End Sub
